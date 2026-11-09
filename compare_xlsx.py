#!/usr/bin/env python3
"""Локальное сравнение XLSX: ZIP, ячейки, формулы, кэш и оформление.

Python 3.10+, только стандартная библиотека. Входные книги не сохраняются.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import posixpath
import re
import sys
import tempfile
import zipfile
from collections import Counter, defaultdict
from contextlib import ExitStack
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from xml.etree import ElementTree as ET

VERSION = "1.0"
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"m": MAIN}
M = "{" + MAIN + "}"
R = "{" + REL + "}"


def local(tag):
    return tag.rsplit("}", 1)[-1]


def canonical(element):
    """Префиксы namespace и порядок атрибутов не меняют смысл XML."""
    text = element.text or ""
    if not text.strip() and local(element.tag) not in ("t", "v", "f", "formula", "formula1", "formula2"):
        text = ""
    return [element.tag, sorted(element.attrib.items()), text,
            [canonical(child) for child in element]]


def style_canonical(element):
    """Свойства шрифта могут менять порядок, значения по умолчанию — появляться."""
    tag = local(element.tag)
    attributes = dict(element.attrib)
    font_flags = {"b", "i", "strike", "outline", "shadow", "condense", "extend"}
    if tag in font_flags:
        attributes["val"] = "0" if attributes.get("val") in ("0", "false") else "1"
    for key in ("pivotButton", "quotePrefix", "diagonalUp", "diagonalDown"):
        if attributes.get(key) in ("0", "false"):
            del attributes[key]
    if tag == "patternFill" and attributes.get("patternType") == "none":
        del attributes["patternType"]
    children = [style_canonical(child) for child in element
                if not (local(child.tag) in font_flags and child.get("val") in ("0", "false"))]
    if tag in {"font", "rPr", "border", "fill", "patternFill", "xf", "dxf"}:
        children.sort(key=encoded)
    text = element.text or ""
    return [element.tag, sorted(attributes.items()), text if text.strip() else "", children]


def encoded(value):
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), sort_keys=True)


def digest(value):
    return hashlib.sha256(encoded(value).encode("utf-8")).hexdigest()


def numeric(value):
    if value is None or value == "":
        return None
    try:
        number = Decimal(value)
        if number.is_finite():
            # normalize() применяет decimal context и может округлить длинное число.
            sign, digits, exponent = number.as_tuple()
            digits = list(digits)
            while len(digits) > 1 and digits[-1] == 0:
                digits.pop()
                exponent += 1
            return [sign if any(digits) else 0, digits if any(digits) else [0], exponent if any(digits) else 0]
    except InvalidOperation:
        pass
    return value


def col_number(letters):
    result = 0
    for letter in letters.upper():
        result = result * 26 + ord(letter) - 64
    return result


def col_letters(number):
    result = ""
    while number:
        number, tail = divmod(number - 1, 26)
        result = chr(65 + tail) + result
    return result


def coordinate(address):
    match = re.fullmatch(r"([A-Za-z]{1,3})([1-9][0-9]*)", address or "")
    if not match:
        raise ValueError(f"Недопустимый адрес ячейки: {address!r}")
    column, row = col_number(match[1]), int(match[2])
    if column > 16384 or row > 1048576:
        raise ValueError(f"Адрес вне границ Excel: {address}")
    return row, column


PROTECTED = re.compile(r'"(?:[^"\n]|"")*"|\'(?:[^\']|\'\')*\'|\[[^\]]*\]')
CELL_REF = re.compile(r"(?<![A-Za-z0-9_.])(?P<col>\$?[A-Za-z]{1,3})(?P<row>\$?[1-9][0-9]*)(?![A-Za-z0-9_!.\[]|\s*\()")
COL_RANGE = re.compile(r"(?<![A-Za-z0-9_.])(?P<a>\$?[A-Za-z]{1,3}):(?P<b>\$?[A-Za-z]{1,3})(?![A-Za-z0-9_!.])")
ROW_RANGE = re.compile(r"(?<![A-Za-z0-9_.])(?P<a>\$?[1-9][0-9]*):(?P<b>\$?[1-9][0-9]*)(?![A-Za-z0-9_.])")


def translate_shared(text, origin, target):
    """Разворачивает обычные A1-ссылки общей формулы, не вычисляя её."""
    row0, col0 = coordinate(origin)
    row1, col1 = coordinate(target)
    dr, dc = row1 - row0, col1 - col0

    def shift(value, delta, column=False):
        absolute = value.startswith("$")
        token = value.lstrip("$")
        number = col_number(token) if column else int(token)
        if number > (16384 if column else 1048576):
            return value
        if not absolute:
            number += delta
        if not 1 <= number <= (16384 if column else 1048576):
            raise ValueError("Общая формула ссылается за пределы листа")
        return ("$" if absolute else "") + (col_letters(number) if column else str(number))

    def replace_plain(part):
        part = COL_RANGE.sub(lambda m: shift(m["a"], dc, True) + ":" + shift(m["b"], dc, True), part)
        part = ROW_RANGE.sub(lambda m: shift(m["a"], dr) + ":" + shift(m["b"], dr), part)
        return CELL_REF.sub(lambda m: shift(m["col"], dc, True) + shift(m["row"], dr), part)

    pieces, start = [], 0
    for match in PROTECTED.finditer(text):
        pieces.extend((replace_plain(text[start:match.start()]), match[0]))
        start = match.end()
    pieces.append(replace_plain(text[start:]))
    return "".join(pieces)


def rich_string(element):
    runs = []
    text = ""
    for child in element:
        if child.tag == M + "t":
            text += child.text or ""
        elif child.tag == M + "r":
            part = child.find(M + "t")
            value = "" if part is None else part.text or ""
            props = child.find(M + "rPr")
            runs.append([value, None if props is None else style_canonical(props)])
            text += value
    # Plain shared strings and inline strings have identical content.
    formatting = runs if any(props is not None for _, props in runs) else None
    return text, formatting


class Package:
    def __init__(self, path):
        self.path = Path(path)
        self.zip = zipfile.ZipFile(path)
        try:
            self.entries = {item.filename: item for item in self.zip.infolist() if not item.is_dir()}
            if len(self.entries) != sum(not item.is_dir() for item in self.zip.infolist()):
                raise ValueError("В архиве повторяются имена внутренних файлов")
            if any(item.flag_bits & 1 for item in self.entries.values()):
                raise ValueError("Зашифрованный ZIP не поддерживается")
            for required in ("xl/workbook.xml", "xl/_rels/workbook.xml.rels"):
                if required not in self.entries:
                    raise ValueError("Не найден " + required + "; нужен обычный XLSX")
            self.strings = self.read_strings()
            self.styles, self.dxf = self.read_styles()
            self.workbook = self.xml("xl/workbook.xml")
            if self.workbook.tag != M + "workbook":
                raise ValueError("Поддерживается обычный Transitional XLSX; namespace книги другой")
            relations = self.relationships("xl/workbook.xml")
            self.sheets = {}
            for sheet in self.workbook.findall("m:sheets/m:sheet", NS):
                relation = relations.get(sheet.get(R + "id"))
                if not relation or relation[2] == "External":
                    raise ValueError("Нельзя определить файл листа " + sheet.get("name", ""))
                kind, target, _ = relation
                if target not in self.entries:
                    raise ValueError("Нет внутреннего файла листа " + target)
                self.sheets[sheet.get("name")] = {"part": target, "kind": kind.rsplit("/", 1)[-1],
                                                  "state": sheet.get("state", "visible")}
        except Exception:
            self.close()
            raise

    def close(self):
        self.zip.close()

    def xml(self, name):
        with self.zip.open(name) as stream:
            return ET.parse(stream).getroot()

    def relationships(self, part):
        folder, name = posixpath.split(part)
        relfile = posixpath.join(folder, "_rels", name + ".rels")
        if relfile not in self.entries:
            return {}
        result = {}
        for relation in self.xml(relfile):
            target = relation.get("Target", "")
            mode = relation.get("TargetMode", "Internal")
            if mode != "External":
                target = posixpath.normpath(target.lstrip("/") if target.startswith("/") else posixpath.join(folder, target))
            result[relation.get("Id")] = [relation.get("Type", ""), target, mode]
        return result

    def read_strings(self):
        if "xl/sharedStrings.xml" not in self.entries:
            return []
        result = []
        with self.zip.open("xl/sharedStrings.xml") as stream:
            iterator = ET.iterparse(stream, events=("start", "end"))
            _, root = next(iterator)
            for event, element in iterator:
                if event == "end" and element.tag == M + "si":
                    result.append(rich_string(element))
                    root.clear()
        return result

    def read_styles(self):
        if "xl/styles.xml" not in self.entries:
            return [None], []
        root = self.xml("xl/styles.xml")
        groups = {local(child.tag): list(child) for child in root}
        number_formats = {item.get("numFmtId"): item.get("formatCode") for item in groups.get("numFmts", [])}

        def xf(element, base=False):
            attributes = dict(element.attrib)
            for key, group in (("fontId", "fonts"), ("fillId", "fills"), ("borderId", "borders")):
                index = int(attributes.pop(key, "0"))
                items = groups.get(group, [])
                attributes[key] = style_canonical(items[index]) if index < len(items) else ["missing", index]
            number = attributes.pop("numFmtId", "0")
            attributes["numFmtId"] = ["custom", number_formats[number]] if number in number_formats else ["builtin", number]
            index = attributes.pop("xfId", None)
            if not base and index is not None:
                items = groups.get("cellStyleXfs", [])
                attributes["base_style"] = xf(items[int(index)], True) if int(index) < len(items) else ["missing", index]
            for key in ("pivotButton", "quotePrefix"):
                if attributes.get(key) in ("0", "false"):
                    del attributes[key]
            return [attributes, [style_canonical(child) for child in element]]

        styles = [digest(xf(element)) for element in groups.get("cellXfs", [])]
        return styles or [None], [digest(style_canonical(element)) for element in groups.get("dxfs", [])]

    def style(self, index):
        index = int(index or 0)
        if not 0 <= index < len(self.styles):
            raise ValueError(f"Несуществующий стиль {index} в {self.path.name}")
        return self.styles[index]


class SheetReader:
    def __init__(self, package, part):
        self.package, self.part = package, part
        self.features = defaultdict(list)
        self.stats = Counter()
        self.shared = {}
        self.relations = package.relationships(part)

    def feature(self, element):
        # Номера стилей и rId могут измениться без изменения оформления/ссылки.
        for child in element.iter():
            if R + "id" in child.attrib:
                ident = child.get(R + "id")
                child.set(R + "id", encoded(self.relations.get(ident, ["unresolved", ident])))
            for key in ("style", "s"):
                if key in child.attrib and local(child.tag) in ("col", "row"):
                    child.set(key, self.package.style(child.get(key)) or "default")
            if "dxfId" in child.attrib:
                index = int(child.get("dxfId"))
                child.set("dxfId", self.package.dxf[index] if index < len(self.package.dxf) else f"missing:{index}")
        return canonical(element)

    def cell(self, element):
        address = element.get("r")
        coordinate(address)
        kind = element.get("t", "n")
        value_node = element.find(M + "v")
        value = None if value_node is None else value_node.text
        f = element.find(M + "f")
        rich = None
        formula = None
        if f is not None:
            formula = f.text or ""
            attributes = dict(f.attrib)
            if attributes.get("t") == "shared":
                index = attributes.get("si")
                if formula:
                    self.shared[index] = (address, formula)
                elif index in self.shared:
                    origin, text = self.shared[index]
                    formula = translate_shared(text, origin, address)
                else:
                    raise ValueError(f"{address}: нет базовой ячейки общей формулы {index}")
                attributes = {k: v for k, v in attributes.items() if k not in ("t", "si", "ref")}
            formula = [formula, attributes]
            content = ["formula", formula]
            cache = [kind, numeric(value) if kind == "n" else value] if value not in (None, "") else None
            self.stats["formulas"] += 1
            self.stats["cached_formulas"] += cache is not None
        else:
            cache = None
            if kind == "s":
                text, rich = self.package.strings[int(value)]
                content = ["string", text]
            elif kind == "inlineStr":
                inline = element.find(M + "is")
                text, rich = rich_string(inline) if inline is not None else ("", None)
                content = ["string", text]
            elif kind == "str":
                content = ["string", value or ""]
            elif value is None:
                content = None
            else:
                content = [kind, numeric(value) if kind == "n" else value]
        self.stats["explicit_cells"] += 1
        self.stats["content_cells"] += content is not None
        attributes = {k: v for k, v in element.attrib.items() if k not in ("r", "s", "t")}
        extra = [canonical(child) for child in element if local(child.tag) not in ("f", "v", "is")]
        # Attribute t belongs to the cache for formulas, not to formula text.
        return {"content": content, "cache": cache, "style": [self.package.style(element.get("s")), rich],
                "extra": [attributes, extra]}

    def rows(self):
        with self.package.zip.open(self.part) as stream:
            stack, previous, data = [], 0, None
            for event, element in ET.iterparse(stream, events=("start", "end")):
                if event == "start":
                    stack.append(element)
                    if element.tag == M + "sheetData":
                        data = element
                    continue
                if element.tag == M + "row":
                    number = int(element.get("r", previous + 1))
                    if number <= previous:
                        raise ValueError(f"Лист {self.part}: строки идут не по порядку")
                    previous = number
                    cells = {}
                    for child in element:
                        if child.tag != M + "c":
                            self.features["rowExtensions"].append([number, canonical(child)])
                            continue
                        address = child.get("r")
                        row, column = coordinate(address)
                        if row != number or column in cells:
                            raise ValueError(f"Лист {self.part}: повторяющаяся/неверная ячейка {address}")
                        cells[column] = self.cell(child)
                    attributes = dict(element.attrib)
                    attributes.pop("r", None)
                    attributes.pop("spans", None)  # Диапазон записанных ячеек — служебная подсказка.
                    if "s" in attributes:
                        attributes["s"] = self.package.style(attributes["s"])
                    self.stats["explicit_rows"] += 1
                    self.stats["data_rows"] += any(c["content"] is not None for c in cells.values())
                    yield number, cells, attributes
                    element.clear()
                    if data is not None:
                        data.clear()
                elif len(stack) == 2 and element.tag not in (M + "sheetData", M + "dimension"):
                    self.features[local(element.tag)].append(self.feature(element))
                    element.clear()
                stack.pop()


def raw_hash(package, part):
    hasher = hashlib.sha256()
    with package.zip.open(part) as stream:
        while block := stream.read(1024 * 1024):
            hasher.update(block)
    return hasher.hexdigest()


def part_kind(name):
    if name == "xl/calcChain.xml":
        return "Цепочка пересчёта формул"
    if name == "xl/sharedStrings.xml":
        return "Общее хранилище строк (возможен переход на inlineStr)"
    if name.startswith("xl/worksheets/") and name.endswith(".xml"):
        return "Данные и настройки листа"
    if name == "xl/styles.xml":
        return "Стили"
    if name.endswith(".rels"):
        return "Связи между частями книги"
    if name.startswith("docProps/"):
        return "Метаданные документа"
    if any(prefix in name for prefix in ("drawings/", "charts/", "media/", "pivot", "externalLinks/", "vbaProject", "connections", "queryTables/", "customXml/", "ctrlProps/", "activeX/", "slicer")):
        return "Объект/связь/расширение: требует проверки"
    return "Прочая часть книги"


def compare_parts(before, after, writer):
    result = []
    for name in sorted(before.entries.keys() | after.entries.keys()):
        a, b = before.entries.get(name), after.entries.get(name)
        hash_a = raw_hash(before, name) if a else None
        hash_b = raw_hash(after, name) if b else None
        status = "added" if a is None else "removed" if b is None else "unchanged" if hash_a == hash_b else "changed"
        if status == "unchanged" and a.compress_size != b.compress_size:
            status = "recompressed"
        xml_equal = None
        if status == "changed" and name.endswith((".xml", ".rels")) and not name.startswith("xl/worksheets/"):
            try:
                xml_equal = canonical(before.xml(name)) == canonical(after.xml(name))
            except ET.ParseError:
                xml_equal = False
        item = {"part": name, "status": status, "kind": part_kind(name),
                "compressed_before": a.compress_size if a else 0, "compressed_after": b.compress_size if b else 0,
                "raw_before": a.file_size if a else 0, "raw_after": b.file_size if b else 0,
                "sha256_before": hash_a, "sha256_after": hash_b, "xml_equal": xml_equal}
        item["compressed_delta"] = item["compressed_after"] - item["compressed_before"]
        writer.writerow(item)
        result.append(item)
    return result


def preview(value, show):
    if not show or value is None:
        return ""
    text = encoded(value)
    return text[:1000] + ("…" if len(text) > 1000 else "")


def compare_sheet(before, after, name, cell_writer, feature_writer, show_values):
    pa, pb = before.sheets.get(name), after.sheets.get(name)
    if (pa and pa["kind"] != "worksheet") or (pb and pb["kind"] != "worksheet"):
        return {"sheet": name, "state": "not_scanned", "reason": "Не обычный лист (например, лист диаграммы); см. parts.tsv"}
    readers = [SheetReader(before, pa["part"]) if pa else None, SheetReader(after, pb["part"]) if pb else None]
    streams = [reader.rows() if reader else iter(()) for reader in readers]
    current = [next(stream, None) for stream in streams]
    counts, examples, changed_rows = Counter(), defaultdict(list), set()
    default_styles = [package.style(0) for package in (before, after)]

    def change(address, kind, old=None, new=None):
        counts[kind] += 1
        changed_rows.add(coordinate(address)[0])
        if len(examples[kind]) < 20:
            examples[kind].append(address)
        cell_writer.writerow([name, address, kind, preview(old, show_values), preview(new, show_values)])

    while any(row is not None for row in current):
        number = min(row[0] for row in current if row is not None)
        rows = [row if row is not None and row[0] == number else (number, {}, {}) for row in current]
        old_cells, new_cells = rows[0][1], rows[1][1]
        if rows[0][2] != rows[1][2]:
            counts["row_properties_changed"] += 1
            feature_writer.writerow([name, f"row:{number}", "changed", digest(rows[0][2]), digest(rows[1][2]),
                                     preview(rows[0][2], show_values), preview(rows[1][2], show_values)])
        for column in sorted(old_cells.keys() | new_cells.keys()):
            old, new = old_cells.get(column), new_cells.get(column)
            a, b = old["content"] if old else None, new["content"] if new else None
            address = col_letters(column) + str(number)
            if a != b:
                kind = "cell_added" if a is None else "cell_removed" if b is None else "formula_changed" if a[0] == "formula" or b[0] == "formula" else "value_changed"
                change(address, kind, a, b)
            elif a is not None and a[0] == "formula":
                ca, cb = old["cache"], new["cache"]
                if ca != cb:
                    kind = "cache_removed" if cb is None else "cache_added" if ca is None else "cache_changed"
                    change(address, kind, ca, cb)
            style_a = old["style"] if old else [default_styles[0], None]
            style_b = new["style"] if new else [default_styles[1], None]
            if style_a != style_b:
                change(address, "cell_format_changed", style_a, style_b)
            extra_a = old["extra"] if old else [{}, []]
            extra_b = new["extra"] if new else [{}, []]
            if extra_a != extra_b:
                change(address, "cell_metadata_changed", extra_a, extra_b)
        for index in (0, 1):
            if current[index] is not None and current[index][0] == number:
                current[index] = next(streams[index], None)
    features = []
    fa, fb = readers[0].features if readers[0] else {}, readers[1].features if readers[1] else {}
    for feature in sorted(fa.keys() | fb.keys()):
        a, b = fa.get(feature, []), fb.get(feature, [])
        if a != b:
            features.append(feature)
            feature_writer.writerow([name, feature, "added" if not a else "removed" if not b else "changed", digest(a), digest(b),
                                     preview(a, show_values), preview(b, show_values)])
    return {"sheet": name, "state": "added" if pa is None else "removed" if pb is None else "compared",
            "counts": dict(counts), "examples": dict(examples), "features_changed": features,
            "rows_with_cell_differences": len(changed_rows),
            "before": dict(readers[0].stats) if readers[0] else {}, "after": dict(readers[1].stats) if readers[1] else {}}


LABELS = {
    "cell_added": "Добавлены непустые ячейки", "cell_removed": "Удалены значения/формулы",
    "value_changed": "Изменены значения", "formula_changed": "Изменены формулы или тип ячейки",
    "cache_removed": "Убран кэш при неизменной формуле", "cache_added": "Добавлен кэш при неизменной формуле",
    "cache_changed": "Изменён кэш при неизменной формуле", "cell_format_changed": "Изменено оформление ячеек",
    "cell_metadata_changed": "Изменены дополнительные свойства ячеек", "row_properties_changed": "Изменены свойства строк",
}


def workbook_features(package):
    result = {}
    groups = defaultdict(list)
    for child in package.workbook:
        if local(child.tag) != "sheets":
            groups[local(child.tag)].append(canonical(child))
    for key, value in groups.items():
        result[key] = digest(value)
    result["sheet_order_state_type"] = digest([[name, spec["state"], spec["kind"]] for name, spec in package.sheets.items()])
    return result


def text_report(summary):
    size_a, size_b = summary["before"]["bytes"], summary["after"]["bytes"]
    lines = ["СРАВНЕНИЕ XLSX — " + VERSION, "", "Исходник: " + summary["before"]["name"],
             "Результат: " + summary["after"]["name"],
             f"Размер: {size_a:,} → {size_b:,} байт; разница {size_b-size_a:+,} байт ({(size_b-size_a)/size_a:+.2%}).",
             "Входные файлы открывались только для чтения. Сеть и Excel не используются.", "",
             "ГДЕ ИЗМЕНИЛСЯ РАЗМЕР", "Изменения ниже — сжатые размеры частей внутри ZIP."]
    changed = sorted((p for p in summary["parts"] if p["compressed_delta"]), key=lambda p: p["compressed_delta"])
    for part in changed[:15]:
        lines.append(f"{part['compressed_delta']:+,} байт | {part['part']} | {part['status']} | {part['kind']}")
    if len(changed) > 15:
        lines.append(f"Ещё {len(changed)-15} частей: см. parts.tsv.")
    lines.append(f"Изменение ZIP-заголовков/служебного обрамления: {summary['zip_overhead_delta']:+,} байт.")
    lines.extend(["", "УДАЛЁННЫЕ ЧАСТИ АРХИВА"])
    removed = [p for p in summary["parts"] if p["status"] == "removed"]
    lines.extend([p["part"] + " — " + p["kind"] for p in removed] or ["Нет."])
    if any(p["part"] == "xl/sharedStrings.xml" for p in removed):
        lines.append("Удаление sharedStrings.xml само по себе не означает потерю текста: текст мог перейти в inlineStr. См. сравнение ячеек.")
    if any(p["part"] == "xl/calcChain.xml" for p in removed):
        lines.append("calcChain.xml — служебная цепочка пересчёта. Её удаление само по себе не означает удаления формул.")
    lines.extend(["", "СРАВНЕНИЕ ЛИСТОВ"])
    totals = Counter()
    for sheet in summary["sheets"]:
        lines.extend(["", f"Лист: {sheet['sheet']} ({sheet['state']})"])
        if sheet["state"] == "not_scanned":
            lines.append(sheet["reason"])
            continue
        counts = sheet["counts"]
        totals.update(counts)
        for key, label in (("data_rows", "Строк с данными"), ("content_cells", "Ячеек с данными"), ("formulas", "Формул"), ("cached_formulas", "Формул с кэшем")):
            lines.append(f"  {label}: {sheet['before'].get(key,0):,} → {sheet['after'].get(key,0):,}")
        for kind, label in LABELS.items():
            count = counts.get(kind, 0)
            if count:
                example = ", ".join(sheet["examples"].get(kind, []))
                lines.append(f"  {label}: {count:,}" + (f". Примеры: {example}" if example else ""))
        if not counts:
            lines.append("  Отличий ячеек и свойств строк не найдено.")
        lines.append("  Изменённые настройки листа: " + (", ".join(sheet["features_changed"]) or "нет"))
    lines.extend(["", "НАСТРОЙКИ КНИГИ", ", ".join(summary["workbook_features_changed"]) or "Отличий не найдено.",
                  "", "КАК ЧИТАТЬ РЕЗУЛЬТАТ"])
    lost_cache = totals["cache_removed"]
    if lost_cache:
        lines.append(f"Подтверждено удаление кэша у {lost_cache:,} неизменённых формул. Сами эти формулы сохранились.")
    core_changes = sum(totals[key] for key in ("cell_added", "cell_removed", "value_changed", "formula_changed"))
    lines.append(f"Отличий содержимого ячеек (без кэша и оформления): {core_changes:,}.")
    lines.append("Добавления и обновления импортёра тоже попадают в этот счётчик. Скрипт не знает, какие из них были запланированы.")
    lines.extend(["Точное распределение байтов дано по частям ZIP. Внутри изменённого XML нельзя надёжно отделить вклад кэша от повторного сжатия.",
                  "Все адреса изменений, без ограничения количества: cells.tsv. Настройки: features.tsv. Внутренние файлы: parts.tsv.",
                  "Проверка не вычисляет формулы, не проверяет бизнес-логику, внешний вид и работоспособность объектов.",
                  "Изменения сложных объектов, расширений и связей нужно рассматривать отдельно; отсутствие изменений ячеек не доказывает сохранность всех функций Excel.",
                  "Общие формулы разворачиваются по обычным A1-ссылкам. Изменение способа записи формулы/нетипичная формула может потребовать ручной проверки.",
                  "Даты, сохранённые разными способами (число Excel и ISO-дата), не приводятся к единому виду. Перестановки строк показываются как изменения адресов.",
                  "Настройки сравниваются после нормализации XML-префиксов, атрибутов, ссылок rId и части индексов стилей; некоторые эквивалентные записи могут отмечаться как изменённые.",
                  "Значения в TSV: " + ("ВКЛЮЧЕНЫ, первые 1000 символов на поле." if summary["show_values"] else "ВЫКЛЮЧЕНЫ. Есть имена листов, адреса и внутренние имена частей."),
                  "Отчёт остаётся на этом компьютере. Файлы отчёта не следует передавать без проверки их содержимого."])
    return "\n".join(lines) + "\n"


def file_info(path):
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            hasher.update(block)
    return {"name": path.name, "bytes": path.stat().st_size, "sha256": hasher.hexdigest()}


def compare_files(before_path, after_path, report_dir, show_values=False, progress=print):
    before_path, after_path, report_dir = Path(before_path), Path(after_path), Path(report_dir)
    if before_path.resolve() == after_path.resolve():
        raise ValueError("Выберите два разных файла")
    # Не перезаписываем ни входные файлы, ни предыдущие отчёты.
    if report_dir.exists():
        raise ValueError("Папка отчёта уже существует; укажите новую: " + str(report_dir))
    info_a, info_b = file_info(before_path), file_info(after_path)
    with ExitStack() as stack:
        before = Package(before_path)
        stack.callback(before.close)
        after = Package(after_path)
        stack.callback(after.close)
        report_dir.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix=".xlsx-compare-", dir=report_dir.parent) as temporary:
            folder = Path(temporary)
            files = [stack.enter_context((folder / name).open("w", encoding="utf-8-sig", newline=""))
                     for name in ("parts.tsv", "cells.tsv", "features.tsv")]
            fieldnames = ["part", "status", "kind", "compressed_before", "compressed_after", "raw_before", "raw_after", "sha256_before", "sha256_after", "xml_equal", "compressed_delta"]
            parts = csv.DictWriter(files[0], fieldnames, delimiter="\t")
            parts.writeheader()
            cells, features = csv.writer(files[1], delimiter="\t"), csv.writer(files[2], delimiter="\t")
            cells.writerow(["sheet", "cell", "change", "before_optional", "after_optional"])
            features.writerow(["sheet", "feature", "change", "before_sha256", "after_sha256", "before_optional", "after_optional"])
            progress("Сравниваю внутренние файлы ZIP…")
            inventory = compare_parts(before, after, parts)
            sheets = []
            for name in dict.fromkeys([*before.sheets, *after.sheets]):
                progress("Сравниваю лист: " + name)
                sheets.append(compare_sheet(before, after, name, cells, features, show_values))
            wa, wb = workbook_features(before), workbook_features(after)
            workbook_changes = sorted(key for key in wa.keys() | wb.keys() if wa.get(key) != wb.get(key))
            for key in workbook_changes:
                features.writerow(["[workbook]", key, "changed", wa.get(key, ""), wb.get(key, ""), "", ""])
            summary = {"version": VERSION, "before": info_a, "after": info_b, "show_values": show_values,
                       "parts": inventory, "sheets": sheets, "workbook_features_changed": workbook_changes,
                       "zip_overhead_delta": (info_b["bytes"] - sum(p["compressed_after"] for p in inventory)) - (info_a["bytes"] - sum(p["compressed_before"] for p in inventory))}
            (folder / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
            (folder / "report.txt").write_text(text_report(summary), encoding="utf-8-sig")
            for stream in files:
                stream.close()
            # mkdir(exist_ok=False) защищает от случайного перезаписывания отчёта.
            report_dir.mkdir()
            for file in folder.iterdir():
                file.rename(report_dir / file.name)
    return summary


def main(argv=None):
    parser = argparse.ArgumentParser(description="Сравнить исходник и результат XLSX локально, без зависимостей и изменения книг.")
    parser.add_argument("before", nargs="?", help="Исходный XLSX")
    parser.add_argument("after", nargs="?", help="Результат XLSX")
    parser.add_argument("-o", "--output", type=Path, help="Новая папка отчёта")
    parser.add_argument("--show-values", action="store_true", help="Включить значения и формулы в cells.tsv (до 1000 символов)")
    args = parser.parse_args(argv)
    try:
        if bool(args.before) != bool(args.after):
            parser.error("Укажите оба файла или запустите без аргументов")
        if args.before is None:
            args.before = input("Путь к исходному .xlsx: ").strip().strip("\"'")
            args.after = input("Путь к результату .xlsx: ").strip().strip("\"'")
        output = args.output or Path.cwd() / ("xlsx_compare_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f"))
        compare_files(args.before, args.after, output, args.show_values)
        print("\nГотово. Откройте: " + str((output / "report.txt").resolve()))
        print("Все изменения ячеек: " + str((output / "cells.tsv").resolve()))
        return 0
    except (OSError, ValueError, KeyError, IndexError, zipfile.BadZipFile, ET.ParseError, EOFError) as exc:
        print("Ошибка: " + str(exc), file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("\nСравнение остановлено.", file=sys.stderr)
        return 130


if __name__ == "__main__":
    sys.exit(main())
