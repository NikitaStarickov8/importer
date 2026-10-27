"""Импорт без запуска Excel. Все изменения сохраняются только после проверки пакета."""
from __future__ import annotations

import json
import os
import re
import tempfile
from collections import defaultdict
from copy import copy
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import get_column_letter, range_boundaries
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange

ROOT = Path(__file__).resolve().parent.parent


class ImportProblem(Exception):
    """Понятная пользователю ошибка; выходной файл не создаётся."""


def norm(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip().casefold()


def header(value):
    return re.sub(r"\s+", "", str(value or "").replace("\xa0", "")).upper()


def blank(value):
    return value is None or (isinstance(value, str) and not value.strip())


def formula(value):
    return isinstance(value, str) and value.startswith("=")


def load_config(path=None):
    with open(path or ROOT / "config.json", encoding="utf-8") as stream:
        config = json.load(stream)
    keys = config["group_key"] + config["row_key"]
    if len(keys) != len(set(keys)):
        raise ImportProblem("В настройке ключа повторяется одно из полей.")
    if any(name in keys for name in ("IFT", "PROM", "VALUE", "TEXT_FOR_CLIENT")):
        raise ImportProblem("Результаты решения и IFT/PROM нельзя включать в ключ записи.")
    return config


def sheet_named(book, name):
    matches = [s for s in book if norm(s.title) == norm(name)]
    # Excel ограничивает имя листа 31 символом; Р7 может сохранять более длинные имена.
    if not matches and len(name) > 31:
        matches = [s for s in book if norm(s.title) == norm(name[:31])]
    if len(matches) != 1:
        raise ImportProblem(f"Не найден единственный лист «{name}».")
    return matches[0]


def columns_at(sheet, row):
    result = {}
    for cells in sheet.iter_rows(min_row=row, max_row=row):
        for cell in cells:
            if not blank(cell.value):
                key = header(cell.value)
                if key in result:
                    raise ImportProblem(f"{sheet.title}, строка {row}: повторный столбец {cell.value}.")
                result[key] = cell.column
    return result


def literal(cell):
    if cell.data_type == "f":
        raise ImportProblem(f"{cell.parent.title}!{cell.coordinate}: неожиданная формула. "
                            "Для этого поля требуется исходное значение.")
    if cell.data_type == "e":
        raise ImportProblem(f"{cell.parent.title}!{cell.coordinate}: ошибка {cell.value}.")
    return cell.value


@dataclass(frozen=True)
class Entry:
    code: object
    description: object


class Catalog:
    def __init__(self, book, config):
        self.book = book
        self.config = config
        self.indices = {}
        self.cache = {}

    def resolve(self, name, value):
        cache_key = (name, norm(value))
        if cache_key in self.cache:
            return self.cache[cache_key]
        if name not in self.indices:
            spec = self.config["lookups"][name]
            sheet = sheet_named(self.book, spec["sheet"])
            index = defaultdict(set)
            for row in sheet.iter_rows(min_row=2, values_only=True):
                code = row[spec["code_col"] - 1] if len(row) >= spec["code_col"] else None
                desc = row[spec["description_col"] - 1] if len(row) >= spec["description_col"] else None
                if blank(code) and blank(desc):
                    continue
                if blank(code) or blank(desc) or formula(code) or formula(desc):
                    raise ImportProblem(f"Справочник «{sheet.title}»: пустой код/описание или формула.")
                entry = Entry(code, desc)
                index[norm(code)].add(entry)
                index[norm(desc)].add(entry)
            self.indices[name] = index
        candidates = self.indices[name].get(norm(value), set())
        if len(candidates) != 1:
            reason = "не найдено" if not candidates else "соответствует нескольким записям"
            raise ImportProblem(f"{name}: значение «{value}» {reason} в справочнике.")
        entry = next(iter(candidates))
        self.cache[cache_key] = entry
        return entry


def amount(value):
    if blank(value):
        return None
    if isinstance(value, bool):
        raise ImportProblem("MAX_AMOUNT должен быть числом, а не логическим значением.")
    try:
        number = Decimal(str(value).replace("\xa0", "").replace(" ", "").replace(",", "."))
    except InvalidOperation as exc:
        raise ImportProblem(f"MAX_AMOUNT: нечисловое значение «{value}».") from exc
    if not number.is_finite() or number < 0:
        raise ImportProblem("MAX_AMOUNT должен быть конечным неотрицательным числом.")
    return int(number) if number == number.to_integral_value() else float(number)


@dataclass
class Record:
    values: dict
    source: str

    def key(self, config):
        fields = config["group_key"] + config["row_key"]
        parts = [norm(self.values.get(field)) for field in fields]
        platform = self.values.get("SBOL_PRO_PLATFORM") if norm(self.values.get("CHANNEL_INDICATOR")) == norm(config["platform_channel"]) else None
        return tuple(parts + [norm(platform)])


def canonical(raw, catalog, config, source, require_payload=True):
    values = dict(raw)
    required = set(config["group_key"] + ["STOP_LIST_STATUS", "MATCH_NEGATIVE_CASE", "INSPECTED_PERSON"])
    if require_payload:
        required.add("VALUE")
    for name, spec in config["lookups"].items():
        if name == "VALUE" and not require_payload:
            continue
        description = spec["description_field"]
        code_value = raw.get(name)
        desc_value = raw.get(description) if name != "VALUE" else None
        value = code_value if not blank(code_value) else desc_value
        if blank(value):
            if name in required:
                raise ImportProblem(f"Не заполнено обязательное поле {name}.")
            values[name] = None
            values[description] = None
            continue
        entry = catalog.resolve(name, value)
        if not blank(code_value) and not blank(desc_value):
            if catalog.resolve(name, desc_value) != entry:
                raise ImportProblem(f"{name}: код «{code_value}» не соответствует описанию «{desc_value}».")
        values[name] = entry.code
        values[description] = entry.description
    if norm(values.get("CHANNEL_INDICATOR")) == norm(config["platform_channel"]) and blank(values.get("SBOL_PRO_PLATFORM")):
        raise ImportProblem("Для канала SBOL_PRO требуется SBOL_PRO_PLATFORM.")
    values["MAX_AMOUNT"] = amount(raw.get("MAX_AMOUNT"))
    return Record(values, source)


TOP_CELLS = {
    "CHANNEL_INDICATOR_DESCRIPTION": ("D3", "B3"),
    "SP_NAME_DESCRIPTION": ("D4", "B4"),
    "SP_NAME": ("D5", "B5"),
    "OPERATION_DESCRIPTION": ("D6", "B6"),
    "OPERATION_TYPE": ("D7", "B7"),
    "CLIENT_TYPE_DESCRIPTION": ("B8",),
    "RELATION_TYPE_DESCRIPTION": ("B10",),
    "RELATION_TYPE": ("D10",),
}


def parse_template(path, catalog, config):
    book = load_workbook(path, read_only=False, data_only=False)
    try:
        matches = [s for s in book if norm(config["template_sheet_contains"]) in norm(s.title)]
        if len(matches) != 1:
            raise ImportProblem("Не найден единственный лист шаблона с «ЗАПОЛНИТЬ» в имени.")
        sheet = matches[0]
        cols, header_row = None, None
        marker_fields = {"STOP_LIST_STATUS_DESCRIPTION", "MATCH_NEGATIVE_CASE_DESCRIPTION", "VALUE"}
        for row in range(1, min(sheet.max_row, config["template_header_scan_rows"]) + 1):
            row_values = [header(c.value) for c in sheet[row]]
            if marker_fields.issubset(row_values):
                cols, header_row = columns_at(sheet, row), row
                break
        if cols is None:
            raise ImportProblem("Не найдены технические заголовки STOP_LIST_STATUS_DESCRIPTION, MATCH_NEGATIVE_CASE_DESCRIPTION, VALUE.")
        missing = {"MAX_AMOUNT", "TITLE_FOR_CLIENT", "TEXT_FOR_CLIENT"} - cols.keys()
        if missing:
            raise ImportProblem("В шаблоне отсутствуют столбцы: " + ", ".join(sorted(missing)))
        common = {}
        for field, addresses in TOP_CELLS.items():
            for address in addresses:
                value = literal(sheet[address])
                if not blank(value):
                    common[field] = value
                    break
        # Известную D8 вычисляем по B8 через справочник матрицы, без кэша Excel.
        client_code = sheet["D8"]
        if client_code.data_type != "f" and not blank(client_code.value):
            common["CLIENT_TYPE"] = literal(client_code)
        result = []
        known = {header(k): k for k in config["lookups"]}
        known.update({header(v["description_field"]): v["description_field"] for k, v in config["lookups"].items() if k != "VALUE"})
        for field in ("VALUE", "MAX_AMOUNT", "TITLE_FOR_CLIENT", "TEXT_FOR_CLIENT"):
            known[header(field)] = field
        bypass_col = cols.get(header("Поле Байпас"))
        for row in range(max(header_row + 1, config["template_data_start_row"]), sheet.max_row + 1):
            value_col = cols["VALUE"]
            signals = [sheet.cell(row, cols[x]).value for x in marker_fields]
            if all(blank(x) for x in signals):
                continue
            source = f"{Path(path).name}, {sheet.title}!{row}"
            try:
                raw = {field: literal(sheet.cell(row, cols[key])) for key, field in known.items() if key in cols}
                # Поля строки приоритетны; связанные общие коды не должны перебить их.
                for field, val in common.items():
                    desc = config["lookups"].get(field, {}).get("description_field")
                    if blank(raw.get(field)) and (not desc or blank(raw.get(desc))):
                        raw[field] = val
                raw["VALUE"] = literal(sheet.cell(row, value_col))
                raw["TITLE_FOR_CLIENT"] = raw.get("TITLE_FOR_CLIENT")
                raw["TEXT_FOR_CLIENT"] = raw.get("TEXT_FOR_CLIENT")
                if bypass_col:
                    bypass = literal(sheet.cell(row, bypass_col))
                    mapping = {norm(k): v for k, v in config["bypass_values"].items()}
                    if blank(bypass):
                        raw[config["bypass_target"]] = None
                    elif norm(bypass) in mapping:
                        raw[config["bypass_target"]] = mapping[norm(bypass)]
                    else:
                        raise ImportProblem(f"Неизвестное значение поля Байпас: «{bypass}».")
                result.append(canonical(raw, catalog, config, source))
            except ImportProblem as exc:
                raise ImportProblem(f"{source}: {exc}") from exc
        if not result:
            raise ImportProblem("В шаблоне нет строк с решениями.")
        return result
    finally:
        book.close()


def read_matrix_row(sheet, cols, row, catalog, config):
    raw = {}
    for field in config["group_key"] + config["row_key"] + ["SBOL_PRO_PLATFORM"]:
        if field in config["lookups"]:
            desc = config["lookups"][field]["description_field"]
            for f in (field, desc):
                if header(f) in cols:
                    cell = sheet.cell(row, cols[header(f)])
                    if cell.data_type != "f":
                        raw[f] = literal(cell)
        elif header(field) in cols:
            raw[field] = literal(sheet.cell(row, cols[header(field)]))
    return canonical(raw, catalog, config, f"{sheet.title}!{row}", require_payload=False)


def payload(record, cols, config):
    result = {}
    for field, spec in config["lookups"].items():
        if field == "VALUE":
            continue
        desc = spec["description_field"]
        if header(desc) in cols:
            result[desc] = record.values.get(desc)
        if header(field) in cols:
            result[field] = record.values.get(field)
    for field in ("VALUE", "MAX_AMOUNT", "TITLE_FOR_CLIENT", "TEXT_FOR_CLIENT", config["bypass_target"]):
        if field in record.values:
            result[field] = record.values[field]
    return result


def text_write(cell, value):
    cell.value = value
    if isinstance(value, str):
        # Введённый пользователем текст, начинающийся на '=', остаётся текстом.
        cell.data_type = "s"


def grow_ranges(sheet, old_end, new_end, start):
    if new_end <= old_end:
        return
    def extend(rng):
        bounds = range_boundaries(str(rng))
        a, b, c, d = bounds
        return f"{get_column_letter(a)}{b}:{get_column_letter(c)}{new_end}" if old_end <= d < new_end and b <= start else str(rng)
    for table in sheet.tables.values():
        table.ref = extend(table.ref)
        if table.autoFilter:
            table.autoFilter.ref = table.ref
    if sheet.auto_filter.ref:
        sheet.auto_filter.ref = extend(sheet.auto_filter.ref)
    for validation in sheet.data_validations.dataValidation:
        validation.sqref = MultiCellRange([CellRange(extend(r)) for r in validation.sqref.ranges])
    conditional = []
    for cf in sheet.conditional_formatting:
        conditional.append((str(cf.sqref), list(sheet.conditional_formatting[cf])))
    for old, rules in conditional:
        new = " ".join(extend(r) for r in old.split())
        if new != old:
            del sheet.conditional_formatting[old]
            for rule in rules:
                sheet.conditional_formatting.add(new, rule)


def append_row(sheet, cols, row, proto, values, config):
    for col in range(1, sheet.max_column + 1):
        src, dst = sheet.cell(proto, col), sheet.cell(row, col)
        dst._style = copy(src._style)
        dst.protection = copy(src.protection)
        if src.data_type == "f":
            if not isinstance(src.value, str):
                raise ImportProblem(f"{src.coordinate}: массивная формула не поддерживается прототипом.")
            dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)
    sheet.row_dimensions[row].height = sheet.row_dimensions[proto].height
    for field, value in config["new_row_defaults"].items():
        if header(field) in cols:
            text_write(sheet.cell(row, cols[header(field)]), value)
    # Если CODE хранится числом, а не формулой, используем последовательность строк.
    if "CODE" in cols and sheet.cell(row, cols["CODE"]).data_type != "f":
        sheet.cell(row, cols["CODE"], row - config["matrix_header_row"])
    for field, value in values.items():
        cell = sheet.cell(row, cols[header(field)])
        if cell.data_type != "f":
            text_write(cell, value)


@dataclass
class Result:
    added: int = 0
    updated: int = 0
    skipped: int = 0
    total: int = 0
    output: str = ""


def import_files(matrix, templates, output, config_path=None, progress: Callable[[str], None] = lambda _: None):
    config = load_config(config_path)
    matrix, output = Path(matrix).resolve(), Path(output).resolve()
    templates = [Path(p).resolve() for p in templates]
    if not templates:
        raise ImportProblem("Выберите хотя бы один шаблон.")
    inputs = [matrix, *templates]
    if any(p.suffix.lower() != ".xlsx" for p in [*inputs, output]):
        raise ImportProblem("Первая версия поддерживает только .xlsx.")
    if output in inputs or output.exists():
        raise ImportProblem("Выберите новое имя результата: входные и существующие файлы не перезаписываются.")
    for path in inputs:
        if not path.is_file():
            raise ImportProblem(f"Файл не найден: {path.name}")
    if matrix in templates:
        raise ImportProblem("Основная матрица выбрана как шаблон.")
    progress("Читаю основную матрицу…")
    book = load_workbook(matrix, data_only=False, keep_links=True, rich_text=True)
    try:
        sheet = sheet_named(book, config["matrix_sheet"])
        if sheet.protection.sheet:
            raise ImportProblem("Основной лист защищён от изменений.")
        cols = columns_at(sheet, config["matrix_header_row"])
        needed = {"VALUE", "MAX_AMOUNT", "TITLE_FOR_CLIENT", "TEXT_FOR_CLIENT", config["bypass_target"]}
        for field, spec in config["lookups"].items():
            if field != "VALUE":
                needed.add(spec["description_field"])
                needed.add(field)
        missing = needed - cols.keys()
        if missing:
            raise ImportProblem("В матрице нет столбцов: " + ", ".join(sorted(missing)))
        catalog = Catalog(book, config)
        incoming = []
        for path in templates:
            progress(f"Читаю шаблон {path.name}…")
            try:
                incoming.extend(parse_template(path, catalog, config))
            except ImportProblem as exc:
                raise ImportProblem(f"{path.name}: {exc}") from exc
        result = Result(total=len(incoming), output=str(output))
        progress("Сопоставляю записи матрицы…")
        index = defaultdict(list)
        last = config["matrix_header_row"]
        for row in range(last + 1, sheet.max_row + 1):
            # Предзаполненные формулы ниже данных не считаем записями.
            active = [sheet.cell(row, cols[header(f)]).value for f in ("SP_NAME_DESCRIPTION", "OPERATION_DESCRIPTION", "STOP_LIST_STATUS_DESCRIPTION")]
            if all(blank(v) for v in active):
                continue
            try:
                record = read_matrix_row(sheet, cols, row, catalog, config)
            except ImportProblem as exc:
                raise ImportProblem(f"Матрица, строка {row}: {exc}") from exc
            index[record.key(config)].append(row)
            last = row
            if row % 5000 == 0:
                progress(f"Проверено строк матрицы: {row:,}")
        old_end = last
        plans, seen = [], {}
        for record in incoming:
            key = record.key(config)
            values = payload(record, cols, config)
            signature = tuple(sorted((k, None if blank(v) else v) for k, v in values.items()))
            if key in seen:
                old_signature, source = seen[key]
                if signature != old_signature:
                    raise ImportProblem(f"Конфликт шаблонов: {source} и {record.source} описывают одну запись разными значениями.")
                result.skipped += 1
                continue
            seen[key] = (signature, record.source)
            matches = index.get(key, [])
            if len(matches) > 1:
                raise ImportProblem(f"{record.source}: в матрице несколько подходящих строк: {matches[:10]}.")
            if matches:
                row = matches[0]
                changes = {}
                for field, value in values.items():
                    cell = sheet.cell(row, cols[header(field)])
                    if cell.data_type == "f":
                        # Формулы кодов остаются; входные поля должны быть редактируемыми.
                        if field not in config["lookups"]:
                            raise ImportProblem(f"{record.source}: поле {field} в матрице вычисляется формулой; автоматическая замена запрещена.")
                        continue
                    current = cell.value
                    equal = (blank(current) and blank(value)) or (type(current) == type(value) and current == value)
                    if isinstance(current, (int, float)) and not isinstance(current, bool) and isinstance(value, (int, float)) and not isinstance(value, bool):
                        equal = current == value
                    if not equal:
                        changes[field] = value
                if changes:
                    plans.append((row, changes, False))
                    result.updated += 1
                else:
                    result.skipped += 1
            else:
                last += 1
                plans.append((last, values, True))
                result.added += 1
        proto = config["prototype_row"]
        if proto <= config["matrix_header_row"] or proto > old_end:
            raise ImportProblem("В матрице нет образцовой строки для копирования формул и оформления.")
        # Проверяем отсутствие вычислений во входных столбцах до добавления строк.
        for field in payload(incoming[0], cols, config):
            if field not in config["lookups"] and sheet.cell(proto, cols[header(field)]).data_type == "f":
                raise ImportProblem(f"Образцовая строка: входное поле {field} содержит формулу.")
        progress("Записываю изменения и сохраняю новый файл…")
        for row, values, is_new in plans:
            if is_new:
                if any(not blank(c.value) and c.data_type != "f" for c in sheet[row]):
                    raise ImportProblem(f"Строка {row} ниже таблицы содержит данные. Добавление поверх неё остановлено.")
                append_row(sheet, cols, row, proto, values, config)
            else:
                for field, value in values.items():
                    text_write(sheet.cell(row, cols[header(field)]), value)
        grow_ranges(sheet, old_end, last, config["matrix_header_row"] + 1)
        book.calculation = CalcProperties(calcId=0, fullCalcOnLoad=True, forceFullCalc=True, calcMode="auto")
        output.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp = tempfile.mkstemp(prefix=".matrix-import-", suffix=".xlsx", dir=output.parent)
        os.close(fd)
        try:
            book.save(tmp)
            # Эксклюзивное создание не позволяет затереть файл, появившийся за время работы.
            with output.open("xb") as dest, open(tmp, "rb") as src:
                import shutil
                try:
                    shutil.copyfileobj(src, dest)
                except BaseException:
                    dest.close()
                    output.unlink(missing_ok=True)
                    raise
        finally:
            Path(tmp).unlink(missing_ok=True)
        return result
    finally:
        book.close()
