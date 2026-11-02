"""Ограниченное, только читающее вычисление английских ключей матрицы.

Поддерживает ссылки и INDEX/MATCH с точным поиском. Это не движок Excel:
никакой eval, кэша Excel, внешних книг или записи вычисленных значений.
"""
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
import re

from openpyxl.formula.tokenizer import Tokenizer, TokenizerError
from openpyxl.utils.cell import range_boundaries


class CodeFormulaError(Exception):
    pass


@dataclass(frozen=True)
class Reference:
    sheet: object
    left: int
    top: int
    right: int
    bottom: int


class CodeFormulaReader:
    def __init__(self, book):
        self.book = book
        self.match_indices = {}
        self.expressions = {}

    def value(self, cell, visited=frozenset()):
        location = (cell.parent.title, cell.row, cell.column)
        if location in visited or len(visited) >= 64:
            raise CodeFormulaError('циклическая или слишком длинная цепочка ссылок')
        if cell.data_type == 'e':
            raise CodeFormulaError(f'{cell.coordinate}: ошибка {cell.value}')
        if cell.data_type != 'f':
            return cell.value
        if not isinstance(cell.value, str) or len(cell.value) > 16000:
            raise CodeFormulaError('неподдерживаемая формула кода')
        expression = self._parse(cell.value)
        return self._scalar(self._evaluate(expression, cell, visited | {location}), cell, visited | {location})

    def _parse(self, formula):
        if formula in self.expressions:
            return self.expressions[formula]
        try:
            tokens = [t for t in Tokenizer(formula).items if t.type != 'WHITE-SPACE']
        except (TokenizerError, IndexError, ValueError) as exc:
            raise CodeFormulaError('нельзя разобрать формулу кода') from exc
        if len(tokens) > 256:
            raise CodeFormulaError('слишком сложная формула кода')
        position = 0

        def parse(depth=0):
            nonlocal position
            if depth > 32 or position >= len(tokens):
                raise CodeFormulaError('неполная или слишком сложная формула кода')
            token = tokens[position]
            position += 1
            if token.type == 'OPERAND':
                if token.subtype == 'RANGE':
                    return ('ref', token.value)
                if token.subtype == 'NUMBER':
                    try:
                        return ('literal', Decimal(token.value))
                    except InvalidOperation as exc:
                        raise CodeFormulaError('неверное число в формуле') from exc
                if token.subtype == 'TEXT':
                    return ('literal', token.value[1:-1].replace('""', '"'))
            if token.type == 'FUNC' and token.subtype == 'OPEN':
                name = token.value[:-1].lstrip('@').upper()
                name = {'ИНДЕКС':'INDEX', 'ПОИСКПОЗ':'MATCH'}.get(name, name)
                if name not in ('INDEX','MATCH'):
                    raise CodeFormulaError(f'неподдерживаемая функция кода: {name}')
                args = [parse(depth + 1)]
                while position < len(tokens) and tokens[position].type == 'SEP':
                    position += 1
                    args.append(parse(depth + 1))
                if position >= len(tokens) or tokens[position].type != 'FUNC' or tokens[position].subtype != 'CLOSE':
                    raise CodeFormulaError('неверные аргументы формулы кода')
                position += 1
                return (name, args)
            raise CodeFormulaError('поддерживаются только ссылки и INDEX/MATCH с точным поиском')

        expression = parse()
        if position != len(tokens):
            raise CodeFormulaError('в формуле есть неподдерживаемые операции')
        if len(self.expressions) < 256:
            self.expressions[formula] = expression
        return expression

    def _reference(self, value, caller):
        if '[' in value or ']' in value:
            raise CodeFormulaError('внешние книги и структурированные ссылки не поддерживаются')
        name, separator, address = value.rpartition('!')
        if not separator:
            sheet, address = caller.parent, value
        else:
            name = name.strip()
            if name.startswith("'") and name.endswith("'"):
                name = name[1:-1].replace("''", "'")
            matches = [s for s in self.book if s.title.casefold() == name.casefold()]
            if not matches and len(name) > 31:
                matches = [s for s in self.book if s.title.casefold() == name[:31].casefold()]
            if len(matches) != 1:
                raise CodeFormulaError(f'не найден единственный лист формулы «{name}»')
            sheet = matches[0]
        address = address.lstrip('@')
        if not re.fullmatch(r'\$?[A-Za-z]{1,3}(?:\$?[1-9][0-9]*)?(?::\$?[A-Za-z]{1,3}(?:\$?[1-9][0-9]*)?)?', address):
            raise CodeFormulaError('неподдерживаемая ссылка в формуле кода')
        try:
            left, top, right, bottom = range_boundaries(address)
        except ValueError as exc:
            raise CodeFormulaError('неверный диапазон формулы кода') from exc
        if (top is None) != (bottom is None) or top is None and ':' not in address:
            raise CodeFormulaError('неверная ссылка на столбец')
        top, bottom = top or 1, bottom or 1048576
        if not (1 <= left <= right <= 16384 and 1 <= top <= bottom <= 1048576):
            raise CodeFormulaError('диапазон вне границ Excel')
        return Reference(sheet, left, top, right, bottom)

    def _cell(self, reference, row, col, visited):
        # Не создаём даже пустые ячейки в загруженной книге при чтении диапазона.
        cell = reference.sheet._cells.get((row, col))
        return None if cell is None else self.value(cell, visited)

    def _scalar(self, value, caller, visited):
        if not isinstance(value, Reference):
            return value
        if value.left != value.right:
            raise CodeFormulaError('неоднозначное пересечение диапазона')
        if value.top == value.bottom:
            row = value.top
        elif value.top <= caller.row <= value.bottom:
            # MATCH('Матрица'!O:O, ...) использует значение O текущей строки.
            row = caller.row
        else:
            raise CodeFormulaError('диапазон не пересекает текущую строку')
        return self._cell(value, row, value.left, visited)

    @staticmethod
    def _integer(value):
        if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
            raise CodeFormulaError('требуется целочисленный аргумент INDEX/MATCH')
        number = Decimal(str(value))
        if not number.is_finite() or abs(number) > 1048576 or number != number.to_integral_value():
            raise CodeFormulaError('целочисленный аргумент INDEX/MATCH вне границ Excel')
        return int(number)

    @staticmethod
    def _key(value):
        if isinstance(value, str):
            # MATCH нечувствителен к регистру, но не удаляет пробелы из ключа.
            return ('text', value.casefold())
        if isinstance(value, bool):
            return ('bool', value)
        if isinstance(value, (int, float, Decimal)):
            return ('number', value)
        return ('other', value)

    def _match(self, needle, reference, visited):
        if not isinstance(reference, Reference) or reference.left != reference.right:
            raise CodeFormulaError('MATCH требует один столбец поиска')
        if needle is None or needle == '':
            raise CodeFormulaError('пустой ключ поиска английского кода')
        if isinstance(needle, str) and any(c in needle for c in '*?~'):
            raise CodeFormulaError('шаблоны с * ? ~ в MATCH требуют ручного разбора')
        key = (reference.sheet.title, reference.left, reference.top, reference.bottom)
        if key not in self.match_indices:
            first, error = {}, None
            for row in range(reference.top, min(reference.bottom, reference.sheet.max_row) + 1):
                index = row - reference.top + 1
                try:
                    value = self._cell(reference, row, reference.left, visited)
                except CodeFormulaError as exc:
                    if error is None:
                        error = (index, str(exc))
                else:
                    if value is not None:
                        first.setdefault(self._key(value), index)
            self.match_indices[key] = (first, error)
        first, error = self.match_indices[key]
        index = first.get(self._key(needle))
        if error is not None and (index is None or error[0] <= index):
            raise CodeFormulaError(f'ошибка в диапазоне MATCH: {error[1]}')
        if index is None:
            raise CodeFormulaError('точное совпадение в диапазоне MATCH не найдено')
        return index

    def _evaluate(self, expression, caller, visited):
        kind, value = expression
        if kind == 'literal':
            return value
        if kind == 'ref':
            return self._reference(value, caller)
        if kind == 'MATCH':
            if len(value) != 3:
                raise CodeFormulaError('MATCH должен явно содержать режим точного поиска 0')
            needle = self._scalar(self._evaluate(value[0], caller, visited), caller, visited)
            reference = self._evaluate(value[1], caller, visited)
            mode = self._integer(self._scalar(self._evaluate(value[2], caller, visited), caller, visited))
            if mode != 0:
                raise CodeFormulaError('приблизительный MATCH не поддерживается')
            return self._match(needle, reference, visited)
        if kind == 'INDEX':
            if len(value) not in (2,3):
                raise CodeFormulaError('INDEX должен содержать два или три аргумента')
            reference = self._evaluate(value[0], caller, visited)
            if not isinstance(reference, Reference):
                raise CodeFormulaError('INDEX требует диапазон исходных кодов')
            row = self._integer(self._scalar(self._evaluate(value[1], caller, visited), caller, visited))
            col = self._integer(self._scalar(self._evaluate(value[2], caller, visited), caller, visited)) if len(value) == 3 else 1
            if not (1 <= row <= reference.bottom-reference.top+1 and 1 <= col <= reference.right-reference.left+1):
                raise CodeFormulaError('INDEX выходит за границы исходного диапазона')
            return self._cell(reference, reference.top + row - 1, reference.left + col - 1, visited)
        raise CodeFormulaError('неподдерживаемая формула английского кода')
