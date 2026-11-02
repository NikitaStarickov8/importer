"""Коды операции/подсистемы первичны; названия берутся из первой строки кода."""
import unittest
import test_business_rules as business
from app.importer import Catalog, ImportProblem, columns_at, load_config, read_matrix_row
from openpyxl import load_workbook


class PrimaryCodeTests(unittest.TestCase):
    setUp = business.BusinessRulesTests.setUp
    prepare_matrix = business.BusinessRulesTests.prepare_matrix
    change = business.BusinessRulesTests.change
    template = business.BusinessRulesTests.template
    run_import = business.BusinessRulesTests.run_import
    rows = business.BusinessRulesTests.rows
    empty_group = business.BusinessRulesTests.empty_group
    specs = {
        'SP_NAME': ('Подсистемы', 'DEMO_LOAN', 'Демонстрационная подсистема', 'B', 'B5', 'D5', 'H', 'G'),
        'OPERATION_TYPE': ('Типы Операций', 'CHECK_BALANCE', 'Проверка баланса', 'C', 'B7', 'D7', 'P', 'O'),
    }

    def primary_template(self, name, description='Опечатка', code=None, edit=lambda s: None, blank_header=False, resolution='З'):
        _, default_code, _, col, manual, calculated, _, _ = self.specs[name]
        template = self.template([('01 — Тестовое ограничение', resolution, self.empty_group(**{col:description}))], name=f'{name}_base.xlsx')
        def modify(book):
            sheet = book.worksheets[0]
            sheet[manual], sheet[calculated] = None if blank_header else code or default_code, None
            edit(sheet)
        return self.change(template, modify, f'{name}_template.xlsx')

    def assert_new_rows(self, template, matrix, name, expected_description=None, literal=False):
        _, code, description, _, _, _, target, desc_target = self.specs[name]
        output = self.folder/f'{name}_out.xlsx'
        result = self.run_import(template, matrix=matrix, output=output)
        self.assertEqual((result.added, len(result.conflicts)), (3, 0), result.conflicts)
        book = load_workbook(output)
        try:
            sheet = book.worksheets[0]
            config = load_config()
            catalog = Catalog(book, config)
            for row in range(6, 9):
                self.assertEqual(sheet[f'{desc_target}{row}'].value, expected_description or description)
                record = read_matrix_row(sheet, columns_at(sheet, 1), row, catalog, config)
                self.assertEqual(record.values[name], code)
                if literal:
                    self.assertEqual(sheet[f'{target}{row}'].value, code)
                    self.assertEqual(sheet[f'{target}{row}'].data_type, 's')
                self.assertEqual(sheet[f'A{row}'].data_type, 'f')
                if record.values['STOP_LIST_STATUS'] in ('NORMAL', 'TECH_ERROR'):
                    self.assertEqual(sheet[f'AA{row}'].value, record.values['STOP_LIST_STATUS'])
                    self.assertEqual(sheet[f'AA{row}'].data_type, 's')
                else:
                    self.assertEqual(sheet[f'AA{row}'].data_type, 'f')
        finally:
            book.close()
        repeat = self.run_import(template, matrix=output, output=self.folder/f'{name}_repeat.xlsx')
        self.assertEqual(repeat.affected_rows, [])
        return output

    def test_same_code_with_multiple_russian_names_uses_first_sheet_row(self):
        for name, (ref, code, _, *_) in self.specs.items():
            with self.subTest(name=name):
                def edit(book):
                    book[ref].append([code, 'Альтернативное название'])
                    book[ref].append([code, 'Ещё одно название'])
                matrix = self.change(self.matrix, edit, f'{name}_aliases.xlsx')
                self.assert_new_rows(self.primary_template(name), matrix, name)

    def test_first_row_is_respected_when_order_changes(self):
        for name, (ref, code, _, *_) in self.specs.items():
            with self.subTest(name=name):
                def edit(book):
                    sheet = book[ref]
                    sheet.insert_rows(2)
                    sheet['A2'], sheet['B2'] = code, 'Первое по порядку название'
                matrix = self.change(self.matrix, edit, f'{name}_first.xlsx')
                self.assert_new_rows(self.primary_template(name), matrix, name, 'Первое по порядку название')

    def test_shared_russian_label_cannot_change_either_english_code(self):
        for name, (ref, code, description, *_) in self.specs.items():
            with self.subTest(name=name):
                def edit(book):
                    sheet = book[ref]
                    sheet.insert_rows(2)
                    sheet['A2'], sheet['B2'] = 'WRONG_CODE', description
                    sheet.append([code,'Дополнительное имя того же кода'])
                matrix = self.change(self.matrix, edit, f'{name}_shared.xlsx')
                output = self.assert_new_rows(self.primary_template(name), matrix, name, literal=True)
                prom = self.run_import(self.primary_template(name), matrix=output, environment='PROM', output=self.folder/f'{name}_prom.xlsx')
                self.assertEqual((prom.added, prom.promoted, len(prom.conflicts)), (0, 3, 0))

    def test_missing_english_code_and_shared_russian_label_stays_manual(self):
        for name, (ref, _, description, *_) in self.specs.items():
            with self.subTest(name=name):
                matrix = self.change(self.matrix, lambda b: b[ref].append(['WRONG_CODE',description]), f'{name}_no_code.xlsx')
                result = self.run_import(self.primary_template(name, description=description, blank_header=True), matrix=matrix, output=self.folder/f'{name}_manual.xlsx')
                self.assertEqual(result.added, 0)
                self.assertEqual(len(result.conflicts), 1)

    def test_alternate_russian_name_with_single_code_uses_first_name(self):
        for name, (ref, code, _, *_) in self.specs.items():
            with self.subTest(name=name):
                matrix = self.change(self.matrix, lambda b: b[ref].append([code,'Альтернативное русское имя']), f'{name}_alternate.xlsx')
                self.assert_new_rows(self.primary_template(name, description='Альтернативное русское имя', blank_header=True), matrix, name)

    def test_subsystem_typo_and_formula_error_use_manual_english_code(self):
        for index, description in enumerate(('Опечатка в названии', '#N/A', '=UNSUPPORTED_LOOKUP(1)')):
            with self.subTest(description=description):
                template = self.primary_template('SP_NAME', description=description)
                result = self.run_import(template, output=self.folder/f'typo_{index}.xlsx')
                self.assertEqual((result.added, len(result.conflicts)), (3, 0))

    def test_subsystem_d5_formula_falls_back_to_b5(self):
        template = self.primary_template('SP_NAME', edit=lambda s: setattr(s['D5'],'value','=UNSUPPORTED_LOOKUP(B4)'))
        self.assert_new_rows(template, self.matrix, 'SP_NAME')

    def test_explicit_subsystem_code_wins_over_bad_header_and_description(self):
        def edit(sheet):
            sheet['R16'], sheet['R18'] = 'SP_NAME', 'DEMO_LOAN'
            sheet['B5'], sheet['D5'] = 'UNKNOWN_CODE', '=BAD_FUNCTION(1)'
        self.assert_new_rows(self.primary_template('SP_NAME', description='#N/A', edit=edit), self.matrix, 'SP_NAME')

    def test_unknown_subsystem_code_does_not_fall_back_to_russian(self):
        with self.assertRaisesRegex(ImportProblem, 'UNKNOWN_CODE'):
            self.run_import(self.primary_template('SP_NAME', description='Демонстрационная подсистема', code='UNKNOWN_CODE'))
        self.assertFalse(self.output.exists())

    def test_subsystem_without_english_code_still_supports_unique_russian_name(self):
        self.assert_new_rows(self.primary_template('SP_NAME', description='Демонстрационная подсистема', blank_header=True), self.matrix, 'SP_NAME')

    def test_skipped_allowed_rows_use_subsystem_english_code_for_mandatory_rows(self):
        result = self.run_import(self.primary_template('SP_NAME', description='#N/A', resolution='Р'))
        self.assertEqual((result.added, result.skipped), (2, 1))
        self.assertTrue(all(row['SP_NAME_DESCRIPTION'] == 'Демонстрационная подсистема' for row in self.rows()[4:]))

    def test_existing_literal_codes_ignore_errors_in_russian_names(self):
        def edit(book):
            sheet = book.worksheets[0]
            for row in range(2, 6):
                sheet[f'H{row}'], sheet[f'G{row}'] = 'DEMO_LOAN', '#N/A'
                sheet[f'P{row}'], sheet[f'O{row}'] = 'ISSUANCE', '#N/A'
        matrix = self.change(self.matrix, edit, 'literal_existing.xlsx')
        template = self.template([('01 — Тестовое ограничение','З',{})])
        result = self.run_import(template, matrix=matrix)
        self.assertEqual(result.affected_rows, [])
        self.assertEqual(len(result.conflicts), 0)

    def test_primary_description_formula_is_replaced_only_in_new_rows(self):
        def edit(book):
            sheet = book.worksheets[0]
            sheet['G2'], sheet['H2'] = '=UNSUPPORTED_FUNCTION(1)', 'DEMO_LOAN'
            sheet['O2'], sheet['P2'] = '=UNSUPPORTED_FUNCTION(1)', 'ISSUANCE'
        matrix = self.change(self.matrix, edit, 'description_formulas.xlsx')
        output = self.assert_new_rows(self.primary_template('SP_NAME'), matrix, 'SP_NAME')
        book = load_workbook(output)
        try:
            self.assertEqual(book.worksheets[0]['G2'].data_type, 'f')
            self.assertEqual(book.worksheets[0]['O2'].data_type, 'f')
            self.assertEqual(book.worksheets[0]['G6'].data_type, 's')
            self.assertEqual(book.worksheets[0]['O6'].data_type, 's')
        finally:
            book.close()

    def test_code_alias_with_multiple_labels_keeps_its_target(self):
        book = load_workbook(self.matrix)
        try:
            book['Подсистемы'].append(['DEMO_LOAN','Другое название'])
            config = load_config()
            config['aliases']['SP_NAME'] = {'Псевдоним':'DEMO_LOAN'}
            catalog = Catalog(book, config)
            self.assertEqual(catalog.resolve('SP_NAME','Псевдоним').description, 'Демонстрационная подсистема')
        finally:
            book.close()
