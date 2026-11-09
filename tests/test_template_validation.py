"""Проверка шапки операции и приоритет платформы только в шаблоне."""
import unittest

import test_business_rules as business
from app.importer import Catalog, ImportProblem, columns_at, import_files, load_config, read_matrix_row
from openpyxl import load_workbook


class TemplateValidationTests(unittest.TestCase):
    setUp = business.BusinessRulesTests.setUp
    prepare_matrix = business.BusinessRulesTests.prepare_matrix
    change = business.BusinessRulesTests.change
    template = business.BusinessRulesTests.template
    run_import = business.BusinessRulesTests.run_import
    rows = business.BusinessRulesTests.rows
    empty_group = business.BusinessRulesTests.empty_group

    def header_template(self, b6, d6, **extra):
        source = self.template([('01 — Тестовое ограничение', 'З', self.empty_group())])
        def edit(book):
            sheet = book.worksheets[0]
            sheet['B6'], sheet['D6'] = b6, d6
            for address, value in extra.items():
                sheet[address] = value
        return self.change(source, edit, 'header.xlsx')

    def test_different_operations_stop_import_even_with_valid_english_code(self):
        template = self.header_template('Тестовая выдача', 'Проверка баланса')
        original = self.matrix.read_bytes()
        with self.assertRaisesRegex(ImportProblem, 'header.xlsx.*B6.*D6.*Тестовая выдача.*Проверка баланса'):
            self.run_import(template)
        self.assertEqual(self.matrix.read_bytes(), original)
        self.assertFalse(self.output.exists())

    def test_invalid_second_template_prevents_any_output(self):
        good = self.template([('01 — Тестовое ограничение', 'З', self.empty_group())], 'good.xlsx')
        bad = self.header_template('Операция 1', 'Операция 2')
        with self.assertRaisesRegex(ImportProblem, 'разные операции'):
            import_files(self.matrix, [good, bad], self.output)
        self.assertFalse(self.output.exists())

    def test_equal_operations_and_simple_links_are_accepted(self):
        for index, (b6, d6) in enumerate((('Тестовая выдача', 'Тестовая выдача'),
                                         ('  Тестовая\xa0выдача ', 'тестовая выдача'),
                                         ('Тестовая выдача', '=$B$6'),
                                         ('=D6', 'Тестовая выдача'),
                                         (None, 'Тестовая выдача'),
                                         ('Тестовая выдача', None))):
            with self.subTest(b6=b6, d6=d6):
                result = self.run_import(self.header_template(b6, d6), output=self.folder/f'equal_{index}.xlsx')
                self.assertEqual((result.added, len(result.conflicts)), (3, 0))

    def test_link_to_other_operation_is_rejected(self):
        with self.assertRaisesRegex(ImportProblem, 'разные операции'):
            self.run_import(self.header_template('Тестовая выдача', '=Z1', Z1='Проверка баланса'))

    def test_b4_d4_are_not_compared(self):
        result = self.run_import(self.header_template('Тестовая выдача', '=B6', B4='Другая подсистема'))
        self.assertEqual((result.added, len(result.conflicts)), (3, 0))

    def test_unreadable_russian_formula_still_uses_english_code(self):
        result = self.run_import(self.header_template('Тестовая выдача', '=UNSUPPORTED_LOOKUP(1)'))
        self.assertEqual((result.added, len(result.conflicts)), (3, 0))

    def platform_template(self, channel='Офис', platform='Web — тестовая платформа', resolution='З', **extra):
        template = self.template([('01 — Тестовое ограничение', resolution,
                                   self.empty_group(A=channel, E=platform))])
        def edit(book):
            for address, value in extra.items():
                book.worksheets[0][address] = value
        return self.change(template, edit, 'platform.xlsx')

    def resolved(self, path=None):
        book = load_workbook(path or self.output)
        try:
            sheet, config = book.worksheets[0], load_config()
            catalog = Catalog(book, config)
            cols = columns_at(sheet, 1)
            return [read_matrix_row(sheet, cols, row, catalog, config).values for row in range(2, sheet.max_row+1)]
        finally:
            book.close()

    def test_platform_overrides_other_unknown_blank_and_formula_channels(self):
        for index, channel in enumerate(('Офис', 'UNKNOWN_CHANNEL', None, '#N/A', '=BAD_FUNCTION(1)')):
            with self.subTest(channel=channel):
                output = self.folder/f'channel_{index}.xlsx'
                template = self.platform_template(channel=channel, D3='=BAD_HEADER(1)', B3='#N/A',
                                                   R16='CHANNEL_INDICATOR', R18='=BAD_CODE(1)')
                result = self.run_import(template, output=output)
                self.assertEqual((result.added, len(result.conflicts)), (3, 0))
                for row in self.resolved(output)[4:]:
                    self.assertEqual(row['CHANNEL_INDICATOR'], 'SBOL_PRO')
                    self.assertEqual(row['CHANNEL_INDICATOR_DESCRIPTION'], 'СБОЛ про')
                    self.assertEqual(row['SBOL_PRO_PLATFORM'], 'WEB')

    def test_platform_code_alone_overrides_channel(self):
        result = self.run_import(self.platform_template(platform=None, R16='SBOL_PRO_PLATFORM', R18='WEB'))
        self.assertEqual((result.added, len(result.conflicts)), (3, 0))
        self.assertTrue(all(row['CHANNEL_INDICATOR'] == 'SBOL_PRO' for row in self.resolved()[4:]))

    def test_permitted_seed_gets_sbol_channel_for_both_mandatory_rows(self):
        result = self.run_import(self.platform_template(resolution='Р'))
        self.assertEqual((result.added, result.skipped), (2, 1))
        self.assertTrue(all(row['CHANNEL_INDICATOR'] == 'SBOL_PRO' for row in self.resolved()[4:]))

    def test_repeat_and_prom_use_same_normalized_group(self):
        template = self.platform_template()
        self.run_import(template)
        repeat = self.run_import(template, matrix=self.output, output=self.folder/'repeat.xlsx')
        self.assertEqual(repeat.affected_rows, [])
        prom = self.run_import(template, matrix=self.output, environment='PROM', output=self.folder/'prom.xlsx')
        self.assertEqual((prom.added, prom.promoted, len(prom.conflicts)), (0, 3, 0))

    def test_without_platform_office_channel_is_kept(self):
        result = self.run_import(self.platform_template(platform=None))
        self.assertEqual((result.added, len(result.conflicts)), (3, 0))
        self.assertTrue(all(row['CHANNEL_INDICATOR'] == 'OFFICE' for row in self.resolved()[4:]))

    def test_unknown_platform_is_not_silently_accepted(self):
        with self.assertRaisesRegex(ImportProblem, 'SBOL_PRO_PLATFORM.*UNKNOWN_PLATFORM'):
            self.run_import(self.platform_template(platform='UNKNOWN_PLATFORM'))
        self.assertFalse(self.output.exists())

    def test_existing_office_matrix_is_not_rewritten_to_sbol(self):
        def edit(book):
            for row in range(2, 6):
                book.worksheets[0][f'E{row}'] = 'Офис'
        matrix = self.change(self.matrix, edit, 'office_matrix.xlsx')
        before = self.rows(matrix)
        result = self.run_import(self.platform_template(), matrix=matrix)
        self.assertEqual((result.added, len(result.conflicts)), (3, 0))
        self.assertEqual(self.rows()[:4], before)

    def test_existing_sbol_record_is_updated_using_overridden_channel(self):
        template = self.template([('01 — Тестовое ограничение', 'ТЗ', {'A':'Офис'})])
        original = self.matrix.read_bytes()
        matrix = self.change(self.matrix, lambda b: setattr(b.worksheets[0]['D4'], 'value', None), 'ift_only.xlsx')
        result = self.run_import(template, matrix=matrix)
        self.assertEqual((result.added, result.updated, len(result.conflicts)), (0, 1, 0))
        self.assertEqual(self.matrix.read_bytes(), original)
        self.assertEqual(self.resolved()[2]['CHANNEL_INDICATOR'], 'SBOL_PRO')
