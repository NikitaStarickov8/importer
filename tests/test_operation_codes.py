"""Операция определяется по английскому коду, не по опечатке в названии."""
import unittest
import test_business_rules as business
from app.importer import ImportProblem


class OperationCodeTests(unittest.TestCase):
    setUp = business.BusinessRulesTests.setUp
    prepare_matrix = business.BusinessRulesTests.prepare_matrix
    change = business.BusinessRulesTests.change
    template = business.BusinessRulesTests.template
    run_import = business.BusinessRulesTests.run_import
    rows = business.BusinessRulesTests.rows
    empty_group = business.BusinessRulesTests.empty_group

    def operation_template(self, description='Опечатка в наименовании', header_code='CHECK_BALANCE', edit=lambda s: None):
        template = self.template([('01 — Тестовое ограничение', 'З', {'C':description, 'E':'Web — тестовая платформа'})])
        def modify(book):
            sheet = book.worksheets[0]
            sheet['D7'], sheet['B7'] = None, header_code
            edit(sheet)
        return self.change(template, modify, 'operation.xlsx')

    def assert_balance_group(self, template):
        result = self.run_import(template)
        self.assertEqual((result.added, len(result.conflicts)), (3, 0))
        for row in self.rows()[4:]:
            self.assertEqual(row['OPERATION_DESCRIPTION'], 'Проверка баланса')
        repeat = self.run_import(template, matrix=self.output, output=self.folder/'repeat.xlsx')
        self.assertEqual(repeat.affected_rows, [])

    def test_russian_typo_uses_code_from_b7_and_is_idempotent(self):
        self.assert_balance_group(self.operation_template())

    def test_valid_but_different_russian_name_does_not_override_english_code(self):
        self.assert_balance_group(self.operation_template(description='Тестовая выдача'))

    def test_russian_formula_error_uses_header_code(self):
        self.assert_balance_group(self.operation_template(description='#N/A'))

    def test_unsupported_russian_lookup_formula_uses_header_code(self):
        self.assert_balance_group(self.operation_template(description='=INDEX(A1:A2,MATCH("опечатка",B1:B2,0))'))

    def test_header_simple_link_is_resolved(self):
        self.assert_balance_group(self.operation_template(edit=lambda s: setattr(s['D7'], 'value', '=$B$7')))

    def test_failed_header_code_formula_falls_back_to_manual_b7(self):
        self.assert_balance_group(self.operation_template(edit=lambda s: setattr(s['D7'], 'value', '=UNSUPPORTED_LOOKUP(B6)')))

    def test_header_error_falls_back_to_manual_b7(self):
        self.assert_balance_group(self.operation_template(edit=lambda s: setattr(s['D7'], 'value', '#N/A')))

    def test_explicit_row_code_wins_over_header_and_description(self):
        def edit(sheet):
            sheet['R16'], sheet['R18'] = 'OPERATION_TYPE', 'CHECK_BALANCE'
            sheet['D7'] = '=UNSUPPORTED_FUNCTION(123)'
        self.assert_balance_group(self.operation_template(description='#N/A', header_code='UNKNOWN_CODE', edit=edit))

    def test_explicit_row_code_simple_link_is_supported(self):
        def edit(sheet):
            sheet['R16'], sheet['R18'] = 'OPERATION_TYPE', '=$B$7'
        self.assert_balance_group(self.operation_template(edit=edit))

    def test_unreadable_row_code_can_use_manual_header_code(self):
        def edit(sheet):
            sheet['R16'], sheet['R18'] = 'OPERATION_TYPE', '=UNSUPPORTED_LOOKUP(C18)'
        self.assert_balance_group(self.operation_template(description='#N/A', edit=edit))

    def test_code_in_legacy_description_column_wins_over_header(self):
        self.assert_balance_group(self.operation_template(description='CHECK_BALANCE', header_code='UNKNOWN_CODE'))

    def test_unknown_code_in_legacy_column_does_not_select_header_operation(self):
        with self.assertRaisesRegex(ImportProblem, 'UNKNOWN_OPERATION'):
            self.run_import(self.operation_template(description='UNKNOWN_OPERATION'))
        self.assertFalse(self.output.exists())

    def test_unknown_explicit_english_code_does_not_fall_back_to_valid_russian(self):
        def edit(sheet):
            sheet['R16'], sheet['R18'] = 'OPERATION_TYPE', 'UNKNOWN_CODE'
        with self.assertRaisesRegex(ImportProblem, 'UNKNOWN_CODE'):
            self.run_import(self.operation_template(description='Проверка баланса', edit=edit))
        self.assertFalse(self.output.exists())

    def test_unknown_header_code_does_not_fall_back_to_valid_russian(self):
        with self.assertRaisesRegex(ImportProblem, 'UNKNOWN_CODE'):
            self.run_import(self.operation_template(description='Проверка баланса', header_code='UNKNOWN_CODE'))

    def test_no_english_code_still_allows_exact_russian_lookup(self):
        self.assert_balance_group(self.operation_template(description='Проверка баланса', header_code=None))

    def test_no_code_and_unknown_russian_name_is_still_an_error(self):
        with self.assertRaisesRegex(ImportProblem, 'Опечатка'):
            self.run_import(self.operation_template(header_code=None))
        self.assertFalse(self.output.exists())

    def test_broken_row_without_code_is_not_replaced_by_header_russian_name(self):
        with self.assertRaisesRegex(ImportProblem, 'C18.*ошибка'):
            self.run_import(self.operation_template(description='#N/A', header_code=None))

    def test_known_code_ignores_bad_russian_header_formula(self):
        def edit(sheet):
            sheet['D6'], sheet['B6'] = '=UNSUPPORTED_FUNCTION(1)', '#N/A'
        self.assert_balance_group(self.operation_template(description=None, edit=edit))

    def test_skipped_permitted_group_also_uses_english_code(self):
        template = self.operation_template(edit=lambda s: setattr(s['L18'], 'value', 'Р'))
        result = self.run_import(template)
        self.assertEqual((result.added, result.skipped), (2, 1))
        self.assertTrue(all(row['OPERATION_DESCRIPTION'] == 'Проверка баланса' for row in self.rows()[4:]))

    def test_other_lookup_code_description_mismatch_remains_an_error(self):
        def edit(sheet):
            sheet['S16'], sheet['S18'] = 'CHANNEL_INDICATOR', 'SBOL_PRO'
            sheet['A18'] = 'Офис'
        with self.assertRaisesRegex(ImportProblem, 'CHANNEL_INDICATOR.*не соответствует'):
            self.run_import(self.operation_template(edit=edit))

    def test_shared_russian_description_cannot_override_known_english_code(self):
        matrix = self.change(self.matrix, lambda b: b['Типы Операций'].append(['ANOTHER_CODE','Проверка баланса']), 'ambiguous.xlsx')
        result = self.run_import(self.operation_template(), matrix=matrix)
        self.assertEqual(result.added, 3)
        self.assertEqual(len(result.conflicts), 0)
        self.assertTrue(all(row['OPERATION_TYPE'] == 'CHECK_BALANCE' for row in self.rows()[4:]))
