"""Политика по умолчанию: обязательные строки, стопкоды, пропуски, ИФТ/ПРОМ."""
import hashlib
import tempfile
import unittest
from pathlib import Path

import test_workflow as workflow
from app.importer import import_files, load_config
from openpyxl import load_workbook


class BusinessRulesTests(unittest.TestCase):
    change = workflow.WorkflowTests.change

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.folder = Path(self.tmp.name)
        root = Path(__file__).resolve().parents[1]
        self.matrix = self.change(root / 'examples/01_matrix.xlsx', self.prepare_matrix, 'matrix.xlsx')
        self.base_template = root / 'examples/02_template_update.xlsx'
        self.output = self.folder / 'result.xlsx'

    def prepare_matrix(self, book):
        statuses = book['Статусы Стоп Листа']
        for row, code in enumerate(('TECH_ERROR', 'NORMAL', 'CODE_309', 'CODE_405'), 2):
            statuses.cell(row, 3, code)

    def template(self, rows, name='template.xlsx'):
        # Каждая строка: статус, резолюция, необязательные изменения по буквам.
        def edit(book):
            sheet = book.worksheets[0]
            base = [cell.value for cell in sheet[18]]
            for row in sheet.iter_rows(min_row=18):
                for cell in row:
                    cell.value = None
            for row, (status, resolution, extra) in enumerate(rows, 18):
                for col, value in enumerate(base, 1):
                    sheet.cell(row, col, value)
                sheet[f'J{row}'], sheet[f'L{row}'] = status, resolution
                sheet[f'O{row}'], sheet[f'P{row}'] = None, None
                for col, value in extra.items():
                    sheet[f'{col}{row}'] = value
        return self.change(self.base_template, edit, name)

    def run_import(self, template, matrix=None, environment='IFT', output=None):
        return import_files(matrix or self.matrix, [template], output or self.output, environment=environment)

    def rows(self, path=None):
        book = load_workbook(path or self.output)
        try:
            sheet = book.worksheets[0]
            return [dict(zip([c.value for c in sheet[1]], [c.value for c in row])) for row in sheet.iter_rows(min_row=2)]
        finally:
            book.close()

    def empty_group(self, **kwargs):
        return {'C': 'CHECK_BALANCE', 'E': 'Web — тестовая платформа', **kwargs}

    def test_config_migrates_old_keys_without_replacing_aliases(self):
        config = load_config()
        self.assertIn('INSPECTED_PERSON', config['group_key'])
        self.assertIn('MATCH_NEGATIVE_CASE', config['group_key'])
        self.assertNotIn('TRANSACTION_TYPE', config['group_key'] + config['row_key'])
        self.assertEqual(config['aliases']['STOP_LIST_STATUS']['Флаг 66'], 'MARKER_66')

    def test_permitted_is_skipped_and_missing_mandatory_rows_created(self):
        template = self.template([('02 — Тестовая проверка', 'Р', self.empty_group())])
        result = self.run_import(template)
        self.assertEqual((result.added, result.updated, result.skipped), (2, 0, 1))
        new = self.rows()[4:]
        self.assertEqual({r['STOP_LIST_STATUS_DESCRIPTION'] for r in new}, {'Технический запрет', 'В стоп-листе не найден'})
        self.assertTrue(all(r['IFT'] == 1 and r['PROM'] is None and r['VALUE_SWITCH'] is None and r['MAX_AMOUNT'] is None for r in new))
        technical = next(r for r in new if r['VALUE'] == 'ТЗ')
        self.assertEqual(technical['TITLE_FOR_CLIENT'], 'Сервис временно недоступен')
        self.assertEqual(technical['TEXT_FOR_CLIENT'], 'Пожалуйста, попробуйте зайти через 5 минут или позже.')
        normal = next(r for r in new if r['VALUE'] == 'Р')
        self.assertIsNone(normal['TEXT_FOR_CLIENT'])
        self.assertIsNone(normal['TITLE_FOR_CLIENT'])
        self.assertEqual(len(result.outcomes), result.total)

    def test_partial_skip_precedes_bad_fields_and_does_not_create_group(self):
        template = self.template([('НЕИЗВЕСТНО', 'З', self.empty_group(K='Частичное совпадение', N='не число', P='=1/0'))])
        result = self.run_import(template)
        self.assertEqual((result.added, result.updated, result.skipped), (0, 0, 1))
        self.assertIn('больше не используется', result.outcomes[0].reason)
        self.assertEqual(len(self.rows()), 4)

    def test_skipped_permitted_ignores_unknown_status_amount_and_text(self):
        template = self.template([('НЕИЗВЕСТНО', 'Р', self.empty_group(N='не число', P='=1/0'))])
        result = self.run_import(template)
        self.assertEqual(result.added, 2)
        self.assertEqual(result.skipped, 1)

    def test_no_mass_cleanup_of_existing_allowed_or_partial_rows(self):
        template = self.template([('02 — Тестовая проверка', 'Р', {})])
        before = self.rows(self.matrix)
        result = self.run_import(template)
        self.assertEqual(result.affected_rows, [])
        self.assertEqual(before, self.rows())

    def test_existing_mandatory_ift_preserves_texts_and_does_not_version(self):
        template = self.template([('Технический запрет', 'ТЗ', {'O':'Новый заголовок', 'P':'Новый текст'}), ('В стоп-листе не найден', 'Р', {})])
        before = self.rows(self.matrix)
        result = self.run_import(template)
        self.assertEqual((result.added, result.updated, result.skipped), (0, 0, 2))
        self.assertEqual(before, self.rows())

    def test_regular_forbidden_is_added_with_two_mandatory_records(self):
        template = self.template([('01 — Тестовое ограничение', 'З', self.empty_group())])
        result = self.run_import(template)
        self.assertEqual(result.added, 3)
        self.assertEqual(len(result.conflicts), 0)
        self.assertEqual(len(result.outcomes), result.total)

    def test_repeat_does_not_duplicate_any_record(self):
        template = self.template([('01 — Тестовое ограничение', 'З', self.empty_group())])
        self.run_import(template)
        result = self.run_import(template, matrix=self.output, output=self.folder/'repeat.xlsx')
        self.assertEqual((result.added, result.updated), (0, 0))
        self.assertEqual(result.affected_rows, [])
        self.assertEqual(len(self.rows(self.folder/'repeat.xlsx')), 7)

    def test_same_stop_code_with_changed_amount_requires_manual_review(self):
        template = self.template([('01 — Тестовое ограничение', 'З', {'N':100})])
        result = self.run_import(template)
        self.assertEqual(result.affected_rows, [])
        self.assertTrue(any('MAX_AMOUNT' in c for c in result.conflicts))

    def test_same_stop_code_with_changed_transaction_requires_manual_review(self):
        template = self.template([('01 — Тестовое ограничение', 'З', {'G':'Финансовые операции'})])
        result = self.run_import(template)
        self.assertEqual(result.affected_rows, [])
        self.assertEqual(len(result.conflicts), 3)
        self.assertTrue(all('TRANSACTION_TYPE' in c for c in result.conflicts))

    def test_conflicting_template_amounts_do_not_create_first_variant(self):
        template = self.template([('01 — Тестовое ограничение', 'З', self.empty_group(N=100)),
                                  ('01 — Тестовое ограничение', 'З', self.empty_group(N=200))])
        result = self.run_import(template)
        self.assertEqual(result.added, 2)
        self.assertEqual(len(result.conflicts), 2)
        self.assertTrue(all('конфликт шаблонов' in c for c in result.conflicts))

    def test_two_statuses_with_same_bo_step_code_are_not_two_entries(self):
        def edit(book):
            book['Статусы Стоп Листа']['C5'] = 'CODE_309'
        matrix = self.change(self.matrix, edit, 'same_step.xlsx')
        template = self.template([('02 — Тестовая проверка', 'З', {})])
        result = self.run_import(template, matrix=matrix)
        self.assertEqual(result.added, 0)
        self.assertTrue(any('разные статусы' in c for c in result.conflicts))

    def test_prom_only_row_is_not_a_substitute_for_mandatory_ift(self):
        matrix = self.change(self.matrix, lambda b: setattr(b.worksheets[0]['C2'], 'value', None), 'prom_only.xlsx')
        template = self.template([('В стоп-листе не найден', 'Р', {})])
        result = self.run_import(template, matrix=matrix)
        self.assertEqual(result.added, 1)
        self.assertEqual(self.rows()[0]['PROM'], 1)
        self.assertEqual(self.rows()[-1]['IFT'], 1)

    def test_prom_never_creates_missing_mandatory_rows(self):
        template = self.template([('01 — Тестовое ограничение', 'З', self.empty_group())])
        result = self.run_import(template, environment='PROM')
        self.assertEqual((result.added, result.updated), (0, 0))
        self.assertEqual(len(result.conflicts), 3)
        self.assertEqual(len(self.rows()), 4)

    def test_prom_promotes_existing_mandatory_preserving_custom_text(self):
        matrix = self.change(self.matrix, lambda b: setattr(b.worksheets[0]['D2'], 'value', None), 'ift_only.xlsx')
        template = self.template([('В стоп-листе не найден', 'Р', {})])
        result = self.run_import(template, matrix=matrix, environment='PROM')
        self.assertEqual(result.promoted, 2)
        self.assertEqual(result.added, 0)
        self.assertEqual(self.rows()[0]['TEXT_FOR_CLIENT'], self.rows(matrix)[0]['TEXT_FOR_CLIENT'])
        self.assertEqual(self.rows()[0]['PROM'], 1)

    def test_existing_mandatory_with_wrong_resolution_is_manual(self):
        matrix = self.change(self.matrix, lambda b: setattr(b.worksheets[0]['AE2'], 'value', 'З'), 'wrong.xlsx')
        result = self.run_import(self.template([('Технический запрет','ТЗ',{})]), matrix=matrix)
        self.assertEqual(result.affected_rows, [])
        self.assertTrue(any('резолюция' in c for c in result.conflicts))

    def test_group_is_separate_for_each_inspected_person(self):
        def edit(book):
            book['Проверяемые лица'].append(['REP', 10, 'Представитель'])
        matrix = self.change(self.matrix, edit, 'representative.xlsx')
        result = self.run_import(self.template([('01 — Тестовое ограничение','З',{'I':'Представитель'})]), matrix=matrix)
        self.assertEqual(result.added, 3)
        self.assertTrue(all(r['INSPECTED_PERSON_DESCRIPTION'] == 'Представитель' for r in self.rows()[4:]))

    def test_group_platform_is_separate_for_sbol_pro(self):
        result = self.run_import(self.template([('01 — Тестовое ограничение','З',{'E':'Web — тестовая платформа'})]))
        self.assertEqual(result.added, 3)

    def test_report_has_three_sections_and_no_code_noise(self):
        template = self.template([('01 — Тестовое ограничение','З',self.empty_group()), ('02 — Тестовая проверка','Р',self.empty_group())])
        result = self.run_import(template)
        report = result.change_report
        self.assertLess(report.index('ПРОПУЩЕНО ('), report.index('ДОБАВЛЕНО ('))
        self.assertLess(report.index('ДОБАВЛЕНО ('), report.index('ИЗМЕНЕНО ('))
        self.assertNotIn('CODE', report)
        self.assertNotIn('пересчёт', report)
        self.assertIn('Автоматически:', report)
        self.assertIn('Источник:', report)
        self.assertIn('Подсистема:', report)
        self.assertIn('Строка Excel 6', report)

    def test_formulas_styles_catalogs_and_inputs_are_preserved(self):
        template = self.template([('01 — Тестовое ограничение','З',self.empty_group())])
        hashes = {p: hashlib.sha256(p.read_bytes()).hexdigest() for p in (self.matrix, template)}
        self.run_import(template)
        before, after = load_workbook(self.matrix), load_workbook(self.output)
        try:
            for a, b in zip(before.worksheets[1:], after.worksheets[1:]):
                self.assertEqual(list(a.values), list(b.values))
            a, b = before.worksheets[0], after.worksheets[0]
            for row in range(1, 6):
                self.assertEqual([c.value for c in a[row]], [c.value for c in b[row]])
            self.assertEqual(b['H6'].data_type, 'f')
            self.assertEqual(b['H6']._style, a['H2']._style)
            self.assertEqual(b.tables['MatrixTable'].ref, 'A1:AQ8')
        finally:
            before.close()
            after.close()
        for p, digest in hashes.items():
            self.assertEqual(hashlib.sha256(p.read_bytes()).hexdigest(), digest)

    def test_regular_ift_only_payload_updates_in_place(self):
        matrix = self.change(self.matrix, lambda b: setattr(b.worksheets[0]['D4'], 'value', None), 'ift.xlsx')
        template = self.template([('01 — Тестовое ограничение','ТЗ',{'O':'Заголовок','P':'Текст'})])
        before = self.rows(matrix)
        result = self.run_import(template, matrix=matrix)
        self.assertEqual((result.added, result.updated), (0, 1))
        self.assertEqual(result.affected_rows, [4])
        after = self.rows()
        expected = dict(before[2], VALUE='ТЗ', TITLE_FOR_CLIENT='Заголовок', TEXT_FOR_CLIENT='Текст')
        self.assertEqual(after[2], expected)

    def test_regular_shared_record_versions_then_promotes(self):
        template = self.template([('01 — Тестовое ограничение','ТЗ',{'P':'Новая версия'})])
        first = self.run_import(template)
        self.assertEqual((first.added, first.versioned), (1, 1))
        rows = self.rows()
        self.assertIsNone(rows[2]['IFT'])
        self.assertEqual(rows[2]['PROM'], 1)
        self.assertEqual(rows[4]['IFT'], 1)
        self.assertIsNone(rows[4]['PROM'])
        prom_path = self.folder/'prom.xlsx'
        result = self.run_import(template, matrix=self.output, environment='PROM', output=prom_path)
        self.assertEqual(result.added, 0)
        self.assertEqual(result.prom_deactivated, 1)
        rows = self.rows(prom_path)
        self.assertIsNone(rows[2]['PROM'])
        self.assertEqual(rows[4]['PROM'], 1)
        self.assertEqual(rows[4]['TEXT_FOR_CLIENT'], 'Новая версия')

    def test_unreadable_active_matching_row_is_not_treated_as_absent(self):
        matrix = self.change(self.matrix, lambda b: setattr(b.worksheets[0]['M4'], 'value', 'НЕИЗВЕСТНО'), 'uncertain.xlsx')
        result = self.run_import(self.template([('01 — Тестовое ограничение','З',{})]), matrix=matrix)
        self.assertEqual(result.added, 0)
        self.assertEqual(result.affected_rows, [])
        self.assertTrue(any('неразобранными' in c for c in result.conflicts))

    def test_duplicate_mandatory_ift_does_not_activate_arbitrary_prom_row(self):
        def edit(book):
            sheet = book.worksheets[0]
            sheet.append([c.value for c in sheet[2]])
            sheet['D6'] = None
        matrix = self.change(self.matrix, edit, 'duplicates.xlsx')
        result = self.run_import(self.template([('Технический запрет','ТЗ',{})]), matrix=matrix, environment='PROM')
        self.assertTrue(any('обязательных ИФТ' in c for c in result.conflicts))
        self.assertNotIn(2, result.affected_rows)
        self.assertNotIn(6, result.affected_rows)

    def test_generated_mandatory_does_not_choose_between_transaction_types(self):
        template = self.template([('01 — Тестовое ограничение','З',self.empty_group()),
                                  ('02 — Тестовая проверка','З',self.empty_group(G='Финансовые операции'))])
        result = self.run_import(template)
        self.assertEqual(result.added, 2)
        self.assertEqual(len(result.conflicts), 2)
        self.assertTrue(all('Обязательная строка не создана' in c for c in result.conflicts))

    def test_two_mandatory_rows_are_not_duplicated_across_templates(self):
        template = self.template([('Технический запрет','ТЗ',self.empty_group()),
                                  ('В стоп-листе не найден','Р',self.empty_group())])
        result = import_files(self.matrix,[template,template],self.output)
        self.assertEqual(result.added, 2)
        self.assertEqual(result.skipped, 2)

    def test_full_and_partial_in_same_template_only_full_generates_mandatory(self):
        template = self.template([('01 — Тестовое ограничение','З',self.empty_group()),
                                  ('02 — Тестовая проверка','З',self.empty_group(K='Частичное совпадение'))])
        result = self.run_import(template)
        self.assertEqual(result.added, 3)
        self.assertEqual(result.skipped, 1)

    def test_filled_platform_overrides_office_and_splits_sbol_group(self):
        template = self.template([('01 — Тестовое ограничение','З',{'A':'Офис','E':'iPad — тестовая платформа'}),
                                  ('02 — Тестовая проверка','З',{'A':'Офис','E':'Web — тестовая платформа'})])
        result = self.run_import(template)
        self.assertEqual(result.added, 3)
        self.assertEqual(self.rows()[-1]['CHANNEL_INDICATOR_DESCRIPTION'], 'СБОЛ про')

    def test_legacy_aliases_are_applied_with_business_rules_enabled(self):
        def edit(book):
            sheet = book['Статусы Стоп Листа']
            sheet.append(['MARKER_66', 'Маркер 66', 'CODE_66'])
        matrix = self.change(self.matrix, edit, 'marker.xlsx')
        result = self.run_import(self.template([('Флаг 66','З',self.empty_group())]), matrix=matrix)
        self.assertEqual(result.added, 3)
        self.assertEqual(self.rows()[4]['STOP_LIST_STATUS_DESCRIPTION'], 'Маркер 66')

    def test_permitted_with_message_is_not_plain_permitted(self):
        def edit(book):
            book['Решения'].append([4,'РС','operationPermittedWithMessage','Разрешено с сообщением',None,20])
        matrix = self.change(self.matrix, edit, 'permitted_message.xlsx')
        result = self.run_import(self.template([('01 — Тестовое ограничение','РС',self.empty_group(P='Сообщение'))]), matrix=matrix)
        self.assertEqual(result.added, 3)
        self.assertEqual(self.rows()[4]['VALUE'], 'РС')
        self.assertEqual(self.rows()[4]['TEXT_FOR_CLIENT'], 'Сообщение')
