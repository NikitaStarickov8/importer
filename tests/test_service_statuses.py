"""NORMAL и TECH_ERROR работают без записей в справочниках стоп-листа."""
import unittest
from unittest.mock import patch

import test_business_rules as business
from app.importer import Catalog, ImportProblem, import_files, load_config, norm, service_entry
from openpyxl import load_workbook


class ServiceStatusTests(unittest.TestCase):
    setUp = business.BusinessRulesTests.setUp
    prepare_matrix = business.BusinessRulesTests.prepare_matrix
    change = business.BusinessRulesTests.change
    template = business.BusinessRulesTests.template
    run_import = business.BusinessRulesTests.run_import
    rows = business.BusinessRulesTests.rows
    empty_group = business.BusinessRulesTests.empty_group

    def without_services(self, book):
        for name in ('Статусы Стоп Листа','COMPLIANCE_STOP_LIST_CODES'):
            sheet = book[name]
            sheet.delete_rows(2, 2)

    def service_matrix(self):
        return self.change(self.matrix, self.without_services, 'no_services.xlsx')

    def test_auto_service_rows_do_not_require_catalog_entries(self):
        matrix = self.service_matrix()
        template = self.template([('01 — Тестовое ограничение','З',self.empty_group())])
        result = self.run_import(template, matrix=matrix)
        self.assertEqual((result.added, len(result.conflicts)), (3, 0))
        rows = self.rows()[4:]
        services = [r for r in rows if r['STOP_LIST_STATUS'] in ('NORMAL','TECH_ERROR')]
        self.assertEqual(len(services), 2)
        for row in services:
            self.assertEqual(row['BO_STEP_CODE_DESCRIPTION'], row['STOP_LIST_STATUS'])
            self.assertIsNone(row['BO_STEP_CODE'])
            self.assertEqual(row['STOP_LIST_STATUS_PRIORITY'], 0)
            self.assertEqual(row['STATUS_CODE'], row['STOP_LIST_STATUS_DESCRIPTION'])
            self.assertEqual(row['INFORMATION_SOURCE'], 'SL')
        book = load_workbook(self.output)
        try:
            for row in (7, 8):
                for column in ('K','L','Z','AA','AB','AO','AP'):
                    self.assertNotEqual(book.worksheets[0][f'{column}{row}'].data_type, 'f')
                for column in ('F','H','P','AI','AF'):
                    self.assertEqual(book.worksheets[0][f'{column}{row}'].data_type, 'f')
        finally:
            book.close()

    def test_repeat_and_prom_work_without_service_entries(self):
        template = self.template([('01 — Тестовое ограничение','З',self.empty_group())])
        self.run_import(template, matrix=self.service_matrix())
        repeat = self.run_import(template, matrix=self.output, output=self.folder/'repeat.xlsx')
        self.assertEqual(repeat.affected_rows, [])
        prom = self.run_import(template, matrix=self.output, environment='PROM', output=self.folder/'prom.xlsx')
        self.assertEqual((prom.added, prom.promoted, len(prom.conflicts)), (0, 3, 0))

    def test_service_rows_in_template_are_imported_without_catalog_entries(self):
        template = self.template([('В стоп-листе не найден','Р',self.empty_group()),
                                  ('Технический запрет','ТЗ',self.empty_group())])
        result = self.run_import(template, matrix=self.service_matrix())
        self.assertEqual((result.added, len(result.conflicts)), (2, 0))
        self.assertEqual({r['STOP_LIST_STATUS'] for r in self.rows()[4:]}, {'NORMAL','TECH_ERROR'})

    def test_service_code_input_works_without_russian_lookup(self):
        template = self.template([('NORMAL','Р',self.empty_group()), ('TECH_ERROR','ТЗ',self.empty_group())])
        result = self.run_import(template, matrix=self.service_matrix())
        self.assertEqual((result.added, len(result.conflicts)), (2, 0))

    def test_russian_service_spelling_with_or_without_hyphen_is_recognized(self):
        for label in ('В Стоп-листе не найден', 'В стоп листе не найден', 'В стоплисте не найден', ' NORMAL '):
            with self.subTest(label=label):
                self.assertEqual(service_entry(label).code, 'NORMAL')

    def test_service_lookup_never_reads_status_reference(self):
        book = load_workbook(self.matrix)
        try:
            catalog = Catalog(book, load_config())
            with patch.object(catalog, '_build_index', side_effect=AssertionError('Reference must not be read')):
                for value in ('NORMAL', 'В стоп-листе не найден', 'TECH_ERROR', 'Технический запрет'):
                    entry = catalog.resolve('STOP_LIST_STATUS', value)
                    self.assertTrue(catalog.accepts('STOP_LIST_STATUS', value, entry))
                    self.assertEqual(catalog.possible_codes('STOP_LIST_STATUS', value), {norm(entry.code)})
                    self.assertEqual(catalog.step_code(entry), norm(entry.code))
        finally:
            book.close()

    def test_service_import_works_even_if_status_reference_sheet_is_absent(self):
        matrix = self.change(self.matrix, lambda b: b.remove(b['Статусы Стоп Листа']), 'missing_sheet.xlsx')
        template = self.template([('NORMAL','Р',self.empty_group())])
        result = self.run_import(template, matrix=matrix)
        self.assertEqual((result.added, len(result.conflicts)), (2, 0))

    def test_service_reference_is_not_rebuilt_or_modified(self):
        matrix = self.service_matrix()
        before = load_workbook(matrix)
        expected = {s.title:list(s.values) for s in before.worksheets[1:]}
        before.close()
        self.run_import(self.template([('NORMAL','Р',self.empty_group())]), matrix=matrix)
        after = load_workbook(self.output)
        try:
            self.assertEqual({s.title:list(s.values) for s in after.worksheets[1:]}, expected)
        finally:
            after.close()

    def test_unknown_regular_status_is_still_rejected(self):
        template = self.template([('UNKNOWN_STATUS','З',self.empty_group())])
        with self.assertRaisesRegex(ImportProblem, 'UNKNOWN_STATUS'):
            self.run_import(template, matrix=self.service_matrix())
        self.assertFalse(self.output.exists())

    def test_existing_service_rows_are_recognized_after_reference_entries_removed(self):
        template = self.template([('NORMAL','Р',{}), ('TECH_ERROR','ТЗ',{})])
        before = self.rows(self.matrix)
        result = self.run_import(template, matrix=self.service_matrix())
        self.assertEqual(result.affected_rows, [])
        self.assertEqual(len(result.conflicts), 0)
        self.assertEqual(self.rows(), before)

    def test_ordinary_permitted_does_not_need_services_in_reference(self):
        template = self.template([('01 — Тестовое ограничение','Р',self.empty_group())])
        result = self.run_import(template, matrix=self.service_matrix())
        self.assertEqual((result.added,result.skipped), (2, 1))

    def test_prom_without_ift_still_does_not_create_service_rows(self):
        result = self.run_import(self.template([('NORMAL','Р',self.empty_group())]), matrix=self.service_matrix(), environment='PROM')
        self.assertEqual(result.added, 0)
        self.assertEqual(len(result.conflicts), 2)

    def test_conflicting_service_reference_rows_do_not_override_builtins(self):
        def edit(book):
            sheet = book['Статусы Стоп Листа']
            sheet.append(['WRONG','В стоп-листе не найден','WRONG_STEP'])
            sheet.append(['TECH_ERROR','Постороннее название',None])
        matrix = self.change(self.matrix, edit, 'conflicting_service_catalog.xlsx')
        result = self.run_import(self.template([('NORMAL','Р',self.empty_group())]), matrix=matrix)
        self.assertEqual((result.added,len(result.conflicts)), (2,0))

    def test_empty_service_template_messages_and_bypass_keep_previous_rules(self):
        template = self.template([('TECH_ERROR','ТЗ',self.empty_group(O='Мой заголовок',P='Мой текст'))])
        result = self.run_import(template, matrix=self.service_matrix())
        self.assertEqual(result.added, 2)
        tech = next(r for r in self.rows()[4:] if r['STOP_LIST_STATUS'] == 'TECH_ERROR')
        self.assertEqual(tech['TITLE_FOR_CLIENT'], 'Мой заголовок')
        self.assertEqual(tech['TEXT_FOR_CLIENT'], 'Мой текст')
        self.assertIsNone(tech['VALUE_SWITCH'])
