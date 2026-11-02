"""Читаем английский код формулы, не меняя существующую книгу."""
import hashlib
import unittest

import test_business_rules as business
from app.code_formulas import CodeFormulaError, CodeFormulaReader
from app.importer import Catalog, columns_at, load_config, read_matrix_row
from openpyxl import load_workbook


class CodeFormulaTests(unittest.TestCase):
    setUp = business.BusinessRulesTests.setUp
    prepare_matrix = business.BusinessRulesTests.prepare_matrix
    change = business.BusinessRulesTests.change
    template = business.BusinessRulesTests.template
    run_import = business.BusinessRulesTests.run_import
    rows = business.BusinessRulesTests.rows
    empty_group = business.BusinessRulesTests.empty_group

    def formula_book(self):
        book = load_workbook(self.matrix)
        self.addCleanup(book.close)
        return book

    def ambiguous_matrix(self, formula=None, extra_edit=lambda b: None):
        def edit(book):
            book['Типы Операций'].append(['SECOND_CODE','Тестовая выдача'])
            sheet = book.worksheets[0]
            for row in range(2, 6):
                sheet[f'P{row}'] = formula or "=INDEX('Типы Операций'!A:A,MATCH('Матрица решений комплаенс'!O:O,'Типы Операций'!B:B,0),1)"
            extra_edit(book)
        return self.change(self.matrix, edit, 'ambiguous_formula.xlsx')

    def test_whole_column_formula_from_screenshot_reads_first_actual_match(self):
        matrix = self.ambiguous_matrix()
        book = load_workbook(matrix)
        try:
            sheet = book.worksheets[0]
            before = {s.title:{p:(c.value,c.data_type) for p,c in s._cells.items()} for s in book}
            config = load_config()
            catalog = Catalog(book, config)
            for row in range(2,6):
                record = read_matrix_row(sheet,columns_at(sheet,1),row,catalog,config)
                self.assertEqual(record.values['OPERATION_TYPE'],'ISSUANCE')
            after = {s.title:{p:(c.value,c.data_type) for p,c in s._cells.items()} for s in book}
            self.assertEqual(before,after)
        finally:
            book.close()

    def test_matching_template_english_code_updates_only_requested_record(self):
        matrix = self.ambiguous_matrix()
        digest = hashlib.sha256(matrix.read_bytes()).digest()
        template = self.template([('01 — Тестовое ограничение','З',{'P':'Новое сообщение'})])
        result = self.run_import(template,matrix=matrix)
        self.assertEqual((result.added,result.versioned,len(result.conflicts)),(1,1,0))
        self.assertEqual(result.affected_rows,[4,6])
        before,after = load_workbook(matrix),load_workbook(self.output)
        try:
            a,b = before.worksheets[0],after.worksheets[0]
            for row in range(1,6):
                expected = [c.value for c in a[row]]
                if row == 4:
                    expected[2] = None
                self.assertEqual([c.value for c in b[row]],expected)
            self.assertEqual(b['P4'].value,a['P4'].value)
            self.assertEqual(b['P4'].data_type,'f')
            for old,new in zip(before.worksheets[1:],after.worksheets[1:]):
                self.assertEqual(list(old.values),list(new.values))
        finally:
            before.close()
            after.close()
        self.assertEqual(hashlib.sha256(matrix.read_bytes()).digest(),digest)
        repeat = self.run_import(template,matrix=self.output,output=self.folder/'repeat.xlsx')
        self.assertEqual(repeat.affected_rows,[])
        self.assertEqual((self.folder/'repeat.xlsx').read_bytes(),self.output.read_bytes())

    def test_other_english_code_with_same_russian_name_adds_without_touching_old_rows(self):
        matrix = self.ambiguous_matrix()
        template = self.template([('01 — Тестовое ограничение','З',{'C':'SECOND_CODE'})])
        before = self.rows(matrix)
        result = self.run_import(template,matrix=matrix)
        self.assertEqual((result.added,len(result.conflicts)),(3,0))
        self.assertEqual(self.rows()[:4],before)
        self.assertTrue(all(row['OPERATION_TYPE']=='SECOND_CODE' for row in self.rows()[4:]))

    def test_prom_cannot_change_other_operation_with_same_russian_name(self):
        matrix=self.ambiguous_matrix()
        template=self.template([('01 — Тестовое ограничение','З',{'C':'SECOND_CODE'})])
        result=self.run_import(template,matrix=matrix,environment='PROM')
        self.assertEqual(result.affected_rows,[])
        self.assertEqual(len(result.conflicts),3)
        self.assertEqual(self.output.read_bytes(),matrix.read_bytes())

    def test_actual_formula_range_not_global_catalog_order_defines_old_code(self):
        formula = "=INDEX('Типы Операций'!$A$3:$A$50,MATCH('Матрица решений комплаенс'!O:O,'Типы Операций'!$B$3:$B$50,0),1)"
        matrix = self.ambiguous_matrix(formula)
        book = load_workbook(matrix)
        try:
            reader = CodeFormulaReader(book)
            self.assertEqual(reader.value(book.worksheets[0]['P4']),'SECOND_CODE')
        finally:
            book.close()
        before = self.rows(matrix)
        result = self.run_import(self.template([('01 — Тестовое ограничение','З',{})]),matrix=matrix)
        self.assertEqual((result.added,len(result.conflicts)),(3,0))
        self.assertEqual(self.rows()[:4],before)

    def test_russian_function_names_and_semicolons_are_supported(self):
        matrix = self.ambiguous_matrix("=ИНДЕКС('Типы Операций'!A:A;ПОИСКПОЗ('Матрица решений комплаенс'!O:O;'Типы Операций'!B:B;0);1)")
        result = self.run_import(self.template([('01 — Тестовое ограничение','З',{})]),matrix=matrix)
        self.assertEqual(result.affected_rows,[])
        self.assertEqual(len(result.conflicts),0)
        self.assertEqual(self.output.read_bytes(),matrix.read_bytes())

    def test_subsystem_formula_with_shared_russian_label_is_read_by_english_result(self):
        def edit(book):
            book['Подсистемы'].append(['OTHER_SP','Демонстрационная подсистема'])
            sheet=book.worksheets[0]
            for row in range(2,6):
                sheet[f'H{row}']="=INDEX(Подсистемы!A:A,MATCH('Матрица решений комплаенс'!G:G,Подсистемы!B:B,0),1)"
        matrix=self.change(self.matrix,edit,'sp_formula.xlsx')
        result=self.run_import(self.template([('01 — Тестовое ограничение','З',{})]),matrix=matrix)
        self.assertEqual((result.added,len(result.conflicts)),(0,0))
        self.assertEqual(self.output.read_bytes(),matrix.read_bytes())

    def test_uncertain_other_field_does_not_block_different_known_english_code(self):
        matrix=self.ambiguous_matrix(extra_edit=lambda b:setattr(b.worksheets[0]['Z4'],'value','Неизвестный статус'))
        result=self.run_import(self.template([('01 — Тестовое ограничение','З',{'C':'SECOND_CODE'})]),matrix=matrix)
        self.assertEqual((result.added,len(result.conflicts)),(3,0))

    def test_unsupported_formula_remains_manual_without_russian_guessing(self):
        matrix=self.ambiguous_matrix('=UNSUPPORTED_FUNCTION(1)')
        result=self.run_import(self.template([('01 — Тестовое ограничение','З',{})]),matrix=matrix)
        self.assertEqual(result.affected_rows,[])
        self.assertTrue(any('неподдерживаемая функция' in c for c in result.conflicts))
        self.assertEqual(self.output.read_bytes(),matrix.read_bytes())

    def test_first_match_and_relative_range_positions(self):
        book=self.formula_book()
        sheet=book.create_sheet('Lookup')
        sheet['A2'],sheet['B2']='FIRST','Одинаковое название'
        sheet['A3'],sheet['B3']='SECOND','Одинаковое название'
        sheet['D4']='Одинаковое название'
        sheet['E4']='=INDEX(A:A,MATCH(D:D,B:B,0),1)'
        reader=CodeFormulaReader(book)
        self.assertEqual(reader.value(sheet['E4']),'FIRST')
        sheet['E4']='=INDEX(A3:A10,MATCH(D4,B3:B10,0))'
        self.assertEqual(reader.value(sheet['E4']),'SECOND')

    def test_excel_match_keeps_spaces_and_ignores_case(self):
        book=self.formula_book()
        sheet=book.create_sheet('Lookup')
        sheet['A2'],sheet['B2']='ONE','Название '
        sheet['A3'],sheet['B3']='TWO','Название'
        sheet['D4']='название'
        sheet['E4']='=INDEX(A:A,MATCH(D4,B:B,0))'
        self.assertEqual(CodeFormulaReader(book).value(sheet['E4']),'TWO')

    def test_match_respects_numeric_and_text_types(self):
        book=self.formula_book()
        sheet=book.create_sheet('Lookup')
        sheet['A2'],sheet['B2']='NUMBER',123
        sheet['A3'],sheet['B3']='TEXT','123'
        sheet['E4']='=INDEX(A:A,MATCH("123",B:B,0))'
        self.assertEqual(CodeFormulaReader(book).value(sheet['E4']),'TEXT')

    def test_match_errors_before_result_are_manual_and_errors_after_result_do_not_block(self):
        book=self.formula_book()
        sheet=book.create_sheet('Lookup')
        sheet['A2'],sheet['B2']='EN','Ключ'
        sheet['B3']='#N/A'
        sheet['E4']='=INDEX(A:A,MATCH("Ключ",B:B,0))'
        self.assertEqual(CodeFormulaReader(book).value(sheet['E4']),'EN')
        sheet['B1']='#N/A'
        with self.assertRaises(CodeFormulaError):
            CodeFormulaReader(book).value(sheet['E4'])

    def test_simple_links_cycles_and_blank_reference_do_not_mutate_cells(self):
        book=self.formula_book()
        sheet=book.create_sheet('Links')
        sheet['A1'],sheet['A2']='CODE','=A1'
        reader=CodeFormulaReader(book)
        self.assertEqual(reader.value(sheet['A2']),'CODE')
        sheet['B1'],sheet['B2']='=B2','=B1'
        with self.assertRaisesRegex(CodeFormulaError,'циклическая'):
            reader.value(sheet['B1'])
        sheet['C1']='=D10000'
        before=set(sheet._cells)
        self.assertIsNone(reader.value(sheet['C1']))
        self.assertEqual(set(sheet._cells),before)

    def test_quoted_sheet_with_apostrophe_and_explicit_source_columns(self):
        book=self.formula_book()
        ref=book.create_sheet("Справочник 'коды'")
        ref['C4'],ref['D4']='Ключ','ENGLISH'
        cell=book.worksheets[0]['P4']
        cell.value="=INDEX('Справочник ''коды'''!D4:D10,MATCH(\"Ключ\",'Справочник ''коды'''!C4:C10,0))"
        self.assertEqual(CodeFormulaReader(book).value(cell),'ENGLISH')

    def test_unsupported_modes_and_external_references_are_rejected(self):
        book=self.formula_book()
        cell=book.worksheets[0]['P4']
        formulas=('=INDEX(A:A,MATCH(O4,B:B,1))','=INDEX(A:A,MATCH(O4,B:B))',
                  "='[external.xlsx]Sheet1'!A1",'=DDE("cmd", "x", "y")','=A1&"text"')
        for value in formulas:
            with self.subTest(formula=value):
                cell.value=value
                with self.assertRaises(CodeFormulaError):
                    CodeFormulaReader(book).value(cell)

    def test_errors_missing_matches_and_out_of_bounds_are_rejected(self):
        book=self.formula_book()
        sheet=book.create_sheet('Lookup')
        sheet['A2'],sheet['B2']='EN','Ключ'
        sheet['D4']='Нет'
        cell=sheet['E4']
        cell.value='=INDEX(A:A,MATCH(D4,B:B,0))'
        with self.assertRaisesRegex(CodeFormulaError,'не найдено'):
            CodeFormulaReader(book).value(cell)
        cell.value='=INDEX(A2:A2,2)'
        with self.assertRaisesRegex(CodeFormulaError,'границы'):
            CodeFormulaReader(book).value(cell)
        cell.value='=INDEX(A:A,1E99999)'
        with self.assertRaises(CodeFormulaError):
            CodeFormulaReader(book).value(cell)

    def test_whole_column_match_cache_is_reused_and_unused_cells_not_created(self):
        book=self.formula_book()
        sheet=book.create_sheet('Lookup')
        sheet['A2'],sheet['B2']='EN','Ключ'
        formula='=INDEX(A:A,MATCH(D:D,B:B,0))'
        for row in range(4,104):
            sheet[f'D{row}'],sheet[f'E{row}']='Ключ',formula
        before=set(sheet._cells)
        reader=CodeFormulaReader(book)
        for row in range(4,104):
            self.assertEqual(reader.value(sheet[f'E{row}']),'EN')
        self.assertEqual(len(reader.match_indices),1)
        self.assertEqual(len(reader.expressions),1)
        self.assertEqual(before,set(sheet._cells))

    def test_whole_column_lookup_uses_current_row_21112_and_50000(self):
        book=self.formula_book()
        source=book.create_sheet('Dictionary')
        source['A2'],source['B2']='FIRST','Первый'
        source['A3'],source['B3']='SECOND','Второй'
        sheet=book.worksheets[0]
        formula="=INDEX(Dictionary!A:A,MATCH('Матрица решений комплаенс'!O:O,Dictionary!B:B,0),1)"
        sheet['O21112'],sheet['P21112']='Второй',formula
        sheet['O50000'],sheet['P50000']='Первый',formula
        before=set(sheet._cells)
        reader=CodeFormulaReader(book)
        self.assertEqual(reader.value(sheet['P21112']),'SECOND')
        self.assertEqual(reader.value(sheet['P50000']),'FIRST')
        self.assertEqual(before,set(sheet._cells))
