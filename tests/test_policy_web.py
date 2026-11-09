"""HTTP-проверки новой политики по умолчанию (без отключения business_rules)."""
import test_web as web
import unittest
import io
from openpyxl import load_workbook


class PolicyWebTests(unittest.TestCase):
    tearDown = web.WebTests.tearDown
    request = web.WebTests.request
    json_request = web.WebTests.json_request
    create = web.WebTests.create
    wait = web.WebTests.wait

    def setUp(self):
        web.WebTests.setUp(self)
        self.policy_patch.stop()

    def test_different_header_operations_stop_http_import(self):
        book = load_workbook(self.templates[0])
        try:
            book.worksheets[0]['D6'] = 'Другая операция'
            data = io.BytesIO()
            book.save(data)
        finally:
            book.close()
        ident = self.create(templates=[data.getvalue()])
        self.json_request('POST', f'/api/jobs/{ident}/start', {})
        job = self.wait(ident)
        self.assertEqual(job['state'], 'error', job)
        self.assertIn('B6 и D6 указаны разные операции', job['error'])
        self.assertNotIn('download_url', job)

    def test_platform_overrides_invalid_channel_over_http(self):
        book = load_workbook(self.templates[1])
        try:
            for row in (18, 19):
                book.worksheets[0][f'A{row}'] = '=BAD_CHANNEL(1)'
            data = io.BytesIO()
            book.save(data)
        finally:
            book.close()
        ident = self.create(templates=[data.getvalue()])
        self.json_request('POST', f'/api/jobs/{ident}/start', {})
        job = self.wait(ident)
        self.assertEqual(job['state'], 'done', job)
        self.assertEqual((job['counts']['added'], job['counts']['conflicted']), (2, 0))
        code, saved, _ = self.request('GET', job['download_url'])
        self.assertEqual(code, 200)
        book = load_workbook(io.BytesIO(saved))
        try:
            self.assertEqual(book.worksheets[0]['E6'].value, 'СБОЛ про')
            self.assertEqual(book.worksheets[0]['E7'].value, 'СБОЛ про')
        finally:
            book.close()

    def test_existing_formula_english_code_is_compared_to_template_code_over_http(self):
        book=load_workbook(self.matrix)
        try:
            book['Подсистемы'].append(['SECOND_SP','Демонстрационная подсистема'])
            book['Типы Операций'].append(['SECOND_OPERATION','Тестовая выдача'])
            sheet=book.worksheets[0]
            for row in range(2,6):
                sheet[f'H{row}']="=INDEX(Подсистемы!A:A,MATCH('Матрица решений комплаенс'!G:G,Подсистемы!B:B,0),1)"
                sheet[f'P{row}']="=INDEX('Типы Операций'!A:A,MATCH('Матрица решений комплаенс'!O:O,'Типы Операций'!B:B,0),1)"
            data=io.BytesIO()
            book.save(data)
        finally:
            book.close()
        matrix=data.getvalue()
        ident=self.create(matrix=matrix,templates=[self.templates[0].read_bytes()])
        self.json_request('POST',f'/api/jobs/{ident}/start',{})
        result=self.wait(ident)
        self.assertEqual(result['state'],'done',result)
        self.assertEqual(result['affected_rows'],[])
        self.assertEqual(result['counts']['conflicted'],0,result)
        code,saved,_=self.request('GET',result['download_url'])
        self.assertEqual(code,200)
        self.assertEqual(saved,matrix)

    def test_service_rows_without_reference_entries_work_over_http(self):
        book = load_workbook(self.matrix)
        try:
            for name in ('Статусы Стоп Листа','COMPLIANCE_STOP_LIST_CODES'):
                book[name].delete_rows(2, 2)
            data = io.BytesIO()
            book.save(data)
        finally:
            book.close()
        ident = self.create(matrix=data.getvalue(), templates=[self.templates[1].read_bytes()])
        self.json_request('POST', f'/api/jobs/{ident}/start', {})
        result = self.wait(ident)
        self.assertEqual(result['state'], 'done', result)
        self.assertEqual((result['counts']['added'],result['counts']['conflicted']), (2,0))
        code, saved, _ = self.request('GET', result['download_url'])
        self.assertEqual(code, 200)
        book = load_workbook(io.BytesIO(saved))
        try:
            sheet = book.worksheets[0]
            self.assertEqual(sheet['AA6'].value, 'TECH_ERROR')
            self.assertEqual(sheet['K6'].value, 'TECH_ERROR')
            self.assertEqual(sheet['AA7'].value, 'NORMAL')
            self.assertEqual(sheet['K7'].value, 'NORMAL')
            self.assertEqual(sheet['AB6'].value, 0)
            self.assertEqual(sheet['AB7'].value, 0)
        finally:
            book.close()

    def test_default_policy_upload_download_repeat_and_prom(self):
        ident = self.create()
        self.json_request('POST', f'/api/jobs/{ident}/start', {})
        job = self.wait(ident)
        self.assertEqual(job['state'], 'done', job)
        self.assertEqual((job['counts']['added'], job['counts']['updated']), (2, 0))
        self.assertEqual(job['counts']['conflicted'], 0, job)
        self.assertNotIn('CODE:', job['change_report'])
        self.assertIn('ПРОПУЩЕНО (', job['change_report'])
        self.assertIn('ДОБАВЛЕНО (2)', job['change_report'])
        self.assertIn('ИЗМЕНЕНО (0)', job['change_report'])
        self.assertTrue(all('001_' not in o['source'] for o in job['outcomes']))
        code, matrix, _ = self.request('GET', job['download_url'])
        self.assertEqual(code, 200)
        ident = self.create(matrix=matrix)
        self.json_request('POST', f'/api/jobs/{ident}/start', {})
        repeat = self.wait(ident)
        self.assertEqual(repeat['state'], 'done', repeat)
        self.assertEqual(repeat['affected_rows'], [])
        ident = self.create(matrix=matrix, environment='PROM')
        self.json_request('POST', f'/api/jobs/{ident}/start', {})
        prom = self.wait(ident)
        self.assertEqual(prom['state'], 'done', prom)
        self.assertEqual(prom['counts']['added'], 0)
        self.assertEqual(prom['counts']['promoted'], 3)
        self.assertEqual(prom['counts']['conflicted'], 0, prom)

    def test_russian_operation_typo_does_not_block_http_import(self):
        book = load_workbook(self.templates[1])
        try:
            sheet = book.worksheets[0]
            sheet['C18'], sheet['C19'] = 'Опечатка в названии операции', '#N/A'
            data = io.BytesIO()
            book.save(data)
        finally:
            book.close()

        ident = self.create(templates=[data.getvalue()])
        self.json_request('POST', f'/api/jobs/{ident}/start', {})
        result = self.wait(ident)
        self.assertEqual(result['state'], 'done', result)
        self.assertEqual(result['counts']['added'], 2)
        self.assertEqual(result['counts']['conflicted'], 0)
        code, saved, _ = self.request('GET', result['download_url'])
        self.assertEqual(code, 200)
        book = load_workbook(io.BytesIO(saved))
        try:
            self.assertEqual(book.worksheets[0]['O6'].value, 'Проверка баланса')
            self.assertEqual(book.worksheets[0]['O7'].value, 'Проверка баланса')
        finally:
            book.close()

    def test_duplicate_english_codes_and_subsystem_typo_work_over_http(self):
        book = load_workbook(self.matrix)
        try:
            book['Подсистемы'].append(['DEMO_LOAN','Второе русское имя'])
            book['Типы Операций'].append(['CHECK_BALANCE','Второе имя операции'])
            matrix_data = io.BytesIO()
            book.save(matrix_data)
        finally:
            book.close()
        book = load_workbook(self.templates[1])
        try:
            sheet = book.worksheets[0]
            sheet['B18'], sheet['B19'] = 'Опечатка в подсистеме', '#N/A'
            data = io.BytesIO()
            book.save(data)
        finally:
            book.close()
        ident = self.create(matrix=matrix_data.getvalue(), templates=[data.getvalue()])
        self.json_request('POST', f'/api/jobs/{ident}/start', {})
        result = self.wait(ident)
        self.assertEqual(result['state'], 'done', result)
        self.assertEqual((result['counts']['added'], result['counts']['conflicted']), (2, 0))
        code, saved, _ = self.request('GET', result['download_url'])
        self.assertEqual(code, 200)
        book = load_workbook(io.BytesIO(saved))
        try:
            for row in (6, 7):
                self.assertEqual(book.worksheets[0][f'G{row}'].value, 'Демонстрационная подсистема')
                self.assertEqual(book.worksheets[0][f'O{row}'].value, 'Проверка баланса')
        finally:
            book.close()
