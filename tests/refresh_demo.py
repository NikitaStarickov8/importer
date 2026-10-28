"""Обновление только искусственного справочника демо для версии 0.3.0."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'vendor'))
from openpyxl import load_workbook


def main():
    path = ROOT / 'examples/01_matrix.xlsx'
    book = load_workbook(path)
    sheet = book['Типы транзакций']
    existing = {row[0] for row in sheet.values}
    for code, description in [('NON_FINANCIAL', 'Не финансовые операции'), ('FINANCIAL', 'Финансовые операции')]:
        if code not in existing:
            sheet.append([code, description])
    book.save(path)
    book.close()


if __name__ == '__main__':
    main()
