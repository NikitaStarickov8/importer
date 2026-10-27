#!/usr/bin/env python3
"""Запуск: python3 run.py; CLI: python3 run.py --help."""
import argparse
import json
import sys
from dataclasses import asdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "vendor"))


def main():
    parser = argparse.ArgumentParser(description="Перенос потреб из Excel-шаблонов в матрицу.")
    parser.add_argument("--matrix", help="Основная матрица .xlsx")
    parser.add_argument("--templates", nargs="+", help="Один или несколько шаблонов .xlsx")
    parser.add_argument("--output", help="Новое имя результата .xlsx")
    parser.add_argument("--config", help="Другой config.json")
    args = parser.parse_args()
    try:
        from app.importer import ImportProblem, import_files
    except ModuleNotFoundError as exc:
        print(f"Не хватает зависимости: {exc.name}. Выполните: python3 -m pip install -r requirements.txt", file=sys.stderr)
        return 1
    if any((args.matrix, args.templates, args.output)):
        if not all((args.matrix, args.templates, args.output)):
            parser.error("Для CLI нужны --matrix, --templates и --output вместе.")
        try:
            result = import_files(args.matrix, args.templates, args.output, args.config,
                                  progress=lambda s: print(s, file=sys.stderr))
            print(json.dumps(asdict(result), ensure_ascii=False, indent=2))
            return 0
        except (ImportProblem, OSError, ValueError) as exc:
            print(f"Импорт не выполнен: {exc}", file=sys.stderr)
            return 1
    try:
        from app.gui import start
        start(args.config)
    except ImportError:
        print("Для окна требуется Python с Tkinter. Можно использовать CLI: python3 run.py --help", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
