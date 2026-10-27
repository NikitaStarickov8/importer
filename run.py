#!/usr/bin/env python3
"""Локальная страница: python3 run.py; CLI: python3 run.py --help."""
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
    parser.add_argument("--port", type=int, default=8765, help="Порт локального интерфейса (по умолчанию 8765)")
    parser.add_argument("--no-browser", action="store_true", help="Не открывать браузер автоматически")
    parser.add_argument("--desktop", action="store_true", help="Старое окно Tkinter (необязательно)")
    args = parser.parse_args()
    try:
        from app.importer import ImportProblem, import_files
    except ModuleNotFoundError as exc:
        print(f"Не хватает зависимости: {exc.name}. Проверьте, что папка vendor перенесена целиком.", file=sys.stderr)
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
    if args.desktop:
        try:
            from app.gui import start
            start(args.config)
        except ImportError as exc:
            print(f"Окно Tkinter недоступно: {exc}. Запустите python3 run.py для браузерного интерфейса.", file=sys.stderr)
            return 1
    else:
        if not 0 <= args.port <= 65535:
            parser.error("Порт должен находиться в диапазоне 0…65535.")
        from app.web import start
        try:
            start(args.config, args.port, open_browser=not args.no_browser)
        except OSError as exc:
            print(f"Не удалось открыть локальный интерфейс: {exc}", file=sys.stderr)
            return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
