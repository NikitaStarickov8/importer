#!/bin/sh
cd "$(dirname "$0")" || exit 1
if ! command -v python3 >/dev/null 2>&1; then
  echo 'Для запуска нужен Python 3.10 или новее с Tkinter.'
  read -r answer
  exit 1
fi
python3 run.py
status=$?
if [ "$status" -ne 0 ]; then
  echo 'Нажмите Enter, чтобы закрыть окно.'
  read -r answer
fi
exit "$status"
