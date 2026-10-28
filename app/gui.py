"""Небольшое локальное окно без веб-сервера."""
import queue
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .importer import ROOT, import_files


def start(config_path=None):
    root = tk.Tk()
    root.title("MatrixImporter — импорт потреб")
    root.geometry("820x620")
    root.minsize(700, 580)
    style = ttk.Style(root)
    style.configure("Title.TLabel", font=("Arial", 21, "bold"))
    style.configure("Hint.TLabel", foreground="#526173")
    outer = ttk.Frame(root, padding=24)
    outer.pack(fill="both", expand=True)
    outer.columnconfigure(0, weight=1)
    ttk.Label(outer, text="Импорт потреб в матрицу", style="Title.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
    ttk.Label(outer, text="Выберите актуальную матрицу и заполненные шаблоны .xlsx.", style="Hint.TLabel").grid(row=1, column=0, sticky="w", pady=(0, 20))
    matrix_var = tk.StringVar()
    environment_var = tk.StringVar()
    status = tk.StringVar(value="Готов к работе")
    controls = []
    files = []
    events = queue.Queue()
    busy = False

    def pick_matrix():
        path = filedialog.askopenfilename(title="Основная матрица", filetypes=[("Excel", "*.xlsx")])
        if path:
            matrix_var.set(path)

    def show_files():
        file_list.delete(0, tk.END)
        for path in files:
            file_list.insert(tk.END, Path(path).name)

    def pick_templates():
        selected = filedialog.askopenfilenames(title="Заполненные шаблоны", filetypes=[("Excel", "*.xlsx")])
        for path in selected:
            if path not in files:
                files.append(path)
        show_files()

    def clear_templates():
        files.clear()
        show_files()

    def demo():
        environment_var.set("IFT")
        matrix_var.set(str(ROOT / "examples" / "01_matrix.xlsx"))
        files[:] = [str(ROOT / "examples" / "02_template_update.xlsx"), str(ROOT / "examples" / "03_template_new_group.xlsx")]
        show_files()
        status.set("Демо ИФТ: 4 добавленные строки (включая 1 новую версию), 1 без изменений.")

    ttk.Label(outer, text="1. Основная матрица").grid(row=2, column=0, sticky="w", pady=(0, 6))
    matrix_row = ttk.Frame(outer)
    matrix_row.grid(row=3, column=0, sticky="ew")
    matrix_row.columnconfigure(0, weight=1)
    entry = ttk.Entry(matrix_row, textvariable=matrix_var)
    entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))
    controls.append(entry)
    button = ttk.Button(matrix_row, text="Выбрать…", command=pick_matrix)
    button.grid(row=0, column=1)
    controls.append(button)
    ttk.Label(outer, text="2. Шаблоны потреб").grid(row=4, column=0, sticky="w", pady=(20, 6))
    list_row = ttk.Frame(outer)
    list_row.grid(row=5, column=0, sticky="nsew")
    list_row.columnconfigure(0, weight=1)
    list_row.rowconfigure(0, weight=1)
    outer.rowconfigure(5, weight=1)
    file_list = tk.Listbox(list_row, height=7, borderwidth=1, relief="solid", activestyle="none", font=("Arial", 12))
    file_list.grid(row=0, column=0, sticky="nsew")
    scrollbar = ttk.Scrollbar(list_row, orient="vertical", command=file_list.yview)
    scrollbar.grid(row=0, column=1, sticky="ns")
    file_list.configure(yscrollcommand=scrollbar.set)
    actions = ttk.Frame(outer)
    actions.grid(row=6, column=0, sticky="w", pady=(8, 16))
    ttk.Label(actions, text="Среда: ").pack(side="left")
    for value, label in (("IFT", "ИФТ"), ("PROM", "ПРОМ")):
        button = ttk.Radiobutton(actions, text=label, variable=environment_var, value=value)
        button.pack(side="left", padx=(0, 8))
        controls.append(button)
    for text, command in (("Добавить шаблоны…", pick_templates), ("Очистить", clear_templates), ("Загрузить демо", demo)):
        button = ttk.Button(actions, text=text, command=command)
        button.pack(side="left", padx=(0, 8))
        controls.append(button)
    ttk.Label(outer, text="Меняется только выбранная среда. Байпас проверяется вручную.\nФормулы пересчитаются при открытии результата в Р7-Офис или Excel.", style="Hint.TLabel", wraplength=740).grid(row=7, column=0, sticky="w", pady=(0, 18))
    progress = ttk.Progressbar(outer, mode="indeterminate")
    progress.grid(row=8, column=0, sticky="ew", pady=(0, 8))
    ttk.Label(outer, textvariable=status, wraplength=740).grid(row=9, column=0, sticky="w", pady=(0, 14))

    def begin():
        nonlocal busy
        if not matrix_var.get() or not files:
            messagebox.showinfo("Выберите файлы", "Нужны основная матрица и хотя бы один шаблон.")
            return
        if not environment_var.get():
            messagebox.showinfo("Выберите среду", "Укажите ИФТ или ПРОМ для задачи.")
            return
        output = filedialog.asksaveasfilename(title="Новый файл результата", defaultextension=".xlsx",
                                            initialfile=f"matrix_result_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                                            filetypes=[("Excel", "*.xlsx")], confirmoverwrite=True)
        if not output:
            return
        matrix, templates = matrix_var.get(), list(files)
        environment = environment_var.get()
        busy = True
        for control in controls:
            control.configure(state="disabled")
        progress.start()

        def worker():
            try:
                result = import_files(matrix, templates, output, config_path, lambda text: events.put(("progress", text)), environment=environment)
                events.put(("done", result))
            except Exception as exc:
                events.put(("error", str(exc)))
        threading.Thread(target=worker, daemon=True).start()

    run = ttk.Button(outer, text="Импортировать и сохранить…", command=begin)
    run.grid(row=10, column=0, sticky="e")
    controls.append(run)

    def poll():
        nonlocal busy
        try:
            while True:
                kind, value = events.get_nowait()
                if kind == "progress":
                    status.set(value)
                    continue
                busy = False
                progress.stop()
                for control in controls:
                    control.configure(state="normal")
                if kind == "done":
                    text = f"Добавлено: {value.added}. Обновлено: {value.updated}. Без изменений: {value.skipped}. На ручной разбор: {len(value.conflicts)}."
                    text += f" Новых ИФТ-версий: {value.versioned}. Включено в ПРОМ: {value.promoted}; прежних ПРОМ-версий отключено: {value.prom_deactivated}. Среда: {value.environment}. Байпас проверьте вручную."
                    status.set(text)
                    details = "\n\n".join(value.conflicts)
                    messagebox.showinfo("Готово", f"{text}\n\nСохранено:\n{value.output}\n\n{details}")
                else:
                    status.set("Импорт остановлен. Выходной файл не сохранён.")
                    messagebox.showerror("Не удалось выполнить импорт", value)
        except queue.Empty:
            pass
        root.after(100, poll)

    def close():
        if busy:
            messagebox.showinfo("Импорт выполняется", "Дождитесь завершения записи файла.")
        else:
            root.destroy()
    root.protocol("WM_DELETE_WINDOW", close)
    poll()
    root.mainloop()
