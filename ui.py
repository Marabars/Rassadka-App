"""Tkinter UI для генератора рассадки офиса."""
from __future__ import annotations

import contextlib
import io
import os
import queue
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk

sys.path.insert(0, str(Path(__file__).parent))

_APP_DIR = Path(__file__).parent
_SAMPLE_DIR = _APP_DIR.parent / "Rassadka v0.1"
_CONFIG = _APP_DIR / "config.yaml"

_CHOICES_HINT = "График Май гугл Таблица Обезличенный.xlsx"
_TEMPLATE_HINT = "График Май Обезличенный.xlsx"
_OUTPUT_DEFAULT = "Рассадка_следующий_месяц.xlsx"


def _get_sheet_names(path: str) -> list[str]:
    """Return sheet names from an xlsx file without loading data."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        names = wb.sheetnames
        wb.close()
        return names
    except Exception:
        return []


class SeatingApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Генератор рассадки офиса")
        self.resizable(False, False)

        self._log_queue: queue.Queue[tuple[str, str | None]] = queue.Queue()
        self._output_path: Path | None = None

        self._var_choices = tk.StringVar()
        self._var_template = tk.StringVar()
        self._var_output = tk.StringVar()
        self._var_omap = tk.StringVar()
        self._var_choices_sheet = tk.StringVar()
        self._var_template_sheet = tk.StringVar()

        self._prefill_demo_paths()
        self._build_ui()

        # Auto-fill sheet names for pre-filled paths
        self._on_choices_path_change()
        self._on_template_path_change()

        # Watch path fields for changes
        self._var_choices.trace_add("write", lambda *_: self._on_choices_path_change())
        self._var_template.trace_add("write", lambda *_: self._on_template_path_change())

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------

    def _build_ui(self):
        self.columnconfigure(1, minsize=320)

        # Title
        ttk.Label(self, text="Генератор рассадки офиса", font=("Segoe UI", 13, "bold")).grid(
            row=0, column=0, columnspan=3, pady=(14, 8)
        )
        ttk.Separator(self, orient="horizontal").grid(
            row=1, column=0, columnspan=3, sticky="ew", padx=12
        )

        # File + sheet rows
        self._build_file_row(2, "Файл выборов *", self._var_choices, save=False)
        self._cb_choices_sheet = self._build_sheet_row(3, "Лист (выборы) *", self._var_choices_sheet)

        self._build_file_row(4, "Рассадка прошлого месяца *", self._var_template, save=False)
        self._cb_template_sheet = self._build_sheet_row(5, "Лист (шаблон) *", self._var_template_sheet)

        self._build_file_row(6, "Выходной файл *", self._var_output, save=True)
        self._build_file_row(7, "Карта офиса", self._var_omap, save=False, optional=True)

        ttk.Separator(self, orient="horizontal").grid(
            row=8, column=0, columnspan=3, sticky="ew", padx=12, pady=(4, 0)
        )

        # Seats section
        self._build_seats_section(9)

        ttk.Separator(self, orient="horizontal").grid(
            row=13, column=0, columnspan=3, sticky="ew", padx=12, pady=(4, 0)
        )

        # Generate button
        self._btn_generate = ttk.Button(
            self, text="  Сгенерировать  ", command=self._on_generate
        )
        self._btn_generate.grid(row=14, column=0, columnspan=3, pady=10)

        ttk.Separator(self, orient="horizontal").grid(
            row=15, column=0, columnspan=3, sticky="ew", padx=12
        )

        # Log
        ttk.Label(self, text="Лог:", font=("Segoe UI", 9)).grid(
            row=16, column=0, columnspan=3, sticky="w", padx=14, pady=(6, 2)
        )
        self._log = scrolledtext.ScrolledText(
            self, width=72, height=10, state="disabled",
            font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4",
            insertbackground="white", relief="flat", bd=0,
        )
        self._log.grid(row=17, column=0, columnspan=3, padx=14, pady=(0, 6))

        # Status + open button
        bottom = ttk.Frame(self)
        bottom.grid(row=18, column=0, columnspan=3, sticky="ew", padx=12, pady=(0, 12))
        bottom.columnconfigure(0, weight=1)

        self._status_var = tk.StringVar(value="Готов к работе")
        ttk.Label(bottom, textvariable=self._status_var, font=("Segoe UI", 9)).grid(
            row=0, column=0, sticky="w"
        )
        self._btn_open = ttk.Button(
            bottom, text="Открыть результат", command=self._on_open_result, state="disabled"
        )
        self._btn_open.grid(row=0, column=1, sticky="e")

    def _build_file_row(self, row: int, label: str, var: tk.StringVar, save: bool, optional: bool = False):
        lbl = label if not optional else f"{label} (необяз.)"
        ttk.Label(self, text=lbl, width=24, anchor="e").grid(
            row=row, column=0, sticky="e", padx=(12, 4), pady=3
        )
        ttk.Entry(self, textvariable=var, width=42).grid(
            row=row, column=1, sticky="ew", pady=3
        )
        ttk.Button(
            self, text="Обзор...",
            command=lambda s=save, v=var: self._browse(v, s)
        ).grid(row=row, column=2, padx=(4, 12), pady=3)

    def _build_sheet_row(self, row: int, label: str, var: tk.StringVar) -> ttk.Combobox:
        ttk.Label(self, text=label, width=24, anchor="e").grid(
            row=row, column=0, sticky="e", padx=(12, 4), pady=2
        )
        cb = ttk.Combobox(self, textvariable=var, width=40, state="readonly")
        cb.grid(row=row, column=1, sticky="ew", pady=2)
        return cb

    def _build_seats_section(self, start_row: int):
        ttk.Label(self, text="Доступные места:", font=("Segoe UI", 9)).grid(
            row=start_row, column=0, columnspan=3, sticky="w", padx=14, pady=(6, 2)
        )

        frame = ttk.Frame(self)
        frame.grid(row=start_row + 1, column=0, columnspan=3, sticky="ew", padx=14)
        frame.columnconfigure(0, weight=1)

        self._seats_listbox = tk.Listbox(
            frame, height=6, selectmode=tk.EXTENDED,
            font=("Consolas", 9), activestyle="none",
        )
        self._seats_listbox.grid(row=0, column=0, sticky="ew")
        sb = ttk.Scrollbar(frame, orient="vertical", command=self._seats_listbox.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self._seats_listbox.config(yscrollcommand=sb.set)

        ctrl = ttk.Frame(self)
        ctrl.grid(row=start_row + 2, column=0, columnspan=3, sticky="ew", padx=14, pady=(2, 4))

        ttk.Label(ctrl, text="Добавить:").pack(side="left")
        self._var_new_seat = tk.StringVar()
        entry = ttk.Entry(ctrl, textvariable=self._var_new_seat, width=14)
        entry.pack(side="left", padx=(4, 4))
        entry.bind("<Return>", lambda _: self._on_add_seat())
        ttk.Button(ctrl, text="Добавить", command=self._on_add_seat).pack(side="left")
        ttk.Button(ctrl, text="Удалить выбранное", command=self._on_delete_seat).pack(side="right")

    # ------------------------------------------------------------------
    # Seats management
    # ------------------------------------------------------------------

    def _populate_seats_from_template(self, path: str):
        """Load seat list from template file and fill the listbox."""
        if not path or not Path(path).exists():
            return
        try:
            sys.path.insert(0, str(_APP_DIR))
            from app.readers.template_reader import read_template
            sheet = self._var_template_sheet.get() or None
            if not sheet:
                return
            td = read_template(Path(path), sheet)
            self._seats_listbox.delete(0, tk.END)
            for seat in td.all_seats:
                self._seats_listbox.insert(tk.END, seat)
        except Exception:
            pass  # silently skip — user can still edit manually

    def _on_add_seat(self):
        raw = self._var_new_seat.get().strip()
        if not raw:
            return
        existing = list(self._seats_listbox.get(0, tk.END))
        if raw not in existing:
            self._seats_listbox.insert(tk.END, raw)
        self._var_new_seat.set("")

    def _on_delete_seat(self):
        for idx in reversed(self._seats_listbox.curselection()):
            self._seats_listbox.delete(idx)

    def _get_seats(self) -> list[str]:
        return list(self._seats_listbox.get(0, tk.END))

    # ------------------------------------------------------------------
    # File dialogs + sheet auto-detect
    # ------------------------------------------------------------------

    def _browse(self, var: tk.StringVar, save: bool):
        opts = {"filetypes": [("Excel files", "*.xlsx"), ("All files", "*.*")]}
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", **opts) if save \
               else filedialog.askopenfilename(**opts)
        if path:
            var.set(path)

    def _on_choices_path_change(self):
        self._refresh_sheet_combo(self._var_choices.get(), self._cb_choices_sheet, self._var_choices_sheet, "2026")

    def _on_template_path_change(self):
        path = self._var_template.get()
        self._refresh_sheet_combo(path, self._cb_template_sheet, self._var_template_sheet, "График")
        self._populate_seats_from_template(path)

    def _refresh_sheet_combo(self, path: str, cb: ttk.Combobox, var: tk.StringVar, preferred: str):
        """Load sheet names from file and update combobox. Pick preferred name if present."""
        if not path or not Path(path).exists():
            cb["values"] = []
            return
        sheets = _get_sheet_names(path)
        # Filter out service sheets unlikely to be the main data sheet
        cb["values"] = sheets
        if not sheets:
            return
        current = var.get()
        if current in sheets:
            return  # keep existing selection
        # Auto-select: prefer the configured name, else first non-service sheet
        pick = preferred if preferred in sheets else sheets[0]
        var.set(pick)

    # ------------------------------------------------------------------
    # Generate
    # ------------------------------------------------------------------

    def _on_generate(self):
        choices = self._var_choices.get().strip()
        template = self._var_template.get().strip()
        output = self._var_output.get().strip()
        choices_sheet = self._var_choices_sheet.get().strip()
        template_sheet = self._var_template_sheet.get().strip()
        omap = self._var_omap.get().strip() or None

        errors = []
        if not choices:
            errors.append("Укажите файл выборов.")
        if not template:
            errors.append("Укажите шаблон рассадки.")
        if not output:
            errors.append("Укажите путь к выходному файлу.")
        if not choices_sheet:
            errors.append("Выберите лист в файле выборов.")
        if not template_sheet:
            errors.append("Выберите лист в шаблоне.")
        if errors:
            self._log_clear()
            for e in errors:
                self._log_append(e, tag="error")
            return

        self._btn_generate.config(state="disabled")
        self._btn_open.config(state="disabled")
        self._status_var.set("Генерация...")
        self._log_clear()
        self._output_path = None

        seats = self._get_seats() or None

        thread = threading.Thread(
            target=self._run_pipeline,
            args=(
                Path(choices), Path(template), Path(output),
                Path(omap) if omap else None,
                choices_sheet, template_sheet,
                seats,
            ),
            daemon=True,
        )
        thread.start()
        self.after(100, self._poll_log_queue)

    def _run_pipeline(
        self,
        choices: Path, template: Path, output: Path,
        omap: Path | None, choices_sheet: str, template_sheet: str,
        seats: list[str] | None = None,
    ):
        buf = io.StringIO()
        try:
            with contextlib.redirect_stdout(buf):
                from generate_seating import run
                exit_code = run(
                    choices_path=choices,
                    template_path=template,
                    output_path=output,
                    config_path=_CONFIG,
                    office_map_path=omap,
                    choices_sheet_override=choices_sheet,
                    template_sheet_override=template_sheet,
                    seats_override=seats,
                )
            for line in buf.getvalue().splitlines():
                if line.strip():
                    self._log_queue.put(("ok", line))

            tag = "__done_ok__" if exit_code == 0 else "__done_err__"
            msg = f"Файл сохранён: {output}" if exit_code == 0 \
                  else "Завершено с ошибками. Проверьте Validation Report."
            self._log_queue.put(("success" if exit_code == 0 else "error", msg))
            self._log_queue.put((tag, str(output)))

        except Exception as exc:
            self._log_queue.put(("error", f"Ошибка: {exc}"))
            self._log_queue.put(("__done_err__", None))

    def _poll_log_queue(self):
        try:
            while True:
                tag, text = self._log_queue.get_nowait()
                if tag == "__done_ok__":
                    self._output_path = Path(text)
                    self._btn_open.config(state="normal")
                    self._btn_generate.config(state="normal")
                    self._status_var.set("Готово!")
                    return
                elif tag == "__done_err__":
                    self._btn_generate.config(state="normal")
                    self._status_var.set("Завершено с ошибками")
                    return
                else:
                    self._log_append(text, tag=tag if tag != "ok" else None)
        except queue.Empty:
            pass
        self.after(100, self._poll_log_queue)

    # ------------------------------------------------------------------
    # Log helpers
    # ------------------------------------------------------------------

    def _log_clear(self):
        self._log.config(state="normal")
        self._log.delete("1.0", tk.END)
        self._log.config(state="disabled")

    def _log_append(self, text: str, tag: str | None = None):
        self._log.config(state="normal")
        if tag == "error":
            self._log.insert(tk.END, "> " + text + "\n", "error")
            self._log.tag_config("error", foreground="#f48771")
        elif tag == "success":
            self._log.insert(tk.END, "> " + text + "\n", "success")
            self._log.tag_config("success", foreground="#4ec9b0")
        else:
            self._log.insert(tk.END, "> " + text + "\n")
        self._log.see(tk.END)
        self._log.config(state="disabled")

    # ------------------------------------------------------------------
    # Open result
    # ------------------------------------------------------------------

    def _on_open_result(self):
        if self._output_path and self._output_path.exists():
            os.startfile(self._output_path)

    # ------------------------------------------------------------------
    # Demo prefill
    # ------------------------------------------------------------------

    def _prefill_demo_paths(self):
        choices = _SAMPLE_DIR / _CHOICES_HINT
        template = _SAMPLE_DIR / _TEMPLATE_HINT
        if choices.exists():
            self._var_choices.set(str(choices))
        if template.exists():
            self._var_template.set(str(template))
        if choices.exists() or template.exists():
            self._var_output.set(str(_SAMPLE_DIR / _OUTPUT_DEFAULT))


def main():
    SeatingApp().mainloop()


if __name__ == "__main__":
    main()
