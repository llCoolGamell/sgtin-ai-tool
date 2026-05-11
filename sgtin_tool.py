"""SGTIN AI 01-21 tool.

Графический инструмент для пакетной обработки SGTIN кодов маркировки.

Возможности:
- Верхнее поле для ввода SGTIN (по одному в строке).
- Кнопка "УДАЛИТЬ AI 01-21": убирает идентификаторы применения
  AI(01) и AI(21) из строк длиной 31 символ.
  - 31 символ: "01" + GTIN(14) + "21" + Serial(13) -> Serial+GTIN без AI = 27 символов.
- Кнопка "Добавить AI 01-21": добавляет AI(01) и AI(21) к строкам
  длиной 27 символов: GTIN(14) + Serial(13) -> "01" + GTIN + "21" + Serial = 31 символ.
- Нижнее поле с результатом, поддерживает вставку, копирование, удаление, редактирование.
- Кнопка "Выгрузить в Excel" сохраняет содержимое нижнего поля в .xlsx файл.
- Рассчитано на тысячи строк: обработка идёт построчно без лишних копий.
"""
from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Iterable, List, Tuple

try:
    from openpyxl import Workbook
except ImportError:  # pragma: no cover - openpyxl is a runtime dep
    Workbook = None  # type: ignore[assignment]


REMOVE_LEN = 31
ADD_LEN = 27
GTIN_LEN = 14
SERIAL_LEN = 13
AI_01 = "01"
AI_21 = "21"


def remove_ai(code: str) -> str:
    """Удалить AI(01) и AI(21) из SGTIN длиной 31 символ.

    Ожидается формат: "01" + GTIN(14) + "21" + Serial(13).
    Возвращает GTIN + Serial (27 символов).
    """
    if len(code) != REMOVE_LEN:
        raise ValueError(
            f"ожидалось {REMOVE_LEN} символов, получено {len(code)}"
        )
    if not code[: len(AI_01)] == AI_01:
        raise ValueError("строка не начинается с AI(01)")
    if code[len(AI_01) + GTIN_LEN : len(AI_01) + GTIN_LEN + len(AI_21)] != AI_21:
        raise ValueError("AI(21) не найден в ожидаемой позиции")
    gtin = code[len(AI_01) : len(AI_01) + GTIN_LEN]
    serial = code[len(AI_01) + GTIN_LEN + len(AI_21) :]
    return gtin + serial


def add_ai(code: str) -> str:
    """Добавить AI(01) и AI(21) к SGTIN длиной 27 символов.

    Ожидается формат: GTIN(14) + Serial(13).
    Возвращает "01" + GTIN + "21" + Serial (31 символ).
    """
    if len(code) != ADD_LEN:
        raise ValueError(
            f"ожидалось {ADD_LEN} символов, получено {len(code)}"
        )
    gtin = code[:GTIN_LEN]
    serial = code[GTIN_LEN:]
    return AI_01 + gtin + AI_21 + serial


def process_lines(lines: Iterable[str], mode: str) -> Tuple[List[str], List[str]]:
    """Применить выбранную операцию к каждой непустой строке.

    Возвращает (результаты, ошибки). Каждая ошибка содержит номер строки
    (1-based) и описание проблемы.
    """
    if mode == "remove":
        op = remove_ai
    elif mode == "add":
        op = add_ai
    else:  # pragma: no cover - defensive
        raise ValueError(f"неизвестная операция: {mode}")

    results: List[str] = []
    errors: List[str] = []
    for idx, raw in enumerate(lines, start=1):
        code = raw.strip()
        if not code:
            continue
        try:
            results.append(op(code))
        except ValueError as exc:
            errors.append(f"строка {idx}: '{code}' — {exc}")
    return results, errors


class SgtinApp(tk.Tk):
    """Главное окно приложения."""

    def __init__(self) -> None:
        super().__init__()
        self.title("SGTIN AI 01-21 — инструмент")
        self.geometry("900x700")
        self.minsize(640, 480)

        self._build_ui()
        self._bind_shortcuts()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=8)
        root.pack(fill=tk.BOTH, expand=True)
        root.rowconfigure(1, weight=1)
        root.rowconfigure(3, weight=1)
        root.columnconfigure(0, weight=1)

        ttk.Label(root, text="Входные SGTIN (по одному в строке):").grid(
            row=0, column=0, sticky="w"
        )
        self.input_text = self._make_text_area(root, row=1)

        middle = ttk.Frame(root)
        middle.grid(row=2, column=0, sticky="ew", pady=8)
        for col in range(4):
            middle.columnconfigure(col, weight=1)

        ttk.Button(
            middle, text="УДАЛИТЬ AI 01-21", command=self.on_remove
        ).grid(row=0, column=0, padx=4, sticky="ew")
        ttk.Button(
            middle, text="Добавить AI 01-21", command=self.on_add
        ).grid(row=0, column=1, padx=4, sticky="ew")
        ttk.Button(
            middle, text="Очистить результат", command=self.on_clear_output
        ).grid(row=0, column=2, padx=4, sticky="ew")
        ttk.Button(
            middle, text="Выгрузить в Excel", command=self.on_export_excel
        ).grid(row=0, column=3, padx=4, sticky="ew")

        ttk.Label(root, text="Результат:").grid(row=3, column=0, sticky="nw")
        self.output_text = self._make_text_area(root, row=3, label_row=True)

        self.status_var = tk.StringVar(value="Готово")
        ttk.Label(root, textvariable=self.status_var, anchor="w").grid(
            row=4, column=0, sticky="ew", pady=(4, 0)
        )

    def _make_text_area(
        self, parent: ttk.Frame, row: int, label_row: bool = False
    ) -> tk.Text:
        frame = ttk.Frame(parent)
        if label_row:
            frame.grid(row=row, column=0, sticky="nsew", pady=(20, 0))
        else:
            frame.grid(row=row, column=0, sticky="nsew", pady=(2, 0))
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        text = tk.Text(frame, wrap="none", undo=True, font=("Consolas", 10))
        text.grid(row=0, column=0, sticky="nsew")

        yscroll = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll = ttk.Scrollbar(frame, orient="horizontal", command=text.xview)
        xscroll.grid(row=1, column=0, sticky="ew")
        text.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self._attach_context_menu(text)
        return text

    def _attach_context_menu(self, widget: tk.Text) -> None:
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(
            label="Вырезать", command=lambda: widget.event_generate("<<Cut>>")
        )
        menu.add_command(
            label="Копировать", command=lambda: widget.event_generate("<<Copy>>")
        )
        menu.add_command(
            label="Вставить", command=lambda: widget.event_generate("<<Paste>>")
        )
        menu.add_separator()
        menu.add_command(
            label="Выделить всё",
            command=lambda: widget.event_generate("<<SelectAll>>"),
        )
        menu.add_command(
            label="Удалить выделенное",
            command=lambda: self._delete_selection(widget),
        )

        def show(event: tk.Event) -> None:
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

        widget.bind("<Button-3>", show)

    @staticmethod
    def _delete_selection(widget: tk.Text) -> None:
        try:
            widget.delete("sel.first", "sel.last")
        except tk.TclError:
            pass

    def _bind_shortcuts(self) -> None:
        for widget in (self.input_text, self.output_text):
            widget.bind("<Control-a>", self._select_all)
            widget.bind("<Control-A>", self._select_all)

    @staticmethod
    def _select_all(event: tk.Event) -> str:
        widget: tk.Text = event.widget  # type: ignore[assignment]
        widget.tag_add("sel", "1.0", "end-1c")
        widget.mark_set("insert", "1.0")
        widget.see("insert")
        return "break"

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------
    def _get_input_lines(self) -> List[str]:
        raw = self.input_text.get("1.0", "end-1c")
        return raw.splitlines()

    def _set_output(self, results: List[str]) -> None:
        self.output_text.delete("1.0", "end")
        if results:
            self.output_text.insert("1.0", "\n".join(results))

    def _run(self, mode: str) -> None:
        lines = self._get_input_lines()
        if not any(line.strip() for line in lines):
            messagebox.showwarning(
                "Нет данных", "Введите хотя бы один SGTIN в верхнем поле."
            )
            return

        results, errors = process_lines(lines, mode)
        self._set_output(results)

        if errors:
            preview = "\n".join(errors[:20])
            extra = "" if len(errors) <= 20 else f"\n… и ещё {len(errors) - 20}"
            messagebox.showerror(
                "Ошибки в данных",
                f"Не обработано строк: {len(errors)}\n\n{preview}{extra}",
            )
        self.status_var.set(
            f"Обработано: {len(results)}; ошибок: {len(errors)}"
        )

    def on_remove(self) -> None:
        self._run("remove")

    def on_add(self) -> None:
        self._run("add")

    def on_clear_output(self) -> None:
        self.output_text.delete("1.0", "end")
        self.status_var.set("Результат очищен")

    def on_export_excel(self) -> None:
        if Workbook is None:
            messagebox.showerror(
                "openpyxl не установлен",
                "Установите пакет openpyxl: pip install openpyxl",
            )
            return

        data = self.output_text.get("1.0", "end-1c").splitlines()
        data = [line for line in data if line.strip()]
        if not data:
            messagebox.showwarning(
                "Нет данных", "Нижнее поле пустое — нечего выгружать."
            )
            return

        path = filedialog.asksaveasfilename(
            title="Сохранить как",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialfile="sgtin_export.xlsx",
        )
        if not path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "SGTIN"
            ws.append(["SGTIN"])
            for row in data:
                # Сохраняем как текст, чтобы Excel не превращал в число
                # и не терял ведущие нули.
                ws.append([str(row)])
            for cell in ws["A"]:
                cell.number_format = "@"
            wb.save(path)
        except OSError as exc:
            messagebox.showerror("Ошибка записи", str(exc))
            return

        self.status_var.set(
            f"Выгружено {len(data)} строк в {os.path.basename(path)}"
        )
        messagebox.showinfo(
            "Готово", f"Сохранено строк: {len(data)}\nФайл: {path}"
        )


def main() -> None:
    app = SgtinApp()
    app.mainloop()


if __name__ == "__main__":
    main()
