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


# Цветовая палитра (Tailwind-подобная).
BG = "#f1f5f9"          # фон окна — slate-100
CARD = "#ffffff"        # фон текстовых полей
BORDER = "#cbd5e1"      # рамка / линии
TEXT = "#0f172a"        # основной текст
MUTED = "#475569"       # вторичный текст / статус-бар
DANGER = "#dc2626"      # «УДАЛИТЬ» — красный
DANGER_HOVER = "#b91c1c"
SUCCESS = "#16a34a"     # «Добавить» — зелёный
SUCCESS_HOVER = "#15803d"
NEUTRAL = "#64748b"     # «Очистить» — slate
NEUTRAL_HOVER = "#475569"
PRIMARY = "#2563eb"     # «Excel» — синий
PRIMARY_HOVER = "#1d4ed8"
FONT_UI = ("Segoe UI", 10)
FONT_UI_BOLD = ("Segoe UI", 10, "bold")
FONT_HEADING = ("Segoe UI", 11, "bold")
FONT_MONO = ("Consolas", 11)


class SgtinApp(tk.Tk):
    """Главное окно приложения."""

    def __init__(self) -> None:
        super().__init__()
        self.title("SGTIN AI 01-21 — инструмент")
        self.geometry("1000x760")
        self.minsize(700, 520)
        self.configure(bg=BG)

        self._setup_style()
        self._build_ui()
        self._bind_shortcuts()

    # ------------------------------------------------------------------
    # Style
    # ------------------------------------------------------------------
    def _setup_style(self) -> None:
        style = ttk.Style(self)
        # «clam» позволяет красить ttk.Button фоном на всех платформах.
        if "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure(".", background=BG, foreground=TEXT, font=FONT_UI)
        style.configure("TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD, relief="flat")
        style.configure("TLabel", background=BG, foreground=TEXT, font=FONT_UI)
        style.configure(
            "Heading.TLabel", background=BG, foreground=TEXT, font=FONT_HEADING
        )
        style.configure(
            "Status.TLabel", background=BG, foreground=MUTED, font=FONT_UI
        )

        def _btn(name: str, bg: str, hover: str) -> None:
            style.configure(
                name,
                background=bg,
                foreground="white",
                bordercolor=bg,
                lightcolor=bg,
                darkcolor=bg,
                focuscolor=bg,
                padding=(14, 10),
                font=FONT_UI_BOLD,
                borderwidth=0,
                relief="flat",
            )
            style.map(
                name,
                background=[("active", hover), ("pressed", hover)],
                foreground=[("disabled", "#e2e8f0")],
            )

        _btn("Danger.TButton", DANGER, DANGER_HOVER)
        _btn("Success.TButton", SUCCESS, SUCCESS_HOVER)
        _btn("Neutral.TButton", NEUTRAL, NEUTRAL_HOVER)
        _btn("Primary.TButton", PRIMARY, PRIMARY_HOVER)

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=14, style="TFrame")
        root.pack(fill=tk.BOTH, expand=True)
        root.rowconfigure(1, weight=1)
        root.rowconfigure(4, weight=1)
        root.columnconfigure(0, weight=1)

        ttk.Label(
            root,
            text="Входные SGTIN (по одному в строке):",
            style="Heading.TLabel",
        ).grid(row=0, column=0, sticky="w", pady=(0, 6))
        self.input_text = self._make_text_area(root, row=1)

        middle = ttk.Frame(root, style="TFrame")
        middle.grid(row=2, column=0, sticky="ew", pady=14)
        for col in range(4):
            middle.columnconfigure(col, weight=1, uniform="btn")

        ttk.Button(
            middle,
            text="УДАЛИТЬ AI 01-21",
            command=self.on_remove,
            style="Danger.TButton",
            takefocus=False,
        ).grid(row=0, column=0, padx=(0, 6), sticky="ew")
        ttk.Button(
            middle,
            text="Добавить AI 01-21",
            command=self.on_add,
            style="Success.TButton",
            takefocus=False,
        ).grid(row=0, column=1, padx=6, sticky="ew")
        ttk.Button(
            middle,
            text="Очистить результат",
            command=self.on_clear_output,
            style="Neutral.TButton",
            takefocus=False,
        ).grid(row=0, column=2, padx=6, sticky="ew")
        ttk.Button(
            middle,
            text="Выгрузить в Excel",
            command=self.on_export_excel,
            style="Primary.TButton",
            takefocus=False,
        ).grid(row=0, column=3, padx=(6, 0), sticky="ew")

        ttk.Label(root, text="Результат:", style="Heading.TLabel").grid(
            row=3, column=0, sticky="w", pady=(0, 6)
        )
        self.output_text = self._make_text_area(root, row=4)

        self.status_var = tk.StringVar(value="Готово")
        ttk.Label(
            root,
            textvariable=self.status_var,
            anchor="w",
            style="Status.TLabel",
        ).grid(row=5, column=0, sticky="ew", pady=(10, 0))

    def _make_text_area(self, parent: ttk.Frame, row: int) -> tk.Text:
        frame = ttk.Frame(parent, style="Card.TFrame")
        frame.grid(row=row, column=0, sticky="nsew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        frame.configure(borderwidth=1, relief="solid")

        text = tk.Text(
            frame,
            wrap="none",
            undo=True,
            maxundo=-1,
            font=FONT_MONO,
            bg=CARD,
            fg=TEXT,
            insertbackground=TEXT,
            selectbackground=PRIMARY,
            selectforeground="white",
            relief="flat",
            borderwidth=0,
            padx=8,
            pady=6,
            highlightthickness=0,
        )
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
        # Маппинг действий: и латинская раскладка, и кириллическая (на тех же
        # клавишах). В Tk при русской раскладке keysym становится Cyrillic_*
        # и стандартные Control-a/c/v/x/z не срабатывают, поэтому
        # подвязываем оба варианта.
        select_all_keys = (
            "<Control-a>", "<Control-A>",
            "<Control-Cyrillic_ef>", "<Control-Cyrillic_EF>",  # Ф
        )
        copy_keys = (
            "<Control-c>", "<Control-C>",
            "<Control-Cyrillic_es>", "<Control-Cyrillic_ES>",  # С
            "<Control-Insert>",
        )
        paste_keys = (
            "<Control-v>", "<Control-V>",
            "<Control-Cyrillic_em>", "<Control-Cyrillic_EM>",  # М
            "<Shift-Insert>",
        )
        cut_keys = (
            "<Control-x>", "<Control-X>",
            "<Control-Cyrillic_che>", "<Control-Cyrillic_CHE>",  # Ч
            "<Shift-Delete>",
        )
        undo_keys = (
            "<Control-z>", "<Control-Z>",
            "<Control-Cyrillic_ya>", "<Control-Cyrillic_YA>",  # Я
        )
        redo_keys = (
            "<Control-y>", "<Control-Y>",
            "<Control-Cyrillic_en>", "<Control-Cyrillic_EN>",  # Н (на клавише Y)
            "<Control-Shift-Z>", "<Control-Shift-z>",
            "<Control-Shift-Cyrillic_ya>", "<Control-Shift-Cyrillic_YA>",
        )

        def safe_bind(widget: tk.Text, seqs: tuple, handler) -> None:
            for seq in seqs:
                try:
                    widget.bind(seq, handler)
                except tk.TclError:
                    # Платформа не знает такой keysym (например, Cyrillic_*
                    # на минимальной сборке Tk). Просто пропускаем.
                    pass

        copy_h = lambda e: (e.widget.event_generate("<<Copy>>"), "break")[1]
        paste_h = lambda e: (e.widget.event_generate("<<Paste>>"), "break")[1]
        cut_h = lambda e: (e.widget.event_generate("<<Cut>>"), "break")[1]

        for widget in (self.input_text, self.output_text):
            safe_bind(widget, select_all_keys, self._select_all)
            safe_bind(widget, copy_keys, copy_h)
            safe_bind(widget, paste_keys, paste_h)
            safe_bind(widget, cut_keys, cut_h)
            safe_bind(widget, undo_keys, self._undo)
            safe_bind(widget, redo_keys, self._redo)
            widget.bind("<Delete>", self._on_delete)
            widget.bind("<BackSpace>", self._on_backspace)

    @staticmethod
    def _select_all(event: tk.Event) -> str:
        widget: tk.Text = event.widget  # type: ignore[assignment]
        widget.tag_add("sel", "1.0", "end-1c")
        widget.mark_set("insert", "1.0")
        widget.see("insert")
        return "break"

    @staticmethod
    def _undo(event: tk.Event) -> str:
        widget: tk.Text = event.widget  # type: ignore[assignment]
        try:
            widget.edit_undo()
        except tk.TclError:
            pass
        return "break"

    @staticmethod
    def _redo(event: tk.Event) -> str:
        widget: tk.Text = event.widget  # type: ignore[assignment]
        try:
            widget.edit_redo()
        except tk.TclError:
            pass
        return "break"

    @staticmethod
    def _on_delete(event: tk.Event) -> str | None:
        widget: tk.Text = event.widget  # type: ignore[assignment]
        try:
            widget.delete("sel.first", "sel.last")
            return "break"
        except tk.TclError:
            return None  # нет выделения — пусть отработает стандартный Delete

    @staticmethod
    def _on_backspace(event: tk.Event) -> str | None:
        widget: tk.Text = event.widget  # type: ignore[assignment]
        try:
            widget.delete("sel.first", "sel.last")
            return "break"
        except tk.TclError:
            return None

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
