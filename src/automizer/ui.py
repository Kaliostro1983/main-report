# src/automizer/ui.py
from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Any

import pandas as pd

from src.automizer.conclusions import (
    load_conclusions,
    try_autopick_conclusion,
    render_template,
    ConclusionTemplate,
)
from src.armorkit.xlsxutils.writer import save_df_over_original


# константи з твого файлу
COL_TEXT = "р\\обмін"
COL_NOTES = "примітки"
COL_FREQ = "Частота"


class AutomizerApp:
    def __init__(
        self,
        intercepts_df: pd.DataFrame,
        reference_df: pd.DataFrame,
        report_path: str | Path,
        config_path: str | Path,
        conclusions_path: str | Path,
    ) -> None:
        self.df = intercepts_df
        self.ref_df = reference_df
        self.report_path = Path(report_path)
        self.config_path = Path(config_path)
        self.conclusions_path = Path(conclusions_path)

        # службовий стовпчик для "затверджено"
        if "__approved" not in self.df.columns:
            self.df["__approved"] = False

        self.templates: list[ConclusionTemplate] = load_conclusions(self.conclusions_path)

        self.index: int = 0  # поточний рядок

        # створення вікна
        self.root = tk.Tk()
        self.root.title("Automizer – перехоплення")
        self.root.geometry("1200x800")

        # головні елементи
        self._build_ui()

        # заповнюємо перший запис
        self._load_record(0)

        # гарячі клавіші
        self._bind_hotkeys()

    # ------------------------------
    # UI
    # ------------------------------
    def _build_ui(self) -> None:
        # ліва частина (текст перехоплення + поле приміток)
        left_frame = tk.Frame(self.root)
        left_frame.place(x=10, y=10, width=850, height=760)  # було height=980

        # поле перехоплення (велике)
        lbl_intercept = tk.Label(left_frame, text="Текст перехоплення", anchor="w")
        lbl_intercept.pack(fill="x")

        self.txt_intercept = tk.Text(left_frame, wrap="word", font=("Arial", 18))
        self.txt_intercept.configure(state="disabled")
        self.txt_intercept.pack(fill="both", expand=True, pady=(0, 10))

        # combobox з шаблонами
        bottom_frame = tk.Frame(left_frame)
        bottom_frame.pack(fill="x", pady=(0, 5))

        tk.Label(bottom_frame, text="Шаблонний висновок:").pack(side="left")

        self.cbo_templates = ttk.Combobox(
            bottom_frame,
            state="readonly",
            values=[t.name for t in self.templates],
        )
        self.cbo_templates.pack(side="left", fill="x", expand=True, padx=(5, 5))
        self.cbo_templates.bind("<<ComboboxSelected>>", self._on_template_selected)

        # поле приміток (куди пишемо остаточний висновок)
        lbl_notes = tk.Label(left_frame, text='Примітки (колонка "примітки")', anchor="w")
        lbl_notes.pack(fill="x")

        self.txt_notes = tk.Text(left_frame, height=4, wrap="word")
        self.txt_notes.pack(fill="x")

        # права частина
        right_frame = tk.Frame(self.root)
        right_frame.place(x=880, y=10, width=300, height=760)  # було height=980


        # індикатор X/XXX
        self.lbl_pos = tk.Label(right_frame, text="0/0", font=("Arial", 16))
        self.lbl_pos.pack(pady=(0, 20))

        # кнопка "Завантажити"
        self.btn_save = tk.Button(right_frame, text="Завантажити", command=self._on_save_clicked, height=2)
        self.btn_save.pack(fill="x", pady=(0, 10))

        # кнопка "Очистити"
        self.btn_clear = tk.Button(right_frame, text="Очистити", command=self._on_clear_clicked, height=2)
        self.btn_clear.pack(fill="x", pady=(0, 30))

        # кнопки навігації
        nav_frame = tk.Frame(right_frame)
        nav_frame.pack()

        self.btn_prev = tk.Button(nav_frame, text="<", width=8, command=self._on_prev)
        self.btn_prev.grid(row=0, column=0, padx=5)

        self.btn_next = tk.Button(nav_frame, text=">", width=8, command=self._on_next)
        self.btn_next.grid(row=0, column=1, padx=5)

        # невеликий напис з підказками
        tips = tk.Label(
            right_frame,
            text="←/→ – попередній/наступний\n↓ – очистити\n1,2,3… – вставити шаблон",
            justify="left",
        )
        tips.pack(pady=(30, 0), anchor="w")

    def _bind_hotkeys(self) -> None:
        self.root.bind("<Left>", lambda e: self._on_prev())
        self.root.bind("<Right>", lambda e: self._on_next())
        self.root.bind("<Down>", lambda e: self._on_clear_clicked())

        # гарячі клавіші з conclusions.json
        for tmpl in self.templates:
            if tmpl.shortcut:
                # наприклад, "1" → "<Key-1>"
                self.root.bind(f"<Key-{tmpl.shortcut}>", lambda e, t=tmpl: self._apply_template(t))

    # ------------------------------
    # Робота з поточним записом
    # ------------------------------
    def _load_record(self, index: int) -> None:
        total = len(self.df)
        if total == 0:
            return

        self.index = index % total  # навігація по колу

        row = self.df.iloc[self.index]

        # 1. показати текст перехоплення
        text_val = row.get(COL_TEXT, "")
        self.txt_intercept.configure(state="normal")
        self.txt_intercept.delete("1.0", tk.END)
        self.txt_intercept.insert(tk.END, "" if pd.isna(text_val) else str(text_val))
        self.txt_intercept.configure(state="disabled")

        # 2. показати примітки (якщо вже були)
        notes_val = row.get(COL_NOTES, "")
        self.txt_notes.delete("1.0", tk.END)
        if not (notes_val is None or (isinstance(notes_val, float) and pd.isna(notes_val))):
            self.txt_notes.insert(tk.END, str(notes_val))

        # 3. якщо не затверджено – спробувати автопідбір
        approved = bool(row.get("__approved", False))
        if not approved:
            candidate = try_autopick_conclusion(
                intercept_text=str(text_val) if text_val is not None else "",
                freq_value=row.get(COL_FREQ),
                reference_df=self.ref_df,
                templates=self.templates,
            )
            if candidate:
                self.txt_notes.delete("1.0", tk.END)
                self.txt_notes.insert(tk.END, candidate)

        # 4. оновити індикатор
        self.lbl_pos.config(text=f"{self.index + 1}/{total}")

    def _save_current_to_df(self) -> None:
        # беремо те, що в полі приміток, і пишемо в df
        notes_text = self.txt_notes.get("1.0", tk.END).strip()
        self.df.at[self.index, COL_NOTES] = notes_text
        # якщо є текст – вважаємо затвердженим
        self.df.at[self.index, "__approved"] = bool(notes_text)

    # ------------------------------
    # Обробники
    # ------------------------------
    def _on_prev(self) -> None:
        self._save_current_to_df()
        self._load_record(self.index - 1)

    def _on_next(self) -> None:
        self._save_current_to_df()
        self._load_record(self.index + 1)

    def _on_clear_clicked(self) -> None:
        self.txt_notes.delete("1.0", tk.END)
        # позначаємо як не затверджений
        self.df.at[self.index, "__approved"] = False

    def _on_save_clicked(self) -> None:
        # зберегти перед записом
        self._save_current_to_df()
        try:
            save_df_over_original(self.df, self.report_path)
            messagebox.showinfo("Готово", f"Файл оновлено:\n{self.report_path}")
        except Exception as ex:
            messagebox.showerror("Помилка", f"Не вдалось зберегти файл:\n{ex}")

    def _on_template_selected(self, event: Any) -> None:
        idx = self.cbo_templates.current()
        if idx < 0:
            return
        tmpl = self.templates[idx]
        self._apply_template(tmpl)

    def _apply_template(self, tmpl: ConclusionTemplate) -> None:
        # будуємо текст за шаблоном
        row = self.df.iloc[self.index]
        freq_val = row.get(COL_FREQ)
        text = render_template(tmpl, freq_val, self.ref_df)
        self.txt_notes.delete("1.0", tk.END)
        self.txt_notes.insert(tk.END, text)
        # відмічаємо як затверджене
        self.df.at[self.index, "__approved"] = True

    # ------------------------------
    def run(self) -> None:
        self.root.mainloop()
