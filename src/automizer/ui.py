# src/automizer/ui.py
from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import font as tkfont
from pathlib import Path
from typing import Any, List

from src.automizer.conclusions import (
    load_conclusions, ConclusionTemplate,
    STATUS_EMPTY, STATUS_NEED_APPROVE, STATUS_APPROVED,
    COL_STATUS, COL_NOTES, COL_MATCHED_TEMPLATE, COL_MATCHED_WORD, COL_MULTI_MATCH,
    apply_autopick_to_df,
    find_template_candidates,   # ← ДОДАТИ ОЦЕ
)


import pandas as pd

# from src.automizer.conclusions import (
#     load_conclusions,
#     ConclusionTemplate,
#     STATUS_EMPTY,
#     STATUS_NEED_APPROVE,
#     STATUS_APPROVED,
#     COL_STATUS,
#     COL_NOTES,
#     COL_MATCHED_TEMPLATE,
#     COL_MATCHED_WORD,
#     COL_MULTI_MATCH,
#     apply_autopick_to_df,
# )
from src.automizer.freqdict import COL_FREQ
from src.automizer.service import initialize_data

# ---------- Допоміжні утиліти для збереження ----------
def save_df_over_original(df: pd.DataFrame, path: Path) -> None:
    df.to_excel(path, index=False)


class AutomizerApp:
    # Єдина назва колонки з текстом перехоплення (бекслеш треба екранувати)
    TEXT_COL = "р\\обмін"

    @staticmethod
    def row_text_safe(row: pd.Series, text_col: str) -> str:
        """Взяти текст із заданої колонки; якщо порожньо — найдовше строкове поле."""
        val = row.get(text_col, "")
        s = ("" if pd.isna(val) else str(val)).strip()
        if s:
            return s
        best = ""
        for v in row.values:
            if isinstance(v, str) and len(v) > len(best):
                best = v
        return best

    def __init__(
        self,
        *,
        intercepts_df: pd.DataFrame,
        reference_df: pd.DataFrame,
        report_path: str | Path,
        config_path: str | Path,
        conclusions_path: str | Path,
        templates: List[ConclusionTemplate] | None = None,
    ) -> None:
        # Дані
        self.df = intercepts_df
        self.ref_df = reference_df
        self.report_path = Path(report_path)
        self.config_path = Path(config_path)
        self.conclusions_path = Path(conclusions_path)

        # Створюємо кореневе вікно РАНІШЕ за будь-які Tk-змінні
        self.root = tk.Tk()
        self.root.title("Automizer – перехоплення")
        self.root.geometry("1100x700")

        # Зафіксована назва текстової колонки
        self.text_col = self.TEXT_COL

        # Службові колонки (на випадок, якщо ще не додали)
        if COL_STATUS not in self.df.columns:
            self.df[COL_STATUS] = STATUS_EMPTY
        if COL_NOTES not in self.df.columns:
            self.df[COL_NOTES] = ""
        if COL_MATCHED_TEMPLATE not in self.df.columns:
            self.df[COL_MATCHED_TEMPLATE] = ""
        if COL_MATCHED_WORD not in self.df.columns:
            self.df[COL_MATCHED_WORD] = ""
        if COL_MULTI_MATCH not in self.df.columns:
            self.df[COL_MULTI_MATCH] = False

        # Шаблони
        self.templates: list[ConclusionTemplate] = templates or load_conclusions(self.conclusions_path)

        # Фільтри статусів (за промовчанням показуємо все)
        self.filter_show_empty = tk.BooleanVar(self.root, value=True)
        self.filter_show_need = tk.BooleanVar(self.root, value=True)
        self.filter_show_approved = tk.BooleanVar(self.root, value=True)

        # Індексація видимих рядків за фільтрами
        self.visible_indices: list[int] = []
        self.index_pos: int = 0  # позиція в self.visible_indices (НЕ глобальний індекс df)

        # UI
        self._build_ui()
        self._bind_hotkeys()

        # Початкове заповнення списку видимих
        self._rebuild_visible_indices()
        self._load_record()

    # ---------- UI ----------
    def _build_ui(self) -> None:
        # ----- Лейаут вікна: 2 колонки (ліва розтягується, права фіксована) -----
        self.root.grid_columnconfigure(0, weight=1)              # LEFT: тягнеться
        self.root.grid_columnconfigure(1, weight=0, minsize=300) # RIGHT: фікс. ширина ~300
        self.root.grid_rowconfigure(0, weight=1)

        # ----- Ліва панель -----
        left = tk.Frame(self.root)
        left.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        # Велике поле тексту перехоплення
        self.txt_intercept = tk.Text(left, wrap="word", height=26, state="disabled")
        self.txt_intercept.pack(fill=tk.BOTH, expand=True)

        # Збільшити шрифт на 50%
        from tkinter import font as tkfont
        base_font = tkfont.Font(font=self.txt_intercept["font"])
        new_size = max(int(int(base_font["size"]) * 1.5), int(base_font["size"]) + 1)
        base_font.configure(size=new_size)
        self.txt_intercept.configure(font=base_font)

        # Поле "Примітки" (висновок)
        block2 = tk.Frame(left)
        block2.pack(fill=tk.X, pady=(8, 0))
        lbl_notes = tk.Label(block2, text='Примітки (колонка "примітки")')
        lbl_notes.pack(anchor="w")
        self.txt_notes = tk.Text(block2, height=5)
        self.txt_notes.pack(fill=tk.X)

        # ----- Права панель (фіксована ширина) -----
        right = tk.Frame(self.root, width=300)
        right.grid(row=0, column=1, sticky="ns", padx=(0, 10), pady=10)
        right.grid_propagate(False)  # контент не впливає на ширину

        # Позиція
        self.lbl_pos = tk.Label(right, text="0/0", font=("Segoe UI", 11, "bold"))
        self.lbl_pos.pack(pady=(0, 8), anchor="w")

        # Кнопки керування
        btn_frame = tk.Frame(right)
        btn_frame.pack(pady=(0, 8), fill=tk.X)
        self.btn_prev = tk.Button(btn_frame, text="<", width=6, command=self._on_prev)
        self.btn_prev.pack(side=tk.LEFT, padx=(0, 6))
        self.btn_next = tk.Button(btn_frame, text=">", width=6, command=self._on_next)
        self.btn_next.pack(side=tk.LEFT)

        # Зберегти / Оновити
        self.btn_save = tk.Button(right, text="Зберегти", command=self._on_save_clicked)
        self.btn_save.pack(pady=(4, 10), fill=tk.X)
        self.btn_update = tk.Button(right, text="Оновити", command=self._on_update_clicked)
        self.btn_update.pack(pady=(0, 10), fill=tk.X)

        # Фільтри статусів
        tk.Label(right, text="Показувати зі статусом:").pack(anchor="w")
        self.chk_empty = tk.Checkbutton(right, text="empty", variable=self.filter_show_empty, command=self._on_filters_changed)
        self.chk_need = tk.Checkbutton(right, text="need_approve", variable=self.filter_show_need, command=self._on_filters_changed)
        self.chk_ok   = tk.Checkbutton(right, text="approved", variable=self.filter_show_approved, command=self._on_filters_changed)
        self.chk_empty.pack(anchor="w")
        self.chk_need.pack(anchor="w")
        self.chk_ok.pack(anchor="w")

        # Підказка/статус — під чекбоксами, багаторядкова
        self.lbl_auto = tk.Label(right, text="", justify="left", fg="#555", wraplength=280)
        self.lbl_auto.pack(pady=(8, 0), anchor="w")



    def _bind_hotkeys(self) -> None:
        # Enter — зберегти + авто-approve (за правилом) + далі
        self.root.bind("<Return>", lambda e: (self._save_current_to_df(), self._auto_approve_if_needed(), self._on_next()))
        # Ctrl+S — зберегти у файл
        self.root.bind("<Control-s>", lambda e: self._on_save_clicked())
        self.root.bind("<Control-S>", lambda e: self._on_save_clicked())
        # Стрілки ← / → — попередній / наступний (працює навіть коли фокус у полі приміток)
        self.root.bind_all("<Left>", lambda e: self._on_prev())
        self.root.bind_all("<Right>", lambda e: self._on_next())


    # ---------- Навігація та індексація ----------
    def _global_index(self) -> int:
        if not self.visible_indices:
            return -1
        return self.visible_indices[self.index_pos]

    def _rebuild_visible_indices(self) -> None:
        vis: list[int] = []
        for i, st in enumerate(self.df[COL_STATUS].tolist()):
            if (st == STATUS_EMPTY and self.filter_show_empty.get()) \
               or (st == STATUS_NEED_APPROVE and self.filter_show_need.get()) \
               or (st == STATUS_APPROVED and self.filter_show_approved.get()):
                vis.append(i)
        self.visible_indices = vis
        # Підправляємо позицію
        if not self.visible_indices:
            self.index_pos = 0
        else:
            self.index_pos = min(self.index_pos, len(self.visible_indices) - 1)

    def _load_record(self) -> None:
        gi = self._global_index()
        if gi == -1:
            self.lbl_pos.config(text="0/0")
            self.txt_intercept.config(state="normal")
            self.txt_intercept.delete("1.0", tk.END)
            self.txt_intercept.insert(tk.END, "(Немає записів у відфільтрованому наборі)")
            self.txt_intercept.config(state="disabled")
            self.txt_notes.delete("1.0", tk.END)
            self.lbl_auto.config(text="")
            self._update_nav_state()
            return

        row = self.df.iloc[gi]
        total = len(self.visible_indices)
        self.lbl_pos.config(text=f"{self.index_pos + 1}/{total}")

        # Текст перехоплення
        text_val = self.row_text_safe(row, self.text_col)
        self.txt_intercept.config(state="normal")
        self.txt_intercept.delete("1.0", tk.END)
        self.txt_intercept.insert(tk.END, text_val)
        self.txt_intercept.config(state="disabled")

        # Примітки
        # у _load_record() перед insert у self.txt_notes:
        val = row.get(COL_NOTES, "")
        notes = "" if (pd.isna(val) if hasattr(pd, "isna") else val is None) else str(val)
        self.txt_notes.delete("1.0", tk.END)
        self.txt_notes.insert(tk.END, notes)


        # Усі збіги для цього тексту
        cands = find_template_candidates(text_val, self.templates)  # [(tmpl, rule), ...]
        mt = str(row.get(COL_MATCHED_TEMPLATE, "") or "")
        mw = str(row.get(COL_MATCHED_WORD, "") or "")
        st = str(row.get(COL_STATUS, "") or STATUS_EMPTY)
        multi = bool(row.get(COL_MULTI_MATCH, False))

        lines = [f"Статус: {st}"]
        if mt:
            lines.append(f'Обрано: "{mt}" ← "{mw}"')
        if cands:
            if multi and len(cands) > 1:
                lines.append("(було кілька варіантів)")
            lines.append("Усі збіги:")
            for i, (tmpl, rule) in enumerate(cands, start=1):
                star = "★" if i == 1 else " "
                p = getattr(rule, "probability", 0)
                lines.append(f'{star} {i}. "{tmpl.name}" ← "{rule.word}" (p={p})')
        else:
            lines.append("Збігів не знайдено")

        self.lbl_auto.config(text="\n".join(lines))

        self._update_nav_state()


    def _update_nav_state(self) -> None:
        has = len(self.visible_indices) > 0
        self.btn_prev.config(state=("normal" if has else "disabled"))
        self.btn_next.config(state=("normal" if has else "disabled"))
        self.btn_save.config(state=("normal" if has else "disabled"))

    # ---------- Події ----------
    def _on_prev(self) -> None:
        self._save_current_to_df()
        self._auto_approve_if_needed()
        if self.visible_indices:
            self.index_pos = max(0, self.index_pos - 1)
        self._load_record()

    def _on_next(self) -> None:
        self._save_current_to_df()
        self._auto_approve_if_needed()
        if self.visible_indices:
            self.index_pos = min(len(self.visible_indices) - 1, self.index_pos + 1)
        self._load_record()

    def _on_save_clicked(self) -> None:
        self._save_current_to_df()
        try:
            save_df_over_original(self.df, self.report_path)
            messagebox.showinfo("Готово", f"Файл оновлено:\n{self.report_path}")
        except Exception as ex:
            messagebox.showerror("Помилка", f"Не вдалось зберегти файл:\n{ex}")

    def _on_update_clicked(self) -> None:
        # Перечитуємо дані з диска, автопідбір виконується в initialize_data
        data = initialize_data(
            config_path=self.config_path,
            conclusions_path=self.conclusions_path,
            skip_approved_on_reload=True,  # не чіпати вже затверджені
        )
        self.df = data["df"]
        self.ref_df = data["ref_df"]
        self.templates = data["templates"]
        # Перебудувати видимі за фільтрами
        self._rebuild_visible_indices()
        self._load_record()
        messagebox.showinfo("Готово", "Дані оновлено.")

    def _on_filters_changed(self) -> None:
        self._rebuild_visible_indices()
        self._load_record()

    # ---------- Синхронізація UI -> df ----------
    def _save_current_to_df(self) -> None:
        gi = self._global_index()
        if gi == -1:
            return
        notes = self.txt_notes.get("1.0", tk.END).strip()
        self.df.at[gi, COL_NOTES] = notes or ""   # гарантовано не NaN

    def _auto_approve_if_needed(self) -> None:
        """
        Якщо поточний рядок мав статус need_approve і оператор пішов далі —
        встановлюємо approved (ЛИШЕ якщо у нотатках є хоч щось).
        """
        gi = self._global_index()
        if gi == -1:
            return
        st = self.df.at[gi, COL_STATUS]
        if st == STATUS_NEED_APPROVE:
            notes = str(self.df.at[gi, COL_NOTES] or "").strip()
            if notes:
                self.df.at[gi, COL_STATUS] = STATUS_APPROVED

    # ---------- Публічний запуск ----------
    def run(self) -> None:
        self.root.mainloop()
