# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import pandas as pd

# 1) використовуємо існуючі модулі з твого проєкту
from src.reportgen.settings import load_config  # читаємо config.yml (freq_file, reports_dir тощо)
from src.armorkit.data_loader import load_reference  # читаємо довідник XLSX
from src.armorkit.normalize_freq import (
    FREQ_NOT_FOUND,
    get_true_freq_by_mask,
    is_real_freq,
)
from .mgrs import is_valid_mgrs

import re

_space_re = re.compile(r"\s+")

def _sanitize_mgrs_line(line: str) -> str:
    """
    Приводить рядок до виду:
      <token0> <token1> <digits5> <digits5>
    Діє так:
      - trim, згортує кратні пробіли до одного,
      - не чіпає регістр цифро-буквених токенів (окрім того, що токени 0–1 можна аперкейснути за бажанням),
      - помилка лише якщо кількість цифр у двох останніх токенах не дорівнює 5 або вони не цифри.
    """
    s = (line or "").strip()
    if not s:
        raise ValueError("Порожній рядок")

    s = _space_re.sub(" ", s)
    parts = s.split(" ")

    if len(parts) < 4:
        raise ValueError("Неповний рядок (очікується 4+ токени)")

    # беремо останні 2 токени як цифри
    d1, d2 = parts[-2], parts[-1]
    if not (d1.isdigit() and d2.isdigit() and len(d1) == 5 and len(d2) == 5):
        raise ValueError("Цифрові блоки мають бути по 5 цифр")

    # реконструюємо рядок: перші 2 токени (зазвичай '37U' 'DQ') + 2 цифрових
    t0 = parts[0].upper()
    t1 = parts[1].upper()
    return f"{t0} {t1} {d1} {d2}"


FALLBACK_UNIT = "НВ підрозділу"
FALLBACK_LOC  = "ТОРСЬКЕ"

def fmt_date_now(): return datetime.now().strftime("%d.%m.%Y")
def fmt_time_now(): return datetime.now().strftime("%H.%M")

def _norm3(s: str) -> str:
    return f"{float(str(s).replace(',', '.')):.3f}"

def _norm4(s: str) -> str:
    return f"{float(str(s).replace(',', '.')):.4f}"

def _resolve_unit_and_location(freq4: str, reference_df: pd.DataFrame) -> tuple[str, str]:
    """
    Пробуємо витягти Підрозділ + Зона функціонування з довідника.
    Якщо не знайдено — фолбеки.
    """
    if reference_df is None or reference_df.empty:
        return FALLBACK_UNIT, FALLBACK_LOC

    df = reference_df.copy()
    # частоту в довіднику перетворимо у 4 знаки після коми для зіставлення
    def _to4(x):
        try:
            return f"{float(str(x).replace(',', '.')):.4f}"
        except Exception:
            return None

    df["__f4"] = df["Частота"].map(_to4) if "Частота" in df.columns else None
    hit = df[df["__f4"] == freq4] if "__f4" in df.columns else df.iloc[0:0]

    if hit.empty:
        return FALLBACK_UNIT, FALLBACK_LOC

    row = hit.iloc[0]
    unit = str(row.get("Підрозділ", "")).strip() or FALLBACK_UNIT
    loc  = str(row.get("Зона функціонування", "")).strip() or FALLBACK_LOC
    return unit, loc

class Toast(ttk.Frame):
    def __init__(self, master, text: str, ms: int = 1200):
        super().__init__(master)
        ttk.Label(self, text=text, foreground="#0a5").pack(padx=8, pady=4)
        self.after(ms, self.destroy)

class App(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=12)
        self.pack(fill="both", expand=True)

        # 1) Конфіг + довідник
        cfg = load_config("config.yml")                 # шляхи беремо звідти
        self.reference_df = load_reference(cfg.paths.freq_file)  # XLSX у DataFrame

        # 2) Побудова UI за твоїм ескізом
        self.date = tk.StringVar(value=fmt_date_now())
        self.time = tk.StringVar(value=fmt_time_now())
        self.freq = tk.StringVar()  # введення маски/частоти
        self.unit = tk.StringVar(value=FALLBACK_UNIT)
        self.location = tk.StringVar(value=FALLBACK_LOC)

        # ряд 1: дата, час, C
        r1 = ttk.Frame(self); r1.pack(fill="x")
        ttk.Label(r1, text="Дата").pack(side="left")
        ttk.Entry(r1, textvariable=self.date, width=12).pack(side="left", padx=(4,12))
        ttk.Label(r1, text="Час").pack(side="left")
        ttk.Entry(r1, textvariable=self.time, width=8).pack(side="left", padx=(4,12))
        ttk.Button(r1, text="C", width=3, command=lambda: (self.date.set(fmt_date_now()), self.time.set(fmt_time_now()))).pack(side="left")

        # ряд 2: частота/маска + Прийняти
        r2 = ttk.Frame(self); r2.pack(fill="x", pady=(10,0))
        ttk.Label(r2, text="Частота/Маска").pack(side="left")
        ttk.Entry(r2, textvariable=self.freq, width=12).pack(side="left", padx=(4,12))
        ttk.Button(r2, text="Прийняти", command=self.accept_freq).pack(side="left")

        # ряд 3: unit + location
        r3 = ttk.Frame(self); r3.pack(fill="x", pady=(10,0))
        ttk.Label(r3, text="Підрозділ").pack(side="left")
        ttk.Entry(r3, textvariable=self.unit).pack(side="left", fill="x", expand=True, padx=(4,12))
        ttk.Label(r3, text="Location").pack(side="left")
        ttk.Entry(r3, textvariable=self.location, width=28).pack(side="left")

        # MGRS
        ttk.Label(self, text="MGRS coordinates (кожен рядок окремо)").pack(anchor="w", pady=(10,0))
        self.coords = tk.Text(self, height=6, wrap="none")
        self.coords.pack(fill="both", expand=True)

        # Коментар
        ttk.Label(self, text="Коментар / банер").pack(anchor="w", pady=(10,0))
        self.comment = tk.Text(self, height=3, wrap="word")
        self.comment.insert("1.0", "-------  🦁 63 ОМБр 🦁 -------")
        self.comment.pack(fill="x")

        # Вихідний текст
        ttk.Label(self, text="Згенероване повідомлення").pack(anchor="w", pady=(10,0))
        self.output = tk.Text(self, height=7, wrap="word")
        self.output.pack(fill="both", expand=True)

        # Кнопки
        btns = ttk.Frame(self); btns.pack(fill="x", pady=8)
        ttk.Button(btns, text="Згенерувати", command=self.generate).pack(side="left")
        ttk.Button(btns, text="Копіювати", command=self.copy_output).pack(side="left", padx=8)
        ttk.Button(btns, text="Вихід", command=self.master.destroy).pack(side="right")
        
    def _install_clipboard_shortcuts(self):
        """Глобальні шорткати для Copy/Cut/Paste в активне поле."""
        def _gen(seq):
            def handler(event=None):
                w = self.master.focus_get()
                if w: 
                    try: w.event_generate(seq)
                    except Exception: pass
                return "break"
            return handler

        # Вставка
        for seq in ("<Control-v>", "<Control-V>", "<Shift-Insert>", "<Command-v>"):
            self.master.bind_all(seq, _gen("<<Paste>>"), add="+")
        # Копіювання
        for seq in ("<Control-c>", "<Control-C>", "<Command-c>"):
            self.master.bind_all(seq, _gen("<<Copy>>"), add="+")
        # Вирізати
        for seq in ("<Control-x>", "<Control-X>", "<Shift-Delete>", "<Command-x>"):
            self.master.bind_all(seq, _gen("<<Cut>>"), add="+")
        

    def _install_context_menu(self):
        """Правий клік: контекстне меню Cut/Copy/Paste для Entry/Text."""
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Вирізати", command=lambda: self._ctx_action("<<Cut>>"))
        menu.add_command(label="Копіювати", command=lambda: self._ctx_action("<<Copy>>"))
        menu.add_command(label="Вставити", command=lambda: self._ctx_action("<<Paste>>"))

        def show_menu(event):
            w = event.widget
            if isinstance(w, (tk.Entry, tk.Text, ttk.Entry)):
                menu.tk.call("tk_popup", menu, event.x_root, event.y_root)

        self.master.bind_all("<Button-3>", show_menu, add="+")   # Windows/Linux
        self.master.bind_all("<Control-Button-1>", show_menu, add="+")  # macOS (альтернатива)

    def _ctx_action(self, seq):
        w = self.master.focus_get()
        if w:
            try: w.event_generate(seq)
            except Exception: pass


    def accept_freq(self):
        """
        Якщо введено реальну частоту — нормалізуємо до 4 знаків і шукаємо по 'Частота'.
        Якщо введено маску — нормалізуємо маску до 3 знаків, шукаємо справжню частоту,
        але в полі лишаємо МАСКУ (вона піде у фінальний текст).
        """
        raw = (self.freq.get() or "").strip()
        if not raw:
            messagebox.showinfo("Інфо", "Введіть частоту або маску.")
            return

        # Реальна частота?
        if is_real_freq(raw):
            try:
                freq4 = _norm4(raw)                  # нормалізуємо частоту
            except Exception:
                messagebox.showwarning("Помилка", "Невірний формат частоти.")
                return

            # підставляємо саме ЧАСТОТУ в поле (логічно для real freq)
            self.freq.set(freq4)

            # підтягнути unit/location
            unit, loc = _resolve_unit_and_location(freq4, self.reference_df)
            self.unit.set(unit or FALLBACK_UNIT)
            self.location.set(loc or FALLBACK_LOC)
            return

        # Інакше — це МАСКА
        try:
            mask3 = _norm3(raw)                      # нормалізуємо маску
        except Exception:
            messagebox.showwarning("Помилка", "Невірний формат маски.")
            return

        # шукаємо справжню частоту за маскою (щоб підставити unit/location)
        true_f = get_true_freq_by_mask(mask3, self.reference_df)
        if true_f != FREQ_NOT_FOUND:
            try:
                freq4 = _norm4(true_f)
            except Exception:
                freq4 = None
        else:
            freq4 = None

        # УВАГА: у полі лишаємо МАСКУ (а не частоту)
        self.freq.set(mask3)

        # unit/location за знайденою частотою (якщо є), інакше фолбеки
        if freq4:
            unit, loc = _resolve_unit_and_location(freq4, self.reference_df)
            self.unit.set(unit or FALLBACK_UNIT)
            self.location.set(loc or FALLBACK_LOC)
        else:
            self.unit.set(FALLBACK_UNIT)
            self.location.set(FALLBACK_LOC)
            

    def generate(self):
        # 1) санітизуємо/перевіряємо MGRS
        raw_lines = [ln for ln in self.coords.get("1.0", "end").splitlines() if ln.strip()]
        lines = []
        bad_idx = []
        for i, ln in enumerate(raw_lines, 1):
            try:
                lines.append(_sanitize_mgrs_line(ln))
            except Exception:
                bad_idx.append(i)

        if bad_idx:
            messagebox.showwarning(
                "MGRS",
                f"Невірний формат цифр у рядках: {bad_idx}. Очікується два блоки по 5 цифр наприкінці."
            )
            return

        # 2) базові поля
        freq_or_mask = (self.freq.get() or "").strip()          # важливо: це може бути МАСКА!
        date = (self.date.get() or "").strip()
        time = (self.time.get() or "").strip()
        unit = (self.unit.get() or FALLBACK_UNIT).strip()
        loc  = (self.location.get() or FALLBACK_LOC).strip()
        comment = self.comment.get("1.0", "end").strip()

        if not all([freq_or_mask, date, time, unit, loc]):
            messagebox.showwarning("Увага", "Заповніть частоту/маску, дату, час, підрозділ і location.")
            return

        # 3) формування повідомлення — у ПЕРШОМУ РЯДКУ тепер може бути маска
        desc = f"УКХ р/м {unit} ({loc})"
        out_lines = [f"{freq_or_mask} / {date} {time}", f"{desc}", *lines]
        if comment:
            out_lines.append(comment)
        msg = "\n".join(out_lines)

        # 4) показ і копіювання
        self.output.delete("1.0", "end")
        self.output.insert("1.0", msg)
        self.master.clipboard_clear()
        self.master.clipboard_append(msg)
        Toast(self, "Скопійовано у буфер обміну", 1200).place(relx=0.5, rely=0.0, anchor="n")


    def copy_output(self):
        txt = self.output.get("1.0", "end").strip()
        if not txt:
            messagebox.showinfo("Інфо", "Спершу згенеруйте повідомлення."); return
        self.master.clipboard_clear(); self.master.clipboard_append(txt)
        Toast(self, "Скопійовано", 900).place(relx=0.5, rely=0.0, anchor="n")

def main():
    root = tk.Tk()
    root.title("peleng-gen • Формувач повідомлення")
    root.geometry("820x720")
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
