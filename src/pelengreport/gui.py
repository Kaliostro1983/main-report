# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from pathlib import Path
import json
import re
import sys

import pandas as pd

from src.armorkit.settings import load_config
from src.armorkit.data_loader import load_reference
from src.armorkit.normalize_freq import (
    FREQ_NOT_FOUND,
    get_true_freq_by_mask,
    is_real_freq,
)

from .mgrs import is_valid_mgrs  # залишаю як було (може знадобитись)
from .runner import run as run_report  # генерує DOCX + відкриває папку/файл

_space_re = re.compile(r"\s+")

FALLBACK_UNIT = "НВ підрозділу"
FALLBACK_LOC  = "ТОРСЬКЕ"


def fmt_date_now() -> str:
    return datetime.now().strftime("%d.%m.%Y")


def fmt_time_now() -> str:
    return datetime.now().strftime("%H.%M")


def _norm3(s: str) -> str:
    return f"{float(str(s).replace(',', '.')):.3f}"


def _norm4(s: str) -> str:
    return f"{float(str(s).replace(',', '.')):.4f}"


def _sanitize_mgrs_line(line: str) -> str:
    s = (line or "").strip()
    if not s:
        raise ValueError("Порожній рядок")
    s = _space_re.sub(" ", s)
    parts = s.split(" ")
    if len(parts) < 4:
        raise ValueError("Неповний рядок (очікується 4+ токени)")
    d1, d2 = parts[-2], parts[-1]
    if not (d1.isdigit() and d2.isdigit() and len(d1) == 5 and len(d2) == 5):
        raise ValueError("Цифрові блоки мають бути по 5 цифр")
    t0 = parts[0].upper()
    t1 = parts[1].upper()
    return f"{t0} {t1} {d1} {d2}"


def _resolve_unit_and_location(freq4: str, reference_df: pd.DataFrame) -> tuple[str, str]:
    if reference_df is None or reference_df.empty:
        return FALLBACK_UNIT, FALLBACK_LOC

    df = reference_df.copy()

    def _to4(x):
        try:
            return f"{float(str(x).replace(',', '.')):.4f}"
        except Exception:
            return None

    if "Частота" not in df.columns:
        return FALLBACK_UNIT, FALLBACK_LOC

    df["__f4"] = df["Частота"].map(_to4)
    hit = df[df["__f4"] == freq4]

    if hit.empty:
        return FALLBACK_UNIT, FALLBACK_LOC

    row = hit.iloc[0]
    unit = str(row.get("Підрозділ", "")).strip() or FALLBACK_UNIT
    loc  = str(row.get("Зона функціонування", "")).strip() or FALLBACK_LOC
    return unit, loc


# -------------------- Posts store --------------------
def _repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _default_posts() -> list[dict]:
    return [
        {
            "id": "bp0000",
            "active": True,
            "name": "МІКОЛАЇВКА, БП №0000",
            "unit": "А3719\n(63 омбр)",
            "place": "МІКОЛАЇВКА,\nБП №0000",
            "equipment": "“Пластун”",
            "task_text": "Відповідно до плану бойового застосування",
            "match": "МІКОЛАЇВКА",
        },
        {
            "id": "bp0001",
            "active": True,
            "name": "МАЯКИ, БП №0001",
            "unit": "А3719\n(63 омбр)",
            "place": "МАЯКИ,\nБП №0001",
            "equipment": "“Пластун”",
            "task_text": "Відповідно до плану бойового застосування",
            "match": "МАЯКИ",
        },
    ]


def load_posts(posts_path: Path) -> list[dict]:
    posts_path.parent.mkdir(parents=True, exist_ok=True)
    if not posts_path.exists():
        posts = _default_posts()
        posts_path.write_text(json.dumps(posts, ensure_ascii=False, indent=2), encoding="utf-8")
        return posts
    return json.loads(posts_path.read_text(encoding="utf-8"))


def save_posts(posts_path: Path, posts: list[dict]) -> None:
    posts_path.write_text(json.dumps(posts, ensure_ascii=False, indent=2), encoding="utf-8")


# -------------------- UI --------------------
class Toast(ttk.Frame):
    def __init__(self, master, text: str, ms: int = 1200):
        super().__init__(master)
        ttk.Label(self, text=text).pack(padx=8, pady=4)
        self.after(ms, self.destroy)


class PelengTab(ttk.Frame):
    def __init__(self, master, reference_df: pd.DataFrame):
        super().__init__(master, padding=12)
        self.reference_df = reference_df

        self.date = tk.StringVar(value=fmt_date_now())
        self.time = tk.StringVar(value=fmt_time_now())
        self.freq = tk.StringVar()
        self.unit = tk.StringVar(value=FALLBACK_UNIT)
        self.location = tk.StringVar(value=FALLBACK_LOC)

        r1 = ttk.Frame(self); r1.pack(fill="x")
        ttk.Label(r1, text="Дата").pack(side="left")
        ttk.Entry(r1, textvariable=self.date, width=12).pack(side="left", padx=(4,12))
        ttk.Label(r1, text="Час").pack(side="left")
        ttk.Entry(r1, textvariable=self.time, width=8).pack(side="left", padx=(4,12))
        ttk.Button(r1, text="C", width=3, command=lambda: (self.date.set(fmt_date_now()), self.time.set(fmt_time_now()))).pack(side="left")

        r2 = ttk.Frame(self); r2.pack(fill="x", pady=(10,0))
        ttk.Label(r2, text="Частота/Маска").pack(side="left")
        ttk.Entry(r2, textvariable=self.freq, width=12).pack(side="left", padx=(4,12))
        ttk.Button(r2, text="Прийняти", command=self.accept_freq).pack(side="left")

        r3 = ttk.Frame(self); r3.pack(fill="x", pady=(10,0))
        ttk.Label(r3, text="Підрозділ").pack(side="left")
        ttk.Entry(r3, textvariable=self.unit).pack(side="left", fill="x", expand=True, padx=(4,12))
        ttk.Label(r3, text="Location").pack(side="left")
        ttk.Entry(r3, textvariable=self.location, width=28).pack(side="left")

        ttk.Label(self, text="MGRS coordinates (кожен рядок окремо)").pack(anchor="w", pady=(10,0))
        self.coords = tk.Text(self, height=6, wrap="none")
        self.coords.pack(fill="both", expand=True)

        ttk.Label(self, text="Коментар / банер").pack(anchor="w", pady=(10,0))
        self.comment = tk.Text(self, height=3, wrap="word")
        self.comment.insert("1.0", "-------  🦁 63 ОМБр 🦁 -------")
        self.comment.pack(fill="x")

        ttk.Label(self, text="Згенероване повідомлення").pack(anchor="w", pady=(10,0))
        self.output = tk.Text(self, height=7, wrap="word")
        self.output.pack(fill="both", expand=True)

        btns = ttk.Frame(self); btns.pack(fill="x", pady=8)
        ttk.Button(btns, text="Згенерувати", command=self.generate).pack(side="left")
        ttk.Button(btns, text="Копіювати", command=self.copy_output).pack(side="left", padx=8)

    def accept_freq(self):
        raw = (self.freq.get() or "").strip()
        if not raw:
            messagebox.showinfo("Інфо", "Введіть частоту або маску.")
            return

        if is_real_freq(raw):
            try:
                freq4 = _norm4(raw)
            except Exception:
                messagebox.showwarning("Помилка", "Невірний формат частоти.")
                return

            self.freq.set(freq4)
            unit, loc = _resolve_unit_and_location(freq4, self.reference_df)
            self.unit.set(unit or FALLBACK_UNIT)
            self.location.set(loc or FALLBACK_LOC)
            return

        try:
            mask3 = _norm3(raw)
        except Exception:
            messagebox.showwarning("Помилка", "Невірний формат маски.")
            return

        true_f = get_true_freq_by_mask(mask3, self.reference_df)
        freq4 = None
        if true_f != FREQ_NOT_FOUND:
            try:
                freq4 = _norm4(true_f)
            except Exception:
                freq4 = None

        self.freq.set(mask3)

        if freq4:
            unit, loc = _resolve_unit_and_location(freq4, self.reference_df)
            self.unit.set(unit or FALLBACK_UNIT)
            self.location.set(loc or FALLBACK_LOC)
        else:
            self.unit.set(FALLBACK_UNIT)
            self.location.set(FALLBACK_LOC)

    def generate(self):
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

        freq_or_mask = (self.freq.get() or "").strip()
        date = (self.date.get() or "").strip()
        time = (self.time.get() or "").strip()
        unit = (self.unit.get() or FALLBACK_UNIT).strip()
        loc  = (self.location.get() or FALLBACK_LOC).strip()
        comment = self.comment.get("1.0", "end").strip()

        if not all([freq_or_mask, date, time, unit, loc]):
            messagebox.showwarning("Увага", "Заповніть частоту/маску, дату, час, підрозділ і location.")
            return

        desc = f"УКХ р/м {unit} ({loc})"
        out_lines = [f"{freq_or_mask} / {date} {time}", f"{desc}", *lines]
        if comment:
            out_lines.append(comment)
        msg = "\n".join(out_lines)

        self.output.delete("1.0", "end")
        self.output.insert("1.0", msg)
        self.clipboard_copy(msg)

        Toast(self, "Скопійовано у буфер обміну", 1200).place(relx=0.5, rely=0.0, anchor="n")

    def clipboard_copy(self, text: str) -> None:
        self.clipboard_clear()
        self.clipboard_append(text)

    def copy_output(self):
        txt = self.output.get("1.0", "end").strip()
        if not txt:
            messagebox.showinfo("Інфо", "Спершу згенеруйте повідомлення.")
            return
        self.clipboard_copy(txt)
        Toast(self, "Скопійовано", 900).place(relx=0.5, rely=0.0, anchor="n")


class ReportTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=12)

        self.posts_path = _repo_root() / "posts.json"
        self.posts: list[dict] = load_posts(self.posts_path)

        # input file row
        r0 = ttk.Frame(self); r0.pack(fill="x")
        ttk.Label(r0, text="Вхідний txt").pack(side="left")
        self.input_path_var = tk.StringVar(value="")
        ttk.Entry(r0, textvariable=self.input_path_var).pack(side="left", fill="x", expand=True, padx=(8,8))
        ttk.Button(r0, text="Обрати…", command=self.pick_input).pack(side="left")
        ttk.Button(r0, text="Авто", command=self.pick_latest).pack(side="left", padx=(8,0))

        ttk.Label(self, text="Пости (у звіт потрапляють лише активні)").pack(anchor="w", pady=(10,4))

        self.tree = ttk.Treeview(self, columns=("active", "name"), show="headings", height=10)
        self.tree.heading("active", text="Активний")
        self.tree.heading("name", text="Назва поста")
        self.tree.column("active", width=90, anchor="center")
        self.tree.column("name", width=520, anchor="w")
        self.tree.pack(fill="both", expand=True)

        self.tree.bind("<Button-1>", self.on_click)
        self.tree.bind("<Double-1>", self.on_double_click)

        btns = ttk.Frame(self); btns.pack(fill="x", pady=(10,0))
        ttk.Button(btns, text="Додати", command=self.add_post).pack(side="left")
        ttk.Button(btns, text="Видалити", command=self.delete_post).pack(side="left", padx=8)
        ttk.Button(btns, text="Зберегти", command=self.save).pack(side="left")
        ttk.Button(btns, text="Згенерувати звіт", command=self.generate_report).pack(side="right")

        self._editor: tk.Entry | None = None
        self.refresh()

    def refresh(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for p in self.posts:
            mark = "☑" if p.get("active") else "☐"
            self.tree.insert("", "end", iid=str(p.get("id")), values=(mark, p.get("name", "")))

    def pick_input(self):
        p = filedialog.askopenfilename(
            title="Обрати txt",
            filetypes=[("Text files", "*.txt *.TXT"), ("All files", "*.*")]
        )
        if p:
            self.input_path_var.set(p)

    def pick_latest(self):
        p = self._resolve_latest_txt()
        if p:
            self.input_path_var.set(str(p))
        else:
            messagebox.showinfo("Інфо", "Не знайдено *.txt у стандартних каталогах.")

    def _resolve_latest_txt(self) -> Path | None:
        root = _repo_root()
        candidates_dirs = [
            root / "pelengreport" / "data",
            root / "src" / "pelengreport" / "data",
            Path.cwd(),
        ]
        txt_files: list[Path] = []
        for d in candidates_dirs:
            if d.exists():
                txt_files += list(d.glob("*.txt")) + list(d.glob("*.TXT"))
        if not txt_files:
            return None
        txt_files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        return txt_files[0]

    def on_click(self, event):
        # toggle active if click in first column
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)  # '#1' active, '#2' name
        if not item or col != "#1":
            return
        for p in self.posts:
            if str(p.get("id")) == str(item):
                p["active"] = not bool(p.get("active"))
                break
        self.refresh()

    def on_double_click(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item or col != "#2":
            return
        bbox = self.tree.bbox(item, "name")
        if not bbox:
            return
        x, y, w, h = bbox
        value = self.tree.set(item, "name")

        if self._editor is not None:
            self._editor.destroy()
            self._editor = None

        e = tk.Entry(self.tree)
        e.place(x=x, y=y, width=w, height=h)
        e.insert(0, value)
        e.focus_set()

        def commit(_evt=None):
            new_val = e.get().strip()
            for p in self.posts:
                if str(p.get("id")) == str(item):
                    p["name"] = new_val
                    # якщо match порожній — підхопимо з назви
                    if not str(p.get("match") or "").strip():
                        p["match"] = new_val
                    break
            e.destroy()
            self._editor = None
            self.refresh()

        e.bind("<Return>", commit)
        e.bind("<Escape>", lambda _e: (e.destroy(), setattr(self, "_editor", None)))
        e.bind("<FocusOut>", commit)
        self._editor = e

    def add_post(self):
        # простий генератор id
        base = "bp"
        i = 1
        ids = {str(p.get("id")) for p in self.posts}
        while True:
            pid = f"{base}{i:04d}"
            if pid not in ids:
                break
            i += 1
        self.posts.append({
            "id": pid,
            "active": True,
            "name": "НОВИЙ ПОСТ",
            "unit": "А3719\n(63 омбр)",
            "place": "НОВИЙ РАЙОН,\nБП №____",
            "equipment": "“Пластун”",
            "task_text": "Відповідно до плану бойового застосування",
            "match": "НОВИЙ ПОСТ",
        })
        self.refresh()

    def delete_post(self):
        sel = self.tree.selection()
        if not sel:
            return
        pid = sel[0]
        self.posts = [p for p in self.posts if str(p.get("id")) != str(pid)]
        self.refresh()

    def save(self):
        save_posts(self.posts_path, self.posts)
        Toast(self, "Збережено posts.json", 1200).place(relx=0.5, rely=0.0, anchor="n")

    def generate_report(self):
        self.save()
        p = (self.input_path_var.get() or "").strip()
        if not p:
            # auto
            latest = self._resolve_latest_txt()
            if latest is None:
                messagebox.showwarning("Помилка", "Не вказано txt і авто-пошук нічого не знайшов.")
                return
            p = str(latest)

        try:
            out_path = run_report(Path(p))
            messagebox.showinfo("Готово", f"Звіт згенеровано:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Помилка", str(e))


class App(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=0)
        self.pack(fill="both", expand=True)

        cfg = load_config("config.yml")
        reference_df = load_reference(cfg.paths.freq_file)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)

        tab1 = PelengTab(nb, reference_df=reference_df)
        tab2 = ReportTab(nb)

        nb.add(tab1, text="Пеленги")
        nb.add(tab2, text="Звіт")


def main():
    root = tk.Tk()
    root.title("Peleng tools")
    root.geometry("600x600")
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
