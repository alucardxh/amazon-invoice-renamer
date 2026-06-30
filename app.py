# -*- coding: utf-8 -*-
"""
Amazon 发票批量重命名工具
命名格式: {发行日期YYYYMMDD}_Amazon_{类目}_{含税合计金额}.pdf
"""
import json
import re
import sys
import threading
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pdfplumber

APP_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
# 配置文件放在exe/脚本同目录，方便用户自己编辑、且打包后依然可写
CONFIG_PATH = Path(sys.argv[0]).resolve().parent / "category_rules.json"

DEFAULT_RULES = {
    "本部お菓子": ["ウエハース", "コンフェクト", "クッキー", "チョコ", "お菓子", "スナック"],
    "本部文房具": ["ノート", "ペン", "クリップ", "ファイル"],
    "本部日用品": ["タオル", "洗剤", "ティッシュ"],
}

FILENAME_PATTERN = re.compile(r"^\d{8}_Amazon_.+_\d+\.pdf$", re.IGNORECASE)


def load_rules() -> dict:
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    save_rules(DEFAULT_RULES)
    return dict(DEFAULT_RULES)


def save_rules(rules: dict):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(rules, f, ensure_ascii=False, indent=2)


def extract_text(pdf_path: Path) -> str:
    text = ""
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or "") + "\n"
    return text


def parse_invoice(text: str, rules: dict) -> dict:
    info = {}

    m = re.search(r"請求書発行日\s*[:\s]*\s*(\d{4})-(\d{2})-(\d{2})", text)
    if m:
        info["date"] = f"{m.group(1)}{m.group(2)}{m.group(3)}"

    m = re.search(r"請求書番号\s*([A-Z0-9]+)\s*合計\s*￥\s*([\d,]+)", text)
    if m:
        info["total"] = m.group(2).replace(",", "")
    else:
        amounts = [int(a.replace(",", "")) for a in re.findall(r"合計\s*￥\s*([\d,]+)", text)]
        if amounts:
            info["total"] = str(max(amounts))

    items = re.findall(r"^(.*?)\s*\|\s*[A-Z0-9]{10}\b", text, flags=re.MULTILINE)
    info["items"] = items

    category = "本部未分类"
    for cat, keywords in rules.items():
        if any(kw in item for item in items for kw in keywords):
            category = cat
            break
    info["category"] = category
    return info


def build_filename(info: dict) -> str:
    date = info.get("date", "00000000")
    total = info.get("total", "0")
    category = info.get("category", "本部未分类")
    return f"{date}_Amazon_{category}_{total}.pdf"


def unique_path(path: Path) -> Path:
    """若目标文件名已存在，追加 _1 _2 避免覆盖"""
    if not path.exists():
        return path
    stem, suffix, parent = path.stem, path.suffix, path.parent
    i = 1
    while True:
        candidate = parent / f"{stem}_{i}{suffix}"
        if not candidate.exists():
            return candidate
        i += 1


class CategoryEditor(tk.Toplevel):
    def __init__(self, master, rules: dict, on_save):
        super().__init__(master)
        self.title("编辑类目关键词")
        self.geometry("520x420")
        self.rules = {k: list(v) for k, v in rules.items()}
        self.on_save = on_save

        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        ttk.Label(frame, text="类目名称").grid(row=0, column=0, sticky="w")
        ttk.Label(frame, text="关键词（逗号分隔）").grid(row=0, column=1, sticky="w")

        self.tree_frame = ttk.Frame(frame)
        self.tree_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=5)
        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=2)

        self.rows = []  # list of (cat_entry, kw_entry, row_frame)
        self.rows_container = ttk.Frame(self.tree_frame)
        self.rows_container.pack(fill="both", expand=True)

        for cat, kws in self.rules.items():
            self._add_row(cat, ", ".join(kws))

        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(btn_frame, text="+ 新增类目", command=lambda: self._add_row("", "")).pack(side="left")
        ttk.Button(btn_frame, text="保存", command=self._save).pack(side="right")
        ttk.Button(btn_frame, text="取消", command=self.destroy).pack(side="right", padx=5)

    def _add_row(self, cat_val, kw_val):
        row = ttk.Frame(self.rows_container)
        row.pack(fill="x", pady=2)
        cat_entry = ttk.Entry(row, width=18)
        cat_entry.insert(0, cat_val)
        cat_entry.pack(side="left", padx=(0, 5))
        kw_entry = ttk.Entry(row)
        kw_entry.insert(0, kw_val)
        kw_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        del_btn = ttk.Button(row, text="删除", width=6,
                              command=lambda: self._remove_row(row))
        del_btn.pack(side="left")
        self.rows.append((cat_entry, kw_entry, row))

    def _remove_row(self, row):
        self.rows = [r for r in self.rows if r[2] is not row]
        row.destroy()

    def _save(self):
        new_rules = {}
        for cat_entry, kw_entry, _ in self.rows:
            cat = cat_entry.get().strip()
            kws = [k.strip() for k in kw_entry.get().split(",") if k.strip()]
            if cat and kws:
                new_rules[cat] = kws
        if not new_rules:
            messagebox.showwarning("提示", "至少需要保留一个有效类目")
            return
        save_rules(new_rules)
        self.on_save(new_rules)
        messagebox.showinfo("已保存", "类目关键词已更新")
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Amazon 发票批量重命名工具")
        self.geometry("760x520")

        self.rules = load_rules()
        self.folder_path = tk.StringVar()

        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=10)

        ttk.Entry(top, textvariable=self.folder_path).pack(side="left", fill="x", expand=True)
        ttk.Button(top, text="选择文件夹", command=self.choose_folder).pack(side="left", padx=5)
        ttk.Button(top, text="编辑类目关键词", command=self.open_editor).pack(side="left", padx=5)

        action = ttk.Frame(self)
        action.pack(fill="x", padx=10)
        self.start_btn = ttk.Button(action, text="开始处理", command=self.start_processing)
        self.start_btn.pack(side="left")
        self.progress_label = ttk.Label(action, text="")
        self.progress_label.pack(side="left", padx=10)

        columns = ("original", "new", "status")
        self.tree = ttk.Treeview(self, columns=columns, show="headings")
        self.tree.heading("original", text="原文件名")
        self.tree.heading("new", text="新文件名 / 结果")
        self.tree.heading("status", text="状态")
        self.tree.column("original", width=260)
        self.tree.column("new", width=320)
        self.tree.column("status", width=120)
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

    def choose_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_path.set(path)

    def open_editor(self):
        def on_save(new_rules):
            self.rules = new_rules
        CategoryEditor(self, self.rules, on_save)

    def start_processing(self):
        folder = self.folder_path.get().strip()
        if not folder or not Path(folder).is_dir():
            messagebox.showwarning("提示", "请先选择一个有效的文件夹")
            return
        self.start_btn.config(state="disabled")
        self.tree.delete(*self.tree.get_children())
        threading.Thread(target=self._process_folder, args=(folder,), daemon=True).start()

    def _process_folder(self, folder):
        pdf_files = sorted(Path(folder).glob("*.pdf"))
        total = len(pdf_files)
        if total == 0:
            self._set_progress("文件夹内没有找到 PDF 文件")
            self.start_btn.config(state="normal")
            return

        for idx, pdf_path in enumerate(pdf_files, 1):
            self._set_progress(f"处理中 {idx}/{total}: {pdf_path.name}")

            if FILENAME_PATTERN.match(pdf_path.name):
                self._add_row(pdf_path.name, "(已是目标格式，跳过)", "跳过")
                continue

            try:
                text = extract_text(pdf_path)
                info = parse_invoice(text, self.rules)
                if "date" not in info or "total" not in info:
                    self._add_row(pdf_path.name, "缺少日期或金额，未重命名", "失败")
                    continue

                new_name = build_filename(info)
                new_path = unique_path(pdf_path.parent / new_name)
                pdf_path.rename(new_path)

                status = "成功" if info.get("category") != "本部未分类" else "成功(未分类)"
                self._add_row(pdf_path.name, new_path.name, status)
            except Exception as e:
                self._add_row(pdf_path.name, f"出错: {e}", "失败")

        self._set_progress(f"完成，共处理 {total} 个文件")
        self.start_btn.config(state="normal")

    def _add_row(self, original, new, status):
        self.tree.insert("", "end", values=(original, new, status))

    def _set_progress(self, text):
        self.progress_label.config(text=text)


if __name__ == "__main__":
    App().mainloop()
