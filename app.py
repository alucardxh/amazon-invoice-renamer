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
    "お菓子": ["ウエハース", "コンフェクト", "クッキー", "チョコ", "お菓子", "スナック"],
    "文房具": ["ノート", "ペン", "クリップ", "ファイル"],
    "日用品": ["タオル", "洗剤", "ティッシュ", "ハンドウォッシュ", "除菌", "洗浄", "キュキュット"],
    "電化製品": ["チャイム", "インターホン", "ワイヤレス", "Bluetooth", "充電"],
}


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


def _parse_blocks(pdf_path: Path, rules: dict) -> list:
    """逐页解析，同一发票号只取第一页（多页发票后续页跳过）。
    一个PDF里若有多张独立发票（不同发票号），每张各自生成一个block。"""
    blocks = []
    seen_inv = set()
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            m = re.search(r"請求書番号\s*([A-Z0-9]+)\s*合計\s*￥\s*([\d,]+)", text)
            if not m:
                continue

            invoice_no = m.group(1)
            if invoice_no in seen_inv:
                continue  # 多页发票后续页重复同一发票号，跳过
            seen_inv.add(invoice_no)

            total = int(m.group(2).replace(",", ""))

            m_date = re.search(r"請求書発行日\s*[:\s]*\s*(\d{4})-(\d{2})-(\d{2})", text)
            date = f"{m_date.group(1)}{m_date.group(2)}{m_date.group(3)}" if m_date else None

            # 商品名常跨多行，单独按行匹配会漏判，所以直接在整页文本里查找关键词
            category = "未分类"
            for cat, keywords in rules.items():
                if any(kw in text for kw in keywords):
                    category = cat
                    break

            blocks.append({"invoice_no": invoice_no, "date": date, "total": total, "category": category})
    return blocks


def _build_filename(blocks: list) -> str:
    if not blocks:
        return "00000000_Amazon_未分类_0.pdf"

    date = next((b["date"] for b in blocks if b["date"]), "00000000")
    total_sum = sum(b["total"] for b in blocks)

    categories = []
    for b in blocks:
        if b["category"] not in categories:
            categories.append(b["category"])
    category_str = "・".join(categories)

    return f"{date}_Amazon_{category_str}_{total_sum}.pdf"


def _unique_path(path: Path) -> Path:
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


class InvoiceFile:
    """封装单个 PDF 的路径、解析结果和处理状态。"""

    def __init__(self, path: Path):
        self.path = path
        self.blocks: list = []
        self.error: str = ""
        self.dup_warnings: list = []  # 由外部跨文件检测后写入
        self.new_path: Path = None

    # ── 属性 ──────────────────────────────────────────────

    @property
    def name(self) -> str:
        return self.path.name

    @property
    def invoice_numbers(self) -> list:
        return [b["invoice_no"] for b in self.blocks]

    @property
    def new_name(self) -> str:
        base = _build_filename(self.blocks)
        if self.dup_warnings:
            return base[:-4] + "_重複.pdf"
        return base

    @property
    def has_unclassified(self) -> bool:
        return any(b["category"] == "未分类" for b in self.blocks)

    # ── 操作 ──────────────────────────────────────────────

    def parse(self, rules: dict):
        try:
            self.blocks = _parse_blocks(self.path, rules)
        except Exception as e:
            self.error = str(e)

    def rename(self):
        stem = self.new_name[:-4]
        if re.match(rf"^{re.escape(stem)}(_\d+|_重複)?\.pdf$", self.path.name):
            self.new_path = self.path  # 已是目标名或其 _N/_重複 变体，无需改动
            return
        target = _unique_path(self.path.parent / self.new_name)
        self.path.rename(target)
        self.new_path = target


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
        self.tree.tag_configure("warning", foreground="red")
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

        # 步骤一：解析所有文件
        invoices = []
        for idx, path in enumerate(pdf_files, 1):
            self._set_progress(f"解析中 {idx}/{total}: {path.name}")
            inv = InvoiceFile(path)
            inv.parse(self.rules)
            invoices.append(inv)

        # 步骤二：跨文件重复检测（同一发票号出现在不同文件里）
        seen_inv = {}
        for inv in invoices:
            for no in inv.invoice_numbers:
                if no in seen_inv:
                    inv.dup_warnings.append(f"{no}(与 {seen_inv[no]} 重复)")
                else:
                    seen_inv[no] = inv.name

        # 步骤三：重命名并展示结果
        for idx, inv in enumerate(invoices, 1):
            self._set_progress(f"重命名 {idx}/{total}: {inv.name}")

            if inv.error:
                self._add_row(inv.name, f"出错: {inv.error}", "失败")
                continue

            if not inv.blocks:
                self._add_row(inv.name, "未识别到日期/金额，未重命名", "失败")
                continue

            try:
                inv.rename()
                if inv.new_path == inv.path:
                    self._add_row(inv.name, inv.name, "无需改动")
                else:
                    status = "成功(含未分类)" if inv.has_unclassified else "成功"
                    self._add_row(inv.name, inv.new_path.name, status, warning=bool(inv.dup_warnings))
            except Exception as e:
                self._add_row(inv.name, f"出错: {e}", "失败")

        self._set_progress(f"完成，共处理 {total} 个文件")
        self.start_btn.config(state="normal")

    def _add_row(self, original, new, status, warning=False):
        tags = ("warning",) if warning else ()
        self.tree.insert("", "end", values=(original, new, status), tags=tags)

    def _set_progress(self, text):
        self.progress_label.config(text=text)


if __name__ == "__main__":
    App().mainloop()
