"""奢侈品 eBay 标题批量校验工具。

依赖：openpyxl。打包示例：pyinstaller --onefile --windowed cost_calc.py
"""

import re
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    from openpyxl import load_workbook
except ImportError:  # 让 exe 或缺少依赖时给出可读提示
    load_workbook = None


# ============================== SOP 配置区 ==============================
SOP_CONFIG = {
    "max_title_length": 80,
    "uppercase_allowed": {"LV"},
    "accent_terms": {
        "Matelassé": "Matelasse",
        "Bandoulière": "Bandouliere",
        "NéoNoé": "NeoNoe",
        "J'adior": "Jadior",
    },
    "old_size_replacements": {"bb": "Mini", "pm": "Small", "mm": "Medium", "gm": "Large"},
    "brand_pairs": {
        "LV": "Louis Vuitton",
        "YSL": "Yves Saint Laurent",
        "BV": "Bottega Veneta",
    },
    "duplicate_colors": {
        "Black", "White", "Brown", "Red", "Blue", "Green", "Pink", "Grey", "Gray",
        "Yellow", "Orange", "Purple", "Beige", "Gold", "Silver", "Navy", "Ivory",
        "Burgundy", "Emerald Green", "Sherry", "Natural", "Tan",
    },
}
# ========================================================================


def _contains_word(text, word):
    return re.search(r"(?<![A-Za-z])" + re.escape(word) + r"(?![A-Za-z])", text, re.I) is not None


def validate_title(title):
    """返回字符数、问题列表和总结果文本。"""
    title = str(title or "").strip()
    issues = []
    config = SOP_CONFIG

    length = len(title)
    if length > config["max_title_length"]:
        issues.append(f"错误：含空格字符数 {length}，超过 80")

    uppercase_words = re.findall(r"(?<![A-Za-z])([A-Z][A-Z0-9'-]*)(?![A-Za-z])", title)
    # 两字符缩写（如 GG）属于官方系列写法；旧尺寸代号由下方规则单独提示。
    unexpected_uppercase = sorted({word for word in uppercase_words if len(word) >= 3 and word not in config["uppercase_allowed"]})
    if unexpected_uppercase:
        issues.append("告警：存在全大写词（仅 LV 允许）: " + ", ".join(unexpected_uppercase))

    if _contains_word(title, "Gucci") and _contains_word(title, "Monogram"):
        issues.append("错误：Gucci 老花应写 GG Supreme canvas，不应写 Monogram")

    if _contains_word(title, "canvas") and _contains_word(title, "leather"):
        issues.append("错误：canvas 与 leather 不能同时出现")

    if re.search(r"\bmonogram(?:[- ]+[A-Za-z]+)+\b", title, re.I):
        issues.append("告警：Monogram 不能拼接后缀，应单独使用")

    for accented, replacement in config["accent_terms"].items():
        if accented.casefold() in title.casefold():
            issues.append(f"告警：{accented} 含重音符号，应抹平为 {replacement}")

    for color in sorted(config["duplicate_colors"], key=len, reverse=True):
        if re.search(r"\b" + re.escape(color) + r"\s+" + re.escape(color) + r"\b", title, re.I):
            issues.append(f"错误：重复颜色 {color} {color}")

    for old_size, replacement in config["old_size_replacements"].items():
        if re.search(r"\b" + old_size + r"\b", title, re.I):
            issues.append(f"告警：旧尺寸代号 {old_size}，应替换为 {replacement}")

    for abbreviation, full_name in config["brand_pairs"].items():
        has_abbreviation = _contains_word(title, abbreviation)
        has_full_name = full_name.casefold() in title.casefold()
        if has_abbreviation != has_full_name:
            missing = full_name if has_abbreviation else abbreviation
            issues.append(f"告警：品牌名称与缩写不完整，应同时包含 {full_name} {abbreviation}（缺少 {missing}）")

    result = "通过" if not issues else "\n".join(issues)
    return length, issues, result


class TitleCheckerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("奢侈品 eBay 标题校验工具")
        self.root.geometry("1180x720")
        self.root.minsize(860, 540)
        self.root.configure(bg="#f4f6f8")
        self.workbook = None
        self.source_path = None
        self.source_rows = []
        self.title_column_index = None
        self._build_style()
        self._build_ui()

    def _build_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background="#f4f6f8")
        style.configure("Header.TLabel", background="#18324b", foreground="white", font=("Microsoft YaHei UI", 16, "bold"))
        style.configure("Hint.TLabel", background="#f4f6f8", foreground="#5b6875")
        style.configure("Treeview", rowheight=48, font=("Microsoft YaHei UI", 10))
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 10, "bold"))

    def _build_ui(self):
        header = ttk.Label(self.root, text="  奢侈品 eBay 标题校验工具", style="Header.TLabel", anchor="w")
        header.pack(fill="x", ipady=12)

        toolbar = ttk.Frame(self.root, padding=(14, 12, 14, 6))
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="导入 Excel", command=self.import_excel).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="导出校验结果", command=self.export_excel).pack(side="left", padx=(0, 18))
        self.file_label = ttk.Label(toolbar, text="尚未导入文件", style="Hint.TLabel")
        self.file_label.pack(side="left")

        manual = ttk.LabelFrame(self.root, text="单条手动校验", padding=10)
        manual.pack(fill="x", padx=14, pady=(0, 10))
        self.manual_entry = ttk.Entry(manual)
        self.manual_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.manual_entry.bind("<Return>", lambda _event: self.check_manual())
        ttk.Button(manual, text="立即校验", command=self.check_manual).pack(side="left")
        self.manual_result = ttk.Label(manual, text="请输入标题后校验", style="Hint.TLabel")
        self.manual_result.pack(side="left", padx=(14, 0))

        self.column_frame = ttk.Frame(self.root, padding=(14, 0, 14, 8))
        self.column_frame.pack(fill="x")
        self.column_label = ttk.Label(self.column_frame, text="标题列：导入 Excel 后选择", style="Hint.TLabel")
        self.column_label.pack(side="left")
        self.column_var = tk.StringVar()
        self.column_combo = ttk.Combobox(self.column_frame, textvariable=self.column_var, state="readonly", width=24)
        self.column_combo.pack(side="left", padx=10)
        self.column_combo.bind("<<ComboboxSelected>>", lambda _event: self.run_excel_validation())

        result_frame = ttk.Frame(self.root, padding=(14, 0, 14, 14))
        result_frame.pack(fill="both", expand=True)
        columns = ("row", "title", "result", "length")
        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings")
        headings = {"row": "行号", "title": "原始标题", "result": "校验结果 / 错误告警", "length": "字符数"}
        widths = {"row": 65, "title": 420, "result": 560, "length": 70}
        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], anchor="w", stretch=column in {"title", "result"})
        self.tree.tag_configure("pass", foreground="#197044")
        self.tree.tag_configure("issue", foreground="#a43d32")
        y_scroll = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(result_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        result_frame.rowconfigure(0, weight=1)
        result_frame.columnconfigure(0, weight=1)

    def import_excel(self):
        if load_workbook is None:
            messagebox.showerror("缺少依赖", "请先安装 openpyxl：\npip install openpyxl")
            return
        path = filedialog.askopenfilename(title="选择 Excel 文件", filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            self.workbook = load_workbook(path)
            self.source_path = Path(path)
            sheet = self.workbook.active
            rows = list(sheet.iter_rows(values_only=True))
            if not rows:
                raise ValueError("Excel 文件没有数据")
            headers = [str(value).strip() if value is not None else f"未命名列{i + 1}" for i, value in enumerate(rows[0])]
            self.source_rows = rows
            self.column_combo["values"] = headers
            self.column_combo.current(0)
            self.file_label.configure(text=f"已导入：{self.source_path.name}")
            self.run_excel_validation()
        except Exception as error:
            messagebox.showerror("导入失败", f"无法读取 Excel：{error}")

    def run_excel_validation(self):
        if not self.source_rows:
            return
        self.title_column_index = self.column_combo.current()
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row_number, row in enumerate(self.source_rows[1:], start=2):
            title = row[self.title_column_index] if self.title_column_index < len(row) else ""
            length, issues, result = validate_title(title)
            tag = "pass" if not issues else "issue"
            self.tree.insert("", "end", values=(row_number, str(title or ""), result, length), tags=(tag,))

    def check_manual(self):
        title = self.manual_entry.get()
        length, issues, result = validate_title(title)
        color = "#197044" if not issues else "#a43d32"
        self.manual_result.configure(text=f"字符数：{length} | {result.replace(chr(10), '；')}", foreground=color)

    def export_excel(self):
        if self.workbook is None or self.title_column_index is None:
            messagebox.showwarning("无法导出", "请先导入 Excel 并完成标题列选择")
            return
        target = filedialog.asksaveasfilename(
            title="导出校验结果", defaultextension=".xlsx", initialfile=f"{self.source_path.stem}_校验结果.xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not target:
            return
        try:
            sheet = self.workbook.active
            last_column = sheet.max_column
            result_column = last_column + 1
            length_column = last_column + 2
            sheet.cell(1, result_column, "校验结果")
            sheet.cell(1, length_column, "字符数")
            for row_number in range(2, sheet.max_row + 1):
                value = sheet.cell(row_number, self.title_column_index + 1).value
                length, _issues, result = validate_title(value)
                sheet.cell(row_number, result_column, result)
                sheet.cell(row_number, length_column, length)
            self.workbook.save(target)
            messagebox.showinfo("导出成功", f"已导出：{Path(target).name}")
        except Exception as error:
            messagebox.showerror("导出失败", f"无法保存 Excel：{error}")


def main():
    root = tk.Tk()
    TitleCheckerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()