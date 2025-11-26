"""Simple Tkinter GUI for summarize_excels.py."""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, Tk, Label, Button

SCRIPT_DIR = Path(__file__).resolve().parent
PY_SCRIPT = SCRIPT_DIR / "summarize_excels.py"


def choose_input() -> str:
    path = filedialog.askopenfilename(
        title="选择 ZIP 数据包",
        filetypes=[("ZIP 文件", "*.zip"), ("所有文件", "*.*")],
    )
    if path:
        return path
    folder = filedialog.askdirectory(title="或选择一个包含 Excel 的文件夹")
    return folder or ""


def choose_output() -> str:
    return filedialog.asksaveasfilename(
        title="选择输出文件",
        defaultextension=".xlsx",
        filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
    )


def collect_detail_cells() -> list[str]:
    text = simpledialog.askstring(
        "抽检单元格",
        "格式：表名:单元格，多个用逗号分隔\n示例：局食堂:N18,中心食堂:D14",
    )
    if not text:
        return []
    return [item.strip() for item in text.split(",") if item.strip()]


def run_summarize() -> None:
    input_path = choose_input()
    if not input_path:
        messagebox.showinfo("提示", "已取消（未选择输入路径）")
        return

    output_path = choose_output()
    if not output_path:
        messagebox.showinfo("提示", "已取消（未选择输出路径）")
        return

    detail_cells = collect_detail_cells()

    cmd = [sys.executable, str(PY_SCRIPT), "--input", input_path, "--output", output_path]
    for cell in detail_cells:
        cmd.extend(["--detail-cell", cell])

    try:
        subprocess.check_call(cmd)
        messagebox.showinfo("成功", f"汇总完成：\n{output_path}")
    except subprocess.CalledProcessError as exc:
        messagebox.showerror("错误", f"执行失败：{exc}")


def main() -> None:
    if not PY_SCRIPT.exists():
        messagebox.showerror("错误", f"未找到 {PY_SCRIPT}")
        return

    root = Tk()
    root.title("Excel Summarizer")
    root.geometry("360x200")

    Label(root, text="Summarize Excel Files", font=("Arial", 16)).pack(pady=20)
    Button(root, text="开始汇总", width=20, command=run_summarize).pack()
    Button(root, text="退出", width=20, command=root.destroy).pack(pady=10)

    root.mainloop()


if __name__ == "__main__":
    main()
