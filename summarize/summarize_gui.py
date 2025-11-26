"""Tkinter GUI launcher for summarize_excels."""
from __future__ import annotations

import shutil
from pathlib import Path
from tkinter import (
    Button,
    Label,
    Tk,
    filedialog,
    messagebox,
    simpledialog,
)

try:
    from . import summarize_excels as summarizer
except ImportError:  # When running as standalone script
    import summarize_excels as summarizer  # type: ignore

SCRIPT_DIR = Path(__file__).resolve().parent


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


def choose_detail_output(default_dir: Path) -> str:
    return filedialog.asksaveasfilename(
        title="选择抽检明细输出文件",
        initialdir=str(default_dir),
        defaultextension=".xlsx",
        filetypes=[("Excel 文件", "*.xlsx"), ("CSV", "*.csv"), ("所有文件", "*.*")],
    )


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
    detail_out = ""
    if detail_cells:
        if messagebox.askyesno("抽检明细", "是否导出抽检明细文件？"):
            default_dir = Path(output_path).resolve().parent
            detail_out = choose_detail_output(default_dir)

    try:
        tmp_dir, excel_files = summarizer.list_excels_from_input(input_path)
        try:
            detail_target = detail_out or None
            summarizer.summarize_excels(excel_files, output_path, detail_cells, detail_target)
        finally:
            if tmp_dir and Path(tmp_dir).exists():
                shutil.rmtree(tmp_dir)
    except Exception as exc:  # noqa: BLE001
        messagebox.showerror("错误", f"执行失败：{exc}")
        return

    messagebox.showinfo("成功", f"汇总完成：\n{output_path}")


def main() -> None:
    if not (SCRIPT_DIR / "summarize_excels.py").exists():
        messagebox.showwarning(
            "提示",
            "未找到 summarize_excels.py，确保 GUI 与核心脚本位于 summarize/ 目录。",
        )

    root = Tk()
    root.title("Excel Summarizer")
    root.geometry("360x200")

    Label(root, text="Summarize Excel Files", font=("Arial", 16)).pack(pady=20)
    Button(root, text="开始汇总", width=20, command=run_summarize).pack()
    Button(root, text="退出", width=20, command=root.destroy).pack(pady=10)

    root.mainloop()


if __name__ == "__main__":
    main()
