# -*- coding: utf-8 -*-
"""
Excel 批量汇总脚本（按相同单元格位置求和，空白视为 0；保留标题文本）
Author: ChatGPT
依赖: pandas, openpyxl, numpy
Python: 3.9+

用法示例：
1) 从压缩包读取并输出结果：
   python summarize_excels.py --input data.zip --output 汇总结果.xlsx

2) 另存一份你想抽检的单元格逐文件明细（可多次传入）：
   python summarize_excels.py --input data.zip --output 汇总结果.xlsx \
       --detail-cell "局食堂:N18" --detail-cell "中心食堂:D14" \
       --detail-out 逐文件明细.xlsx
"""

import argparse
import os
import pathlib
import shutil
import sys
import tempfile
import warnings
import zipfile

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")
EXCEL_EXTS = {".xlsx", ".xls", ".xlsm", ".xlsb"}


def is_excel_file(path: str) -> bool:
    """Return True if path looks like an Excel file."""
    return pathlib.Path(path).suffix.lower() in EXCEL_EXTS


def list_excels_from_input(input_path: str):
    """支持文件夹或 zip。返回解压后的临时目录和 Excel 文件列表"""
    if os.path.isdir(input_path):
        excel_files = []
        for root, _, files in os.walk(input_path):
            for filename in files:
                if is_excel_file(filename):
                    excel_files.append(os.path.join(root, filename))
        return None, sorted(excel_files)

    if zipfile.is_zipfile(input_path):
        tmpdir = tempfile.mkdtemp(prefix="excel_zip_")
        with zipfile.ZipFile(input_path, "r") as zf:
            zf.extractall(tmpdir)
        excel_files = []
        for root, _, files in os.walk(tmpdir):
            for filename in files:
                if is_excel_file(filename):
                    excel_files.append(os.path.join(root, filename))
        return tmpdir, sorted(excel_files)

    raise FileNotFoundError(f"无法识别输入: {input_path}")


def coerce_numeric(value):
    """Convert cell value to float when possible; return NaN otherwise."""
    if value is None:
        return np.nan
    if isinstance(value, (int, float, np.number)):
        return float(value)
    if isinstance(value, str):
        stripped = value.strip().replace(",", "").replace("，", "").replace(" ", "")
        if stripped == "" or stripped.lower() in {"nan", "none"}:
            return np.nan
        if stripped.endswith("%"):
            try:
                return float(stripped[:-1]) / 100.0
            except ValueError:
                return np.nan
        try:
            return float(stripped)
        except ValueError:
            return np.nan
    return np.nan


def to_num_zero(value):
    """Convert to numeric, treating NaN as 0."""
    numeric = coerce_numeric(value)
    if isinstance(numeric, float) and np.isnan(numeric):
        return 0.0
    return numeric if numeric is not None else 0.0


def col_letter_to_index(text: str) -> int:
    """Convert Excel column letters to zero-based index."""
    text = text.strip().upper()
    total = 0
    for ch in text:
        total = total * 26 + (ord(ch) - ord("A") + 1)
    return total - 1


def parse_cell_rc(cell: str):
    """Parse Excel cell reference (e.g., A1) to zero-based row/column."""
    letters, digits = "", ""
    for ch in cell.upper():
        if ch.isalpha():
            letters += ch
        elif ch.isdigit():
            digits += ch
    row = int(digits) - 1
    col = col_letter_to_index(letters)
    return row, col


def summarize_excels(excel_files, output_path, detail_cells=None, detail_out=None):
    sheet_templates = {}
    for path in excel_files:
        try:
            excel = pd.ExcelFile(path, engine="openpyxl")
            for sheet_name in excel.sheet_names:
                if sheet_name not in sheet_templates:
                    sheet_templates[sheet_name] = pd.read_excel(
                        path, sheet_name=sheet_name, header=None, engine="openpyxl"
                    )
        except Exception:
            continue

    acc_numeric = {}
    has_numeric = {}
    for sheet_name, template in sheet_templates.items():
        acc_numeric[sheet_name] = np.zeros(template.shape, float)
        has_numeric[sheet_name] = np.zeros(template.shape, bool)

    for path in excel_files:
        try:
            excel = pd.ExcelFile(path, engine="openpyxl")
            for sheet_name, template in sheet_templates.items():
                if sheet_name in excel.sheet_names:
                    df = pd.read_excel(path, sheet_name=sheet_name, header=None, engine="openpyxl")
                else:
                    df = pd.DataFrame(np.full(template.shape, np.nan))

                df = df.reindex(index=range(template.shape[0]), columns=range(template.shape[1]))
                for i in range(template.shape[0]):
                    for j in range(template.shape[1]):
                        numeric = coerce_numeric(df.iat[i, j])
                        if isinstance(numeric, float) and not np.isnan(numeric):
                            acc_numeric[sheet_name][i, j] += numeric
                            has_numeric[sheet_name][i, j] = True
        except Exception:
            continue

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, template in sheet_templates.items():
            rows, cols = template.shape
            output = np.empty((rows, cols), object)
            for i in range(rows):
                for j in range(cols):
                    if has_numeric[sheet_name][i, j]:
                        output[i, j] = acc_numeric[sheet_name][i, j]
                    else:
                        output[i, j] = template.iat[i, j]
            pd.DataFrame(output).to_excel(writer, sheet_name=sheet_name, header=None, index=None)

    if detail_cells:
        detail_rows = []
        for path in excel_files:
            row = {"文件名": os.path.basename(path)}
            try:
                excel = pd.ExcelFile(path, engine="openpyxl")
                for item in detail_cells:
                    sheet_name, cell = item.split(":", 1)
                    row_idx, col_idx = parse_cell_rc(cell)
                    if sheet_name in excel.sheet_names:
                        df = pd.read_excel(path, sheet_name=sheet_name, header=None, engine="openpyxl")
                        value = (
                            df.iat[row_idx, col_idx]
                            if df.shape[0] > row_idx and df.shape[1] > col_idx
                            else None
                        )
                        row[f"{sheet_name}_{cell}"] = to_num_zero(value)
                    else:
                        row[f"{sheet_name}_{cell}"] = 0.0
            except Exception:
                pass
            detail_rows.append(row)

        detail_df = pd.DataFrame(detail_rows)
        summary_row = {"文件名": "合计"}
        for column in detail_df.columns:
            if column != "文件名":
                summary_row[column] = pd.to_numeric(detail_df[column], errors="coerce").fillna(0).sum()
        detail_df = pd.concat([detail_df, pd.DataFrame([summary_row])], ignore_index=True)

        if detail_out and detail_out.endswith(".csv"):
            detail_df.to_csv(detail_out, index=False, encoding="utf-8-sig")
        elif detail_out:
            detail_df.to_excel(detail_out, index=False)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--detail-cell", action="append", default=[])
    parser.add_argument("--detail-out", default=None)
    args = parser.parse_args()

    tmp_dir, excel_files = list_excels_from_input(args.input)
    try:
        summarize_excels(excel_files, args.output, args.detail_cell, args.detail_out)
    finally:
        if tmp_dir and os.path.isdir(tmp_dir):
            shutil.rmtree(tmp_dir)


if __name__ == "__main__":
    main()
