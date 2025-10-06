#!/bin/bash
# macOS 一键汇总脚本（配合 summarize_excels.py 使用）
# 功能：支持拖拽 ZIP/文件夹 到脚本图标，或双击后图形化选择
# 语言：中文提示

set -euo pipefail

# 定位脚本与 Python 脚本
DIR="$(cd "$(dirname "$0")" && pwd)"
PY_SCRIPT="$DIR/summarize_excels.py"

# 优先使用 Homebrew/Framework Python，保证依赖位置一致
if [ -x "/usr/local/bin/python3" ]; then
  PYTHON_BIN="/usr/local/bin/python3"
elif command -v python3 >/dev/null 2>&1; then
  PYTHON_BIN="$(command -v python3)"
elif command -v python >/dev/null 2>&1; then
  PYTHON_BIN="$(command -v python)"
else
  echo "未找到 Python，请先安装 python3"
  exit 1
fi

echo "使用 Python 解释器：$PYTHON_BIN"

if [ ! -f "$PY_SCRIPT" ]; then
  echo "未在脚本目录找到 summarize_excels.py，请把它放到：$DIR"
  exit 1
fi

INPUT_PATH="${1:-}"   # 支持把 zip/文件夹拖到脚本上
OUTPUT_PATH="${2:-}"  # 可选：第二个参数指定输出路径
DETAIL_CELLS_RAW="${3:-}"  # 可选：第三个参数传入抽检单元格，逗号分隔，如 局食堂:N18,中心食堂:D14

# AppleScript 选择 ZIP 文件
choose_zip() {
  osascript <<'OSA'
try
  set f to choose file with prompt "选择数据包（ZIP）" of type {"zip"}
  return POSIX path of f
on error
  return ""
end try
OSA
}

# AppleScript 选择文件夹
choose_folder() {
  osascript <<'OSA'
try
  set d to choose folder with prompt "选择包含多个 Excel 的文件夹"
  return POSIX path of d
on error
  return ""
end try
OSA
}

# AppleScript 选择输出文件名
choose_output_path() {
  osascript <<'OSA'
try
  set f to choose file name with prompt "选择输出文件名" default name "汇总结果.xlsx"
  return POSIX path of f
on error
  return ""
end try
OSA
}

# AppleScript 输入抽检单元格
ask_detail_cells() {
  osascript <<'OSA'
try
  display dialog "输入抽检单元格（可留空），格式：表名:单元格，多个用英文逗号分隔\n例如：局食堂:N18,中心食堂:D14" default answer "" buttons {"确定"} default button "确定"
  set ans to text returned of result
  return ans
on error
  return ""
end try
OSA
}

# AppleScript 询问是否导出抽检明细
ask_yes_no() {
  osascript <<'OSA'
try
  set dlg to display dialog "是否导出抽检明细文件？" buttons {"否", "是"} default button "否"
  set btn to button returned of dlg
  return btn
on error
  return "否"
end try
OSA
}

# AppleScript 选择抽检明细输出路径
choose_detail_out() {
  osascript <<'OSA'
try
  set f to choose file name with prompt "选择抽检明细输出文件名" default name "逐文件明细.xlsx"
  return POSIX path of f
on error
  return ""
end try
OSA
}

# 1) 处理输入路径
if [ -z "$INPUT_PATH" ]; then
  # 先尝试选择 ZIP，取消则选择文件夹
  INPUT_PATH="$(choose_zip)"
  if [ -z "$INPUT_PATH" ]; then
    INPUT_PATH="$(choose_folder)"
  fi
  if [ -z "$INPUT_PATH" ]; then
    echo "未选择输入数据，退出。"
    exit 1
  fi
fi

# 2) 处理输出路径
if [ -z "$OUTPUT_PATH" ]; then
  OUTPUT_PATH="$(choose_output_path)"
  if [ -z "$OUTPUT_PATH" ]; then
    # 默认放到输入同目录
    base_dir="$(dirname "$INPUT_PATH")"
    OUTPUT_PATH="$base_dir/汇总结果.xlsx"
    echo "未选择输出文件名，默认：$OUTPUT_PATH"
  fi
fi

# 3) 处理抽检单元格与明细导出
DETAIL_FLAGS=()
if [ -z "$DETAIL_CELLS_RAW" ]; then
  DETAIL_CELLS_RAW="$(ask_detail_cells)"
fi

if [ -n "$DETAIL_CELLS_RAW" ]; then
  # 分割逗号
  IFS=',' read -r -a CELLS_ARR <<< "$DETAIL_CELLS_RAW"
  for cell in "${CELLS_ARR[@]}"; do
    trimmed="$(echo "$cell" | sed 's/^[[:space:]]*//;s/[[:space:]]*$//')"
    if [ -n "$trimmed" ]; then
      DETAIL_FLAGS+=( "--detail-cell" "$trimmed" )
    fi
  done
fi

DETAIL_OUT_PATH=""
if [ "${#DETAIL_FLAGS[@]}" -gt 0 ]; then
  yn="$(ask_yes_no)"
  if [ "$yn" = "是" ]; then
    DETAIL_OUT_PATH="$(choose_detail_out)"
    if [ -z "$DETAIL_OUT_PATH" ]; then
      DETAIL_OUT_PATH="$(dirname "$OUTPUT_PATH")/逐文件明细.xlsx"
      echo "未选择明细输出文件名，默认：$DETAIL_OUT_PATH"
    fi
    DETAIL_FLAGS+=( "--detail-out" "$DETAIL_OUT_PATH" )
  fi
fi

echo "====== 开始汇总 ======"
echo "输入：$INPUT_PATH"
echo "输出：$OUTPUT_PATH"
if [ "${#DETAIL_FLAGS[@]}" -gt 0 ]; then
  echo "抽检：${DETAIL_FLAGS[*]}"
fi

# 4) 执行 Python 汇总
if [ "${#DETAIL_FLAGS[@]}" -gt 0 ]; then
  "$PYTHON_BIN" "$PY_SCRIPT" --input "$INPUT_PATH" --output "$OUTPUT_PATH" "${DETAIL_FLAGS[@]}"
else
  "$PYTHON_BIN" "$PY_SCRIPT" --input "$INPUT_PATH" --output "$OUTPUT_PATH"
fi

echo "====== 完成 ======"
# 自动在 Finder 中展示输出文件
if [ -f "$OUTPUT_PATH" ]; then
  open -R "$OUTPUT_PATH"
fi
if [ -n "${DETAIL_OUT_PATH:-}" ] && [ -f "$DETAIL_OUT_PATH" ]; then
  open -R "$DETAIL_OUT_PATH"
fi
