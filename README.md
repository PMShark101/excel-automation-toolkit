# Excel Automation Toolkit

一个用于存放 Excel 自动化脚本与工具的仓库。每个工具放置在独立子目录中，包含执行脚本、底层 Python 脚本以及使用说明，便于后续扩展更多能力。

- 🚀 **快速上手**：请阅读 [`docs/getting-started.md`](docs/getting-started.md)，了解环境安装、脚本运行以及扩展建议。

## 目录结构

- `summarize/`
  - `SummarizeExcels.command`：macOS 一键脚本，支持拖拽 ZIP/文件夹或图形化选择。
  - `summarize_excels.py`：底层汇总核心逻辑，依赖 `pandas`、`openpyxl`、`numpy`。
  - `summarize_gui.py`：基于 Tkinter 的 GUI 启动器，可在 Windows 打包成 `.exe` 与他人分享。

## 使用方法

1. 安装依赖：
   ```bash
   python3 -m pip install pandas openpyxl numpy
   ```
2. 双击 `SummarizeExcels.command` 或拖拽数据到脚本图标，根据提示选择输入/输出。
3. 若在命令行调用：
   ```bash
   python3 summarize/summarize_excels.py --input data.zip --output 汇总结果.xlsx
   ```

## 后续扩展

- 为新的工具创建类似的子目录，例如 `cleaner/`、`reporter/` 等。
- 在各子目录下附上 README 或注释说明使用方式。可在根目录维护一个总览文档。

欢迎继续添加更多脚本并保持结构清晰。
