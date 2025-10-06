# 快速上手指南

本指南面向第一次使用 **Excel Automation Toolkit** 的用户，帮助你在几分钟内完成环境准备、运行示例脚本，并了解如何扩展出更多工具。

## 1. 仓库结构速览

```
excel-automation-toolkit/
├── README.md                # 仓库简介
├── docs/                    # 文档区
│   └── getting-started.md   # 当前指南
└── summarize/               # 汇总工具目录
    ├── SummarizeExcels.command  # macOS 一键脚本
    └── summarize_excels.py      # Python 核心逻辑
```

后续如果新增其它工具（例如批量清洗、自动生成报表等），建议像 `summarize/` 一样创建新的子目录，并在其中放置命令行脚本、说明文档。

## 2. 环境准备

1. **安装 Python**：建议使用 Python 3.10 及以上版本。
2. **安装依赖库**（仅需一次）：
   ```bash
   python3 -m pip install pandas openpyxl numpy
   ```
3. **克隆或下载仓库**：
   ```bash
   git clone git@github.com:PMShark101/excel-automation-toolkit.git
   cd excel-automation-toolkit
   ```

> 如果你不想使用命令行，也可以在 GitHub 页面点击 “Code → Download ZIP”，下载后解压即可。

## 3. 运行汇总工具（Summarize）

### 3.1 macOS 图形化方式

1. 在 Finder 中打开 `summarize/` 目录。
2. 双击 `SummarizeExcels.command`，或直接把包含 Excel 的 ZIP / 文件夹拖拽到该图标上。
3. 按提示选择输入数据、输出文件名，可选输入抽检单元格。
4. 执行完成后脚本会自动弹出 Finder 定位生成的结果文件。

> 需要授予终端或 Finder 运行脚本的权限；首次运行时如弹出安全提示，请选择允许。

### 3.2 命令行方式

```
python3 summarize/summarize_excels.py \
    --input /path/to/data.zip \
    --output /path/to/汇总结果.xlsx \
    --detail-cell "局食堂:N18" \
    --detail-cell "中心食堂:D14" \
    --detail-out /path/to/逐文件明细.xlsx
```

- `--input` 可以是 ZIP 或文件夹路径。
- `--output` 指定汇总结果文件名。
- `--detail-cell` 可重复出现，用于抽检多个单元格位置；若省略则不生成明细。
- `--detail-out` 设定抽检明细文件路径（支持 `.xlsx` 或 `.csv`）。

## 4. 常见问题

| 问题 | 解决方案 |
| ---- | -------- |
| 运行时报错 `No module named 'pandas'` | 确认已经执行 `python3 -m pip install pandas openpyxl numpy` |
| 终端显示 `Permission denied` | 确认脚本具有执行权限：`chmod +x summarize/SummarizeExcels.command` |
| Excel 文件中含有公式或合并单元格 | 当前脚本按单元格位置读取数值，公式会自动计算后的结果被读取，合并单元格会按左上角的值处理 |
| 想切换到特定 Python 解释器 | 编辑 `SummarizeExcels.command` 中的 `PYTHON_BIN` 设置即可 |

## 5. 如何扩展新的工具

1. 在仓库根目录新建子目录，例如 `cleaner/`。
2. 将脚本和相关资源放入该目录；如需图形化便捷操作，可以仿照 `SummarizeExcels.command` 编写新的 `.command` 文件。
3. 在根目录的 `README.md` 中补充该工具的简介与使用方式，并在 `docs/` 中撰写更详细的指南。

## 6. 下一步建议

- 集成自动化测试：为关键 Python 脚本编写单元测试，确保批量处理的稳定性。
- 支持更多数据格式：例如增加对 CSV、ODS 的支持，或与数据库互通。
- 加入 GitHub Actions：在推送代码后自动运行检查、打包或发布。

---

如有使用问题或改进建议，欢迎在仓库的 Issues 区反馈，共同完善这个 Excel 自动化工具集。
