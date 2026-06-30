# Excel Merge

**简体中文** | [English](./README.md)

[![开源许可](https://img.shields.io/github/license/RedMaple96/ExcelMerge?style=for-the-badge)](./LICENSE)
[![版本发布](https://img.shields.io/github/v/release/RedMaple96/ExcelMerge?style=for-the-badge)](https://github.com/RedMaple96/ExcelMerge/releases)
[![下载量](https://img.shields.io/github/downloads/RedMaple96/ExcelMerge/total?style=for-the-badge)](https://github.com/RedMaple96/ExcelMerge/releases)

`Excel Merge` 是一个面向 Excel 差异比对与合并场景的桌面应用，适用于需要加载两个 `.xlsx` 文件、识别行列差异、执行覆盖或追加合并，并以可视化方式辅助人工确认的项目。当前文件可以直接作为仓库的中文说明文档使用，你只需要替换其中的仓库地址、许可证、联系方式与发布链接占位符即可。

## 目录

- [项目介绍](#introduction)
- [功能特性](#features)
- [快速开始](#quick-start)
- [环境依赖](#requirements)
- [安装部署](#installation)
- [使用示例](#usage-examples)
- [API 文档](#api-reference)
- [贡献指南](#contributing)
- [开源许可](#license)
- [联系方式](#contact)
- [更新日志](#changelog)

<a id="introduction"></a>
## 项目介绍

本项目当前采用 Python、PySide6 与 openpyxl 技术栈，提供图形界面形式的 Excel 对比与合并能力。它适合内部数据校对、版本文件核验、业务台账整合等常见场景。文档结构也适合作为同类桌面工具项目的参考模板，但这里的内容已经按当前仓库实际情况进行了落地。

<a id="features"></a>
## 功能特性

- 支持加载两个 `.xlsx` 工作簿并切换工作表进行左右对比。
- 支持基于标题行和内容对齐的差异识别。
- 支持右覆盖、左覆盖、追加差异行等常见合并策略。
- 支持在图形界面中查看差异、导航变更位置并执行人工确认。
- 保留 GitHub 常用徽章、链接与占位符，适合继续扩展和发布。

<a id="quick-start"></a>
## 快速开始

如果你只想快速运行当前仓库，请使用默认入口 `main.py`，按以下步骤完成环境准备与应用启动：

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python main.py
```

如果你希望将应用打包为桌面分发版本，可以继续结合 `packaging/mac_build.sh` 或 `packaging/windows_build.bat` 完成构建流程。

<a id="requirements"></a>
## 环境依赖

当前仓库依赖 Python 运行环境、GUI 框架与 Excel 处理库，建议基线如下：

| 项目 | 说明 |
| --- | --- |
| Python | 建议使用 `3.11+`，并支持虚拟环境 |
| GUI | `PySide6>=6.5.0` |
| Excel 引擎 | `openpyxl>=3.1.0` |
| 打包工具 | `pyinstaller>=6.0.0` |
| 操作系统 | macOS / Windows，其他平台请自行验证 |

<a id="installation"></a>
## 安装部署

建议按照“克隆仓库、创建虚拟环境、安装依赖、启动应用”的顺序完成本地部署：

```bash
git clone https://github.com/RedMaple96/ExcelMerge.git
cd ExcelMerge
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
python main.py
```

如果需要生成可分发构建产物，请先在本地验证 GUI 行为、文件读写权限与目标平台兼容性，再执行打包脚本或 CI 工作流。

<a id="usage-examples"></a>
## 使用示例

以下示例基于当前仓库的实际结构组织，可直接用于本地运行或二次开发场景。

### 示例 1：启动图形界面

```bash
python main.py
```

### 示例 2：在 Python 中比较工作表

```python
from src.core.comparator import ExcelComparator
from src.core.excel_loader import ExcelLoader

left_wb = ExcelLoader.load_workbook("left.xlsx")
right_wb = ExcelLoader.load_workbook("right.xlsx")

left_ws = ExcelLoader.get_worksheet(left_wb, left_wb.sheetnames[0])
right_ws = ExcelLoader.get_worksheet(right_wb, right_wb.sheetnames[0])

left_data = ExcelLoader.extract_sheet_data(left_ws)
right_data = ExcelLoader.extract_sheet_data(right_ws)

diff_result = ExcelComparator.compare_sheets(left_data, right_data)
print(diff_result.stats)
```

### 示例 3：执行合并并保存结果

```python
from src.core.comparator import ExcelComparator
from src.core.excel_loader import ExcelLoader
from src.core.merger import ExcelMerger

left_wb = ExcelLoader.load_workbook("left.xlsx")
right_wb = ExcelLoader.load_workbook("right.xlsx")

left_ws = ExcelLoader.get_worksheet(left_wb, left_wb.sheetnames[0])
right_ws = ExcelLoader.get_worksheet(right_wb, right_wb.sheetnames[0])

left_data = ExcelLoader.extract_sheet_data(left_ws)
right_data = ExcelLoader.extract_sheet_data(right_ws)

diff_result = ExcelComparator.compare_sheets(left_data, right_data)
ExcelMerger.merge_right_to_left(diff_result, left_data, right_data)
ExcelLoader.save_workbook(left_wb, "merged-output.xlsx")
```

<a id="api-reference"></a>
## API 文档

当前仓库不是 HTTP 服务型项目，因此本节采用“模块入口 / 类接口 / 关键方法”的形式描述 API。

### `main.py`

应用启动入口，负责初始化 `QApplication`、加载主题并创建主窗口。

### `src.core.excel_loader.ExcelLoader`

- `load_workbook(path: str)`：加载 `.xlsx` 工作簿。
- `get_sheet_names(wb)`：返回工作表名称列表。
- `get_worksheet(wb, name: str)`：按名称获取工作表对象。
- `extract_sheet_data(ws)`：提取结构化表格数据，供比较与合并流程复用。
- `save_workbook(wb, path: str)`：将工作簿保存到指定路径。

### `src.core.comparator.ExcelComparator`

- `compare_sheets(left, right, key_cols=None, ignore_cols=None)`：比较两个工作表的数据快照。
- 默认支持基于标题与内容的智能对齐。
- 输出包含行对齐结果、列对齐结果、差异单元格集合与统计信息。

### `src.core.merger.ExcelMerger`

- `merge_right_to_left(...)`：以右侧结果覆盖左侧差异。
- `merge_left_to_right(...)`：以左侧结果覆盖右侧差异。
- `append_rows(...)`：将差异行追加到目标工作簿。
- 其余复制、插入方法可作为二次开发扩展点。

<a id="contributing"></a>
## 贡献指南

欢迎通过 Issue、Discussion 或 Pull Request 参与贡献。提交改动前，建议先确认需求范围、完成最小本地验证，并为审阅者提供足够上下文。

1. Fork 本仓库
2. 创建功能分支，例如 `feat/your-feature`
3. 提交改动并补充清晰说明
4. 运行相关测试与自检
5. 发起 Pull Request，并说明动机、实现方式与验证结果

<a id="license"></a>
## 开源许可

当前仓库使用 `MIT` 许可证。正式发布前，请在仓库根目录补充对应的 `LICENSE` 文件。

- 当前许可证：`MIT`
- 许可证文件路径：`./LICENSE`

<a id="contact"></a>
## 联系方式

当前维护者信息如下：

- 维护者：`RedMaple96`
- GitHub：[https://github.com/RedMaple96](https://github.com/RedMaple96)
- 项目主页：[https://github.com/RedMaple96/ExcelMerge](https://github.com/RedMaple96/ExcelMerge)

<a id="changelog"></a>
## 更新日志

建议将面向用户的版本变化记录统一维护在本节，或拆分到独立的 `CHANGELOG.md` 文件并从此处链接。

### `v0.1.0` - `2026-06-30`

- 初始化项目说明文档。
- 补充项目结构、环境依赖、安装部署与使用示例。
- 增加适配 GitHub 的徽章、贡献指南与许可证占位信息。

## 自定义提示

正式上传到 GitHub 前，建议确认根目录已存在 `LICENSE` 文件，并确保发布页链接能对应到你实际创建的版本标签。
