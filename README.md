# Excel Merge

[简体中文](./README.zh-CN.md) | **English**

[![License](https://img.shields.io/github/license/RedMaple96/ExcelMerge?style=for-the-badge)](./LICENSE)
[![Release](https://img.shields.io/github/v/release/RedMaple96/ExcelMerge?style=for-the-badge)](https://github.com/RedMaple96/ExcelMerge/releases)
[![Downloads](https://img.shields.io/github/downloads/RedMaple96/ExcelMerge/total?style=for-the-badge)](https://github.com/RedMaple96/ExcelMerge/releases)

`Excel Merge` is a desktop application for Excel diffing and merge workflows. It is designed for projects that need to load two `.xlsx` files, identify row and column differences, apply overwrite or append merge strategies, and support visual confirmation before saving. This file is ready to use as the default GitHub homepage for the repository, and you only need to replace the placeholder repository links, license, contact details, and release references.

## Table of Contents

- [Introduction](#introduction)
- [Features](#features)
- [Quick Start](#quick-start)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage Examples](#usage-examples)
- [API Reference](#api-reference)
- [Contributing](#contributing)
- [License](#license)
- [Contact](#contact)
- [Changelog](#changelog)

<a id="introduction"></a>
## Introduction

This project currently uses Python, PySide6, and openpyxl to provide a GUI-based Excel comparison and merge experience. It fits common scenarios such as internal data reconciliation, workbook version verification, and spreadsheet consolidation. The structure of this document is also suitable as a reference for similar desktop tools, while the content here is written specifically for the current repository.

<a id="features"></a>
## Features

- Load two `.xlsx` workbooks and switch worksheets for side-by-side comparison.
- Detect differences with header-aware and content-aware alignment logic.
- Support common merge strategies such as right-to-left overwrite, left-to-right overwrite, and append-only merge.
- Review differences visually in the GUI, navigate changed regions, and confirm merge actions manually.
- Keep the document GitHub-ready with standard badges, links, and reusable placeholders.

<a id="quick-start"></a>
## Quick Start

If you want to run the current repository as quickly as possible, use the default entry point `main.py` and follow the steps below:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python main.py
```

If you want to package the application for desktop distribution, continue with `packaging/mac_build.sh` or `packaging/windows_build.bat`.

<a id="requirements"></a>
## Requirements

The current repository depends on a Python runtime, a GUI framework, and an Excel processing library. The table below shows a practical baseline.

| Item | Notes |
| --- | --- |
| Python | Recommended version: `3.11+` with virtual environment support |
| GUI | `PySide6>=6.5.0` |
| Excel Engine | `openpyxl>=3.1.0` |
| Packaging | `pyinstaller>=6.0.0` |
| Operating System | macOS / Windows, validate other platforms in your own environment |

<a id="installation"></a>
## Installation

Use the following flow for repository setup, dependency installation, and local startup:

```bash
git clone https://github.com/RedMaple96/ExcelMerge.git
cd ExcelMerge
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
python main.py
```

If you need distributable artifacts, validate the GUI behavior, file permissions, and target platform compatibility locally before running your packaging script or CI workflow.

<a id="usage-examples"></a>
## Usage Examples

The examples below are based on the actual structure of this repository and can be used directly or adapted for automation scenarios.

### Example 1: Launch the GUI

```bash
python main.py
```

### Example 2: Compare worksheets in Python

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

### Example 3: Merge and save the result

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
## API Reference

This repository is not an HTTP service, so the API section is documented as module entry points, class interfaces, and key methods.

### `main.py`

Application entry point. It initializes `QApplication`, applies the theme, and creates the main window.

### `src.core.excel_loader.ExcelLoader`

- `load_workbook(path: str)`: Load an `.xlsx` workbook.
- `get_sheet_names(wb)`: Return the worksheet name list.
- `get_worksheet(wb, name: str)`: Get a worksheet by name.
- `extract_sheet_data(ws)`: Extract structured sheet data for comparison and merge flows.
- `save_workbook(wb, path: str)`: Save a workbook to a target path.

### `src.core.comparator.ExcelComparator`

- `compare_sheets(left, right, key_cols=None, ignore_cols=None)`: Compare two worksheet snapshots.
- Supports intelligent alignment based on headers and row content by default.
- Returns row alignment, column alignment, diff cell sets, and summary statistics.

### `src.core.merger.ExcelMerger`

- `merge_right_to_left(...)`: Overwrite left-side differences with right-side values.
- `merge_left_to_right(...)`: Overwrite right-side differences with left-side values.
- `append_rows(...)`: Append diff rows into the target workbook.
- Additional copy and insert helpers can act as extension points for custom workflows.

<a id="contributing"></a>
## Contributing

Contributions through Issues, Discussions, and Pull Requests are welcome. Before submitting changes, review the scope, prepare a minimal local verification, and include enough context for reviewers.

1. Fork the repository
2. Create a feature branch such as `feat/your-feature`
3. Commit your changes with clear context
4. Run the relevant tests and self-checks
5. Open a Pull Request with motivation, implementation notes, and validation results

<a id="license"></a>
## License

This repository uses the `MIT` license. Add the matching `LICENSE` file at the repository root before publishing.

- Current license: `MIT`
- License file path: `./LICENSE`

<a id="contact"></a>
## Contact

The maintainer information is listed below:

- Maintainer: `RedMaple96`
- GitHub: [https://github.com/RedMaple96](https://github.com/RedMaple96)
- Project homepage: [https://github.com/RedMaple96/ExcelMerge](https://github.com/RedMaple96/ExcelMerge)

<a id="changelog"></a>
## Changelog

Keep user-facing release notes in this section or move them into a dedicated `CHANGELOG.md` file and link it here.

### `v0.1.0` - `2026-06-30`

- Initialize the project documentation.
- Add project structure, requirements, installation, and usage examples.
- Add GitHub-friendly badges, contribution guidance, and license placeholders.

## Customization Notes

Before publishing this file to GitHub, make sure the `LICENSE` file exists and that your release links match the tags you publish in the repository.
