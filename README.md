# Excel Merge

[简体中文](./README.zh-CN.md) | **English**

[![License](https://img.shields.io/github/license/RedMaple96/ExcelMerge?style=for-the-badge)](./LICENSE)
[![Python](https://img.shields.io/badge/Python-3.11%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![PySide6](https://img.shields.io/badge/PySide6-6.5%2B-41CD52?style=for-the-badge&logo=qt&logoColor=white)](https://doc.qt.io/qtforpython-6/)
[![openpyxl](https://img.shields.io/badge/openpyxl-3.1%2B-217346?style=for-the-badge)](https://openpyxl.readthedocs.io/)

`Excel Merge` is a desktop application for Excel diffing and merge workflows. It is designed for projects that need to load two `.xlsx` files, identify row and column differences, apply overwrite or append merge strategies, and support visual confirmation before saving. 

## Table of Contents

- [Introduction](#introduction)
- [Features](#features)
- [Core Modules](#core-modules)
- [Quick Start](#quick-start)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage Examples](#usage-examples)
- [Testing](#testing)
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
- Package desktop builds for macOS and Windows with the scripts under `packaging/`.
- Cover the core pipeline with automated tests under `tests/`.

<a id="core-modules"></a>
## Core Modules

The main window is the central workspace of `Excel Merge`. It combines file loading, side-by-side comparison, diff navigation, and merge actions in one focused view.

<p>
  <img src="./iScreen%20Shoter%20-%20Python%20-%20260703141432.png" alt="Excel Merge main window showing a side-by-side worksheet comparison and merge workspace" width="100%" />
</p>

### 1. Workbook Loading and Sheet Switching

- **What it does:** Loads the left and right `.xlsx` workbooks and lets users switch worksheets independently to compare the exact tabs they care about.
- **Core value:** Reduces setup friction and makes it practical to inspect multi-sheet Excel files without rewriting or pre-processing the source files.
- **Typical scenarios:** Comparing monthly reports, validating two workbook versions from different teams, or checking whether a newly delivered workbook matches an existing baseline.

### 2. Intelligent Difference Detection

- **What it does:** Compares worksheet content with header-aware and content-aware alignment so row or column changes can still be identified even when data has shifted.
- **Core value:** Helps users locate real business differences instead of manually scanning large tables cell by cell.
- **Typical scenarios:** Detecting inserted rows, changed metrics, missing records, reordered columns, or unexpected value changes in structured spreadsheets.

### 3. Visual Diff Review and Navigation

- **What it does:** Highlights changed areas in the GUI and supports fast navigation across difference positions in the current comparison session.
- **Core value:** Makes large spreadsheets readable and reviewable, especially when the changed cells are scattered across hundreds of rows.
- **Typical scenarios:** Reviewing audit changes, checking operational data corrections, or quickly locating the next issue during manual verification.

### 4. Merge Decision and Data Synchronization

- **What it does:** Supports merge actions such as right-to-left overwrite, left-to-right overwrite, and append-based updates after users review the detected differences.
- **Core value:** Turns comparison results into actionable updates, reducing repetitive copy-paste work and lowering the risk of manual merge mistakes.
- **Typical scenarios:** Merging vendor updates into a master workbook, applying reviewed corrections, or appending newly added records into the target file.

### 5. Local Packaging and Verification Support

- **What it does:** Provides local packaging scripts for macOS and Windows, along with automated tests for the core comparison and merge pipeline.
- **Core value:** Improves maintainability and helps contributors validate functionality before distributing or extending the application.
- **Typical scenarios:** Preparing internal delivery builds, validating a change before release, or onboarding new contributors who need a reliable local verification flow.

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
The current repository does not include a GitHub Actions workflow yet, so packaging is expected to run locally for now.

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

If you need distributable artifacts, validate the GUI behavior, file permissions, and target platform compatibility locally before running the packaging script.
The macOS script builds `dist/Excel Merge.app` and optionally `dist/Excel Merge.dmg`, while the Windows script builds `dist/Excel Merge.exe`.

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

<a id="testing"></a>
## Testing

The repository already includes automated tests for the core comparison and merge pipeline under `tests/`.
GUI behavior is not covered by automated tests because it requires a graphical environment.

```bash
python -m unittest discover -s tests
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

Project history is maintained in [CHANGELOG.md](./CHANGELOG.md).
At the moment, the repository does not have published Git tags or GitHub Releases yet, so the changelog acts as the current source of truth for documented changes.

### `v0.1.0` - `2026-06-30`

- Initialize the project documentation.
- Add project structure, requirements, installation, and usage examples.
- Add GitHub-friendly badges, contribution guidance, and license placeholders.

## Customization Notes

Before publishing future releases, add Git tags and GitHub Releases if you want to restore release-related badges in this README.
