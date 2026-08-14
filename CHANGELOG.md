# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project follows [Semantic Versioning](https://semver.org/).

## [1.0.2] - 2026-08-14

### Fixed

- 修复列对齐错位：合并区域计入行列边界、列重排按列名对齐、自动选择真实表头行。
- 修复合并单元格在插入/删除行列后错位或丢失，覆盖写前自动解除阻碍合并。
- 修复批量合并时两侧列结构不同导致的错位，合并前先同步独占列。
- 修复仅新增/删除列时 Sheet 标签不显示红点的问题。
- 修复查看差异时字符级高亮 delegate 崩溃。
- 修复退出或换文件时未保存修改被静默丢弃的问题。
- 修复打包依赖与资源问题：spec 纳入版本控制、移除幽灵依赖、内置命令行启动脚本。

### Added

- 新增单元格内字符级差异高亮，只高亮差异片段。
- 新增双击未完全显示的单元格自动调整列宽与行高。
- 新增关闭窗口/覆盖打开文件时的未保存确认。

## [0.1.0] - 2026-06-30

### Added

- Added English and Simplified Chinese project documentation.
- Added GitHub-friendly repository badges and language switch links.
- Added quick start, requirements, installation, usage examples, and API reference sections.
- Added contribution guide, license section, contact section, and release notes structure.
- Added `LICENSE` and `CHANGELOG.md` as baseline repository documents.

### Notes

- This is the initial documented release baseline for the repository.

[0.1.0]: https://github.com/RedMaple96/ExcelMerge/releases/tag/v0.1.0
[1.0.2]: https://github.com/RedMaple96/ExcelMerge/releases/tag/v1.0.2
