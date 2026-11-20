# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project follows [Semantic Versioning](https://semver.org/).

## [1.0.5] - 2026-11-20

### Added

- 新增 Sheet 级差异定位：工具栏“上一个表差异 / 下一个表差异”（Ctrl+PgUp / Ctrl+PgDown），按底部标签顺序在所有“有差异的 Sheet”之间循环跳转，两侧同步切换，状态栏显示 `Sheet 差异 n/m`；逐 Sheet 差异仍在后台预计算时提示“Sheet 差异计算中”，不误报无差异。

### Changed

- 两侧 Sheet 支持独立切换：手动切换任一侧下拉不再把对侧强制同步到同名 Sheet，可以做跨表比较；只在打开/替换文件与点击底部标签栏时两侧对齐到同一个 Sheet。
- 行级差异定位精确到单元格：上一个/下一个差异落在该行第一个真正有差异的单元格，不再停在行首；整行独占（仅左/仅右，含整表缺失）时落在该侧第一个非虚拟列。
- 未保存修改确认按受影响侧收窄：只切换单侧 Sheet 时不再被对侧未保存修改打断，也不会把对侧改动重载丢弃。
- 底部标签激活态改由主窗口在切换成功后设置，取消切换时高亮正确回退。

### Fixed

- 修复跨表比较结果被错记到单个 Sheet 标签红点的问题。

## [1.0.4] - 2026-08-27

### Added

- 新增 Sheet 级整表覆盖：右键菜单“复制当前整表到对侧（覆盖）”，用源侧当前工作表完整内容替换对侧同名工作表（保留对侧工作表顺序）。
- 新增文件级整体替换：编辑菜单“文件级替换：左侧覆盖右侧 / 右侧覆盖左侧”，用一侧工作簿全部内容（含所有工作表与顺序）完整替换另一侧，保存后写回被替换侧原路径。

### Fixed

- 修复底部栏 Sheet 快速切换列表首次点击切换失败、再点箭头误触右键菜单的问题。
- 修复双击单元格自适应后内容末尾被省略号截断的问题，加大列宽与行高边距。
- 修复 diff 高亮代理的文本换行显示问题。

### Changed

- Sheet 快速切换列表去除复选框，当前激活项改为加粗标识。
- 授权模式重构为在线激活。

### Added

- 新增整表复制/删除功能，优化底部标签右键菜单。
- 新增 Sheet 缺失（仅一侧存在）比较支持。

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
[1.0.4]: https://github.com/RedMaple96/ExcelMerge/releases/tag/v1.0.4
[1.0.3]: https://github.com/RedMaple96/ExcelMerge/releases/tag/v1.0.3
[1.0.5]: https://github.com/RedMaple96/ExcelMerge/releases/tag/v1.0.5
