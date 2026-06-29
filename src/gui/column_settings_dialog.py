"""列设置对话框模块 — ColumnSettingsDialog(QDialog)。

用于配置 Key 列（行匹配键）与忽略列（不参与差异判断的列）。
同一列不可同时作为 Key 与 Ignored —— 勾选一侧时自动取消另一侧。
对应需求：FR-04 / Task 8.3。
"""

from __future__ import annotations

from typing import List, Tuple

from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)


class ColumnSettingsDialog(QDialog):
    """列设置对话框：选择 Key 列与忽略列。

    构造参数：
    - header_labels: 列标签列表，如 ["A","B","C"]
    - current_key_cols: 当前已选 Key 列的 0-based 索引列表
    - current_ignore_cols: 当前已选忽略列的 0-based 索引列表
    """

    def __init__(
        self,
        header_labels: List[str],
        current_key_cols: List[int] = None,
        current_ignore_cols: List[int] = None,
        parent: QWidget = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("列设置")
        self.setMinimumWidth(420)

        current_key_cols = list(current_key_cols) if current_key_cols else []
        current_ignore_cols = (
            list(current_ignore_cols) if current_ignore_cols else []
        )

        # 最终结果（OK 时填充）
        self.key_cols: List[int] = []
        self.ignore_cols: List[int] = []

        # 每列对应的 (key_checkbox, ignore_checkbox)
        self._key_checks: List[QCheckBox] = []
        self._ignore_checks: List[QCheckBox] = []

        self._build_ui(header_labels, current_key_cols, current_ignore_cols)

    # ------------------------------------------------------------------ #
    # UI 构建
    # ------------------------------------------------------------------ #
    def _build_ui(
        self,
        header_labels: List[str],
        current_key_cols: List[int],
        current_ignore_cols: List[int],
    ) -> None:
        """构建对话框界面：Key 列组 + 忽略列组 + 按钮盒。"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        hint = QLabel(
            "Key 列用于行匹配；忽略列不参与差异判断。\n"
            "同一列不可同时勾选 Key 与忽略。"
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)

        n = len(header_labels)
        key_set = set(current_key_cols)
        ignore_set = set(current_ignore_cols)

        # 列多时用多列网格，更好利用横向空间
        if n <= 10:
            grid_cols = 1
        elif n <= 30:
            grid_cols = 2
        elif n <= 60:
            grid_cols = 3
        else:
            grid_cols = 4

        # ---- 内容容器：放进滚动区 ----
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(8)

        # ---- Key 列分组（网格）----
        key_group = QGroupBox("Key 列（用于行匹配）")
        key_grid = QGridLayout(key_group)
        key_grid.setSpacing(2)
        key_grid.setColumnStretch(grid_cols, 1)
        for i in range(n):
            label = header_labels[i]
            cb = QCheckBox(f"{label}  (第 {i + 1} 列)")
            cb.setChecked(i in key_set)
            # 同步：勾选 Key 时取消该列的 Ignore
            cb.toggled.connect(
                lambda checked, idx=i: self._on_key_toggled(idx, checked)
            )
            self._key_checks.append(cb)
            key_grid.addWidget(cb, i // grid_cols, i % grid_cols)
        content_layout.addWidget(key_group)

        # ---- 忽略列分组（网格）----
        ignore_group = QGroupBox("忽略列（Unimportant Columns）")
        ignore_grid = QGridLayout(ignore_group)
        ignore_grid.setSpacing(2)
        ignore_grid.setColumnStretch(grid_cols, 1)
        for i in range(n):
            label = header_labels[i]
            cb = QCheckBox(f"{label}  (第 {i + 1} 列)")
            cb.setChecked(i in ignore_set)
            # 同步：勾选 Ignore 时取消该列的 Key
            cb.toggled.connect(
                lambda checked, idx=i: self._on_ignore_toggled(idx, checked)
            )
            self._ignore_checks.append(cb)
            ignore_grid.addWidget(cb, i // grid_cols, i % grid_cols)
        content_layout.addWidget(ignore_group)

        # ---- 滚动区：真正包裹内容容器 ----
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(content)
        layout.addWidget(scroll, 1)

        # 限制最大高度，避免超出屏幕
        screen = QApplication.primaryScreen()
        if screen is not None:
            max_h = int(screen.availableGeometry().height() * 0.8)
            self.setMaximumHeight(max_h)

        # ---- 按钮盒 ----
        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel,
            parent=self,
        )
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    # ------------------------------------------------------------------ #
    # 互斥同步
    # ------------------------------------------------------------------ #
    def _on_key_toggled(self, idx: int, checked: bool) -> None:
        """勾选 Key 时，若该列已被忽略则取消忽略。"""
        if checked and idx < len(self._ignore_checks):
            ig = self._ignore_checks[idx]
            if ig.isChecked():
                # 屏蔽信号避免递归
                ig.blockSignals(True)
                ig.setChecked(False)
                ig.blockSignals(False)

    def _on_ignore_toggled(self, idx: int, checked: bool) -> None:
        """勾选 Ignore 时，若该列已是 Key 则取消 Key。"""
        if checked and idx < len(self._key_checks):
            kc = self._key_checks[idx]
            if kc.isChecked():
                kc.blockSignals(True)
                kc.setChecked(False)
                kc.blockSignals(False)

    # ------------------------------------------------------------------ #
    # 确认与结果读取
    # ------------------------------------------------------------------ #
    def _on_accept(self) -> None:
        """点击 OK 时收集结果并接受对话框。"""
        self.key_cols = [
            i for i, cb in enumerate(self._key_checks) if cb.isChecked()
        ]
        self.ignore_cols = [
            i for i, cb in enumerate(self._ignore_checks) if cb.isChecked()
        ]
        self.accept()

    def get_key_cols(self) -> List[int]:
        """返回用户选定的 Key 列索引列表（0-based）。"""
        return list(self.key_cols)

    def get_ignore_cols(self) -> List[int]:
        """返回用户选定的忽略列索引列表（0-based）。"""
        return list(self.ignore_cols)
