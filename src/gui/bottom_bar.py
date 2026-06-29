"""底部导航栏模块 — SheetTab / BottomBar。

位于窗口最下方，提供：
- 左侧：当前单元格坐标（如 A77）
- 右侧：Sheet 标签栏（左右箭头按钮控制横向滑动，标签带差异红点标记）

对应需求：底部导航栏设计。
"""

from __future__ import annotations

from typing import List, Optional, Set

from PySide6.QtCore import QTimer, Qt, Signal
from PySide6.QtGui import QColor, QIcon, QMouseEvent, QPainter, QPixmap
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QMenu,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QWidget,
)

# 固定的图标尺寸，保持有无红点时文本对齐一致
_ICON_SIZE = 8


def _make_dot_icon(color: str, size: int = _ICON_SIZE) -> QIcon:
    """生成一个圆形实心图标，用于差异红点。"""
    pix = QPixmap(size, size)
    pix.fill(Qt.transparent)
    painter = QPainter(pix)
    painter.setRenderHint(QPainter.Antialiasing)
    painter.setBrush(QColor(color))
    painter.setPen(Qt.NoPen)
    painter.drawEllipse(0, 0, size, size)
    painter.end()
    return QIcon(pix)


def _make_transparent_icon(size: int = _ICON_SIZE) -> QIcon:
    """生成透明图标占位。"""
    pix = QPixmap(size, size)
    pix.fill(Qt.transparent)
    return QIcon(pix)


class SheetTab(QPushButton):
    """单个 Sheet 标签按钮。

    特性：
    - 左侧红点（有差异时显示红色，无差异时透明占位）
    - 可选中（checkable），激活态高亮
    - 双击触发重命名信号
    - 右键菜单提供「关闭」「关闭其他」
    """

    double_clicked = Signal(str)          # 双击 → (sheet_name)
    close_requested = Signal(str, bool)   # 右键关闭 → (sheet_name, close_others)

    def __init__(self, name: str, has_diff: bool = False, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._name = name
        self._has_diff = has_diff

        self.setCheckable(True)
        self.setText(name)
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedHeight(24)
        # 固定图标区域宽度，避免有无红点时文本跳动
        self.setIconSize(QPixmap(_ICON_SIZE, _ICON_SIZE).size())

        self._refresh_icon()
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._on_context_menu)

    @property
    def name(self) -> str:
        return self._name

    def set_name(self, name: str) -> None:
        self._name = name
        self.setText(name)

    def set_has_diff(self, has_diff: bool) -> None:
        if self._has_diff != has_diff:
            self._has_diff = has_diff
            self._refresh_icon()

    def _refresh_icon(self) -> None:
        if self._has_diff:
            self.setIcon(_make_dot_icon("#e53e3e"))
        else:
            self.setIcon(_make_transparent_icon())

    # ------------------------------------------------------------------ #
    # 事件
    # ------------------------------------------------------------------ #
    def mouseDoubleClickEvent(self, event: QMouseEvent) -> None:  # noqa: N802
        if event.button() == Qt.LeftButton:
            self.double_clicked.emit(self._name)
        super().mouseDoubleClickEvent(event)

    def _on_context_menu(self, pos) -> None:
        menu = QMenu(self)
        menu.addAction("关闭", lambda: self.close_requested.emit(self._name, False))
        menu.addAction("关闭其他", lambda: self.close_requested.emit(self._name, True))
        menu.exec(self.mapToGlobal(pos))


class BottomBar(QFrame):
    """底部导航栏。

    布局：[单元格: A77]  ┃  [◀] [ Sheet1 ●  Sheet2  Sheet3 ●  ⇢ ] [▶]

    Sheet 标签放在隐藏滚动条的 QScrollArea 内，标签按钮保持自身固有宽度
    并可在可视区域内横向滑动；左右箭头按钮控制滚动条位移，不改变界面窗口大小。

    信号：
    - sheet_activated(str):       点击标签 → sheet_name
    - sheet_renamed(str, str):    双击重命名 → (old_name, new_name)
    - sheet_close_requested(str, bool): 右键关闭 → (sheet_name, close_others)
    """

    sheet_activated = Signal(str)
    sheet_renamed = Signal(str, str)
    sheet_close_requested = Signal(str, bool)

    # 每次点击箭头滚动的步长（像素）
    _SCROLL_STEP = 120

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setObjectName("BottomBar")
        self.setFixedHeight(30)

        self._tabs: dict = {}          # name -> SheetTab
        self._order: List[str] = []    # 标签显示顺序
        self._active_name: str = ""

        layout = QHBoxLayout(self)
        layout.setContentsMargins(6, 0, 6, 0)
        layout.setSpacing(6)

        # ---- 左侧：坐标 ----
        self.coord_label = QLabel("单元格: -")
        self.coord_label.setMinimumWidth(90)
        layout.addWidget(self.coord_label)

        # ---- 分隔线 ----
        sep = QFrame()
        sep.setFrameShape(QFrame.VLine)
        sep.setFixedWidth(1)
        layout.addWidget(sep)

        # ---- 左箭头按钮 ----
        self.btn_prev = QPushButton("◀")
        self.btn_prev.setFixedSize(22, 22)
        self.btn_prev.setCursor(Qt.PointingHandCursor)
        self.btn_prev.setToolTip("向左滚动")
        self.btn_prev.clicked.connect(self._scroll_left)
        layout.addWidget(self.btn_prev)

        # ---- 标签滚动区域（隐藏滚动条）----
        # setWidgetResizable(True)：容器可拉伸；
        # 通过给容器设置 minimumWidth（=所有标签宽度之和）让内容撑开超出视口，
        # 从而产生可滚动范围。
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._scroll.setFrameShape(QFrame.NoFrame)
        self._scroll.setFixedHeight(26)

        self._tabs_container = QWidget()
        self._tabs_container.setObjectName("TabsContainer")
        self._tabs_layout = QHBoxLayout(self._tabs_container)
        self._tabs_layout.setContentsMargins(0, 0, 0, 0)
        self._tabs_layout.setSpacing(2)
        self._tabs_layout.addStretch()  # 末尾弹簧，使标签左对齐

        self._scroll.setWidget(self._tabs_container)
        layout.addWidget(self._scroll, 1)

        # 滚动条范围/位置变化时更新箭头启用状态
        hbar = self._scroll.horizontalScrollBar()
        hbar.rangeChanged.connect(self._update_nav_buttons)
        hbar.valueChanged.connect(self._update_nav_buttons)

        # ---- 右箭头按钮 ----
        self.btn_next = QPushButton("▶")
        self.btn_next.setFixedSize(22, 22)
        self.btn_next.setCursor(Qt.PointingHandCursor)
        self.btn_next.setToolTip("向右滚动")
        self.btn_next.clicked.connect(self._scroll_right)
        layout.addWidget(self.btn_next)

        self._update_nav_buttons()

    # ------------------------------------------------------------------ #
    # Sheet 标签管理
    # ------------------------------------------------------------------ #
    def set_sheets(self, names: List[str], diff_names: Set[str]) -> None:
        """重设全部 Sheet 标签。diff_names 中为有差异的 sheet 名。"""
        # 清除旧标签
        for tab in self._tabs.values():
            tab.setParent(None)
            tab.deleteLater()
        self._tabs.clear()
        self._order = []
        self._active_name = ""

        for name in names:
            self._create_tab(name, name in diff_names)

        self._update_container_width()

        # 重置滚动位置
        self._scroll.horizontalScrollBar().setValue(0)

        if names:
            self.set_active(names[0])

    def _create_tab(self, name: str, has_diff: bool) -> SheetTab:
        """创建单个标签并加入容器与索引。"""
        tab = SheetTab(name, has_diff)
        tab.clicked.connect(lambda checked=False, n=name: self._on_tab_clicked(n))
        tab.double_clicked.connect(self._on_tab_double_clicked)
        tab.close_requested.connect(
            lambda n, others: self.sheet_close_requested.emit(n, others)
        )
        self._tabs[name] = tab
        self._order.append(name)
        # 插在 stretch 之前
        self._tabs_layout.insertWidget(self._tabs_layout.count() - 1, tab)
        return tab

    def _update_container_width(self) -> None:
        """根据布局 sizeHint 设置容器最小宽度，使内容撑开产生滚动范围。

        QScrollArea 在 setWidgetResizable(True) 下，容器宽度至少为
        minimumWidth 与视口宽度中的较大者；当 minimumWidth > 视口宽度时，
        滚动条范围自动产生。

        由于子控件（按钮）的 sizeHint 在控件显示前可能尚未计算完成，
        这里立即设置一次，并额外延迟重试一次以确保生效。
        """
        hint = self._tabs_layout.sizeHint()
        self._tabs_container.setMinimumWidth(hint.width())
        # 延迟再设置一次：确保按钮完成布局后 sizeHint 返回有效值
        QTimer.singleShot(0, self._retry_update_container_width)

    def _retry_update_container_width(self) -> None:
        """延迟重试：按钮完成布局后重新计算容器最小宽度。"""
        hint = self._tabs_layout.sizeHint()
        self._tabs_container.setMinimumWidth(hint.width())
        self._update_nav_buttons()

    def add_sheet(self, name: str, has_diff: bool = False) -> None:
        """追加单个标签。"""
        if name in self._tabs:
            return
        self._create_tab(name, has_diff)
        self._update_container_width()

    def remove_sheet(self, name: str) -> None:
        """移除单个标签。"""
        tab = self._tabs.pop(name, None)
        if tab is None:
            return
        if name in self._order:
            self._order.remove(name)
        tab.setParent(None)
        tab.deleteLater()
        self._update_container_width()
        if self._active_name == name:
            self._active_name = ""
            if self._order:
                self.set_active(self._order[0])

    def set_active(self, name: str) -> None:
        """高亮指定标签，并确保该标签滚动到可视区域内。"""
        self._active_name = name
        for tab_name, tab in self._tabs.items():
            tab.setChecked(tab_name == name)
        # 确保激活标签可见
        tab = self._tabs.get(name)
        if tab is not None:
            self._ensure_visible(tab)

    def get_active(self) -> str:
        return self._active_name

    def update_diff_status(self, diff_names: Set[str]) -> None:
        """更新各标签的红点状态。"""
        for name, tab in self._tabs.items():
            tab.set_has_diff(name in diff_names)

    def rename_tab(self, old_name: str, new_name: str) -> None:
        """重命名标签。"""
        tab = self._tabs.pop(old_name, None)
        if tab is None:
            return
        tab.set_name(new_name)
        self._tabs[new_name] = tab
        if old_name in self._order:
            idx = self._order.index(old_name)
            self._order[idx] = new_name
        if self._active_name == old_name:
            self._active_name = new_name
        self._update_container_width()

    def get_sheet_names(self) -> List[str]:
        """返回当前标签栏中所有 sheet 名（按显示顺序）。"""
        return list(self._order)

    # ------------------------------------------------------------------ #
    # 坐标显示
    # ------------------------------------------------------------------ #
    def set_coord(self, text: str) -> None:
        self.coord_label.setText(f"单元格: {text}")

    # ------------------------------------------------------------------ #
    # 滚动控制
    # ------------------------------------------------------------------ #
    def _scroll_left(self) -> None:
        """向左滚动一个步长。"""
        hbar = self._scroll.horizontalScrollBar()
        hbar.setValue(hbar.value() - self._SCROLL_STEP)

    def _scroll_right(self) -> None:
        """向右滚动一个步长。"""
        hbar = self._scroll.horizontalScrollBar()
        hbar.setValue(hbar.value() + self._SCROLL_STEP)

    def _ensure_visible(self, tab: SheetTab) -> None:
        """滚动使指定标签完全可见。"""
        # 标签在容器中的左边界
        tab_left = tab.x()
        tab_right = tab_left + tab.width()
        hbar = self._scroll.horizontalScrollBar()
        cur = hbar.value()
        viewport_width = self._scroll.viewport().width()

        if tab_left < cur:
            # 标签在可视区左侧 → 左对齐
            hbar.setValue(tab_left)
        elif tab_right > cur + viewport_width:
            # 标签在可视区右侧 → 右对齐
            hbar.setValue(tab_right - viewport_width)

    def _update_nav_buttons(self, *args) -> None:
        """根据滚动条位置与范围，启用/禁用左右箭头。"""
        hbar = self._scroll.horizontalScrollBar()
        max_val = hbar.maximum()
        cur_val = hbar.value()
        # 没有溢出内容时两个按钮都禁用
        if max_val <= 0:
            self.btn_prev.setEnabled(False)
            self.btn_next.setEnabled(False)
            return
        self.btn_prev.setEnabled(cur_val > 0)
        self.btn_next.setEnabled(cur_val < max_val)

    # ------------------------------------------------------------------ #
    # 内部回调
    # ------------------------------------------------------------------ #
    def _on_tab_clicked(self, name: str) -> None:
        self.set_active(name)
        self.sheet_activated.emit(name)

    def _on_tab_double_clicked(self, name: str) -> None:
        new_name, ok = QInputDialog.getText(
            self, "重命名 Sheet", "新名称:", text=name
        )
        if ok and new_name and new_name != name:
            if new_name in self._tabs:
                QMessageBox.warning(self, "重命名失败", f"名称已存在: {new_name}")
                return
            self.rename_tab(name, new_name)
            self.sheet_renamed.emit(name, new_name)

    # ------------------------------------------------------------------ #
    # 事件
    # ------------------------------------------------------------------ #
    def showEvent(self, event) -> None:  # noqa: N802
        """控件首次显示时，重新计算容器宽度（此时子控件 sizeHint 已就绪）。"""
        super().showEvent(event)
        QTimer.singleShot(0, self._retry_update_container_width)

    def resizeEvent(self, event) -> None:  # noqa: N802
        """尺寸变化时刷新箭头启用状态。"""
        super().resizeEvent(event)
        self._update_nav_buttons()
