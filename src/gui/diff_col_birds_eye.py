"""全局差异缩略图 / 导航条（列方向）— DiffColBirdsEyeView(QWidget)。

位于双表格下方的极窄横条，作为整个大表（可能很多列）的"缩略地图"：

- 彩色刻度条标记每一处差异列的相对位置（红=different, 蓝=left_only, 绿=right_only）；
- 半透明矩形指示主视图当前可视列区域；
- 点击彩色区域可让主视图滚动到对应差异列（吸附最近差异列）；
- 拖动可快速翻页定位，避免在大宽表中盲目滚动寻找差异。

视觉映射：差异列索引 i 在缩略图中的横坐标为 ``i / total_cols * W``，
保证缩略图与主表列序一一对应。差异刻度缓存为 QPixmap，滚动时仅重绘
视口指示器，数千列差异也能保持流畅。
"""

from __future__ import annotations

from typing import Optional

from PySide6.QtCore import Qt, QRect, QSize
from PySide6.QtGui import QBrush, QColor, QMouseEvent, QPainter, QPen, QPixmap
from PySide6.QtWidgets import QAbstractItemView, QTableWidget, QWidget

from src.gui.themes import is_dark_mode


class DiffColBirdsEyeView(QWidget):
    """全局差异缩略图导航条（列方向，横向）。

    通过 ``set_diff_cols`` 注入差异列索引（aligned_cols 下标），
    自动随所监听表格的水平滚动刷新视口指示器。
    """

    FIXED_HEIGHT = 18  # 极窄横条高度
    SNAP_THRESHOLD_PX = 8  # 点击吸附到最近差异列的像素阈值

    def __init__(self, table: QTableWidget, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._table = table
        self._diff_cols: list[int] = []
        self._col_types: list[str] = []
        self._hover_x: Optional[int] = None
        self._marks_pixmap: Optional[QPixmap] = None
        self._marks_dirty: bool = True
        self._marks_total: int = -1  # 上次绘制刻度时的总列数，用于检测列数变化

        self.setFixedHeight(self.FIXED_HEIGHT)
        self.setMinimumWidth(40)
        self.setCursor(Qt.PointingHandCursor)
        self.setMouseTracking(True)
        self.setToolTip(
            "差异缩略图：彩色线标记差异列位置\n"
            "点击彩色线可跳转到对应差异列\n"
            "拖动可快速滚动定位"
        )

        # 监听表格水平滚动 / 范围变化，刷新视口指示器
        hbar = self._table.horizontalScrollBar()
        hbar.valueChanged.connect(lambda _v: self.update())
        hbar.rangeChanged.connect(lambda *_a: self.update())

    # ------------------------------------------------------------------ #
    # 对外接口
    # ------------------------------------------------------------------ #
    def set_diff_cols(self, cols: list[int], col_types: Optional[list[str]] = None) -> None:
        """设置差异列索引列表（aligned_cols 下标），触发重绘。

        col_types: 可选，每列对应的类型 ("different"|"left_only"|"right_only")，
                   用于缩略图中用不同颜色区分。长度应与 cols 一致。
                   未提供时全部按 "different" 处理（红色）。
        """
        self._diff_cols = list(cols)
        self._col_types = list(col_types) if col_types else ["different"] * len(cols)
        self._marks_dirty = True
        self.update()

    def clear(self) -> None:
        """清空差异标记。"""
        self._diff_cols = []
        self._col_types = []
        self._marks_dirty = True
        self.update()

    # ------------------------------------------------------------------ #
    # 主题配色
    # ------------------------------------------------------------------ #
    def _colors(self) -> dict:
        """根据当前主题返回缩略图配色。

        刻度颜色按差异类型区分（与表格着色一致）：
        - different: 红色
        - left_only: 蓝色
        - right_only: 绿色
        """
        dark = is_dark_mode()
        if dark:
            return {
                "bg": QColor("#252525"),
                "border": QColor("#3a3a3a"),
                "mark_diff": QColor("#FF5252"),
                "mark_left": QColor("#5B9BD5"),
                "mark_right": QColor("#5DCAA5"),
                "viewport_fill": QColor(255, 255, 255, 30),
                "viewport_border": QColor(255, 255, 255, 100),
                "hover": QColor(255, 255, 255, 60),
            }
        return {
            "bg": QColor("#fafafa"),
            "border": QColor("#d0d0d0"),
            "mark_diff": QColor("#E53935"),
            "mark_left": QColor("#185FA5"),
            "mark_right": QColor("#0F6E56"),
            "viewport_fill": QColor(80, 80, 80, 30),
            "viewport_border": QColor(80, 80, 80, 140),
            "hover": QColor(0, 0, 0, 40),
        }

    # ------------------------------------------------------------------ #
    # 绘制
    # ------------------------------------------------------------------ #
    def _build_marks_pixmap(self, colors: dict) -> Optional[QPixmap]:
        """将差异刻度条预渲染到 QPixmap（滚动时直接 blit，避免重复绘制）。

        按列类型用不同颜色绘制：different=红、left_only=蓝、right_only=绿。
        """
        total = self._table.columnCount()
        w = self.width()
        h = self.height()
        if total <= 0 or w <= 0 or h <= 0:
            self._marks_total = total
            return None
        pm = QPixmap(w, h)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing, False)
        p.setPen(Qt.NoPen)
        mark_y = 2
        mark_h = h - 4
        for idx, i in enumerate(self._diff_cols):
            if i < 0 or i >= total:
                continue
            # 按类型选色
            ctype = self._col_types[idx] if idx < len(self._col_types) else "different"
            if ctype == "left_only":
                p.setBrush(QBrush(colors["mark_left"]))
            elif ctype == "right_only":
                p.setBrush(QBrush(colors["mark_right"]))
            else:
                p.setBrush(QBrush(colors["mark_diff"]))
            x1 = int(i / total * w)
            x2 = int((i + 1) / total * w)
            if x2 <= x1:
                x2 = x1 + 1
            p.drawRect(QRect(x1, mark_y, x2 - x1, mark_h))
        p.end()
        self._marks_total = total
        return pm

    def _draw_viewport(self, painter: QPainter, colors: dict) -> None:
        """绘制当前可视列区域指示矩形。"""
        total = self._table.columnCount()
        w = self.width()
        if total <= 0 or w <= 0:
            return
        vw = self._table.viewport().width()
        first = self._table.columnAt(0)
        last = self._table.columnAt(max(0, vw - 1))
        if first < 0:
            first = 0
        if last < 0:
            last = total - 1
        x1 = int(first / total * w)
        x2 = int((last + 1) / total * w)
        if x2 <= x1:
            x2 = x1 + 1
        painter.setBrush(QBrush(colors["viewport_fill"]))
        painter.setPen(QPen(colors["viewport_border"], 1))
        painter.drawRect(QRect(x1, 0, x2 - x1, self.height() - 1))

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, False)
        colors = self._colors()

        # 背景
        painter.fillRect(self.rect(), colors["bg"])

        total = self._table.columnCount()
        # 列数变化时强制重建缓存
        if total != self._marks_total:
            self._marks_dirty = True

        # 差异刻度条（缓存）+ 视口指示器
        if total > 0:
            if self._marks_dirty or self._marks_pixmap is None:
                self._marks_pixmap = self._build_marks_pixmap(colors)
                self._marks_dirty = False
            if self._marks_pixmap is not None:
                painter.drawPixmap(0, 0, self._marks_pixmap)
            self._draw_viewport(painter, colors)

        # 悬停高亮线
        if self._hover_x is not None and 0 <= self._hover_x < self.width():
            painter.setPen(QPen(colors["hover"], 1))
            painter.drawLine(self._hover_x, 0, self._hover_x, self.height() - 1)

        # 外边框
        painter.setPen(QPen(colors["border"], 1))
        painter.setBrush(Qt.NoBrush)
        painter.drawRect(self.rect().adjusted(0, 0, -1, -1))

    # ------------------------------------------------------------------ #
    # 交互
    # ------------------------------------------------------------------ #
    def _col_at_x(self, x: int) -> int:
        """将缩略图横坐标映射回表格列索引。"""
        total = self._table.columnCount()
        w = self.width()
        if total <= 0 or w <= 0:
            return -1
        col = int(x / w * total)
        return max(0, min(col, total - 1))

    def _nearest_diff_within(self, x: int) -> Optional[int]:
        """返回与横坐标 x 距离在 SNAP_THRESHOLD_PX 内的最近差异列索引，无则 None。"""
        if not self._diff_cols:
            return None
        total = self._table.columnCount()
        w = self.width()
        if total <= 0 or w <= 0:
            return None
        best: Optional[int] = None
        best_dist: Optional[float] = None
        for d in self._diff_cols:
            center = (d + 0.5) / total * w
            dist = abs(center - x)
            if dist <= self.SNAP_THRESHOLD_PX:
                if best_dist is None or dist < best_dist:
                    best_dist = dist
                    best = d
        return best

    def _scroll_to_col(self, col: int, select: bool = False) -> None:
        """滚动主表格至指定列（居中），可选同步选中。"""
        tbl = self._table
        if col < 0 or col >= tbl.columnCount():
            return
        item = tbl.item(0, col)
        if item is not None:
            tbl.scrollToItem(item, QAbstractItemView.ScrollHint.PositionAtCenter)
        else:
            model = tbl.model()
            if model is not None:
                tbl.scrollTo(model.index(0, col), QAbstractItemView.ScrollHint.PositionAtCenter)
        # 仅在列未被过滤隐藏时同步选中
        if select and not tbl.isColumnHidden(col):
            tbl.setCurrentCell(0, col)

    def mousePressEvent(self, event: QMouseEvent) -> None:  # noqa: N802
        if event.button() != Qt.LeftButton:
            return
        x = int(event.position().x())
        # 点击彩色区域：吸附到最近的差异列并选中
        nearest = self._nearest_diff_within(x)
        if nearest is not None:
            self._scroll_to_col(nearest, select=True)
            return
        # 否则滚动到对应比例位置（不选中，避免干扰）
        col = self._col_at_x(x)
        if col >= 0:
            self._scroll_to_col(col, select=False)

    def mouseMoveEvent(self, event: QMouseEvent) -> None:  # noqa: N802
        x = int(event.position().x())
        # 拖动滚动（按比例定位）
        if event.buttons() & Qt.LeftButton:
            col = self._col_at_x(x)
            if col >= 0:
                self._scroll_to_col(col, select=False)
        # 悬停指示
        if x != self._hover_x:
            self._hover_x = x
            self.update()

    def leaveEvent(self, event) -> None:  # noqa: N802
        if self._hover_x is not None:
            self._hover_x = None
            self.update()

    def resizeEvent(self, event) -> None:  # noqa: N802
        self._marks_dirty = True
        super().resizeEvent(event)

    def sizeHint(self) -> QSize:  # noqa: N802
        return QSize(200, self.FIXED_HEIGHT)
