"""全局差异缩略图 / 导航条 — DiffBirdsEyeView(QWidget)。

位于主视图最左侧的极窄纵列，作为整个大表（可能几千行）的"缩略地图"：

- 红色刻度条标记每一处差异行的相对位置；
- 半透明矩形指示主视图当前可视区域；
- 点击红色区域可让主视图滚动到对应差异行（吸附最近差异行）；
- 拖动可快速翻页定位，避免在大文件中盲目滚动寻找差异。

视觉映射：差异行索引 i 在缩略图中的纵坐标为 ``i / total_rows * H``，
保证缩略图与主表行序一一对应。差异刻度缓存为 QPixmap，滚动时仅重绘
视口指示器，数千行差异也能保持流畅。

对应需求：全局差异缩略图/导航条 (Global Diff Bird's-Eye View)。
"""

from __future__ import annotations

from typing import Optional

from PySide6.QtCore import Qt, QRect, QSize
from PySide6.QtGui import QBrush, QColor, QMouseEvent, QPainter, QPen, QPixmap
from PySide6.QtWidgets import QAbstractItemView, QTableWidget, QWidget

from src.gui.themes import is_dark_mode


class DiffBirdsEyeView(QWidget):
    """全局差异缩略图导航条。

    通过 ``set_diff_rows`` 注入差异行索引（aligned_rows 下标），
    自动随所监听表格的滚动刷新视口指示器。
    """

    FIXED_WIDTH = 18  # 极窄纵列宽度
    SNAP_THRESHOLD_PX = 8  # 点击吸附到最近差异行的像素阈值

    def __init__(self, table: QTableWidget, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._table = table
        self._diff_rows: list[int] = []
        self._hover_y: Optional[int] = None
        self._marks_pixmap: Optional[QPixmap] = None
        self._marks_dirty: bool = True
        self._marks_total: int = -1  # 上次绘制刻度时的总行数，用于检测行数变化

        self.setFixedWidth(self.FIXED_WIDTH)
        self.setMinimumHeight(40)
        self.setCursor(Qt.PointingHandCursor)
        self.setMouseTracking(True)
        self.setToolTip(
            "差异缩略图：红线标记差异行位置\n"
            "点击红线可跳转到对应差异行\n"
            "拖动可快速滚动定位"
        )

        # 监听表格滚动 / 范围变化，刷新视口指示器（lambda 忽略参数以匹配 update() 签名）
        vbar = self._table.verticalScrollBar()
        vbar.valueChanged.connect(lambda _v: self.update())
        vbar.rangeChanged.connect(lambda *_a: self.update())

    # ------------------------------------------------------------------ #
    # 对外接口
    # ------------------------------------------------------------------ #
    def set_diff_rows(self, rows: list[int]) -> None:
        """设置差异行索引列表（aligned_rows 下标），触发重绘。"""
        self._diff_rows = list(rows)
        self._marks_dirty = True
        self.update()

    def clear(self) -> None:
        """清空差异标记。"""
        self._diff_rows = []
        self._marks_dirty = True
        self.update()

    # ------------------------------------------------------------------ #
    # 主题配色
    # ------------------------------------------------------------------ #
    def _colors(self) -> dict:
        """根据当前主题返回缩略图配色。"""
        dark = is_dark_mode()
        if dark:
            return {
                "bg": QColor("#252525"),
                "border": QColor("#3a3a3a"),
                "mark": QColor("#FF5252"),
                "viewport_fill": QColor(255, 255, 255, 30),
                "viewport_border": QColor(255, 255, 255, 100),
                "hover": QColor(255, 255, 255, 60),
            }
        return {
            "bg": QColor("#fafafa"),
            "border": QColor("#d0d0d0"),
            "mark": QColor("#E53935"),
            "viewport_fill": QColor(80, 80, 80, 30),
            "viewport_border": QColor(80, 80, 80, 140),
            "hover": QColor(0, 0, 0, 40),
        }

    # ------------------------------------------------------------------ #
    # 绘制
    # ------------------------------------------------------------------ #
    def _build_marks_pixmap(self, colors: dict) -> Optional[QPixmap]:
        """将差异刻度条预渲染到 QPixmap（滚动时直接 blit，避免重复绘制）。"""
        total = self._table.rowCount()
        w = self.width()
        h = self.height()
        if total <= 0 or w <= 0 or h <= 0:
            self._marks_total = total
            return None
        pm = QPixmap(w, h)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing, False)
        p.setBrush(QBrush(colors["mark"]))
        p.setPen(Qt.NoPen)
        mark_x = 2
        mark_w = w - 4
        for i in self._diff_rows:
            if i < 0 or i >= total:
                continue
            y1 = int(i / total * h)
            y2 = int((i + 1) / total * h)
            if y2 <= y1:  # 保证至少 1px 高，相邻差异自然合并
                y2 = y1 + 1
            p.drawRect(QRect(mark_x, y1, mark_w, y2 - y1))
        p.end()
        self._marks_total = total
        return pm

    def _draw_viewport(self, painter: QPainter, colors: dict) -> None:
        """绘制当前可视区域指示矩形。"""
        total = self._table.rowCount()
        h = self.height()
        if total <= 0 or h <= 0:
            return
        vh = self._table.viewport().height()
        first = self._table.rowAt(0)
        last = self._table.rowAt(max(0, vh - 1))
        if first < 0:
            first = 0
        if last < 0:
            last = total - 1
        y1 = int(first / total * h)
        y2 = int((last + 1) / total * h)
        if y2 <= y1:
            y2 = y1 + 1
        painter.setBrush(QBrush(colors["viewport_fill"]))
        painter.setPen(QPen(colors["viewport_border"], 1))
        painter.drawRect(QRect(0, y1, self.width() - 1, y2 - y1))

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, False)
        colors = self._colors()

        # 背景
        painter.fillRect(self.rect(), colors["bg"])

        total = self._table.rowCount()
        # 行数变化时强制重建缓存
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
        if self._hover_y is not None and 0 <= self._hover_y < self.height():
            painter.setPen(QPen(colors["hover"], 1))
            painter.drawLine(0, self._hover_y, self.width() - 1, self._hover_y)

        # 外边框
        painter.setPen(QPen(colors["border"], 1))
        painter.setBrush(Qt.NoBrush)
        painter.drawRect(self.rect().adjusted(0, 0, -1, -1))

    # ------------------------------------------------------------------ #
    # 交互
    # ------------------------------------------------------------------ #
    def _row_at_y(self, y: int) -> int:
        """将缩略图纵坐标映射回表格行索引。"""
        total = self._table.rowCount()
        h = self.height()
        if total <= 0 or h <= 0:
            return -1
        row = int(y / h * total)
        return max(0, min(row, total - 1))

    def _nearest_diff_within(self, y: int) -> Optional[int]:
        """返回与纵坐标 y 距离在 SNAP_THRESHOLD_PX 内的最近差异行索引，无则 None。"""
        if not self._diff_rows:
            return None
        total = self._table.rowCount()
        h = self.height()
        if total <= 0 or h <= 0:
            return None
        best: Optional[int] = None
        best_dist: Optional[float] = None
        for d in self._diff_rows:
            center = (d + 0.5) / total * h
            dist = abs(center - y)
            if dist <= self.SNAP_THRESHOLD_PX:
                if best_dist is None or dist < best_dist:
                    best_dist = dist
                    best = d
        return best

    def _scroll_to_row(self, row: int, select: bool = False) -> None:
        """滚动主表格至指定行（居中），可选同步选中。"""
        tbl = self._table
        if row < 0 or row >= tbl.rowCount():
            return
        item = tbl.item(row, 0)
        if item is not None:
            tbl.scrollToItem(item, QAbstractItemView.ScrollHint.PositionAtCenter)
        else:
            model = tbl.model()
            if model is not None:
                tbl.scrollTo(model.index(row, 0), QAbstractItemView.ScrollHint.PositionAtCenter)
        # 仅在行未被过滤隐藏时同步选中
        if select and not tbl.isRowHidden(row):
            tbl.setCurrentCell(row, 0)

    def mousePressEvent(self, event: QMouseEvent) -> None:  # noqa: N802
        if event.button() != Qt.LeftButton:
            return
        y = int(event.position().y())
        # 点击红色区域：吸附到最近的差异行并选中
        nearest = self._nearest_diff_within(y)
        if nearest is not None:
            self._scroll_to_row(nearest, select=True)
            return
        # 否则滚动到对应比例位置（不选中，避免干扰）
        row = self._row_at_y(y)
        if row >= 0:
            self._scroll_to_row(row, select=False)

    def mouseMoveEvent(self, event: QMouseEvent) -> None:  # noqa: N802
        y = int(event.position().y())
        # 拖动滚动（按比例定位）
        if event.buttons() & Qt.LeftButton:
            row = self._row_at_y(y)
            if row >= 0:
                self._scroll_to_row(row, select=False)
        # 悬停指示
        if y != self._hover_y:
            self._hover_y = y
            self.update()

    def leaveEvent(self, event) -> None:  # noqa: N802
        if self._hover_y is not None:
            self._hover_y = None
            self.update()

    def resizeEvent(self, event) -> None:  # noqa: N802
        self._marks_dirty = True
        super().resizeEvent(event)

    def sizeHint(self) -> QSize:  # noqa: N802
        return QSize(self.FIXED_WIDTH, 200)
