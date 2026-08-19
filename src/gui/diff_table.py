"""Model/View 版差异表格：替代 QTableWidget 的逐格物化实现。

- DiffTableModel: QAbstractTableModel，文本/背景/字体/高亮区间/折行标记
  按字典存储，仅保存实际有值的格式，内存占用与差异规模成正比而非行列积。
- DiffTableView: QTableView 子类，兼容主窗口对 QTableWidget 的既有调用面
  （item/setItem/selectedItems/currentItemChanged/cellDoubleClicked 等），
  使 GUI 层无需大改即可获得 Model/View 的按需渲染收益。
- 单元格代理 CellRef 提供 text/setBackground/setData/setFont/row/column，
  与 QTableWidgetItem 在本项目中的用法子集完全一致。
"""

from __future__ import annotations

from typing import Dict, List, Optional, Set, Tuple

from PySide6.QtCore import (
    QAbstractTableModel,
    QModelIndex,
    Qt,
    Signal,
)
from PySide6.QtGui import QBrush, QFont
from PySide6.QtWidgets import QTableView, QTableWidgetItem

from src.gui.diff_highlight_delegate import CELL_WRAP_ROLE, HIGHLIGHT_RANGES_ROLE

# 虚拟空行列标记（与旧 QTableWidgetItem UserRole 数据一致）
VIRTUAL_ROLE = Qt.ItemDataRole.UserRole
VIRTUAL_MARK = "__virtual__"


class DiffTableModel(QAbstractTableModel):
    """差异表格数据模型：行列网格文本 + 稀疏格式字典。"""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._texts: List[List[Optional[str]]] = []
        self._n_cols: int = 0
        self._h_labels: List[str] = []
        self._v_labels: List[str] = []
        self._backgrounds: Dict[Tuple[int, int], QBrush] = {}
        self._fonts: Dict[Tuple[int, int], QFont] = {}
        self._highlights: Dict[Tuple[int, int], list] = {}
        self._wraps: Set[Tuple[int, int]] = set()
        self._virtuals: Set[Tuple[int, int]] = set()
        self._batch = False  # 批量填充模式下抑制逐格 dataChanged

    # ---------------- 基础接口 ----------------
    def rowCount(self, parent=QModelIndex()) -> int:  # noqa: N802
        return 0 if parent.isValid() else len(self._texts)

    def columnCount(self, parent=QModelIndex()) -> int:  # noqa: N802
        return 0 if parent.isValid() else self._n_cols

    def headerData(self, section, orientation, role=Qt.DisplayRole):  # noqa: N802
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal:
            if 0 <= section < len(self._h_labels):
                return self._h_labels[section]
            return ""
        if 0 <= section < len(self._v_labels):
            return self._v_labels[section]
        return ""

    def data(self, index: QModelIndex, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        r, c = index.row(), index.column()
        key = (r, c)
        if role in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.EditRole,
                    Qt.ItemDataRole.ToolTipRole):
            return self._texts[r][c]
        if role == Qt.ItemDataRole.BackgroundRole:
            return self._backgrounds.get(key)
        if role == Qt.ItemDataRole.FontRole:
            return self._fonts.get(key)
        if role == HIGHLIGHT_RANGES_ROLE:
            return self._highlights.get(key)
        if role == CELL_WRAP_ROLE:
            return True if key in self._wraps else None
        if role == VIRTUAL_ROLE:
            return VIRTUAL_MARK if key in self._virtuals else None
        return None

    # ---------------- 尺寸调整 ----------------
    def setRowCount(self, n: int) -> None:  # noqa: N802
        """调整行数：截断或扩展（扩展行内容为 None=未设置）。"""
        if n < len(self._texts):
            self.beginRemoveRows(QModelIndex(), n, len(self._texts) - 1)
            self._texts = self._texts[:n]
            self._v_labels = self._v_labels[:n]
            self._shrink_formats(n, self._n_cols)
            self.endRemoveRows()
        elif n > len(self._texts):
            self.beginInsertRows(QModelIndex(), len(self._texts), n - 1)
            new_rows = [[None] * self._n_cols for _ in range(n - len(self._texts))]
            self._texts.extend(new_rows)
            self._v_labels.extend([""] * (n - len(self._v_labels)))
            self.endInsertRows()

    def setColumnCount(self, n: int) -> None:  # noqa: N802
        if n == self._n_cols:
            return
        if n < self._n_cols:
            self.beginRemoveColumns(QModelIndex(), n, self._n_cols - 1)
            self._n_cols = n
            self._h_labels = self._h_labels[:n]
            for row in self._texts:
                if row is not None:
                    del row[n:]
            self._shrink_formats(len(self._texts), n)
            self.endRemoveColumns()
        else:
            self.beginInsertColumns(QModelIndex(), self._n_cols, n - 1)
            self._h_labels.extend([""] * (n - len(self._h_labels)))
            # 同步扩展已有行的内层列表，保持网格矩形完整
            for row in self._texts:
                if row is not None:
                    row.extend([None] * (n - len(row)))
            self._n_cols = n
            self.endInsertColumns()

    def _shrink_formats(self, n_rows: int, n_cols: int) -> None:
        for store in (self._backgrounds, self._fonts, self._highlights):
            for key in [k for k in store if k[0] >= n_rows or k[1] >= n_cols]:
                del store[key]
        self._wraps = {
            k for k in self._wraps if k[0] < n_rows and k[1] < n_cols
        }
        self._virtuals = {
            k for k in self._virtuals if k[0] < n_rows and k[1] < n_cols
        }

    # ---------------- 数据写入 ----------------
    def begin_batch(self) -> None:
        """开启批量填充：抑制逐格 dataChanged，结束后统一刷新。"""
        self._batch = True

    def end_batch(self) -> None:
        """结束批量填充：整体发一次 dataChanged + 布局刷新。"""
        self._batch = False
        if self._texts:
            self.dataChanged.emit(
                self.index(0, 0),
                self.index(len(self._texts) - 1, max(0, self._n_cols - 1)),
            )
        self.layoutChanged.emit()
    def setHorizontalHeaderLabels(self, labels: List[str]) -> None:  # noqa: N802
        self._h_labels = list(labels)
        self.headerDataChanged.emit(
            Qt.Orientation.Horizontal, 0, max(0, self._n_cols - 1)
        )

    def setVerticalHeaderLabels(self, labels: List[str]) -> None:  # noqa: N802
        self._v_labels = list(labels)
        self.headerDataChanged.emit(
            Qt.Orientation.Vertical, 0, max(0, len(self._texts) - 1)
        )

    def set_item(self, r: int, c: int, text: Optional[str],
                 virtual: bool = False) -> None:
        """设置单元格文本（None 表示清除为未设置状态）。"""
        if not (0 <= r < len(self._texts) and 0 <= c < self._n_cols):
            return
        self._texts[r][c] = text
        if virtual:
            self._virtuals.add((r, c))
        elif (r, c) in self._virtuals:
            self._virtuals.discard((r, c))
        if not self._batch:
            ix = self.index(r, c)
            self.dataChanged.emit(ix, ix)

    def set_background(self, r: int, c: int, brush: Optional[QBrush]) -> None:
        if not (0 <= r < len(self._texts) and 0 <= c < self._n_cols):
            return
        if brush is None:
            self._backgrounds.pop((r, c), None)
        else:
            self._backgrounds[(r, c)] = brush
        ix = self.index(r, c)
        self.dataChanged.emit(ix, ix)

    def set_font(self, r: int, c: int, font: Optional[QFont]) -> None:
        if not (0 <= r < len(self._texts) and 0 <= c < self._n_cols):
            return
        if font is None:
            self._fonts.pop((r, c), None)
        else:
            self._fonts[(r, c)] = font
        ix = self.index(r, c)
        self.dataChanged.emit(ix, ix)

    def set_highlight(self, r: int, c: int, ranges) -> None:
        if not (0 <= r < len(self._texts) and 0 <= c < self._n_cols):
            return
        if ranges:
            self._highlights[(r, c)] = ranges
        else:
            self._highlights.pop((r, c), None)
        ix = self.index(r, c)
        self.dataChanged.emit(ix, ix)

    def set_wrap(self, r: int, c: int, on: bool) -> None:
        if not (0 <= r < len(self._texts) and 0 <= c < self._n_cols):
            return
        if on:
            self._wraps.add((r, c))
        else:
            self._wraps.discard((r, c))
        ix = self.index(r, c)
        self.dataChanged.emit(ix, ix)

    def set_item_data(self, r: int, c: int, role, value) -> None:
        """按角色写入单元格数据（setData 代理入口）。"""
        if role == Qt.ItemDataRole.BackgroundRole:
            if isinstance(value, QBrush):
                self.set_background(r, c, value)
            elif isinstance(value, Qt.GlobalColor) or value is None:
                self.set_background(r, c, None)
        elif role == Qt.ItemDataRole.FontRole:
            self.set_font(r, c, value)
        elif role == HIGHLIGHT_RANGES_ROLE:
            self.set_highlight(r, c, value)
        elif role == CELL_WRAP_ROLE:
            self.set_wrap(r, c, bool(value))
        elif role == VIRTUAL_ROLE:
            if value == VIRTUAL_MARK:
                self._virtuals.add((r, c))
            else:
                self._virtuals.discard((r, c))
        # 其余角色不存储（项目未使用）


class CellRef:
    """QTableWidgetItem 兼容代理：指向模型中的一个单元格。"""

    __slots__ = ("_view", "_row", "_col")

    def __init__(self, view: "DiffTableView", row: int, col: int) -> None:
        self._view = view
        self._row = row
        self._col = col

    def row(self) -> int:
        return self._row

    def column(self) -> int:
        return self._col

    def text(self) -> str:
        val = self._view.model().data(
            self._view.model().index(self._row, self._col),
            Qt.ItemDataRole.DisplayRole,
        )
        return val if isinstance(val, str) else ""

    def setBackground(self, brush) -> None:  # noqa: N802
        self._view.model().set_background(self._row, self._col, brush)

    def background(self) -> Optional[QBrush]:
        """读取背景（QTableWidgetItem 兼容，测试与查询用）。"""
        return self._view.model().data(
            self._view.model().index(self._row, self._col),
            Qt.ItemDataRole.BackgroundRole,
        ) or QBrush()

    def setFont(self, font) -> None:  # noqa: N802
        self._view.model().set_font(self._row, self._col, font)

    def setData(self, role, value) -> None:  # noqa: N802
        self._view.model().set_item_data(self._row, self._col, role, value)

    def data(self, role):
        return self._view.model().data(
            self._view.model().index(self._row, self._col), role
        )


class DiffTableView(QTableView):
    """QTableWidget 兼容面封装的 QTableView。"""

    # 兼容 QTableWidget 信号签名
    currentItemChanged = Signal(object, object)  # (CellRef, CellRef)
    cellDoubleClicked = Signal(int, int)         # (row, col)
    cellChanged = Signal(int, int)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._model = DiffTableModel(self)
        self.setModel(self._model)
        # 信号桥接：selectionModel 在 setModel 后有效
        self.selectionModel().currentChanged.connect(self._bridge_current)
        self.doubleClicked.connect(
            lambda ix: self.cellDoubleClicked.emit(ix.row(), ix.column())
        )

    # ---- 兼容 QTableWidget 的方法 ----
    def item(self, r: int, c: int) -> Optional[CellRef]:
        m = self._model
        if 0 <= r < m.rowCount() and 0 <= c < m.columnCount():
            if m.data(m.index(r, c), Qt.ItemDataRole.DisplayRole) is not None:
                return CellRef(self, r, c)
        return None

    def setItem(self, r: int, c: int, qtwi: QTableWidgetItem) -> None:  # noqa: N802
        """接受 QTableWidgetItem（取其 text/UserRole），避免调用方改动。"""
        virtual = (
            qtwi.data(VIRTUAL_ROLE) == VIRTUAL_MARK
            if isinstance(qtwi, QTableWidgetItem) else False
        )
        text = qtwi.text() if isinstance(qtwi, QTableWidgetItem) else None
        font = qtwi.font() if isinstance(qtwi, QTableWidgetItem) else None
        self._model.set_item(r, c, text, virtual)
        if font is not None:
            self._model.set_font(r, c, font)

    def setCellWidget(self, *a):  # noqa: N802, D102 - 未使用，占位
        raise NotImplementedError

    def beginPopulate(self) -> None:  # noqa: N802
        """进入批量填充：抑制逐格刷新，配合 endPopulate 统一刷新。"""
        self._model.begin_batch()

    def endPopulate(self) -> None:  # noqa: N802
        """结束批量填充：统一刷新整个视图。"""
        self._model.end_batch()

    def setCellText(self, r: int, c: int, text: str,
                    virtual: bool = False) -> None:
        """直接写入单元格文本（避免构造 QTableWidgetItem 的逐格开销）。"""
        self._model.set_item(r, c, text, virtual)

    def setRowCount(self, n: int) -> None:  # noqa: N802
        self._model.setRowCount(n)

    def rowCount(self) -> int:  # noqa: N802
        return self._model.rowCount()

    def setColumnCount(self, n: int) -> None:  # noqa: N802
        self._model.setColumnCount(n)

    def columnCount(self) -> int:  # noqa: N802
        return self._model.columnCount()

    def setHorizontalHeaderLabels(self, labels) -> None:  # noqa: N802
        self._model.setHorizontalHeaderLabels(list(labels))

    def setVerticalHeaderLabels(self, labels) -> None:  # noqa: N802
        self._model.setVerticalHeaderLabels(list(labels))

    def selectedItems(self) -> List[CellRef]:  # noqa: N802
        return [
            CellRef(self, ix.row(), ix.column())
            for ix in self.selectionModel().selectedIndexes()
        ]

    def currentRow(self) -> int:  # noqa: N802
        return self.currentIndex().row()

    def currentColumn(self) -> int:  # noqa: N802
        return self.currentIndex().column()

    def setCurrentCell(self, r: int, c: int) -> None:  # noqa: N802
        self.setCurrentIndex(self._model.index(r, c))

    def scrollToItem(self, item: CellRef) -> None:  # noqa: N802
        if item is None:
            return
        self.scrollTo(self._model.index(item.row(), item.column()))

    def clearContents(self) -> None:  # noqa: N802
        n_rows, n_cols = self._model.rowCount(), self._model.columnCount()
        for r in range(n_rows):
            for c in range(n_cols):
                self._model.set_item(r, c, None)

    def blockSignals(self, block: bool) -> bool:  # noqa: N802
        """同时阻塞自身与 selectionModel 的信号，等价 QTableWidget 行为。"""
        self.selectionModel().blockSignals(block)
        return super().blockSignals(block)

    def _bridge_current(self, cur: QModelIndex, prev: QModelIndex) -> None:
        def to_ref(ix: QModelIndex) -> Optional[CellRef]:
            if not ix.isValid():
                return None
            return CellRef(self, ix.row(), ix.column())

        self.currentItemChanged.emit(to_ref(cur), to_ref(prev))
