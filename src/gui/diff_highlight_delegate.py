"""单元格内字符级差异高亮 delegate。

与 QTableWidget 配合：item 通过 HIGHLIGHT_RANGES_ROLE 数据角色携带
[(start, end), ...] 半开区间列表，delegate 在绘制时只给这些子串加
背景色，其余字符保持原样，实现"只高亮差异内容"。
"""

from __future__ import annotations

from html import escape as html_escape

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor, QTextDocument, QTextOption
from PySide6.QtWidgets import (
    QStyle,
    QStyledItemDelegate,
    QStyleOptionViewItem,
)

from src.gui.themes import is_dark_mode


# 单元格携带的字符级高亮区间数据角色
HIGHLIGHT_RANGES_ROLE = Qt.ItemDataRole.UserRole + 1
# 单元格是否启用自动折行显示
CELL_WRAP_ROLE = Qt.ItemDataRole.UserRole + 2


def _escape_segment(text: str) -> str:
    """HTML 转义并保留空白与换行的显示效果。"""
    escaped = html_escape(text)
    escaped = escaped.replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;")
    escaped = escaped.replace("\n", "<br/>")
    return escaped.replace(" ", "&nbsp;")


# 供主窗口测量折行高度时复用，确保与 delegate 实际渲染完全一致
escape_segment = _escape_segment


class DiffHighlightDelegate(QStyledItemDelegate):
    """按字符区间渲染部分高亮文本的表格 delegate。"""

    def paint(self, painter, option, index) -> None:  # noqa: N802
        ranges = index.data(HIGHLIGHT_RANGES_ROLE)
        wrap = bool(index.data(CELL_WRAP_ROLE))
        if not ranges and not wrap:
            super().paint(painter, option, index)
            return

        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        # 注意：initStyleOption 会把 "\n" 转成 U+2028，直接用于 HTML 会丢失
        # 换行，因此这里取原始 item 文本，保证换行能正确转为 <br/>。
        text = index.data(Qt.ItemDataRole.DisplayRole)
        if not isinstance(text, str):
            text = ""

        selected = bool(opt.state & QStyle.State_Selected)
        highlight_bg = (
            QColor("#B22222") if is_dark_mode() else QColor("#FF6B6B")
        )
        text_color = (
            opt.palette.highlightedText().color()
            if selected
            else opt.palette.text().color()
        )

        painter.save()
        painter.setClipRect(opt.rect)

        # 背景：选中态用高亮色，否则尊重 item 背景，最后回退到基色
        if selected:
            painter.fillRect(opt.rect, opt.palette.highlight())
        else:
            bg = index.data(Qt.ItemDataRole.BackgroundRole)
            if isinstance(bg, QBrush) and bg.style() != Qt.BrushStyle.NoBrush:
                painter.fillRect(opt.rect, bg)
            else:
                painter.fillRect(opt.rect, opt.palette.base())

        # 按原文本索引切分，逐段转义，避免转义后长度变化导致错位
        parts = []
        pos = 0
        if ranges:
            for start, end in sorted(ranges):
                start = max(start, 0)
                end = min(end, len(text))
                if end <= start:
                    continue
                parts.append(_escape_segment(text[pos:start]))
                parts.append(
                    '<span style="background-color:{};color:{};">{}</span>'.format(
                        highlight_bg.name(),
                        text_color.name(),
                        _escape_segment(text[start:end]),
                    )
                )
                pos = end
        parts.append(_escape_segment(text[pos:]))

        doc = QTextDocument()
        doc.setDefaultFont(opt.font)
        doc.setDocumentMargin(0)
        wrap_option = QTextOption()
        wrap_option.setWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
        doc.setDefaultTextOption(wrap_option)
        doc.setHtml("".join(parts))
        if wrap:
            # 按当前列宽折行，超长内容不再需要把列宽拉到极大值
            doc.setTextWidth(max(0, opt.rect.width() - 4))

        painter.translate(opt.rect.left() + 2, opt.rect.top())
        doc.drawContents(painter)
        painter.restore()
