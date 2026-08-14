"""单元格内字符级差异高亮 delegate。

与 QTableWidget 配合：item 通过 HIGHLIGHT_RANGES_ROLE 数据角色携带
[(start, end), ...] 半开区间列表，delegate 在绘制时只给这些子串加
背景色，其余字符保持原样，实现"只高亮差异内容"。
"""

from __future__ import annotations

from html import escape as html_escape

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor, QTextDocument
from PySide6.QtWidgets import (
    QStyle,
    QStyledItemDelegate,
    QStyleOptionViewItem,
)

from src.gui.themes import is_dark_mode


# 单元格携带的字符级高亮区间数据角色
HIGHLIGHT_RANGES_ROLE = Qt.ItemDataRole.UserRole + 1


def _escape_segment(text: str) -> str:
    """HTML 转义并保留空白与换行的显示效果。"""
    escaped = html_escape(text)
    escaped = escaped.replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;")
    escaped = escaped.replace("\n", "<br/>")
    return escaped.replace(" ", "&nbsp;")


class DiffHighlightDelegate(QStyledItemDelegate):
    """按字符区间渲染部分高亮文本的表格 delegate。"""

    def paint(self, painter, option, index) -> None:  # noqa: N802
        ranges = index.data(HIGHLIGHT_RANGES_ROLE)
        if not ranges:
            super().paint(painter, option, index)
            return

        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        text = opt.text

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
        doc.setHtml("".join(parts))
        doc.setTextWidth(max(0, opt.rect.width() - 4))

        painter.translate(opt.rect.left() + 2, opt.rect.top())
        doc.drawContents(painter)
        painter.restore()
