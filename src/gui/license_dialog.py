"""软件激活对话框 — LicenseDialog(QDialog)（一机一码）。

应用启动时若本机未激活，弹出此对话框：
1. 展示本机「机器码」，用户将其发送给开发者换取专属密匙；
2. 用户输入密匙，校验通过且绑定本机后激活并进入主界面；
3. 取消则退出应用。
"""

from __future__ import annotations

from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)

from src.utils import license


class LicenseDialog(QDialog):
    """密匙激活对话框（一机一码）。"""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("ExcelMerge — 软件激活")
        self.setModal(True)
        self.setMinimumWidth(480)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 16)
        layout.setSpacing(12)

        title = QLabel("软件激活")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        intro = QLabel(
            "本软件采用一机一码授权。请按以下步骤激活：\n"
            "1. 复制下方「本机机器码」并发送给开发者；\n"
            "2. 将开发者提供的专属密匙填入输入框并点击「激活」。\n"
            "密匙仅对当前机器有效，激活后本机无需再次输入。"
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)

        # ---- 机器码展示区 ----
        mc_layout = QHBoxLayout()
        mc_layout.setSpacing(8)
        mc_label = QLabel("本机机器码:")
        mc_label.setStyleSheet("font-weight: 600;")
        self.machine_code_edit = QLineEdit(license.get_machine_code())
        self.machine_code_edit.setReadOnly(True)
        self.machine_code_edit.setStyleSheet(
            "font-family: Menlo, Consolas, monospace; font-size: 14px;"
        )
        copy_btn = QPushButton("复制")
        copy_btn.clicked.connect(self._copy_machine_code)
        mc_layout.addWidget(mc_label)
        mc_layout.addWidget(self.machine_code_edit, 1)
        mc_layout.addWidget(copy_btn)
        layout.addLayout(mc_layout)

        # ---- 密匙输入区 ----
        form = QFormLayout()
        form.setSpacing(8)
        self.key_edit = QLineEdit()
        self.key_edit.setPlaceholderText("XXXX-XXXX-XXXX-XXXX-XXXX")
        self.key_edit.setMaxLength(24)  # 5*4 + 4 连字符
        self.key_edit.textChanged.connect(self._on_text_changed)
        self.key_edit.returnPressed.connect(self._on_activate)
        form.addRow("密匙:", self.key_edit)
        layout.addLayout(form)

        self.hint_label = QLabel(" ")
        self.hint_label.setStyleSheet("color: #B22222;")
        self.hint_label.setWordWrap(True)
        layout.addWidget(self.hint_label)

        layout.addStretch(1)

        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        self._ok_button = buttons.button(QDialogButtonBox.Ok)
        self._ok_button.setText("激活")
        cancel_button = buttons.button(QDialogButtonBox.Cancel)
        cancel_button.setText("退出")
        self._ok_button.clicked.connect(self._on_activate)
        cancel_button.clicked.connect(self.reject)
        layout.addWidget(buttons)

        self._ok_button.setEnabled(False)

    # ------------------------------------------------------------------ #
    # 机器码
    # ------------------------------------------------------------------ #
    def _copy_machine_code(self) -> None:
        """复制本机机器码到系统剪贴板。"""
        QApplication.clipboard().setText(self.machine_code_edit.text())
        self.hint_label.setStyleSheet("color: #2A7A2A;")
        self.hint_label.setText("机器码已复制，请发送给开发者")
        self.hint_label.setStyleSheet("color: #B22222;")

    # ------------------------------------------------------------------ #
    # 输入处理
    # ------------------------------------------------------------------ #
    def _on_text_changed(self, text: str) -> None:
        """自动格式化输入（大写、分组、仅十六进制）并实时校验。"""
        cleaned = license._normalize_hex(text)
        cleaned = cleaned[:20]
        formatted = license._format_key(cleaned)
        if formatted != text:
            self.key_edit.blockSignals(True)
            self.key_edit.setText(formatted)
            self.key_edit.blockSignals(False)

        valid = license.validate_key(formatted)
        self._ok_button.setEnabled(valid)
        if valid:
            self.hint_label.setText(" ")
        elif len(cleaned) == 20:
            self.hint_label.setText("密匙签名无效")
        else:
            self.hint_label.setText("密匙格式不正确")

    def _on_activate(self) -> None:
        """点击激活：校验格式、机器绑定并持久化，成功后关闭对话框。"""
        key = self.key_edit.text()
        if not license.validate_key(key):
            self.hint_label.setText("密匙无效，请检查后重试")
            return
        if not license.key_matches_machine(key, license.get_machine_code()):
            self.hint_label.setText("此密匙非本机专属，无法激活")
            return
        if license.activate(key):
            self.accept()
        else:
            QMessageBox.critical(
                self,
                "激活失败",
                "无法写入激活信息，请检查文件权限后重试。",
            )


if __name__ == "__main__":
    import sys

    app = QApplication(sys.argv)
    dlg = LicenseDialog()
    if dlg.exec() == QDialog.Accepted:
        print("已激活")
