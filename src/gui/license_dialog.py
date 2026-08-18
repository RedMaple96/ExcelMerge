"""ExcelMerge 在线授权激活对话框。"""

from __future__ import annotations

from typing import Callable

from PySide6.QtCore import QThread, Signal
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
from src.utils.license import ActivationError, NetworkError


class _ActivationWorker(QThread):
    finished_with_result = Signal(dict)

    def __init__(self, task: Callable[[], dict], parent: QDialog) -> None:
        super().__init__(parent)
        self._task = task

    def run(self) -> None:
        try:
            result = self._task()
        except ActivationError as exc:
            result = {"ok": False, "error": str(exc), "code": exc.code}
        except NetworkError as exc:
            result = {"ok": False, "error": str(exc), "code": "NETWORK_ERROR"}
        except Exception as exc:  # noqa: BLE001
            result = {"ok": False, "error": f"未知错误: {exc}", "code": "UNKNOWN"}
        self.finished_with_result.emit(result)


class LicenseDialog(QDialog):
    """在线申请授权码并激活当前设备。"""

    def __init__(self, initial_message: str = "", parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("在线授权激活")
        self.setMinimumWidth(480)
        self._worker: _ActivationWorker | None = None

        layout = QVBoxLayout(self)
        title = QLabel("ExcelMerge 在线授权")
        title.setStyleSheet("font-size: 16px; font-weight: 600;")
        layout.addWidget(title)
        layout.addWidget(QLabel("可自助申请 30 天授权码，授权会绑定当前设备。"))

        fingerprint_layout = QHBoxLayout()
        fingerprint_label = QLabel("设备指纹:")
        self.fingerprint_edit = QLineEdit(license.get_machine_fingerprint())
        self.fingerprint_edit.setReadOnly(True)
        copy_button = QPushButton("复制")
        copy_button.clicked.connect(self._copy_fingerprint)
        fingerprint_layout.addWidget(fingerprint_label)
        fingerprint_layout.addWidget(self.fingerprint_edit, 1)
        fingerprint_layout.addWidget(copy_button)
        layout.addLayout(fingerprint_layout)

        form = QFormLayout()
        self.code_edit = QLineEdit()
        self.code_edit.setPlaceholderText("EM-XXXXXXXXXXXXXXXX")
        self.code_edit.setMaxLength(19)
        self.code_edit.returnPressed.connect(self._on_activate)
        form.addRow("授权码:", self.code_edit)
        layout.addLayout(form)

        self.hint_label = QLabel(initial_message or " ")
        self.hint_label.setWordWrap(True)
        self.hint_label.setStyleSheet("color: #B22222;")
        layout.addWidget(self.hint_label)
        layout.addStretch(1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.get_code_button = QPushButton("获取授权码")
        buttons.addButton(self.get_code_button, QDialogButtonBox.ActionRole)
        self.activate_button = buttons.button(QDialogButtonBox.Ok)
        self.activate_button.setText("在线激活")
        exit_button = buttons.button(QDialogButtonBox.Cancel)
        exit_button.setText("退出")
        self.get_code_button.clicked.connect(self._on_request_code)
        self.activate_button.clicked.connect(self._on_activate)
        exit_button.clicked.connect(self.reject)
        layout.addWidget(buttons)

    def _set_busy(self, busy: bool, message: str = "") -> None:
        self.get_code_button.setEnabled(not busy)
        self.activate_button.setEnabled(not busy)
        if message:
            self.hint_label.setText(message)

    def _copy_fingerprint(self) -> None:
        QApplication.clipboard().setText(self.fingerprint_edit.text())
        self.hint_label.setStyleSheet("color: #2A7A2A;")
        self.hint_label.setText("设备指纹已复制")

    def _start_worker(self, task: Callable[[], dict]) -> None:
        self._worker = _ActivationWorker(task, self)
        self._worker.finished_with_result.connect(self._on_worker_result)
        self._worker.start()

    def _on_request_code(self) -> None:
        self._set_busy(True, "正在获取授权码…")
        self._start_worker(license.request_code)

    def _on_activate(self) -> None:
        code = self.code_edit.text().strip()
        if not code:
            self.hint_label.setText("请输入授权码")
            return
        self._set_busy(True, "正在在线激活…")
        self._start_worker(lambda: license.activate(code))

    def _on_worker_result(self, result: dict) -> None:
        if result.get("ok") and result.get("licenseCode"):
            self.code_edit.setText(str(result["licenseCode"]))
            self.hint_label.setStyleSheet("color: #2A7A2A;")
            self.hint_label.setText("授权码已获取，请点击“在线激活”")
            self._set_busy(False)
            return
        if result.get("ok"):
            self.hint_label.setStyleSheet("color: #2A7A2A;")
            self.hint_label.setText("激活成功")
            self.accept()
            return
        self.hint_label.setStyleSheet("color: #B22222;")
        self._set_busy(False, str(result.get("error") or "操作失败"))
