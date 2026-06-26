"""主窗口模块 — MainWindow(QMainWindow)。

应用外壳：菜单栏 / 工具栏 / 状态栏 / 左右并排双表格视图。
负责文件加载、Sheet 切换、比较触发与同步滚动；差异可视化（Task 6）、
合并交互（Task 8）、保存备份（Task 9）等在此预留方法桩。
对应需求：FR-03（主窗口框架与双表格视图）。
"""

from __future__ import annotations

import os
import shutil
from datetime import datetime
from typing import List, Optional

from PySide6.QtCore import Qt
from PySide6.QtGui import (
    QAction,
    QBrush,
    QColor,
    QCursor,
    QGuiApplication,
    QKeySequence,
    QShortcut,
)
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMenu,
    QMessageBox,
    QProgressBar,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

from src.core.comparator import DiffResult, ExcelComparator
from src.core.excel_loader import ExcelLoader, SheetData
from src.core.merger import ExcelMerger
from src.gui.column_settings_dialog import ColumnSettingsDialog
from src.gui.themes import is_dark_mode
from src.gui.workers import CompareWorker


class MainWindow(QMainWindow):
    """应用程序主窗口。"""

    # ------------------------------------------------------------------ #
    # 生命周期
    # ------------------------------------------------------------------ #
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("表格比较 — 无标题")
        self.resize(1200, 750)

        # ---- 共享状态 ----
        self.left_path: Optional[str] = None
        self.right_path: Optional[str] = None
        self.left_wb = None
        self.right_wb = None
        self.left_sheet_data: Optional[SheetData] = None
        self.right_sheet_data: Optional[SheetData] = None
        self.left_sheet_name: str = ""
        self.right_sheet_name: str = ""
        self.diff_result: Optional[DiffResult] = None
        self.key_cols: List[int] = []
        self.ignore_cols: List[int] = []
        self.show_row_numbers: bool = True
        self._sync_scroll_blocked: bool = False  # 同步滚动防递归守卫
        self._current_diff_pos: int = -1  # 差异导航游标（指向 diff_row_indices）
        self._compare_worker: Optional[CompareWorker] = None  # 比较后台线程
        self._dirty: bool = False  # 是否有未保存的修改（合并后置 True）
        self._backup_done: bool = True  # 当前脏状态是否已备份；初始无需备份
        self._last_backup_path: Optional[str] = None  # 最近一次备份路径
        # 右键菜单上下文（由 on_context_menu 填充，_ctx_* 回调读取）
        self._ctx_table: Optional[QTableWidget] = None
        self._ctx_aligned_idx: int = -1  # 对齐行索引（aligned_rows 下标）
        self._ctx_row: int = -1  # 视觉行
        self._ctx_col: int = -1  # 视觉列

        # ---- 构建界面 ----
        self._init_ui()
        self._init_menubar()
        self._init_toolbar()
        self._init_statusbar()
        self._init_connections()
        self._init_shortcuts()
        self._init_theme_listener()

    # ------------------------------------------------------------------ #
    # 中央界面（左右双面板）
    # ------------------------------------------------------------------ #
    def _init_ui(self) -> None:
        """构建中央控件：左右两个面板，各含标题区/表格/统计区。"""
        central = QWidget(self)
        self.setCentralWidget(central)
        central.setAcceptDrops(True)  # 启用窗口级拖拽接收
        outer = QHBoxLayout(central)
        outer.setContentsMargins(4, 4, 4, 4)

        # 左侧面板
        (
            self.left_file_label,
            self.left_sheet_combo,
            self.left_table,
            self.left_stat_label,
        ) = self._make_panel(outer, "左侧")

        # 右侧面板
        (
            self.right_file_label,
            self.right_sheet_combo,
            self.right_table,
            self.right_stat_label,
        ) = self._make_panel(outer, "右侧")

        self._configure_table(self.left_table)
        self._configure_table(self.right_table)

    def _make_panel(self, parent_layout: QHBoxLayout, title: str):
        """创建单个面板，返回 (file_label, sheet_combo, table, stat_label)。"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(4)

        # 标题区：文件信息 + Sheet 下拉
        header = QFrame()
        header.setFrameShape(QFrame.StyledPanel)
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(6, 4, 6, 4)
        header_layout.setSpacing(2)

        file_label = QLabel("未加载")
        file_label.setStyleSheet("font-weight: 600;")
        sheet_combo = QComboBox()
        sheet_combo.setToolTip(f"{title} Sheet 选择")
        header_layout.addWidget(file_label)
        header_layout.addWidget(sheet_combo)
        layout.addWidget(header)

        # 表格
        table = QTableWidget()
        layout.addWidget(table, 1)

        # 统计区
        stat_label = QLabel("差异: 0 | 相同: 0 | 仅左: 0 仅右: 0")
        stat_label.setContentsMargins(6, 0, 6, 0)
        layout.addWidget(stat_label)

        parent_layout.addWidget(panel, 1)
        return file_label, sheet_combo, table, stat_label

    def _configure_table(self, table: QTableWidget) -> None:
        """配置 QTableWidget 的通用属性。"""
        table.setSelectionMode(QTableWidget.SingleSelection)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setAlternatingRowColors(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        # 行号：使用 verticalHeader，可由 show_row_numbers 切换显隐
        table.verticalHeader().setVisible(self.show_row_numbers)
        table.verticalHeader().setDefaultAlignment(Qt.AlignCenter)
        hheader = table.horizontalHeader()
        hheader.setStretchLastSection(True)
        table.setHorizontalScrollMode(QTableWidget.ScrollPerPixel)
        table.setVerticalScrollMode(QTableWidget.ScrollPerPixel)
        # 拖拽接收：实际 drop 处理由 Task 7 完成，此处先开启标志
        table.setAcceptDrops(True)
        table.setDragEnabled(False)

    # ------------------------------------------------------------------ #
    # 菜单栏
    # ------------------------------------------------------------------ #
    def _init_menubar(self) -> None:
        """构建 5 个菜单：会话/编辑/视图/工具/帮助。"""
        menubar = self.menuBar()

        # ---- 会话(S) ----
        m_session = menubar.addMenu("会话(&S)")

        self.action_open_left = QAction("打开左", self)
        self.action_open_left.setShortcut(QKeySequence("Ctrl+L"))
        self.action_open_left.triggered.connect(self._on_open_left)
        m_session.addAction(self.action_open_left)

        self.action_open_right = QAction("打开右", self)
        self.action_open_right.setShortcut(QKeySequence("Ctrl+R"))
        self.action_open_right.triggered.connect(self._on_open_right)
        m_session.addAction(self.action_open_right)

        m_session.addSeparator()

        self.action_save = QAction("保存", self)
        self.action_save.setShortcut(QKeySequence.Save)
        self.action_save.triggered.connect(self.save)
        m_session.addAction(self.action_save)

        self.action_save_as = QAction("另存为", self)
        self.action_save_as.setShortcut(QKeySequence("Ctrl+Shift+S"))
        self.action_save_as.triggered.connect(self.save_as)
        m_session.addAction(self.action_save_as)

        m_session.addSeparator()

        self.action_swap = QAction("交换左右", self)
        self.action_swap.setShortcut(QKeySequence("Ctrl+Shift+X"))
        self.action_swap.triggered.connect(self.swap_sides)
        m_session.addAction(self.action_swap)

        self.action_reload = QAction("重载文件", self)
        self.action_reload.triggered.connect(self.reload_files)
        m_session.addAction(self.action_reload)

        self.action_recompare = QAction("重新比较", self)
        self.action_recompare.setShortcut(QKeySequence("F5"))
        self.action_recompare.triggered.connect(self.recompare)
        m_session.addAction(self.action_recompare)

        # ---- 编辑(E) ----
        m_edit = menubar.addMenu("编辑(&E)")

        self.action_undo = QAction("撤销", self)
        self.action_undo.setShortcut(QKeySequence.Undo)
        self.action_undo.triggered.connect(self.undo)
        m_edit.addAction(self.action_undo)

        self.action_redo = QAction("重做", self)
        self.action_redo.setShortcut(QKeySequence.Redo)
        self.action_redo.triggered.connect(self.redo)
        m_edit.addAction(self.action_redo)

        m_edit.addSeparator()

        self.action_align_rows = QAction("对齐行", self)
        self.action_align_rows.triggered.connect(self.align_rows)
        m_edit.addAction(self.action_align_rows)

        self.action_copy_to_left = QAction("复制到左侧", self)
        self.action_copy_to_left.triggered.connect(self.copy_to_left)
        m_edit.addAction(self.action_copy_to_left)

        self.action_copy_to_right = QAction("复制到右侧", self)
        self.action_copy_to_right.triggered.connect(self.copy_to_right)
        m_edit.addAction(self.action_copy_to_right)

        m_edit.addSeparator()

        self.action_copy_cell = QAction("复制单元格", self)
        self.action_copy_cell.setShortcut(QKeySequence.Copy)
        self.action_copy_cell.triggered.connect(self.copy_cell)
        m_edit.addAction(self.action_copy_cell)

        self.action_paste = QAction("粘贴", self)
        self.action_paste.setShortcut(QKeySequence.Paste)
        self.action_paste.triggered.connect(self.paste)
        m_edit.addAction(self.action_paste)

        # ---- 视图(V) ----
        m_view = menubar.addMenu("视图(&V)")

        self.action_show_row_numbers = QAction("行号", self)
        self.action_show_row_numbers.setCheckable(True)
        self.action_show_row_numbers.setChecked(self.show_row_numbers)
        self.action_show_row_numbers.toggled.connect(self._on_toggle_row_numbers)
        m_view.addAction(self.action_show_row_numbers)

        self.action_column_settings = QAction("列设置", self)
        self.action_column_settings.triggered.connect(self.open_column_settings)
        m_view.addAction(self.action_column_settings)

        self.action_resize_columns = QAction("调整列宽", self)
        self.action_resize_columns.triggered.connect(self.resize_columns)
        m_view.addAction(self.action_resize_columns)

        m_view.addSeparator()

        self.action_show_diff = QAction("显示差异", self)
        self.action_show_diff.setCheckable(True)
        self.action_show_diff.setChecked(True)
        self.action_show_diff.toggled.connect(self._on_view_filter_changed)
        m_view.addAction(self.action_show_diff)

        self.action_show_same = QAction("显示相同", self)
        self.action_show_same.setCheckable(True)
        self.action_show_same.setChecked(True)
        self.action_show_same.toggled.connect(self._on_view_filter_changed)
        m_view.addAction(self.action_show_same)

        self.action_show_left_only = QAction("显示仅左侧", self)
        self.action_show_left_only.setCheckable(True)
        self.action_show_left_only.setChecked(True)
        self.action_show_left_only.toggled.connect(self._on_view_filter_changed)
        m_view.addAction(self.action_show_left_only)

        self.action_show_right_only = QAction("显示仅右侧", self)
        self.action_show_right_only.setCheckable(True)
        self.action_show_right_only.setChecked(True)
        self.action_show_right_only.toggled.connect(self._on_view_filter_changed)
        m_view.addAction(self.action_show_right_only)

        # ---- 工具(T) ----
        m_tools = menubar.addMenu("工具(&T)")

        self.action_file_format = QAction("文件格式", self)
        self.action_file_format.triggered.connect(self.show_file_format)
        m_tools.addAction(self.action_file_format)

        self.action_options = QAction("选项", self)
        self.action_options.triggered.connect(self.show_options)
        m_tools.addAction(self.action_options)

        # ---- 帮助(H) ----
        m_help = menubar.addMenu("帮助(&H)")

        self.action_about = QAction("关于", self)
        self.action_about.triggered.connect(self.about)
        m_help.addAction(self.action_about)

        self.action_docs = QAction("文档", self)
        self.action_docs.triggered.connect(self.show_docs)
        m_help.addAction(self.action_docs)

        self.action_shortcuts = QAction("快捷键列表", self)
        self.action_shortcuts.triggered.connect(self.show_shortcuts)
        m_help.addAction(self.action_shortcuts)

    # ------------------------------------------------------------------ #
    # 工具栏
    # ------------------------------------------------------------------ #
    def _init_toolbar(self) -> None:
        """构建水平工具栏：打开左/打开右/保存/交换/上一个差异/下一个差异/统计/列设置/合并。"""
        self.toolbar = QToolBar("主工具栏", self)
        self.toolbar.setMovable(False)
        self.toolbar.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        self.addToolBar(self.toolbar)

        style = self.style()
        icon = lambda sp: style.standardIcon(sp)  # noqa: E731

        self.tb_open_left = self.toolbar.addAction(
            icon(QStyle.SP_DirOpenIcon), "打开左", self._on_open_left
        )
        self.tb_open_right = self.toolbar.addAction(
            icon(QStyle.SP_DirOpenIcon), "打开右", self._on_open_right
        )
        self.tb_save = self.toolbar.addAction(
            icon(QStyle.SP_DialogSaveButton), "保存", self.save
        )
        self.toolbar.addSeparator()
        self.tb_swap = self.toolbar.addAction(
            icon(QStyle.SP_BrowserReload), "交换", self.swap_sides
        )
        self.tb_prev_diff = self.toolbar.addAction(
            icon(QStyle.SP_ArrowUp), "上一个差异", self.prev_diff
        )
        self.tb_next_diff = self.toolbar.addAction(
            icon(QStyle.SP_ArrowDown), "下一个差异", self.next_diff
        )
        self.toolbar.addSeparator()
        self.tb_statistics = self.toolbar.addAction(
            icon(QStyle.SP_FileDialogListView), "统计", self.show_statistics
        )
        self.tb_column_settings = self.toolbar.addAction(
            icon(QStyle.SP_FileDialogDetailedView), "列设置", self.open_column_settings
        )
        self.tb_merge = self.toolbar.addAction(
            icon(QStyle.SP_DialogApplyButton), "合并", self.merge
        )

    # ------------------------------------------------------------------ #
    # 状态栏
    # ------------------------------------------------------------------ #
    def _init_statusbar(self) -> None:
        """构建状态栏：就绪标签 + 行信息 + 差异数 + 备份状态 + 进度条。"""
        bar = self.statusBar()

        self.status_label = QLabel("就绪")
        bar.addWidget(self.status_label, 1)

        self.row_label = QLabel("左: 0 行 | 右: 0 行")
        bar.addPermanentWidget(self.row_label)

        self.diff_label = QLabel("差异: 0")
        bar.addPermanentWidget(self.diff_label)

        self.backup_label = QLabel("备份: 无")
        bar.addPermanentWidget(self.backup_label)

        # 比较进度条（默认隐藏，比较期间显示）
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximumWidth(200)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.hide()
        bar.addPermanentWidget(self.progress_bar)

    # ------------------------------------------------------------------ #
    # 信号连接
    # ------------------------------------------------------------------ #
    def _init_connections(self) -> None:
        """连接 Sheet 下拉、同步滚动与右键菜单。"""
        self.left_sheet_combo.currentIndexChanged.connect(
            lambda _: self._on_sheet_changed("left")
        )
        self.right_sheet_combo.currentIndexChanged.connect(
            lambda _: self._on_sheet_changed("right")
        )

        # 同步滚动（双向，带防递归守卫）
        self.left_table.verticalScrollBar().valueChanged.connect(
            lambda v: self._sync_scroll(self.left_table, self.right_table, "v", v)
        )
        self.right_table.verticalScrollBar().valueChanged.connect(
            lambda v: self._sync_scroll(self.right_table, self.left_table, "v", v)
        )
        self.left_table.horizontalScrollBar().valueChanged.connect(
            lambda v: self._sync_scroll(self.left_table, self.right_table, "h", v)
        )
        self.right_table.horizontalScrollBar().valueChanged.connect(
            lambda v: self._sync_scroll(self.right_table, self.left_table, "h", v)
        )

        # 右键上下文菜单
        for table in (self.left_table, self.right_table):
            table.setContextMenuPolicy(Qt.CustomContextMenu)
            table.customContextMenuRequested.connect(self.on_context_menu)

    def _init_shortcuts(self) -> None:
        """注册无菜单项的额外快捷键（Task 11）。"""
        # Cmd/Ctrl+O 智能打开
        QShortcut(QKeySequence("Ctrl+O"), self, activated=self._smart_open)
        # Cmd/Ctrl+Down 下一个差异 / Cmd/Ctrl+Up 上一个差异
        QShortcut(QKeySequence("Ctrl+Down"), self, activated=self.next_diff)
        QShortcut(QKeySequence("Ctrl+Up"), self, activated=self.prev_diff)

    def _init_theme_listener(self) -> None:
        """监听系统外观变化（PySide6 6.5+），自动切换主题。"""
        sh = QGuiApplication.styleHints()
        if sh is not None and hasattr(sh, "colorSchemeChanged"):
            try:
                sh.colorSchemeChanged.connect(self._on_color_scheme_changed)
            except Exception:  # noqa: BLE001
                pass

    def _on_color_scheme_changed(self) -> None:
        """系统外观变化时重新应用主题并重渲染差异。"""
        app = QApplication.instance()
        if app is not None:
            from src.gui.themes import apply_theme

            apply_theme(app)
        self._render_diffs()

    # ------------------------------------------------------------------ #
    # 文件加载（完整实现）
    # ------------------------------------------------------------------ #
    def open_left(self) -> None:
        """打开左侧文件：弹出原生文件选择对话框，加载后若两侧齐全则比较。"""
        path, _ = QFileDialog.getOpenFileName(
            self, "选择左侧 Excel 文件", "", "Excel Files (*.xlsx)"
        )
        if not path:
            return
        self.load_file("left", path)
        if self.left_path and self.right_path:
            self._run_compare()

    def open_right(self) -> None:
        """打开右侧文件：弹出原生文件选择对话框，加载后若两侧齐全则比较。"""
        path, _ = QFileDialog.getOpenFileName(
            self, "选择右侧 Excel 文件", "", "Excel Files (*.xlsx)"
        )
        if not path:
            return
        self.load_file("right", path)
        if self.left_path and self.right_path:
            self._run_compare()

    def _smart_open(self) -> None:
        """智能打开：优先打开空侧（左优先），两侧均已加载时打开左侧。"""
        if self.left_path is None:
            self.open_left()
        elif self.right_path is None:
            self.open_right()
        else:
            self.open_left()

    def _on_open_left(self) -> None:
        self.open_left()

    def _on_open_right(self) -> None:
        self.open_right()

    def load_file(self, side: str, path: str) -> None:
        """加载指定侧文件：校验格式 -> 加载工作簿 -> 填充 Sheet 下拉 -> 切换 Sheet。

        side: "left" 或 "right"。
        """
        if not str(path).lower().endswith(".xlsx"):
            QMessageBox.warning(
                self, "文件格式不支持",
                f"仅支持 .xlsx 格式文件:\n{path}\n如需支持旧版 .xls，请先另存为 .xlsx。",
            )
            return

        try:
            wb = ExcelLoader.load_workbook(path)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "加载失败", f"无法加载文件:\n{path}\n\n{exc}")
            return

        sheet_names = ExcelLoader.get_sheet_names(wb)
        if not sheet_names:
            QMessageBox.warning(self, "空工作簿", f"文件中没有工作表:\n{path}")
            return

        if side == "left":
            self.left_path = path
            self.left_wb = wb
            combo = self.left_sheet_combo
        else:
            self.right_path = path
            self.right_wb = wb
            combo = self.right_sheet_combo

        # 填充下拉（屏蔽信号避免重复触发）
        combo.blockSignals(True)
        combo.clear()
        for name in sheet_names:
            combo.addItem(name)
        if sheet_names:
            combo.setCurrentIndex(0)
        combo.blockSignals(False)

        self._update_file_label(side)
        self._on_sheet_changed(side)
        self.update_status(f"已加载 {side} 侧: {os.path.basename(path)}")

    def _update_file_label(self, side: str) -> None:
        """根据当前路径刷新标题区的文件名/大小/修改时间。"""
        path = self.left_path if side == "left" else self.right_path
        label = self.left_file_label if side == "left" else self.right_file_label
        if not path:
            label.setText("未加载")
            return
        try:
            size = os.path.getsize(path)
            mtime = os.path.getmtime(path)
            mtime_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
            label.setText(
                f"{os.path.basename(path)}  |  {self._format_size(size)}  |  {mtime_str}"
            )
        except OSError:
            label.setText(os.path.basename(path))

    @staticmethod
    def _format_size(n: float) -> str:
        """将字节数格式化为易读字符串。"""
        for unit in ("B", "KB", "MB", "GB"):
            if n < 1024:
                return f"{n:.1f} {unit}"
            n /= 1024
        return f"{n:.1f} TB"

    def reload_files(self) -> None:
        """重载左右两侧已打开的文件。"""
        if self.left_path:
            self.load_file("left", self.left_path)
        if self.right_path:
            self.load_file("right", self.right_path)
        self.update_status("已重载文件")

    def recompare(self) -> None:
        """使用当前 key_cols/ignore_cols 重新比较。"""
        self._run_compare()
        self.update_status("已重新比较")

    # ------------------------------------------------------------------ #
    # 拖拽支持（Task 7）
    # ------------------------------------------------------------------ #
    def dragEnterEvent(self, event) -> None:  # noqa: N802
        """拖入事件：仅接受含 .xlsx 文件的拖拽。"""
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().lower().endswith(".xlsx"):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event) -> None:  # noqa: N802
        """拖放事件：根据落点位置决定加载到左侧或右侧。

        - 拖到左表上方 -> 左侧；拖到右表上方 -> 右侧；
        - 单文件且无法判定落点时：左空加载左，否则右空加载右；
        - 双文件：第一个进左，第二个进右。
        """
        paths = [
            url.toLocalFile()
            for url in event.mimeData().urls()
            if url.toLocalFile().lower().endswith(".xlsx")
        ]
        if not paths:
            event.ignore()
            return
        event.acceptProposedAction()

        if len(paths) >= 2:
            self.load_file("left", paths[0])
            self.load_file("right", paths[1])
        else:
            path = paths[0]
            if self.left_table.underMouse():
                side = "left"
            elif self.right_table.underMouse():
                side = "right"
            elif self.left_path is None:
                side = "left"
            elif self.right_path is None:
                side = "right"
            else:
                side = "left"
            self.load_file(side, path)

        if self.left_path and self.right_path:
            self._run_compare()

    # ------------------------------------------------------------------ #
    # Sheet 切换 / 表格渲染 / 比较
    # ------------------------------------------------------------------ #
    def _on_sheet_changed(self, side: str) -> None:
        """下拉切换 Sheet：提取 SheetData -> 刷新表格 -> 触发比较。"""
        wb = self.left_wb if side == "left" else self.right_wb
        combo = self.left_sheet_combo if side == "left" else self.right_sheet_combo
        if wb is None or combo.count() == 0:
            return
        name = combo.currentText()
        if not name:
            return
        try:
            ws = ExcelLoader.get_worksheet(wb, name)
            sheet_data = ExcelLoader.extract_sheet_data(ws)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "Sheet 读取失败", f"无法读取工作表 {name}:\n{exc}")
            return

        if side == "left":
            self.left_sheet_data = sheet_data
            self.left_sheet_name = name
        else:
            self.right_sheet_data = sheet_data
            self.right_sheet_name = name

        # 切换 Sheet 时清除旧的差异结果，避免显示错位的对齐行
        self.diff_result = None
        self._refresh_tables()
        self._run_compare()

    def _refresh_tables(self) -> None:
        """刷新左右表格内容。

        - 若 diff_result 可用且两侧数据齐全：按对齐行渲染（两侧行数一致）
        - 否则：按原始行渲染，行数取 max(左, 右) 保证视觉对齐
        """
        left = self.left_sheet_data
        right = self.right_sheet_data

        if self.diff_result and left and right:
            max_col = self.diff_result.max_col
            left_indices = [p.left_row for p in self.diff_result.aligned_rows]
            right_indices = [p.right_row for p in self.diff_result.aligned_rows]
        else:
            n_left = len(left.values) if left else 0
            n_right = len(right.values) if right else 0
            n = max(n_left, n_right)
            max_col = max(left.max_col if left else 0, right.max_col if right else 0)
            left_indices = [i if i < n_left else None for i in range(n)]
            right_indices = [i if i < n_right else None for i in range(n)]

        self._populate_table(self.left_table, left, left_indices, max_col)
        self._populate_table(self.right_table, right, right_indices, max_col)

    def _populate_table(
        self,
        table: QTableWidget,
        sheet_data: Optional[SheetData],
        row_indices: List[Optional[int]],
        max_col: int,
    ) -> None:
        """按给定的源行索引列表填充表格。

        row_indices 中 None 表示该显示行为空（对应侧无此行）。
        """
        table.setRowCount(0)  # 清空旧内容与旧格式
        table.setColumnCount(max_col)

        # 列头
        if sheet_data and sheet_data.header_labels:
            labels = list(sheet_data.header_labels)
            while len(labels) < max_col:
                labels.append("")
            table.setHorizontalHeaderLabels(labels[:max_col])
        else:
            table.setHorizontalHeaderLabels([""] * max_col)

        table.setRowCount(len(row_indices))
        # 行号标签
        table.setVerticalHeaderLabels(
            [str(i + 1) for i in range(len(row_indices))]
        )

        for r, src_idx in enumerate(row_indices):
            if src_idx is None or sheet_data is None:
                continue
            values = (
                sheet_data.values[src_idx]
                if src_idx < len(sheet_data.values)
                else []
            )
            for c in range(max_col):
                text = values[c] if c < len(values) else ""
                table.setItem(r, c, QTableWidgetItem(text))

        table.verticalHeader().setVisible(self.show_row_numbers)

    def _run_compare(self) -> None:
        """两侧数据齐全时启动后台比较线程，避免阻塞 UI。

        比较完成后由 worker 信号驱动 _on_diff_ready / _on_compare_finished
        更新表格与状态。若已有比较在跑则跳过本次请求。
        """
        if not (self.left_sheet_data and self.right_sheet_data):
            self.diff_result = None
            self._current_diff_pos = -1
            self._update_stat_labels()
            self._render_diffs()
            return

        # 防止重叠：已有比较线程在跑则跳过
        if self._compare_worker is not None and self._compare_worker.isRunning():
            return

        self._set_compare_actions_enabled(False)
        self.update_status("比较中...")
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.show()

        self._compare_worker = CompareWorker(
            self.left_sheet_data,
            self.right_sheet_data,
            self.key_cols,
            self.ignore_cols,
        )
        self._compare_worker.diff_ready.connect(self._on_diff_ready)
        self._compare_worker.progress.connect(self._on_compare_progress)
        self._compare_worker.finished_ok.connect(self._on_compare_finished)
        # 线程结束后清理对象
        self._compare_worker.finished.connect(self._compare_worker.deleteLater)
        self._compare_worker.start()

    def _set_compare_actions_enabled(self, enabled: bool) -> None:
        """比较期间禁用/恢复会触发比较的相关动作，避免重入。"""
        for act in (
            self.action_recompare,
            self.action_swap,
            self.action_reload,
            self.action_column_settings,
        ):
            if act is not None:
                act.setEnabled(enabled)

    def _on_diff_ready(self, result: DiffResult) -> None:
        """比较结果就绪：保存结果、重置游标、刷新表格与差异渲染。"""
        self.diff_result = result
        self._current_diff_pos = -1
        # 重新渲染为对齐行布局，再叠加差异标识
        self._refresh_tables()
        self._update_stat_labels()
        self._render_diffs()

    def _on_compare_progress(self, cur: int, total: int) -> None:
        """更新状态栏进度条。"""
        if total <= 0:
            self.progress_bar.setValue(100)
            return
        self.progress_bar.setValue(int(cur / total * 100))

    def _on_compare_finished(self) -> None:
        """比较结束：恢复动作可用、隐藏进度条、更新状态。"""
        self._set_compare_actions_enabled(True)
        self.progress_bar.hide()
        self._compare_worker = None
        self.update_status("就绪")

    def _diff_colors(self) -> dict:
        """根据当前主题返回差异配色字典。

        返回键：
        - diff_row: 不同行的浅红底色
        - diff_cell: 差异单元格的深红底色
        - only_row: 单侧独占行的灰色底色
        """
        dark = is_dark_mode()
        if dark:
            return {
                "diff_row": QColor("#4A2A2A"),
                "diff_cell": QColor("#8B0000"),
                "only_row": QColor("#3A3A3A"),
            }
        return {
            "diff_row": QColor("#FFE0E0"),
            "diff_cell": QColor("#FF8888"),
            "only_row": QColor("#E0E0E0"),
        }

    def _render_diffs(self) -> None:
        """差异可视化 —— 根据 aligned_rows 为两侧表格着色并应用视图过滤。

        表格行索引 i 与 aligned_rows[i] 一一对应（由 _refresh_tables 保证）。
        - 先复位全部背景与行隐藏状态；
        - 再按状态着色：different 浅红 + 差异列深红；left_only/right_only 灰色；
        - 视图过滤：被关闭的类别将对应行隐藏。
        """
        colors = self._diff_colors()
        transparent = QBrush(Qt.transparent)
        n_rows = self.left_table.rowCount()

        # 1. 复位背景与行隐藏
        for table in (self.left_table, self.right_table):
            for r in range(table.rowCount()):
                table.setRowHidden(r, False)
                for c in range(table.columnCount()):
                    item = table.item(r, c)
                    if item is not None:
                        item.setBackground(transparent)

        # 2. 无比较结果时直接返回
        if not self.diff_result:
            return

        aligned = self.diff_result.aligned_rows
        # 视图过滤开关
        show_diff = self.action_show_diff.isChecked()
        show_same = self.action_show_same.isChecked()
        show_left_only = self.action_show_left_only.isChecked()
        show_right_only = self.action_show_right_only.isChecked()

        # 3. 逐行着色 / 过滤
        for i in range(min(n_rows, len(aligned))):
            pair = aligned[i]
            status = pair.status

            # 行可见性过滤
            if status == "different" and not show_diff:
                self._hide_row_pair(i)
                continue
            if status == "same" and not show_same:
                self._hide_row_pair(i)
                continue
            if status == "left_only" and not show_left_only:
                self._hide_row_pair(i)
                continue
            if status == "right_only" and not show_right_only:
                self._hide_row_pair(i)
                continue

            # 着色
            if status == "different":
                self._paint_row(self.left_table, i, colors["diff_row"])
                self._paint_row(self.right_table, i, colors["diff_row"])
                # 差异单元格深红
                for c in pair.diff_cells:
                    self._paint_cell(self.left_table, i, c, colors["diff_cell"])
                    self._paint_cell(self.right_table, i, c, colors["diff_cell"])
            elif status == "left_only":
                self._paint_row(self.left_table, i, colors["only_row"])
            elif status == "right_only":
                self._paint_row(self.right_table, i, colors["only_row"])
            # same: 不着色

    def _hide_row_pair(self, row: int) -> None:
        """同时隐藏两侧表格的指定行。"""
        self.left_table.setRowHidden(row, True)
        self.right_table.setRowHidden(row, True)

    def _paint_row(self, table: QTableWidget, row: int, color: QColor) -> None:
        """为指定行的所有单元格设置背景色。"""
        brush = QBrush(color)
        for c in range(table.columnCount()):
            item = table.item(row, c)
            if item is not None:
                item.setBackground(brush)

    def _paint_cell(
        self, table: QTableWidget, row: int, col: int, color: QColor
    ) -> None:
        """为指定单元格设置背景色。"""
        item = table.item(row, col)
        if item is not None:
            item.setBackground(QBrush(color))

    def _update_stat_labels(self) -> None:
        """更新底部统计标签与状态栏的行/差异数。"""
        if self.diff_result:
            s = self.diff_result.stats
            text = (
                f"差异: {s['different']} | 相同: {s['same']} | "
                f"仅左: {s['left_only']} 仅右: {s['right_only']}"
            )
            self.left_stat_label.setText(text)
            self.right_stat_label.setText(text)
            self.diff_label.setText(
                f"差异: {s['different'] + s['left_only'] + s['right_only']}"
            )
        else:
            self.left_stat_label.setText("差异: 0 | 相同: 0 | 仅左: 0 仅右: 0")
            self.right_stat_label.setText("差异: 0 | 相同: 0 | 仅左: 0 仅右: 0")
            self.diff_label.setText("差异: 0")

        n_left = len(self.left_sheet_data.values) if self.left_sheet_data else 0
        n_right = len(self.right_sheet_data.values) if self.right_sheet_data else 0
        self.row_label.setText(f"左: {n_left} 行 | 右: {n_right} 行")

    # ------------------------------------------------------------------ #
    # 保存与备份（Task 9 实装）
    # ------------------------------------------------------------------ #
    def save(self) -> None:
        """保存左侧工作簿（目标文件 A）到原路径，并做完整性校验。

        流程：备份（若有未保存修改且尚未备份）-> 保存 -> 完整性校验 -> 复位脏标记。
        """
        if not (self.left_wb and self.left_path):
            QMessageBox.warning(self, "无可保存", "无可保存的文件")
            return
        if not self._dirty:
            self.update_status("无未保存更改")
        # FR-05-03：保存前自动备份原始文件（仅当有未保存修改且尚未备份时）
        self._backup_if_needed()
        try:
            ExcelLoader.save_workbook(self.left_wb, self.left_path)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(
                self, "保存失败", f"保存文件失败:\n{self.left_path}\n\n{exc}"
            )
            return

        # FR-05 可靠性：保存后完整性校验
        if not self._verify_integrity(self.left_path):
            return

        self._dirty = False
        self._backup_done = True
        if self._last_backup_path:
            self.backup_label.setText(
                f"备份: {os.path.basename(self._last_backup_path)}"
            )
        self.update_status("已保存")

    def save_as(self) -> None:
        """另存左侧工作簿到用户指定路径，不改变 left_path。

        另存前若存在未保存修改，先备份原始文件（FR-05-02：不破坏原文件）。
        """
        if not self.left_wb:
            QMessageBox.warning(self, "无可保存", "无可保存的文件")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "另存为", "", "Excel Files (*.xlsx)"
        )
        if not path:
            return
        # 确保 .xlsx 后缀
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        # 有未保存修改时先备份原始文件
        self._backup_if_needed()
        try:
            ExcelLoader.save_workbook(self.left_wb, path)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(
                self, "另存失败", f"另存文件失败:\n{path}\n\n{exc}"
            )
            return
        if not self._verify_integrity(path):
            return
        # 注意：不修改 self.left_path，原始文件路径保持不变
        self.update_status(f"已另存为 {path}")

    def _backup_if_needed(self) -> None:
        """保存前自动备份左侧原文件（FR-05-03）。

        - left_path 为空则跳过；
        - 已备份过当前脏状态（_backup_done=True）则跳过；
        - 备份命名 ``<stem>_backup_<YYYYMMDD_HHMMSS>.xlsx``，与原文件同目录；
        - 失败仅警告，不阻断保存流程。
        """
        if not self.left_path:
            return
        if self._backup_done:
            return
        try:
            dir_name = os.path.dirname(self.left_path)
            stem = os.path.splitext(os.path.basename(self.left_path))[0]
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = os.path.join(
                dir_name, f"{stem}_backup_{timestamp}.xlsx"
            )
            shutil.copy2(self.left_path, backup_path)
            self._backup_done = True
            self._last_backup_path = backup_path
            self.backup_label.setText(f"备份: {os.path.basename(backup_path)}")
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(
                self, "备份失败",
                f"创建备份失败，保存将继续:\n{exc}",
            )

    def _verify_integrity(self, path: str) -> bool:
        """保存后完整性校验：重新加载文件确认未损坏（FR-05 可靠性）。

        成功返回 True；失败弹警告并返回 False。
        """
        try:
            ExcelLoader.load_workbook(path)
            return True
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(
                self, "完整性校验",
                f"保存文件可能损坏: {exc}",
            )
            return False

    def _backup_before_merge(self, side: str) -> None:
        """合并前备份指定侧文件。

        side 为 "left" 或 "right"。备份文件与原文件同目录，
        命名形如 ``<stem>_backup_<YYYYMMDD_HHMMSS><ext>``。
        失败时仅警告，不中断合并流程。
        """
        path = self.left_path if side == "left" else self.right_path
        if not path:
            return
        try:
            dir_name, name = os.path.split(path)
            stem, ext = os.path.splitext(name)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = os.path.join(
                dir_name, f"{stem}_backup_{timestamp}{ext}"
            )
            shutil.copy2(path, backup_path)
            self.backup_label.setText(f"备份: {os.path.basename(backup_path)}")
            self._backup_done = True
            self._last_backup_path = backup_path
            self.update_status(f"已创建备份 {backup_path}")
        except Exception as exc:  # noqa: BLE001
            # 备份失败不阻断合并，仅提示
            QMessageBox.warning(
                self, "备份失败",
                f"创建备份失败，合并将继续:\n{exc}",
            )

    # ------------------------------------------------------------------ #
    # 交换左右
    # ------------------------------------------------------------------ #
    def swap_sides(self) -> None:
        """交换左右两侧的路径/工作簿/数据/Sheet 名称与下拉。"""
        (
            self.left_path,
            self.right_path,
        ) = self.right_path, self.left_path
        self.left_wb, self.right_wb = self.right_wb, self.left_wb
        (
            self.left_sheet_data,
            self.right_sheet_data,
        ) = self.right_sheet_data, self.left_sheet_data
        (
            self.left_sheet_name,
            self.right_sheet_name,
        ) = self.right_sheet_name, self.left_sheet_name

        self._swap_combos()
        self._update_file_label("left")
        self._update_file_label("right")

        self._refresh_tables()
        self._run_compare()
        self.update_status("已交换左右两侧")

    def _swap_combos(self) -> None:
        """交换左右 Sheet 下拉的项目与当前选择（屏蔽信号）。"""
        lc, rc = self.left_sheet_combo, self.right_sheet_combo
        lc.blockSignals(True)
        rc.blockSignals(True)
        try:
            left_items = [lc.itemText(i) for i in range(lc.count())]
            left_idx = lc.currentIndex()
            right_items = [rc.itemText(i) for i in range(rc.count())]
            right_idx = rc.currentIndex()

            lc.clear()
            lc.addItems(right_items)
            if right_items:
                lc.setCurrentIndex(min(right_idx, len(right_items) - 1))
            rc.clear()
            rc.addItems(left_items)
            if left_items:
                rc.setCurrentIndex(min(left_idx, len(left_items) - 1))
        finally:
            lc.blockSignals(False)
            rc.blockSignals(False)

    # ------------------------------------------------------------------ #
    # 差异导航（Task 6 实装）
    # ------------------------------------------------------------------ #
    def next_diff(self) -> None:
        """跳转到下一个差异行（循环），并在两侧表格同步选中与滚动。"""
        indices = self.diff_result.diff_row_indices if self.diff_result else []
        if not indices:
            self.update_status("无差异")
            return
        self._current_diff_pos = (self._current_diff_pos + 1) % len(indices)
        self._goto_diff(indices[self._current_diff_pos], len(indices))

    def prev_diff(self) -> None:
        """跳转到上一个差异行（循环），并在两侧表格同步选中与滚动。"""
        indices = self.diff_result.diff_row_indices if self.diff_result else []
        if not indices:
            self.update_status("无差异")
            return
        if self._current_diff_pos <= 0:
            self._current_diff_pos = len(indices) - 1
        else:
            self._current_diff_pos -= 1
        self._goto_diff(indices[self._current_diff_pos], len(indices))

    def _goto_diff(self, target_row: int, total: int) -> None:
        """选中并滚动两侧表格至指定对齐行，更新状态栏。"""
        for table in (self.left_table, self.right_table):
            if 0 <= target_row < table.rowCount():
                table.setCurrentCell(target_row, 0)
                item = table.item(target_row, 0)
                if item is not None:
                    table.scrollToItem(item)
        pos = self._current_diff_pos + 1
        self.update_status(f"差异 {pos}/{total}")

    def show_statistics(self) -> None:
        """弹窗展示当前比较统计。"""
        if not self.diff_result:
            QMessageBox.information(self, "统计", "尚未进行比较，无统计数据。")
            return
        s = self.diff_result.stats
        total = len(self.diff_result.aligned_rows)
        info = (
            f"对齐行总数: {total}\n"
            f"相同: {s['same']}\n"
            f"不同: {s['different']}\n"
            f"仅左: {s['left_only']}\n"
            f"仅右: {s['right_only']}\n"
            f"差异单元格: {len(self.diff_result.diff_cell_set)}"
        )
        QMessageBox.information(self, "差异统计", info)

    # ------------------------------------------------------------------ #
    # 列设置 / 合并 / 上下文菜单（Task 8 实装）
    # ------------------------------------------------------------------ #
    def open_column_settings(self) -> None:
        """打开列设置对话框，配置 Key 列与忽略列后触发重新比较。"""
        if not self.left_sheet_data and not self.right_sheet_data:
            QMessageBox.warning(self, "列设置", "请先加载文件")
            return
        # 列标签：取两侧列数较多者，保证覆盖所有列
        left_labels = (
            self.left_sheet_data.header_labels if self.left_sheet_data else []
        )
        right_labels = (
            self.right_sheet_data.header_labels if self.right_sheet_data else []
        )
        if len(left_labels) >= len(right_labels):
            labels = left_labels
        else:
            labels = right_labels

        dlg = ColumnSettingsDialog(labels, self.key_cols, self.ignore_cols, self)
        if dlg.exec() == QDialog.Accepted:
            self.key_cols = dlg.get_key_cols()
            self.ignore_cols = dlg.get_ignore_cols()
            self.update_status(
                f"Key 列: {self.key_cols} | 忽略列: {self.ignore_cols}"
            )
            self._run_compare()

    def merge(self) -> None:
        """弹出合并策略选择菜单。"""
        if not self.diff_result:
            self.update_status("无差异结果，无法合并")
            return
        menu = QMenu(self)
        menu.addAction("右覆盖（以右侧为准）", self.merge_right_wins)
        menu.addAction("左覆盖（以左侧为准）", self.merge_left_wins)
        menu.addAction("追加差异行", self.merge_append)
        menu.addAction("手动合并", self.merge_manual)

        btn = self.sender()
        if isinstance(btn, QAction):
            # 由工具栏 QAction 触发：在工具栏按钮下方弹出
            widget = self.toolbar.widgetForAction(btn)
            if widget is not None:
                pos = widget.mapToGlobal(widget.rect().bottomLeft())
                menu.exec(pos)
                return
        menu.exec(QCursor.pos())

    def merge_right_wins(self) -> None:
        """右覆盖：以右侧为准，将右侧差异覆盖到左侧。"""
        if not self.diff_result:
            self.update_status("无差异结果，无法合并")
            return
        btn = QMessageBox.question(
            self, "右覆盖", "将以右侧为准覆盖左侧差异，继续？"
        )
        if btn != QMessageBox.Yes:
            return
        try:
            ExcelMerger.merge_right_to_left(
                self.diff_result, self.left_sheet_data, self.right_sheet_data
            )
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "合并失败", f"右覆盖合并失败:\n{exc}")
            return
        self._post_merge("left")
        self.update_status("已合并（右覆盖）— 请保存")

    def merge_left_wins(self) -> None:
        """左覆盖：以左侧为准，将左侧差异覆盖到右侧。"""
        if not self.diff_result:
            self.update_status("无差异结果，无法合并")
            return
        btn = QMessageBox.question(
            self, "左覆盖", "将以左侧为准覆盖右侧差异，继续？"
        )
        if btn != QMessageBox.Yes:
            return
        try:
            ExcelMerger.merge_left_to_right(
                self.diff_result, self.left_sheet_data, self.right_sheet_data
            )
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "合并失败", f"左覆盖合并失败:\n{exc}")
            return
        self._post_merge("right")
        self.update_status("已合并（左覆盖）— 请保存")

    def merge_append(self) -> None:
        """追加差异行：仅把右侧独占行追加到左侧末尾。"""
        if not self.diff_result:
            self.update_status("无差异结果，无法合并")
            return
        btn = QMessageBox.question(
            self, "追加差异行", "将把右侧独占行追加到左侧末尾，继续？"
        )
        if btn != QMessageBox.Yes:
            return
        try:
            ExcelMerger.append_rows(
                self.diff_result,
                self.left_sheet_data,
                self.right_sheet_data,
                self.key_cols,
            )
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "合并失败", f"追加差异行失败:\n{exc}")
            return
        self._post_merge("left")
        self.update_status("已合并（追加差异行）— 请保存")

    def merge_manual(self) -> None:
        """手动合并：提示用户通过右键菜单逐项合并。"""
        self.update_status("请使用右键菜单逐项合并")

    def _post_merge(self, side: str) -> None:
        """合并后处理：重新提取数据 -> 刷新表格 -> 重新比较 -> 标记未保存。

        side 为合并写入的目标侧（"left" 或 "right"）。由于合并器就地修改了
        worksheet 对象，需重新提取 SheetData 以反映最新值。
        """
        if side == "left" and self.left_sheet_data is not None:
            self.left_sheet_data = ExcelLoader.extract_sheet_data(
                self.left_sheet_data.worksheet
            )
        elif side == "right" and self.right_sheet_data is not None:
            self.right_sheet_data = ExcelLoader.extract_sheet_data(
                self.right_sheet_data.worksheet
            )
        # 清除旧差异结果，避免按过期对齐渲染
        self.diff_result = None
        self._refresh_tables()
        self._run_compare()
        self._dirty = True
        self._backup_done = False
        self.backup_label.setText("备份: 待保存")

    def on_context_menu(self, pos) -> None:
        """表格右键菜单：根据触发表格决定“复制到对侧”的文案与方向。

        记录上下文（_ctx_table/_ctx_row/_ctx_col）供 _ctx_* 回调读取。
        """
        table = self.sender()
        if table is None:
            return
        self._ctx_table = table
        self._ctx_row = table.currentRow()
        self._ctx_col = table.currentColumn()

        is_left = table is self.left_table
        other_label = "右侧" if is_left else "左侧"
        menu = QMenu(self)
        menu.addAction(f"复制行到{other_label}", self._ctx_copy_row)
        menu.addAction(f"复制单元格到{other_label}", self._ctx_copy_single_cell)
        menu.addAction("对齐行", self._ctx_align_rows)
        menu.addAction("复制单元格值", self._ctx_copy_cell)
        menu.addSeparator()
        menu.addAction("交换左右", self.swap_sides)
        global_pos = table.viewport().mapToGlobal(pos)
        menu.exec(global_pos)

    def _ctx_resolve_pair(self):
        """根据 _ctx_table/_ctx_row 解析当前对齐行，返回 (source, src_idx, target, tgt_idx, side)。

        若该行仅一侧存在或无差异结果，返回 None 并更新状态。
        """
        if not self.diff_result or self._ctx_row < 0:
            self.update_status("无有效行可操作")
            return None
        if self._ctx_row >= len(self.diff_result.aligned_rows):
            self.update_status("行索引越界")
            return None
        pair = self.diff_result.aligned_rows[self._ctx_row]
        if pair.left_row is None or pair.right_row is None:
            self.update_status("该行仅一侧存在，无法对应复制")
            return None
        is_left = self._ctx_table is self.left_table
        if is_left:
            return (
                self.left_sheet_data,
                pair.left_row,
                self.right_sheet_data,
                pair.right_row,
                "right",
            )
        return (
            self.right_sheet_data,
            pair.right_row,
            self.left_sheet_data,
            pair.left_row,
            "left",
        )

    def _ctx_copy_row(self) -> None:
        """右键：把当前行整行（值与样式）复制到对侧对应行。"""
        resolved = self._ctx_resolve_pair()
        if resolved is None:
            return
        source, src_idx, target, tgt_idx, side = resolved
        max_col = self.diff_result.max_col
        try:
            ExcelMerger.copy_row_to_other(source, src_idx, target, tgt_idx, max_col)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "复制失败", f"复制行失败:\n{exc}")
            return
        self._post_merge(side)
        self.update_status("已复制行到对侧")

    def _ctx_copy_single_cell(self) -> None:
        """右键：把当前单元格（值与样式）复制到对侧对应位置。"""
        if self._ctx_col < 0:
            self.update_status("无有效单元格可复制")
            return
        resolved = self._ctx_resolve_pair()
        if resolved is None:
            return
        source, src_idx, target, tgt_idx, side = resolved
        try:
            ExcelMerger.copy_single_cell(
                source, src_idx, self._ctx_col, target, tgt_idx
            )
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "复制失败", f"复制单元格失败:\n{exc}")
            return
        self._post_merge(side)
        self.update_status("已复制单元格到对侧")

    def _ctx_align_rows(self) -> None:
        """对齐行：按 Key 列对齐已在比较阶段自动完成，此处仅提示。"""
        self.update_status("对齐行（按 Key 对齐已自动完成）")

    def _ctx_copy_cell(self) -> None:
        """右键：复制当前单元格文本到系统剪贴板。"""
        if self._ctx_table is None:
            return
        item = self._ctx_table.item(self._ctx_row, self._ctx_col)
        if item is not None:
            QApplication.clipboard().setText(item.text())
            self.update_status("已复制单元格值")
        else:
            self.update_status("无单元格值可复制")

    # ------------------------------------------------------------------ #
    # 编辑类操作（多为桩）
    # ------------------------------------------------------------------ #
    def undo(self) -> None:
        self.update_status("TODO: 撤销（Task 8/9）")

    def redo(self) -> None:
        self.update_status("TODO: 重做（Task 8/9）")

    def align_rows(self) -> None:
        self.update_status("TODO: 对齐行（Task 8）")

    def copy_to_left(self) -> None:
        self.update_status("TODO: 复制到左侧（Task 8）")

    def copy_to_right(self) -> None:
        self.update_status("TODO: 复制到右侧（Task 8）")

    def copy_cell(self) -> None:
        self.update_status("TODO: 复制单元格（Task 8）")

    def paste(self) -> None:
        self.update_status("TODO: 粘贴（Task 8）")

    # ------------------------------------------------------------------ #
    # 视图类操作
    # ------------------------------------------------------------------ #
    def _on_toggle_row_numbers(self, checked: bool) -> None:
        self.show_row_numbers = checked
        self.left_table.verticalHeader().setVisible(checked)
        self.right_table.verticalHeader().setVisible(checked)

    def _on_view_filter_changed(self) -> None:
        """显示差异/相同/仅左/仅右 过滤变化 —— 触发重渲染（Task 6 实装过滤）。"""
        self._render_diffs()

    def resize_columns(self) -> None:
        """按内容自动调整列宽。"""
        for table in (self.left_table, self.right_table):
            table.resizeColumnsToContents()
            # 末列保持拉伸
            table.horizontalHeader().setStretchLastSection(True)

    # ------------------------------------------------------------------ #
    # 工具 / 帮助
    # ------------------------------------------------------------------ #
    def show_file_format(self) -> None:
        QMessageBox.information(
            self, "文件格式",
            "本工具仅支持 Microsoft Excel 2007+ 的 .xlsx 格式。\n"
            "旧版 .xls 文件请先在 Excel 中另存为 .xlsx。",
        )

    def show_options(self) -> None:
        self.update_status("TODO: 选项")

    def about(self) -> None:
        QMessageBox.about(
            self, "关于",
            "表格比较与合并工具\n\n"
            "基于 PySide6 + openpyxl 实现 Excel 工作表的智能比较与合并。\n"
            "支持差异可视化、多策略合并、合并单元格与公式保留。",
        )

    def show_docs(self) -> None:
        self.update_status("TODO: 文档")

    def show_shortcuts(self) -> None:
        QMessageBox.information(
            self, "快捷键列表",
            "打开左: Ctrl+L\n"
            "打开右: Ctrl+R\n"
            "保存: Cmd/Ctrl+S\n"
            "另存为: Ctrl+Shift+S\n"
            "交换左右: Ctrl+Shift+X\n"
            "重新比较: F5\n"
            "撤销: Cmd/Ctrl+Z\n"
            "重做: Cmd/Ctrl+Shift+Z\n"
            "复制单元格: Cmd/Ctrl+C\n"
            "粘贴: Cmd/Ctrl+V",
        )

    # ------------------------------------------------------------------ #
    # 同步滚动
    # ------------------------------------------------------------------ #
    def _sync_scroll(self, source: QTableWidget, target: QTableWidget, axis: str, value: int) -> None:
        """同步另一侧表格的滚动条位置（带防递归守卫）。"""
        if self._sync_scroll_blocked:
            return
        self._sync_scroll_blocked = True
        try:
            if axis == "v":
                target.verticalScrollBar().setValue(value)
            else:
                target.horizontalScrollBar().setValue(value)
        finally:
            self._sync_scroll_blocked = False

    # ------------------------------------------------------------------ #
    # 工具方法
    # ------------------------------------------------------------------ #
    def update_status(self, text: str) -> None:
        """更新状态栏左侧提示文本。"""
        self.status_label.setText(text)

    def closeEvent(self, event) -> None:  # noqa: N802
        """关闭窗口：直接接受（Task 9 可能加入未保存确认）。"""
        event.accept()
