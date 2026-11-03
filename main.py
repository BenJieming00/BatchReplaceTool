# -*- coding: utf-8 -*-
"""
批量替换工具（Windows 版） - PySide6 图形界面
布局与 macOS 版一致：
  顶部工具栏（右上角蓝色「运行」按钮）
  左侧：文件区（拖放 / 选择文件 / 选择文件夹 / 列表 / 清空）
  右侧：设置区（文本替换规则 / 4 个转换选项开关 / 输出目录）
  底部：运行日志
  状态栏：Office 引擎指示灯 + 进度 + 状态
"""
import os
import sys
import time

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QLineEdit, QCheckBox, QTextEdit, QListWidget,
    QListWidgetItem, QFileDialog, QFrame, QSplitter, QProgressBar, QToolBar,
)

from batch_runner import BatchRunner, JobConfig
from office_converter import word_available, excel_available

APP_NAME = "批量替换工具"

COLORS = {
    "info": "#8e8e93",
    "success": "#28a745",
    "warning": "#e67e22",
    "error": "#dc3545",
}

STYLE = """
QMainWindow, QWidget { background: #f5f5f7; color: #1d1d1f; font-size: 13px; }
QFrame#dropZone {
    background: #f0f2f7; border: 1.2px dashed #b8becb; border-radius: 10px;
}
QFrame#dropZone[active="true"] { background: #e3edff; border-color: #1777FF; }
QPushButton { padding: 4px 12px; border-radius: 5px; background: #ffffff;
              border: 1px solid #d1d1d6; }
QPushButton:hover { background: #f0f0f3; }
QPushButton:disabled { color: #b0b0b8; }
QPushButton#runButton { background: #1777FF; color: white; font-weight: 600;
                        padding: 5px 22px; border: none; border-radius: 6px; }
QPushButton#runButton:hover { background: #2b86ff; }
QPushButton#runButton:disabled { background: #9dc3f5; }
QLineEdit, QListWidget, QTextEdit { background: #ffffff; border: 1px solid #d9d9de;
                                    border-radius: 5px; padding: 3px 6px; }
QListWidget::item { padding: 3px 2px; }
QLabel#sectionHeader { font-size: 14px; font-weight: 600; }
QLabel#hint { color: #8e8e93; font-size: 11px; }
QStatusBar { background: #ececef; }
"""


class RuleRow(QWidget):
    """一行规则：查找 → 替换 → 删除"""

    def __init__(self, on_delete, can_delete=True, parent=None):
        super().__init__(parent)
        self.find_edit = QLineEdit()
        self.find_edit.setPlaceholderText("查找")
        self.find_edit.setMinimumWidth(120)
        self.replace_edit = QLineEdit()
        self.replace_edit.setPlaceholderText("替换为")
        self.replace_edit.setMinimumWidth(120)
        arrow = QLabel("→")
        arrow.setStyleSheet("color:#8e8e93;")
        self.delete_btn = QPushButton("✕")
        self.delete_btn.setFixedSize(26, 26)
        self.delete_btn.setEnabled(can_delete)
        self.delete_btn.clicked.connect(on_delete)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)
        lay.addWidget(self.find_edit, 1)
        lay.addWidget(arrow)
        lay.addWidget(self.replace_edit, 1)
        lay.addWidget(self.delete_btn)

    def values(self):
        return self.find_edit.text(), self.replace_edit.text()


class Worker(QThread):
    """后台线程执行批处理"""
    log_signal = Signal(str, str)
    progress_signal = Signal(float, str)
    finish_signal = Signal(object)

    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.runner = None

    def run(self):
        self.runner = BatchRunner(self.config)
        self.runner.run(
            on_log=lambda t, lvl: self.log_signal.emit(t, lvl),
            on_progress=lambda p, s: self.progress_signal.emit(p, s),
            on_finish=lambda s: self.finish_signal.emit(s),
        )

    def stop(self):
        if self.runner:
            self.runner.cancel()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setMinimumSize(940, 700)
        self.resize(960, 720)
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.sources = []          # 已添加的路径（去重）
        self.rule_rows = []        # 规则行控件
        self.worker = None
        self.is_running = False

        self._build_ui()
        self._check_engines()

    # MARK: - 界面搭建

    def _build_ui(self):
        # 顶部工具栏：右侧运行按钮
        toolbar = QToolBar()
        toolbar.setMovable(False)
        toolbar.setStyleSheet("QToolBar { background:#f5f5f7; border:none; spacing:8px; }")
        self.addToolBar(toolbar)
        spacer = QWidget()
        spacer.setSizePolicy(spacer.sizePolicy().Expanding, spacer.sizePolicy().Preferred)
        toolbar.addWidget(spacer)
        self.run_btn = QPushButton("▶ 运行")
        self.run_btn.setObjectName("runButton")
        self.run_btn.clicked.connect(self.start)
        toolbar.addWidget(self.run_btn)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # 主区：左文件区 + 右设置区
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self._build_file_column())
        splitter.addWidget(self._build_settings_column())
        splitter.setSizes([560, 400])
        splitter.setHandleWidth(1)
        root.addWidget(splitter, 1)

        # 底部日志
        root.addWidget(self._build_log_panel())

        # 状态栏
        self.status_bar = self.statusBar()
        self.word_badge = self._engine_badge("Word")
        self.excel_badge = self._engine_badge("Excel")
        self.status_bar.addWidget(QLabel("Office引擎"))
        self.status_bar.addWidget(self.word_badge)
        self.status_bar.addWidget(self.excel_badge)
        self.status_bar.addPermanentWidget(QLabel("  "))

        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedWidth(110)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.status_bar.addPermanentWidget(self.progress_bar)

        self.status_label = QLabel("就绪")
        self.status_bar.addPermanentWidget(self.status_label)

    def _engine_badge(self, name):
        w = QWidget()
        lay = QHBoxLayout(w)
        lay.setContentsMargins(4, 0, 6, 0)
        lay.setSpacing(4)
        dot = QLabel("●")
        dot.setObjectName("dot_" + name)
        dot.setStyleSheet("color:#b0b0b8; font-size:9px;")
        lay.addWidget(dot)
        lay.addWidget(QLabel(name))
        w._dot = dot
        return w

    def _build_file_column(self):
        col = QWidget()
        lay = QVBoxLayout(col)
        lay.setContentsMargins(14, 14, 14, 10)
        lay.setSpacing(10)

        # 拖放区
        self.drop_zone = QFrame()
        self.drop_zone.setObjectName("dropZone")
        self.drop_zone.setAcceptDrops(True)
        dz_lay = QVBoxLayout(self.drop_zone)
        dz_lay.setSpacing(6)
        icon = QLabel("⇩")
        icon.setStyleSheet("font-size:26px; color:#b8becb;")
        icon.setAlignment(Qt.AlignCenter)
        hint1 = QLabel("拖拽文件或文件夹到此处")
        hint1.setAlignment(Qt.AlignCenter)
        hint1.setStyleSheet("color:#8e8e93;")
        dz_lay.addStretch(1)
        dz_lay.addWidget(icon)
        dz_lay.addWidget(hint1)
        btns = QHBoxLayout()
        btn_files = QPushButton("选择文件")
        btn_files.clicked.connect(self.choose_files)
        btn_folder = QPushButton("选择文件夹")
        btn_folder.clicked.connect(self.choose_folder)
        btns.addStretch(1)
        btns.addWidget(btn_files)
        btns.addWidget(btn_folder)
        btns.addStretch(1)
        dz_lay.addLayout(btns)
        hint2 = QLabel("支持 .doc / .docx / .xls / .xlsx")
        hint2.setAlignment(Qt.AlignCenter)
        hint2.setStyleSheet("color:#c7c7cc; font-size:11px;")
        dz_lay.addWidget(hint2)
        dz_lay.addStretch(1)
        lay.addWidget(self.drop_zone)

        # 拖放事件
        self.drop_zone.dragEnterEvent = self._on_drag_enter
        self.drop_zone.dragLeaveEvent = self._on_drag_leave
        self.drop_zone.dropEvent = self._on_drop

        # 待处理文件标题栏
        head = QHBoxLayout()
        title = QLabel("待处理文件")
        title.setStyleSheet("font-weight:600;")
        self.count_label = QLabel("共 0 项")
        self.count_label.setStyleSheet("color:#8e8e93;")
        clear_btn = QPushButton("清空")
        clear_btn.setFlat(True)
        clear_btn.clicked.connect(self.clear_sources)
        head.addWidget(title)
        head.addStretch(1)
        head.addWidget(self.count_label)
        head.addWidget(clear_btn)
        lay.addLayout(head)

        # 文件列表
        self.file_list = QListWidget()
        self.file_list.setAcceptDrops(False)
        lay.addWidget(self.file_list, 1)
        return col

    def _build_settings_column(self):
        col = QWidget()
        col.setMinimumWidth(340)
        lay = QVBoxLayout(col)
        lay.setContentsMargins(16, 14, 16, 10)
        lay.setSpacing(8)

        # 文本替换规则
        header = QLabel("文本替换规则")
        header.setObjectName("sectionHeader")
        lay.addWidget(header)
        self.rules_box = QVBoxLayout()
        self.rules_box.setSpacing(6)
        lay.addLayout(self.rules_box)
        self.add_rule_btn = QPushButton("+ 添加规则")
        self.add_rule_btn.setFlat(True)
        self.add_rule_btn.setStyleSheet("color:#1777FF; text-align:left; padding:0;")
        self.add_rule_btn.clicked.connect(self.add_rule)
        lay.addWidget(self.add_rule_btn)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("color:#e2e2e6;")
        lay.addWidget(line)

        # 转换选项
        header2 = QLabel("转换选项")
        header2.setObjectName("sectionHeader")
        lay.addWidget(header2)
        self.chk_backup = QCheckBox("处理前备份源文件")
        self.chk_backup.setChecked(True)
        self.chk_ignore_case = QCheckBox("忽略大小写")
        self.chk_overwrite = QCheckBox("覆盖已存在目标文件")
        self.chk_overwrite.setChecked(True)
        self.chk_recursive = QCheckBox("递归包含子文件夹")
        for c in (self.chk_backup, self.chk_ignore_case, self.chk_overwrite, self.chk_recursive):
            lay.addWidget(c)

        line2 = QFrame()
        line2.setFrameShape(QFrame.HLine)
        line2.setStyleSheet("color:#e2e2e6;")
        lay.addWidget(line2)

        # 输出目录
        header3 = QLabel("输出目录")
        header3.setObjectName("sectionHeader")
        lay.addWidget(header3)
        out_row = QHBoxLayout()
        self.out_edit = QLineEdit()
        self.out_edit.setReadOnly(True)
        self.out_edit.setPlaceholderText("原文件所在目录")
        self.out_edit.setStyleSheet("color:#8e8e93;")
        btn_out = QPushButton("选择…")
        btn_out.clicked.connect(self.choose_output)
        btn_reset = QPushButton("↺")
        btn_reset.setFixedWidth(30)
        btn_reset.clicked.connect(self.reset_output)
        out_row.addWidget(self.out_edit, 1)
        out_row.addWidget(btn_out)
        out_row.addWidget(btn_reset)
        lay.addLayout(out_row)

        lay.addStretch(1)
        return col

    def _build_log_panel(self):
        panel = QWidget()
        panel.setFixedHeight(175)
        lay = QVBoxLayout(panel)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        head = QHBoxLayout()
        head.setContentsMargins(16, 6, 16, 6)
        title = QLabel("运行日志")
        title.setStyleSheet("font-weight:600; font-size:12px;")
        head.addWidget(title)
        head.addStretch(1)
        clear_btn = QPushButton("清空")
        clear_btn.setFlat(True)
        clear_btn.setStyleSheet("color:#8e8e93; font-size:11px;")
        clear_btn.clicked.connect(self.clear_logs)
        head.addWidget(clear_btn)
        lay.addLayout(head)
        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setStyleSheet(
            "border:none; border-top:1px solid #e2e2e6; border-radius:0;"
            "font-family:Menlo, Consolas, monospace; font-size:11px;"
        )
        lay.addWidget(self.log_edit)
        return panel

    # MARK: - 引擎检测

    def _check_engines(self):
        self.word_ok = word_available()
        self.excel_ok = excel_available()
        self.word_badge._dot.setStyleSheet(
            "color:%s; font-size:9px;" % ("#28a745" if self.word_ok else "#dc3545"))
        self.excel_badge._dot.setStyleSheet(
            "color:%s; font-size:9px;" % ("#28a745" if self.excel_ok else "#dc3545"))
        if not self.word_ok:
            self.log("未检测到 Microsoft Word（doc 转换不可用）", "warning")
        if not self.excel_ok:
            self.log("未检测到 Microsoft Excel（xls 转换不可用）", "warning")

    # MARK: - 源文件操作

    def _add_source(self, path):
        p = os.path.abspath(path)
        if p not in self.sources:
            self.sources.append(p)
            self._refresh_list()

    def choose_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "选择文件", "",
            "Office 文件 (*.doc *.docx *.xls *.xlsx)")
        for p in paths:
            self._add_source(p)

    def choose_folder(self):
        d = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if d:
            self._add_source(d)

    def clear_sources(self):
        self.sources.clear()
        self._refresh_list()

    def _refresh_list(self):
        self.file_list.clear()
        for p in self.sources:
            item = QListWidgetItem()
            name = os.path.basename(p)
            ext = os.path.splitext(p)[1].lower()
            tag = "📁" if os.path.isdir(p) else ("📊" if ext in (".xls", ".xlsx") else "📄")
            item.setText("%s %s\n%s" % (tag, name, os.path.dirname(p)))
            item.setData(Qt.UserRole, p)
            item.setToolTip(p)
            self.file_list.addItem(item)
        self.count_label.setText("共 %d 项" % len(self.sources))

    # MARK: - 拖放

    def _on_drag_enter(self, e):
        if e.mimeData().hasUrls():
            self.drop_zone.setProperty("active", "true")
            self.drop_zone.style().unpolish(self.drop_zone)
            self.drop_zone.style().polish(self.drop_zone)
            e.acceptProposedAction()

    def _on_drag_leave(self, e):
        self.drop_zone.setProperty("active", "false")
        self.drop_zone.style().unpolish(self.drop_zone)
        self.drop_zone.style().polish(self.drop_zone)

    def _on_drop(self, e):
        self.drop_zone.setProperty("active", "false")
        self.drop_zone.style().unpolish(self.drop_zone)
        self.drop_zone.style().polish(self.drop_zone)
        for url in e.mimeData().urls():
            p = url.toLocalFile()
            if p:
                self._add_source(p)
        e.acceptProposedAction()

    # MARK: - 规则操作

    def add_rule(self, find="", replace=""):
        row = RuleRow(
            on_delete=lambda: self._delete_rule(row),
            can_delete=True,
            parent=self,
        )
        row.find_edit.setText(find)
        row.replace_edit.setText(replace)
        self.rule_rows.append(row)
        self.rules_box.addWidget(row)
        self._update_delete_state()

    def _delete_rule(self, row):
        if len(self.rule_rows) <= 1:
            row.find_edit.clear()
            row.replace_edit.clear()
            return
        self.rules_box.removeWidget(row)
        row.deleteLater()
        self.rule_rows.remove(row)
        self._update_delete_state()

    def _update_delete_state(self):
        can = len(self.rule_rows) > 1
        for r in self.rule_rows:
            r.delete_btn.setEnabled(can)

    # MARK: - 输出目录

    def choose_output(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if d:
            self.out_edit.setText(d)
            self.out_edit.setStyleSheet("color:#1d1d1f;")

    def reset_output(self):
        self.out_edit.clear()
        self.out_edit.setStyleSheet("color:#8e8e93;")

    # MARK: - 日志

    def log(self, text, level="info"):
        color = COLORS.get(level, COLORS["info"])
        self.log_edit.append('<span style="color:%s">%s</span>' % (color, text))
        sb = self.log_edit.verticalScrollBar()
        sb.setValue(sb.maximum())

    def clear_logs(self):
        self.log_edit.clear()

    # MARK: - 运行

    def start(self):
        if self.is_running:
            return
        if not self.sources:
            self.log("请先选择源文件（拖拽或点击「选择文件」）", "warning")
            return
        if not (self.word_ok or self.excel_ok):
            self.log("未检测到 Office：doc/xls 转换需要 Microsoft Word / Excel，请先安装", "error")
            return

        rules = [r.values() for r in self.rule_rows]
        config = JobConfig(
            sources=list(self.sources),
            recursive=self.chk_recursive.isChecked(),
            backup_before_convert=self.chk_backup.isChecked(),
            ignore_case=self.chk_ignore_case.isChecked(),
            overwrite_existing=self.chk_overwrite.isChecked(),
            output_dir=self.out_edit.text().strip() or None,
            rules=[{"find": f, "replace": r} for f, r in rules],
        )

        self.is_running = True
        self.run_btn.setText("⏹ 停止")
        self.run_btn.setObjectName("runButton")
        self.run_btn.setStyleSheet("background:#e67e22; color:white; font-weight:600;"
                                   "padding:5px 22px; border:none; border-radius:6px;")
        self.progress_bar.setValue(0)
        self.status_label.setText("准备中…")
        self.log_edit.clear()
        self.log("批量替换工具已就绪", "info")

        self.worker = Worker(config)
        self.worker.log_signal.connect(self.log)
        self.worker.progress_signal.connect(self._on_progress)
        self.worker.finish_signal.connect(self._on_finish)
        self.worker.start()

    def _on_progress(self, p, s):
        self.progress_bar.setValue(int(p * 100))
        self.status_label.setText(s)

    def _on_finish(self, summary):
        self.is_running = False
        self.run_btn.setText("▶ 运行")
        self.run_btn.setStyleSheet("background:#1777FF; color:white; font-weight:600;"
                                   "padding:5px 22px; border:none; border-radius:6px;")
        self.progress_bar.setValue(100)
        self.status_label.setText(summary.status_line)
        self.log(summary.detail, "warning" if summary.failed else "success")
        self.worker = None

    def closeEvent(self, e):
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.worker.wait(3000)
        e.accept()


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setStyleSheet(STYLE)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
