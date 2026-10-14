from __future__ import annotations

import os
import sys
import threading
from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal, QUrl, QTimer, QSettings
from PySide6.QtGui import QDesktopServices, QFont, QColor, QIcon
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QFrame, QButtonGroup, QCheckBox, QFileDialog, QTableWidget,
    QTableWidgetItem, QHeaderView, QAbstractItemView, QProgressBar, QSpinBox,
    QLineEdit, QMessageBox, QScrollArea,
)

from .engine import PRESETS, RESULT_FOLDER, Settings, Cancelled, compress_workbook, discover


def size_text(n):
    if n >= 1_000_000_000:
        return f"{n / 1_000_000_000:.2f} GB"
    if n >= 1_000_000:
        return f"{n / 1_000_000:.1f} MB"
    return f"{n / 1000:.1f} KB"


STYLE = """
QWidget { font-family: 'Microsoft YaHei UI', 'PingFang SC', sans-serif; font-size: 13px; color: #213b35; }
QMainWindow, QWidget#page { background: #f3f6f3; }
QLabel#title { font-size: 27px; font-weight: 700; }
QLabel#subtitle, QLabel#muted { color: #667a73; }
QLabel#brand { color: #137153; font-weight: 700; letter-spacing: 2px; }
QLabel#badge { background: #e3eee7; color: #236247; border-radius: 12px; padding: 6px 12px; }
QFrame#drop { background: #ffffff; border: 1px dashed #98b7a6; border-radius: 14px; }
QFrame#settings { background: #ffffff; border: 1px solid #e0e8e2; border-radius: 12px; }
QPushButton { background: #ffffff; border: 1px solid #cedbd1; border-radius: 7px; padding: 9px 17px; }
QPushButton:hover { background: #ecf3ec; border-color: #86a993; }
QPushButton:focus { border: 2px solid #287254; }
QPushButton:disabled { color: #97a39c; background: #edf0ed; border-color: #e1e7e1; }
QPushButton#primary { color: white; background: #186344; border: 1px solid #186344; font-weight: 600; }
QPushButton#primary:hover { background: #104e35; }
QPushButton#primary:disabled { background: #a6bcae; border-color: #a6bcae; }
QPushButton#preset { text-align: left; padding: 12px 14px; background: #fafcf9; }
QPushButton#preset:checked { background: #e8f2e9; border: 2px solid #287254; }
QPushButton#link { border: none; background: transparent; color: #286247; padding: 6px; }
QTableWidget { background: #ffffff; border: 1px solid #dce6dd; border-radius: 8px; gridline-color: #edf2ed; selection-background-color: #e4eee5; selection-color: #213b35; }
QHeaderView::section { background: #eaf0e9; padding: 10px; border: none; color: #52665b; font-weight: 600; }
QTableWidget::item { padding: 8px; }
QLineEdit, QSpinBox { border: 1px solid #d4dfd6; border-radius: 6px; background: white; padding: 6px; }
QProgressBar { background: #e2eae2; border: none; border-radius: 4px; min-height: 8px; max-height: 8px; }
QProgressBar::chunk { background: #39845b; border-radius: 4px; }
QCheckBox { spacing: 7px; }
"""


class BatchWorker(QThread):
    progress = Signal(int, object)
    completed = Signal(int, object)
    failed = Signal(int, str)
    cancelled = Signal(int)

    def __init__(self, jobs, settings, output):
        super().__init__()
        self.jobs, self.settings, self.output = jobs, settings, output
        self.cancel_event = threading.Event()

    def run(self):
        for index, path in self.jobs:
            if self.cancel_event.is_set():
                break
            try:
                result = compress_workbook(path, self.settings, self.output, self.cancel_event,
                                           lambda p: self.progress.emit(index, p))
                self.completed.emit(index, result)
            except Cancelled:
                self.cancelled.emit(index)
                break
            except Exception as error:
                self.failed.emit(index, str(error) or type(error).__name__)


class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("轻表 · Excel 图片压缩")
        self.setWindowIcon(QIcon(str(Path(__file__).parent / "assets" / "logo.png")))
        self.resize(1100, 850)
        self.setMinimumSize(900, 480)
        self.setAcceptDrops(True)
        self.paths = []
        self.results = {}
        self.worker = None
        self.closing = False
        self.saved_bytes = 0
        self.finished_count = 0
        self.success_count = 0
        self.failure_count = 0
        self.prefs = QSettings("QingBiao", "ExcelImageCompressor")
        page = QWidget(objectName="page")
        self.page = page
        page.setMinimumHeight(780)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setWidget(page)
        self.setCentralWidget(scroll)
        layout = QVBoxLayout(page)
        layout.setContentsMargins(30, 24, 30, 24)
        layout.setSpacing(14)
        top = QHBoxLayout()
        logo = QLabel()
        logo.setPixmap(self.windowIcon().pixmap(40, 40))
        logo.setFixedSize(40, 40)
        logo.setAccessibleName("轻表标志：表格与轻叶")
        top.addWidget(logo)
        top.addSpacing(4)
        top.addWidget(QLabel("轻表 / QINGBIAO", objectName="brand"))
        top.addStretch()
        top.addWidget(QLabel("本机处理 · 原文件保留", objectName="badge"))
        layout.addLayout(top)
        layout.addWidget(QLabel("让表格轻一点。", objectName="title"))
        layout.addWidget(QLabel("批量压缩 Excel 图片，保留数据、公式与图片位置。支持 WPS 单元格图片。", objectName="subtitle"))

        drop = QFrame(objectName="drop")
        dl = QVBoxLayout(drop)
        dl.setContentsMargins(18, 15, 18, 15)
        dl.setSpacing(10)
        label = QLabel("拖入 .xlsx 文件或文件夹")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("font-size: 17px; font-weight: 600;")
        dl.addWidget(label)
        row = QHBoxLayout()
        row.addStretch()
        self.add_files = QPushButton("添加文件")
        self.add_files.clicked.connect(self.choose_files)
        self.add_folder = QPushButton("添加文件夹")
        self.add_folder.clicked.connect(self.choose_folder)
        self.recursive = QCheckBox("包含子文件夹")
        row.addWidget(self.add_files)
        row.addWidget(self.add_folder)
        row.addWidget(self.recursive)
        row.addStretch()
        dl.addLayout(row)
        layout.addWidget(drop)

        settings = QFrame(objectName="settings")
        sl = QVBoxLayout(settings)
        sl.setContentsMargins(18, 14, 18, 14)
        sl.setSpacing(12)
        presets = QHBoxLayout()
        self.preset_group = QButtonGroup(self)
        self.preset_buttons = {}
        for key, title, detail in (
            ("clear", "清晰优先", "保留像素尺寸 · 适合放大查看"),
            ("balanced", "均衡压缩 · 推荐", "长边最多 2000 px · 日常查看"),
            ("small", "体积优先", "长边最多 1600 px · 方便传输"),
        ):
            button = QPushButton(f"{title}\n{detail}", objectName="preset")
            button.setMinimumHeight(64)
            button.setCheckable(True)
            self.preset_group.addButton(button)
            button.clicked.connect(lambda checked=False, k=key: self.set_preset(k))
            self.preset_buttons[key] = button
            presets.addWidget(button)
        sl.addLayout(presets)
        self.advanced_toggle = QPushButton("高级设置 ▸", objectName="link")
        self.advanced_toggle.setCheckable(True)
        self.advanced_toggle.toggled.connect(self.toggle_advanced)
        sl.addWidget(self.advanced_toggle, alignment=Qt.AlignmentFlag.AlignLeft)
        self.advanced = QWidget()
        advanced = QHBoxLayout(self.advanced)
        advanced.setContentsMargins(0, 0, 0, 0)
        advanced.addWidget(QLabel("图片质量"))
        self.quality = QSpinBox()
        self.quality.setRange(40, 95)
        self.quality.setAccessibleName("图片质量")
        advanced.addWidget(self.quality)
        advanced.addWidget(QLabel("最长边"))
        self.edge = QSpinBox()
        self.edge.setRange(800, 8000)
        self.edge.setSingleStep(100)
        self.edge.setSuffix(" px")
        self.edge.setAccessibleName("图片最长边")
        advanced.addWidget(self.edge)
        self.original_size = QCheckBox("保留原始分辨率")
        self.original_size.toggled.connect(lambda value: self.edge.setEnabled(not value))
        advanced.addWidget(self.original_size)
        advanced.addStretch()
        self.advanced.hide()
        sl.addWidget(self.advanced)
        output = QHBoxLayout()
        output.addWidget(QLabel("输出位置"))
        self.output = QLineEdit()
        self.output.setReadOnly(True)
        self.output.setPlaceholderText("原文件旁的“压缩结果”文件夹")
        self.output.setText(str(self.prefs.value("output", "")))
        self.browse = QPushButton("更改")
        self.browse.setMinimumHeight(34)
        self.browse.clicked.connect(self.choose_output)
        self.reset_output = QPushButton("默认", objectName="link")
        self.reset_output.clicked.connect(self.output.clear)
        output.addWidget(self.output, 1)
        output.addWidget(self.browse)
        output.addWidget(self.reset_output)
        sl.addLayout(output)
        layout.addWidget(settings)
        self.set_preset(str(self.prefs.value("preset", "balanced")))

        queue = QHBoxLayout()
        self.queue_title = QLabel("文件队列 · 0")
        self.queue_title.setStyleSheet("font-weight: 600;")
        queue.addWidget(self.queue_title)
        queue.addStretch()
        self.remove = QPushButton("移除选中", objectName="link")
        self.remove.clicked.connect(self.remove_selected)
        self.clear = QPushButton("清空", objectName="link")
        self.clear.clicked.connect(self.clear_queue)
        queue.addWidget(self.remove)
        queue.addWidget(self.clear)
        layout.addLayout(queue)
        self.table = QTableWidget(0, 5)
        self.table.setMinimumHeight(100)
        self.table.setHorizontalHeaderLabels(["文件名称", "原大小", "压缩后", "节省", "状态"])
        self.table.verticalHeader().hide()
        self.table.verticalHeader().setDefaultSectionSize(45)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setShowGrid(False)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for col, width in [(1,100),(2,100),(3,80),(4,205)]:
            self.table.setColumnWidth(col, width)
        self.table.cellDoubleClicked.connect(lambda row, _: self.open_result(row))
        layout.addWidget(self.table, 1)
        self.status = QLabel("添加文件后即可开始。压缩完成后，双击文件行打开结果。", objectName="muted")
        self.status.setWordWrap(True)
        layout.addWidget(self.status)
        self.bar = QProgressBar()
        self.bar.setRange(0, 1000)
        self.bar.setTextVisible(False)
        layout.addWidget(self.bar)
        bottom = QHBoxLayout()
        self.open_folder = QPushButton("查看结果文件夹")
        self.open_folder.clicked.connect(self.show_output)
        self.open_folder.setEnabled(False)
        bottom.addWidget(self.open_folder)
        self.retry = QPushButton("重新压缩选中")
        self.retry.setToolTip("始终从原文件重新处理，保留已有压缩版本")
        self.retry.clicked.connect(self.retry_selected)
        bottom.addWidget(self.retry)
        bottom.addStretch()
        self.cancel_button = QPushButton("取消任务")
        self.cancel_button.setEnabled(False)
        self.cancel_button.clicked.connect(self.cancel)
        self.start = QPushButton("开始压缩", objectName="primary")
        self.start.setMinimumWidth(150)
        self.start.setEnabled(False)
        self.start.clicked.connect(self.start_batch)
        bottom.addWidget(self.cancel_button)
        bottom.addWidget(self.start)
        layout.addLayout(bottom)
        self.input_widgets = [drop, settings, self.remove, self.clear, self.retry]

    def toggle_advanced(self, visible):
        self.page.setMinimumHeight(835 if visible else 780)
        self.advanced.setVisible(visible)
        self.advanced_toggle.setText("高级设置 ▾" if visible else "高级设置 ▸")

    def set_preset(self, key):
        key = key if key in PRESETS else "balanced"
        self.preset_buttons[key].setChecked(True)
        setting = PRESETS[key]
        self.quality.setValue(setting.quality)
        self.original_size.setChecked(setting.max_edge == 0)
        self.edge.setValue(setting.max_edge or 2000)
        self.prefs.setValue("preset", key)

    def choose_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "添加 Excel 工作簿", "", "Excel 工作簿 (*.xlsx)")
        self.add_paths(paths)

    def choose_folder(self):
        path = QFileDialog.getExistingDirectory(self, "选择包含 Excel 文件的文件夹")
        if path:
            self.add_paths([path])

    def choose_output(self):
        path = QFileDialog.getExistingDirectory(self, "选择压缩结果保存位置", self.output.text())
        if path:
            self.output.setText(path)

    def dragEnterEvent(self, event):
        if not self.worker and event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        if not self.worker:
            self.add_paths([u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()])
            event.acceptProposedAction()

    def add_paths(self, paths):
        if self.worker or not paths:
            return
        found, errors = discover([Path(p) for p in paths], self.recursive.isChecked())
        existing = {os.path.normcase(str(p)) for p in self.paths}
        added = 0
        for path in found:
            if os.path.normcase(str(path)) in existing:
                continue
            row = self.table.rowCount()
            try:
                size = path.stat().st_size
            except OSError as e:
                errors.append(str(e))
                continue
            self.paths.append(path)
            self.table.insertRow(row)
            for col, value in enumerate([path.name, size_text(size), "—", "—", "等待中"]):
                item = QTableWidgetItem(value)
                item.setToolTip(str(path) if col == 0 else value)
                self.table.setItem(row, col, item)
            added += 1
        self.queue_title.setText(f"文件队列 · {len(self.paths)}")
        self.start.setEnabled(bool(self.paths))
        self.status.setText(f"已添加 {added} 个文件。仅接收 .xlsx，自动跳过重复文件、临时文件及压缩结果。")
        if errors:
            QMessageBox.warning(self, "部分文件未添加", "\n".join(errors[:12]))

    def clear_queue(self):
        if self.worker:
            return
        self.paths.clear()
        self.results.clear()
        self.table.setRowCount(0)
        self.queue_title.setText("文件队列 · 0")
        self.start.setEnabled(False)
        self.open_folder.setEnabled(False)
        self.bar.setValue(0)
        self.status.setText("队列已清空，原文件和已生成结果均保留。")

    def remove_selected(self):
        rows = sorted({i.row() for i in self.table.selectedIndexes()}, reverse=True)
        for row in rows:
            self.results.pop(self.paths[row], None)
            del self.paths[row]
            self.table.removeRow(row)
        self.queue_title.setText(f"文件队列 · {len(self.paths)}")
        self.start.setEnabled(bool(self.paths))

    def retry_selected(self):
        rows = sorted({i.row() for i in self.table.selectedIndexes()})
        if not rows:
            self.status.setText("请先选中要重新压缩的文件。")
            return
        for row in rows:
            self.table.item(row, 4).setText("等待中")
        self.start_batch(rows)

    def start_batch(self, rows=None):
        if self.worker:
            return
        # clicked(bool) must not be mistaken for a row selection.
        if not isinstance(rows, list):
            rows = [i for i in range(len(self.paths)) if self.table.item(i, 4).text() in ("等待中", "失败", "已取消")]
        if not rows:
            self.status.setText("队列已处理完毕。可选中文件，再点击“重新压缩选中”。")
            return
        self.batch_count = len(rows)
        for row in rows:
            self.results.pop(self.paths[row], None)
            self.table.item(row, 2).setText("—")
            self.table.item(row, 3).setText("—")
            self.table.item(row, 4).setForeground(QColor("#213b35"))
            self.table.item(row, 4).setToolTip("")
        self.finished_count = self.saved_bytes = self.success_count = self.failure_count = 0
        self.bar.setValue(0)
        output = Path(self.output.text()) if self.output.text() else None
        self.prefs.setValue("output", self.output.text())
        settings = Settings(self.quality.value(), 0 if self.original_size.isChecked() else self.edge.value())
        self.worker = BatchWorker([(i, self.paths[i]) for i in rows], settings, output)
        self.worker.progress.connect(self.on_progress)
        self.worker.completed.connect(self.on_completed)
        self.worker.failed.connect(self.on_failed)
        self.worker.cancelled.connect(lambda row: self.table.item(row, 4).setText("已取消"))
        self.worker.finished.connect(self.batch_finished)
        for widget in self.input_widgets:
            widget.setEnabled(False)
        self.start.setEnabled(False)
        self.cancel_button.setEnabled(True)
        self.status.setText("正在检查工作簿…")
        self.worker.start()

    def on_progress(self, row, p):
        text = p.phase + (f" {p.done}/{p.total}" if p.phase != "检查文件" else "")
        self.table.item(row, 4).setText(text)
        fraction = p.done / max(1, p.total)
        fraction = .85 * fraction if p.phase == "压缩图片" else .85 + .15 * fraction if p.phase == "校验文件" else 0
        self.bar.setValue(int(1000 * (self.finished_count + fraction) / self.batch_count))
        self.status.setText(f"{self.paths[row].name}  ·  {text}")

    def on_completed(self, row, result):
        self.results[self.paths[row]] = result
        self.finished_count += 1
        self.success_count += bool(result.output)
        self.saved_bytes += result.saved
        self.table.item(row, 2).setText(size_text(result.after) if result.output else "—")
        self.table.item(row, 3).setText(f"{result.saved / result.before:.0%}" if result.output else "—")
        item = self.table.item(row, 4)
        item.setText(result.status + (f" · {len(result.warnings)} 条提示" if result.warnings else ""))
        item.setForeground(QColor("#237447"))
        item.setToolTip("\n".join(result.warnings) or "图片与文件内容校验已通过")
        self.open_folder.setEnabled(any(r.output for r in self.results.values()))

    def on_failed(self, row, message):
        self.finished_count += 1
        self.failure_count += 1
        item = self.table.item(row, 4)
        item.setText("失败")
        item.setForeground(QColor("#ad4338"))
        item.setToolTip(message)
        self.table.item(row, 0).setToolTip(f"{self.paths[row]}\n{message}")

    def batch_finished(self):
        cancelled = self.worker.cancel_event.is_set()
        self.worker.deleteLater()
        self.worker = None
        for widget in self.input_widgets:
            widget.setEnabled(True)
        self.edge.setEnabled(not self.original_size.isChecked())
        self.start.setEnabled(bool(self.paths))
        self.cancel_button.setEnabled(False)
        self.cancel_button.setText("取消任务")
        self.bar.setValue(int(1000 * self.finished_count / self.batch_count) if cancelled else 1000)
        self.status.setText(f"{'已取消' if cancelled else '处理结束'} · 已处理 {self.finished_count}/{self.batch_count} 个文件，生成 {self.success_count} 个结果，失败 {self.failure_count} 个，共节省 {size_text(self.saved_bytes)}。")
        if self.closing:
            self.close()

    def cancel(self):
        if self.worker:
            self.worker.cancel_event.set()
            self.cancel_button.setText("正在取消…")
            self.cancel_button.setEnabled(False)

    def open_result(self, row):
        result = self.results.get(self.paths[row])
        if result and result.output:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(result.output)))
        elif self.table.item(row, 4).text() == "失败":
            QMessageBox.warning(self, "处理失败", self.table.item(row, 4).toolTip())

    def show_output(self):
        rows = sorted({i.row() for i in self.table.selectedIndexes()})
        results = [self.results.get(self.paths[i]) for i in rows] or list(reversed(self.results.values()))
        for result in results:
            if result and result.output:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(result.output.parent)))
                break

    def closeEvent(self, event):
        if self.worker:
            self.closing = True
            self.cancel()
            self.status.setText("正在取消并清理临时文件，完成后自动退出…")
            event.ignore()
        else:
            event.accept()


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("轻表")
    app.setStyle("Fusion")
    app.setStyleSheet(STYLE)
    window = Window()
    available = app.primaryScreen().availableGeometry()
    window.resize(min(1100, available.width() - 40), min(850, available.height() - 60))
    window.show()
    files = [p for p in sys.argv[1:] if not p.startswith("-")]
    if files:
        QTimer.singleShot(0, lambda: window.add_paths(files))
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
