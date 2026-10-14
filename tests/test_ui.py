import os
import time

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication
from excel_compressor.app import Window, STYLE
from conftest import make_workbook


def test_batch_ui_handles_failure_and_retries(tmp_path):
    app = QApplication.instance() or QApplication([])
    app.setStyleSheet(STYLE)
    window = Window()
    window.output.clear()
    window.set_preset("balanced")
    a = make_workbook(tmp_path / "正常.xlsx")
    b = tmp_path / "坏文件.xlsx"
    b.write_bytes(b"broken")
    window.add_paths([a, b, a])
    assert len(window.paths) == 2
    window.start_batch()
    deadline = time.monotonic() + 30
    while window.worker and time.monotonic() < deadline:
        app.processEvents()
        time.sleep(.02)
    assert window.worker is None
    assert window.table.item(0, 4).text() == "完成"
    assert window.table.item(1, 4).text() == "失败"
    assert window.success_count == window.failure_count == 1
    assert window.start.isEnabled()
    assert window.results[a].output.exists()
    window.close()


def test_preset_and_advanced_controls():
    app = QApplication.instance() or QApplication([])
    window = Window()
    window.set_preset("clear")
    assert window.original_size.isChecked() and not window.edge.isEnabled()
    window.set_preset("small")
    assert window.quality.value() == 60 and window.edge.value() == 1600
    assert not window.original_size.isChecked()
    window.close()


def test_cancel_keeps_pending_queue_available(tmp_path):
    app = QApplication.instance() or QApplication([])
    window = Window()
    window.output.clear()
    a = make_workbook(tmp_path / "取消.xlsx")
    b = make_workbook(tmp_path / "等待.xlsx")
    window.add_paths([a, b])
    window.start_batch()
    window.cancel()
    deadline = time.monotonic() + 30
    while window.worker and time.monotonic() < deadline:
        app.processEvents()
        time.sleep(.02)
    assert window.worker is None
    assert window.start.isEnabled()
    assert not list(tmp_path.rglob("*.tmp"))
    assert a.exists() and b.exists()
    window.close()
