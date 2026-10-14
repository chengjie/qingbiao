import io
import threading
import zipfile
from xml.etree import ElementTree as ET

import pytest
from PIL import Image

from conftest import make_workbook
from excel_compressor.engine import compress_workbook, discover, Settings, Cancelled, CompressionError, _patch_types, _compress_image


def test_preserves_wps_formulas_and_updates_both_image_references(workbook):
    original = workbook.read_bytes()
    r = compress_workbook(workbook)
    assert r.output and r.after < r.before * .7
    assert r.images == 1 and r.changed == 1 and r.resized == 1
    assert workbook.read_bytes() == original
    with zipfile.ZipFile(workbook) as a, zipfile.ZipFile(r.output) as b:
        for name in a.namelist():
            if name.startswith("xl/media/") or name.endswith(".rels") or name == "[Content_Types].xml":
                continue
            assert a.read(name) == b.read(name)
        assert b"media/image1-compressed.jpeg" in b.read("xl/_rels/cellimages.xml.rels")
        assert b"../media/image1-compressed.jpeg" in b.read("xl/drawings/_rels/drawing1.xml.rels")
        assert b"image/jpeg" in b.read("[Content_Types].xml")
        im = Image.open(io.BytesIO(b.read("xl/media/image1-compressed.jpeg")))
        assert im.size == (2000, 1455)
        assert b.testzip() is None


def test_transparency_is_preserved(tmp_path):
    r = compress_workbook(make_workbook(tmp_path / "透明.xlsx", transparent=True), Settings(60, 1000))
    with zipfile.ZipFile(r.output) as z:
        im = Image.open(io.BytesIO(z.read("xl/media/image1.png")))
        assert im.mode == "RGBA" and im.width == 1000
        assert im.getchannel("A").getextrema()[0] == 0


def test_clear_preset_keeps_dimensions(workbook):
    r = compress_workbook(workbook, Settings(80, 0))
    assert r.resized == 0


@pytest.mark.parametrize("phase", ["压缩图片", "校验文件"])
def test_cancel_never_publishes_partial_result(workbook, phase):
    cancel = threading.Event()
    def update(p):
        if p.phase == phase:
            cancel.set()
    with pytest.raises(Cancelled):
        compress_workbook(workbook, cancel=cancel, progress=update)
    assert list((workbook.parent / "压缩结果").iterdir()) == []


def test_output_collision_does_not_overwrite(workbook):
    a = compress_workbook(workbook)
    old = a.output.read_bytes()
    b = compress_workbook(workbook)
    assert a.output != b.output and "(1)" in b.output.name
    assert a.output.read_bytes() == old


def test_source_changed_during_processing_is_rejected(workbook):
    import os
    def update(p):
        if p.phase == "压缩图片":
            stat = workbook.stat()
            os.utime(workbook, ns=(stat.st_atime_ns, stat.st_mtime_ns + 1000000000))
    with pytest.raises(CompressionError, match="被修改"):
        compress_workbook(workbook, progress=update)
    assert list((workbook.parent / "压缩结果").iterdir()) == []


def test_no_images_and_broken_images(tmp_path):
    r = compress_workbook(make_workbook(tmp_path / "空.xlsx", no_images=True))
    assert r.output is None and r.status == "无嵌入图片"
    r = compress_workbook(make_workbook(tmp_path / "坏图.xlsx", broken_image=True))
    assert r.output is None and r.warnings


def test_bad_zip_and_wrong_extension(tmp_path):
    p = tmp_path / "坏.xlsx"
    p.write_bytes(b"not a zip")
    with pytest.raises(CompressionError):
        compress_workbook(p)
    p = tmp_path / "old.xls"
    p.touch()
    with pytest.raises(CompressionError, match="仅支持"):
        compress_workbook(p)


def test_discovery_filters_outputs_and_recurses(tmp_path):
    for name in ["a.xlsx", "~$a.xlsx", "a-压缩版.xlsx", "a-压缩版 (1).xlsx", "a.xls"]:
        (tmp_path / name).touch()
    sub = tmp_path / "子目录"
    sub.mkdir()
    (sub / "b.XLSX").touch()
    out = tmp_path / "压缩结果"
    out.mkdir()
    (out / "other.xlsx").touch()
    assert len(discover([tmp_path])[0]) == 1
    assert len(discover([tmp_path, tmp_path / "a.xlsx"], True)[0]) == 2


def test_types_with_namespace_and_override():
    data = b'<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"><ct:Override PartName="/xl/media/a.png" ContentType="image/png"/></ct:Types>'
    result = _patch_types(data, {"xl/media/a.png": "xl/media/a.jpeg"})
    root = ET.fromstring(result)
    assert root[0].get("ContentType") == "image/jpeg"
    assert root[1].get("Extension") == "jpeg"


def test_mpo_primary_picture_supported():
    a = Image.effect_noise((1200, 1000), 50).convert("RGB")
    b = Image.new("RGB", (1200, 1000), "red")
    buf = io.BytesIO()
    a.save(buf, "MPO", save_all=True, append_images=[b], quality=95)
    candidate, warning = _compress_image(buf.getvalue(), Settings(70, 1000))
    assert candidate and not warning
    im = Image.open(io.BytesIO(candidate[0]))
    assert im.format == "JPEG" and im.width == 1000


def test_signed_workbook_rejected(workbook):
    with zipfile.ZipFile(workbook, "a") as z:
        z.writestr("_xmlsignatures/sig1.xml", b"signature")
    with pytest.raises(CompressionError, match="数字签名"):
        compress_workbook(workbook)
