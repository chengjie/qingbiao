import io
import zipfile
from pathlib import Path

import pytest
from PIL import Image


def make_workbook(path: Path, transparent=False, no_images=False, broken_image=False):
    image = Image.effect_noise((2200, 1600), 60).convert("RGBA" if transparent else "RGB")
    if transparent:
        image.putalpha(Image.linear_gradient("L").resize(image.size))
    buf = io.BytesIO()
    image.save(buf, "PNG")
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    entries = {
        "[Content_Types].xml": b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>',
        "xl/workbook.xml": b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>',
        "xl/worksheets/sheet1.xml": b'<worksheet><sheetData><row r="1"><c r="A1"><f>DISPIMG("ID_EXAMPLE",1)</f><v>0</v></c></row></sheetData></worksheet>',
        "xl/cellimages.xml": b'<cellImages xmlns="http://www.wps.cn/officeDocument/2017/etCustomData"><keep/></cellImages>',
        "xl/styles.xml": b'<styleSheet><custom-preserved/></styleSheet>',
        "xl/externalLinks/externalLink1.xml": b'<externalLink><keep/></externalLink>',
        "xl/unknown/vendor.bin": b"DO NOT CHANGE\x00\xff",
    }
    if not no_images:
        entries["xl/media/image1.png"] = b"bad-image" if broken_image else buf.getvalue()
        entries["xl/_rels/cellimages.xml.rels"] = f'<Relationships xmlns="{rel_ns}"><Relationship Id="rId1" Type="{rel_type}" Target="media/image1.png"/></Relationships>'.encode()
        entries["xl/drawings/_rels/drawing1.xml.rels"] = f"<Relationships xmlns='{rel_ns}'><Relationship Id='rId2' Type='{rel_type}' Target='../media/image1.png'/></Relationships>".encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in entries.items():
            z.writestr(name, data)
    return path


@pytest.fixture
def workbook(tmp_path):
    return make_workbook(tmp_path / "供应商 包装.xlsx")
