from __future__ import annotations

import copy
import hashlib
import io
import os
import posixpath
import re
import tempfile
import threading
import warnings
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable
from urllib.parse import quote, unquote, urlsplit
from xml.etree import ElementTree as ET

from PIL import Image, UnidentifiedImageError

RESULT_FOLDER = "压缩结果"
SUFFIX = "-压缩版"
CHUNK = 1024 * 1024
MAX_IMAGE_BYTES = 128 * CHUNK
MAX_IMAGE_PIXELS = 40_000_000
MAX_XML_BYTES = 64 * CHUNK
IMAGE_REL = "/image"


class CompressionError(Exception):
    pass


class Cancelled(CompressionError):
    pass


@dataclass(frozen=True)
class Settings:
    quality: int = 70
    max_edge: int = 2000  # 0 = original dimensions

    def __post_init__(self):
        if not 40 <= self.quality <= 95:
            raise ValueError("图片质量必须在 40–95 之间")
        if self.max_edge != 0 and not 800 <= self.max_edge <= 8000:
            raise ValueError("最长边必须为 0（原尺寸）或 800–8000 像素")


PRESETS = {
    "clear": Settings(80, 0),
    "balanced": Settings(70, 2000),
    "small": Settings(60, 1600),
}


@dataclass
class Progress:
    phase: str
    done: int
    total: int
    detail: str = ""


@dataclass
class Result:
    source: Path
    output: Path | None
    before: int
    after: int
    images: int = 0
    changed: int = 0
    resized: int = 0
    status: str = "完成"
    warnings: list[str] = field(default_factory=list)

    @property
    def saved(self):
        return max(0, self.before - self.after)


def discover(paths: Iterable[Path], recursive=False) -> tuple[list[Path], list[str]]:
    """Do not follow directory symlinks or ingest this application's outputs."""
    found: dict[str, Path] = {}
    errors = []

    def add(p: Path):
        if p.name.startswith("~$") or re.search(r"-压缩版(?: \(\d+\))?\.xlsx$", p.name, re.I):
            return
        if p.suffix.lower() != ".xlsx":
            return
        try:
            resolved = p.resolve(strict=True)
            if resolved.is_file():
                found[os.path.normcase(str(resolved))] = resolved
        except OSError as e:
            errors.append(f"{p.name}：{e}")

    for value in paths:
        p = Path(value)
        if p.is_dir():
            for root, dirs, files in os.walk(p, followlinks=False, onerror=lambda e: errors.append(str(e))):
                dirs[:] = sorted(d for d in dirs if d != RESULT_FOLDER and not d.startswith(".")) if recursive else []
                for name in sorted(files):
                    add(Path(root) / name)
        elif p.exists():
            add(p)
        else:
            errors.append(f"文件不存在：{p}")
    return list(found.values()), errors


def _check(cancel):
    if cancel.is_set():
        raise Cancelled("已取消，未完成的临时文件已清理")


def _digest(stream, cancel):
    h = hashlib.sha256()
    while block := stream.read(CHUNK):
        _check(cancel)
        h.update(block)
    return h.digest()


def _copy(zin, zout, info, cancel):
    clone = copy.copy(info)
    h = hashlib.sha256()
    with zin.open(info) as source, zout.open(clone, "w", force_zip64=True) as target:
        while block := source.read(CHUNK):
            _check(cancel)
            target.write(block)
            h.update(block)
    return h.digest()


def _resolve(relname: str, target: str):
    if urlsplit(target).scheme or "\\" in target:
        return None
    base = posixpath.dirname(posixpath.dirname(relname)) if relname != "_rels/.rels" else ""
    target = unquote(target)
    return posixpath.normpath(target.lstrip("/") if target.startswith("/") else posixpath.join(base, target))


def _patch_rels(name, data, mapping):
    root = ET.fromstring(data)
    replacements = {}
    for rel in root:
        if rel.get("TargetMode") == "External" or not rel.get("Type", "").endswith(IMAGE_REL):
            continue
        target = rel.get("Target", "")
        old = _resolve(name, target)
        if old not in mapping:
            continue
        base = posixpath.dirname(posixpath.dirname(name)) if name != "_rels/.rels" else ""
        new = "/" + mapping[old] if target.startswith("/") else posixpath.relpath(mapping[old], base)
        replacements[target] = quote(new, safe="/.-_")
    # Keep WPS namespaces and every other byte unchanged.
    def replace(match):
        from xml.sax.saxutils import unescape
        target = unescape(match[2].decode("utf-8"), {"&quot;": '"', "&apos;": "'"})
        value = replacements.get(target)
        return b"Target=" + match[1] + value.encode() + match[1] if value else match[0]
    return re.sub(rb"\bTarget=([\"'])(.*?)\1", replace, data)


def _patch_types(data, mapping):
    root = ET.fromstring(data)
    if not mapping:
        return data
    # Existing files use their own namespace prefix (including no prefix).
    for old, new in mapping.items():
        for q in (b'"', b"'"):
            data = data.replace(b"PartName=" + q + ("/" + old).encode() + q,
                                b"PartName=" + q + ("/" + new).encode() + q)
    # A per-image override must describe the new encoding.
    def override(m):
        if any(("/" + new).encode() in m[0] for new in mapping.values()):
            return re.sub(rb"ContentType=([\"']).*?\1", b'ContentType="image/jpeg"', m[0])
        return m[0]
    data = re.sub(rb"<(?:\w+:)?Override\b[^>]*>", override, data)
    jpeg = next((n for n in root if n.get("Extension") == "jpeg"), None)
    if jpeg is not None and jpeg.get("ContentType") != "image/jpeg":
        raise CompressionError("文件中的 JPEG 类型定义异常，未修改原文件")
    if jpeg is None:
        match = re.search(rb"</((?:\w+:)?Types)\s*>", data)
        if not match:
            raise CompressionError("无法识别图片类型定义")
        prefix = match[1].rsplit(b":", 1)[0] + b":" if b":" in match[1] else b""
        addition = b"<" + prefix + b'Default Extension="jpeg" ContentType="image/jpeg"/>'
        data = data[:match.start()] + addition + data[match.start():]
    ET.fromstring(data)
    return data


def _compress_image(data, settings):
    with warnings.catch_warnings():
        warnings.simplefilter("error", Image.DecompressionBombWarning)
        with Image.open(io.BytesIO(data)) as im:
            if im.width * im.height > MAX_IMAGE_PIXELS:
                return None, "像素过大，保留原图"
            if im.format not in ("JPEG", "PNG", "MPO"):
                return None, f"{im.format} 格式保留原图"
            if im.format != "MPO" and getattr(im, "is_animated", False):
                return None, "动态图保留原图"
            im.load()
            old_size = im.size
            transparent = "transparency" in im.info or ("A" in im.getbands() and im.getchannel("A").getextrema()[0] < 255)
            converted = im.convert("RGBA" if transparent else "RGB")
            if settings.max_edge:
                converted.thumbnail((settings.max_edge, settings.max_edge), Image.Resampling.LANCZOS)
            options = {k: im.info[k] for k in ("icc_profile", "exif", "dpi") if k in im.info}
            buf = io.BytesIO()
            fmt = "PNG" if transparent else "JPEG"
            if transparent:
                converted.save(buf, fmt, optimize=True, **options)
            else:
                converted.save(buf, fmt, quality=settings.quality, subsampling=0 if not settings.max_edge else 2,
                               optimize=True, progressive=True, **options)
            payload = buf.getvalue()
            if len(payload) >= len(data) * .95:
                return None, None
            return (payload, fmt, converted.size, old_size != converted.size), None


def _commit(temp, directory, stem):
    # Reserve a name atomically; never replace a user's existing file.
    for count in range(10000):
        suffix = "" if not count else f" ({count})"
        dest = directory / f"{stem}{SUFFIX}{suffix}.xlsx"
        try:
            fd = os.open(dest, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
        except FileExistsError:
            continue
        os.close(fd)
        try:
            os.replace(temp, dest)
        except BaseException:
            dest.unlink(missing_ok=True)
            raise
        return dest
    raise CompressionError("输出目录中的同名文件过多")


def compress_workbook(source: Path, settings=PRESETS["balanced"], output_dir: Path | None = None,
                      cancel: threading.Event | None = None, progress: Callable[[Progress], None] | None = None) -> Result:
    cancel = cancel if cancel is not None else threading.Event()
    progress = progress or (lambda _: None)
    source = Path(source).resolve(strict=True)
    if source.suffix.lower() != ".xlsx":
        raise CompressionError("当前版本仅支持 .xlsx；请先在 WPS 或 Excel 中另存为 .xlsx")
    initial = source.stat()
    result = Result(source, None, initial.st_size, initial.st_size)
    directory = Path(output_dir) if output_dir else source.parent / RESULT_FOLDER
    temp = None
    try:
        _check(cancel)
        progress(Progress("检查文件", 0, 1))
        with zipfile.ZipFile(source) as zin:
            infos = zin.infolist()
            names = set(zin.namelist())
            if len(names) != len(infos) or len(infos) > 100000:
                raise CompressionError("文件包包含重复项目或项目过多")
            if not {"[Content_Types].xml", "xl/workbook.xml"}.issubset(names):
                raise CompressionError("不是有效的 Excel 工作簿")
            if any(i.flag_bits & 1 for i in infos):
                raise CompressionError("请先解除文件密码保护后再压缩")
            if any(n.startswith("_xmlsignatures/") for n in names):
                raise CompressionError("此文件带有数字签名，压缩会使签名失效，已跳过")
            media = [i for i in infos if i.filename.startswith("xl/media/") and not i.is_dir()]
            result.images = len(media)
            if not media:
                result.status = "无嵌入图片"
                return result
            directory.mkdir(parents=True, exist_ok=True)
            fd, path = tempfile.mkstemp(prefix=".excel-compress-", suffix=".tmp", dir=directory)
            os.close(fd)
            temp = Path(path)
            mapping, hashes, dimensions = {}, {}, {}
            with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED, compresslevel=6, allowZip64=True) as zout:
                for index, info in enumerate(media):
                    _check(cancel)
                    name = info.filename
                    candidate = None
                    if info.file_size > MAX_IMAGE_BYTES:
                        result.warnings.append(f"{name}：体积过大，保留原图")
                    else:
                        data = zin.read(info)
                        try:
                            candidate, note = _compress_image(data, settings)
                        except (UnidentifiedImageError, OSError, ValueError, Image.DecompressionBombError, Image.DecompressionBombWarning) as e:
                            note = f"无法处理，保留原图（{type(e).__name__}）"
                        if note:
                            result.warnings.append(f"{name}：{note}")
                    if candidate:
                        payload, fmt, size, resized = candidate
                        if fmt == "JPEG" and Path(name).suffix.lower() not in (".jpg", ".jpeg"):
                            new = str(Path(name).with_suffix("")).replace("\\", "/") + "-compressed.jpeg"
                            while new in names or new in mapping.values():
                                new = new[:-5] + "-copy.jpeg"
                            mapping[name] = new
                        else:
                            new = name
                        clone = copy.copy(info)
                        clone.filename = new
                        clone.compress_type = zipfile.ZIP_DEFLATED
                        zout.writestr(clone, payload)
                        hashes[new] = hashlib.sha256(payload).digest()
                        dimensions[new] = size
                        result.changed += 1
                        result.resized += int(resized)
                    else:
                        hashes[name] = _copy(zin, zout, info, cancel)
                    progress(Progress("压缩图片", index + 1, len(media), name))
                media_names = {i.filename for i in media}
                for info in infos:
                    _check(cancel)
                    if info.filename in media_names:
                        continue
                    name = info.filename
                    if name.endswith(".rels") or name == "[Content_Types].xml":
                        if info.file_size > MAX_XML_BYTES:
                            raise CompressionError("图片引用信息过大，已停止处理")
                        data = zin.read(info)
                        data = _patch_types(data, mapping) if name == "[Content_Types].xml" else _patch_rels(name, data, mapping)
                        zout.writestr(copy.copy(info), data)
                        hashes[name] = hashlib.sha256(data).digest()
                    else:
                        hashes[name] = _copy(zin, zout, info, cancel)
                zout.comment = zin.comment
            # Read every output entry (CRC + SHA256), decode modified images,
            # and independently compare relationship semantics to the source.
            with zipfile.ZipFile(temp) as check:
                if set(check.namelist()) != {mapping.get(n, n) for n in names}:
                    raise CompressionError("校验失败：文件包项目不一致")
                for index, info in enumerate(check.infolist()):
                    _check(cancel)
                    name = info.filename
                    with check.open(info) as stream:
                        if _digest(stream, cancel) != hashes[name]:
                            raise CompressionError(f"校验失败：{name}")
                    if name in dimensions:
                        with Image.open(io.BytesIO(check.read(name))) as im:
                            im.load()
                            if im.size != dimensions[name]:
                                raise CompressionError("校验失败：图片尺寸不一致")
                    if name.endswith(".rels"):
                        before, after = ET.fromstring(zin.read(name)), ET.fromstring(check.read(name))
                        if len(before) != len(after):
                            raise CompressionError("校验失败：图片引用数量变化")
                        for a, b in zip(before, after):
                            if a.tag != b.tag or {k:v for k,v in a.attrib.items() if k != "Target"} != {k:v for k,v in b.attrib.items() if k != "Target"}:
                                raise CompressionError("校验失败：引用属性变化")
                            if a.get("Target") != b.get("Target"):
                                old, new = _resolve(name, a.get("Target", "")), _resolve(name, b.get("Target", ""))
                                if new != mapping.get(old) or new not in hashes:
                                    raise CompressionError("校验失败：图片引用缺失")
                    progress(Progress("校验文件", index + 1, len(infos)))
        _check(cancel)
        final = source.stat()
        if (initial.st_size, initial.st_mtime_ns) != (final.st_size, final.st_mtime_ns):
            raise CompressionError("原文件在压缩期间被修改，请保存后重试")
        size = temp.stat().st_size
        if result.changed == 0 or size >= initial.st_size * .99:
            result.status = "无需压缩"
            return result
        result.output = _commit(temp, directory, source.stem)
        result.after = size
        return result
    except zipfile.BadZipFile as e:
        raise CompressionError("文件已损坏、带有打开密码，或不是有效的 .xlsx 文件") from e
    except ET.ParseError as e:
        raise CompressionError("文件内部的图片引用信息无法解析") from e
    finally:
        if temp is not None:
            temp.unlink(missing_ok=True)
