import argparse
import json
import signal
import threading
from dataclasses import asdict
from pathlib import Path

from .engine import PRESETS, Settings, Cancelled, discover, compress_workbook


def main():
    parser = argparse.ArgumentParser(description="批量压缩 Excel 内嵌图片，保留原文件")
    parser.add_argument("paths", nargs="+", type=Path)
    parser.add_argument("--preset", choices=PRESETS, default="balanced")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--recursive", action="store_true")
    parser.add_argument("--quality", type=int)
    parser.add_argument("--max-edge", type=int, help="0 表示原始尺寸")
    args = parser.parse_args()
    preset = PRESETS[args.preset]
    try:
        settings = Settings(args.quality if args.quality is not None else preset.quality,
                            args.max_edge if args.max_edge is not None else preset.max_edge)
    except ValueError as e:
        parser.error(str(e))
    cancel = threading.Event()
    signal.signal(signal.SIGINT, lambda *_: cancel.set())
    paths, errors = discover(args.paths, args.recursive)
    code = int(bool(errors) or not paths)
    for error in errors:
        print(json.dumps({"error": error}, ensure_ascii=False), flush=True)
    for path in paths:
        try:
            result = compress_workbook(path, settings, args.output, cancel)
            print(json.dumps(asdict(result), default=str, ensure_ascii=False), flush=True)
        except Cancelled:
            return 130
        except Exception as e:
            code = 1
            print(json.dumps({"source": str(path), "error": str(e)}, ensure_ascii=False), flush=True)
    return code


if __name__ == "__main__":
    raise SystemExit(main())
