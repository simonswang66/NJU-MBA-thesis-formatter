#!/usr/bin/env python3
"""极简 OOXML 解包/打包（自实现，无外部依赖）。

docx 本质是 zip：unpack＝原样解压（不做美化与 run 合并，保真优先），
pack＝目录回 zip（ZIP_DEFLATED）。签名与 document-skills 的 office.unpack/pack
对齐，仅实现本 skill 用到的行为，其余关键字参数接受并忽略。
"""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path


def unpack(input_file: str, output_directory: str, **_ignored) -> tuple[None, str]:
    src, dst = Path(input_file), Path(output_directory)
    if dst.exists():
        shutil.rmtree(dst)
    dst.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src) as zf:
        zf.extractall(dst)
    return None, f"Unpacked {src.name} -> {dst}"


def pack(input_directory: str, output_file: str, **_ignored) -> tuple[None, str]:
    src, out = Path(input_directory), Path(output_file)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in sorted(src.rglob("*")):
            if p.is_file():
                zf.write(p, p.relative_to(src).as_posix())
    return None, f"Packed {src} -> {out.name}"
