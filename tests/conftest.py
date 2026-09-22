from __future__ import annotations

import struct
import zlib
from pathlib import Path

import pytest


def _png(width: int = 320, height: int = 200) -> bytes:
    """Минимальный валидный PNG — чтобы тесты не тянули Pillow."""

    def chunk(tag: bytes, data: bytes) -> bytes:
        return (
            struct.pack(">I", len(data))
            + tag
            + data
            + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        )

    header = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    raw = b"".join(b"\x00" + b"\x7f\x7f\x7f" * width for _ in range(height))
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", header)
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


@pytest.fixture
def workdir(tmp_path: Path) -> Path:
    (tmp_path / "images").mkdir()
    (tmp_path / "images" / "scheme.png").write_bytes(_png())
    return tmp_path
