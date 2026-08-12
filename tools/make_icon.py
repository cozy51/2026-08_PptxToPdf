"""アプリアイコンのPNGを生成し、Base64テキストとして書き出すスクリプト。

リポジトリにバイナリファイルを置かずにアイコンを同梱するため、生成結果は
`assets/app_icon_<サイズ>.png.b64` へBase64で保存する。タイトルバーの小さな
表示でも輪郭が崩れないよう、縮小に頼らず各サイズを個別に描画する。
アイコンを描き変えたいときはこのファイルの図形定義を編集して再実行する。

    python tools/make_icon.py

標準ライブラリだけで動作し、Pillowなどの追加パッケージは不要。
"""

from __future__ import annotations

import base64
import struct
import zlib
from pathlib import Path

# 生成する図形はすべて256×256の座標系で定義し、出力サイズへ拡大縮小する。
DESIGN_SIZE = 256
# 1ピクセルあたりの副標本数。輪郭を滑らかにするために使う。
SUPERSAMPLE = 4
# tkのiconphotoへまとめて渡すサイズ。用途に応じて最適なものが選ばれる。
ICON_SIZES = (16, 32, 48, 256)

BADGE_TOP = (0x35, 0x67, 0x9B)
BADGE_BOTTOM = (0x27, 0x4C, 0x77)
SHEET = (0xFF, 0xFF, 0xFF)
FOLD = (0xC3, 0xD8, 0xEF)
TEXT_LINE = (0xA8, 0xC4, 0xE0)
ARROW = (0xF0, 0x5D, 0x4E)
ARROW_HALO = (0xFF, 0xFF, 0xFF)

BADGE_RECT = (8.0, 8.0, 248.0, 248.0)
BADGE_RADIUS = 54.0

SHEET_LEFT = 58.0
SHEET_TOP = 44.0
SHEET_RIGHT = 176.0
SHEET_BOTTOM = 212.0
SHEET_FOLD = 38.0

TEXT_LINES = (
    (82.0, 100.0, 144.0, 110.0),
    (82.0, 126.0, 126.0, 136.0),
)
TEXT_LINE_RADIUS = 5.0

ARROW_POLYGON = (
    (172.0, 122.0),
    (200.0, 122.0),
    (200.0, 170.0),
    (222.0, 170.0),
    (186.0, 222.0),
    (150.0, 170.0),
    (172.0, 170.0),
)
ARROW_HALO_SCALE = 1.16


def _scale_polygon(
    polygon: tuple[tuple[float, float], ...], scale: float
) -> tuple[tuple[float, float], ...]:
    """重心を中心に多角形を拡大し、縁取り用の形として使う。"""

    center_x = sum(point[0] for point in polygon) / len(polygon)
    center_y = sum(point[1] for point in polygon) / len(polygon)
    return tuple(
        (
            center_x + (x - center_x) * scale,
            center_y + (y - center_y) * scale,
        )
        for x, y in polygon
    )


def _sheet_polygon() -> tuple[tuple[float, float], ...]:
    """右上の角を折った書類の外形。"""

    return (
        (SHEET_LEFT, SHEET_TOP),
        (SHEET_RIGHT - SHEET_FOLD, SHEET_TOP),
        (SHEET_RIGHT, SHEET_TOP + SHEET_FOLD),
        (SHEET_RIGHT, SHEET_BOTTOM),
        (SHEET_LEFT, SHEET_BOTTOM),
    )


def _fold_polygon() -> tuple[tuple[float, float], ...]:
    """折り返した紙片。"""

    return (
        (SHEET_RIGHT - SHEET_FOLD, SHEET_TOP),
        (SHEET_RIGHT, SHEET_TOP + SHEET_FOLD),
        (SHEET_RIGHT - SHEET_FOLD, SHEET_TOP + SHEET_FOLD),
    )


def _in_rounded_rect(
    x: float,
    y: float,
    rect: tuple[float, float, float, float],
    radius: float,
) -> bool:
    left, top, right, bottom = rect
    dx = max(left + radius - x, x - (right - radius), 0.0)
    dy = max(top + radius - y, y - (bottom - radius), 0.0)
    if x < left or x > right or y < top or y > bottom:
        return False
    return dx * dx + dy * dy <= radius * radius


def _in_polygon(x: float, y: float, polygon: tuple[tuple[float, float], ...]) -> bool:
    """交差数判定による内外判定。"""

    inside = False
    count = len(polygon)
    for index in range(count):
        x0, y0 = polygon[index]
        x1, y1 = polygon[(index + 1) % count]
        if (y0 > y) != (y1 > y):
            crossing_x = x0 + (y - y0) * (x1 - x0) / (y1 - y0)
            if x < crossing_x:
                inside = not inside
    return inside


def _badge_color(y: float) -> tuple[int, int, int]:
    top = BADGE_RECT[1]
    bottom = BADGE_RECT[3]
    ratio = (y - top) / (bottom - top)
    ratio = min(max(ratio, 0.0), 1.0)
    return tuple(
        round(start + (end - start) * ratio)
        for start, end in zip(BADGE_TOP, BADGE_BOTTOM)
    )


def _sample(x: float, y: float) -> tuple[int, int, int, int] | None:
    """設計座標の1点の色を返す。どの図形にも属さない場合はNone。"""

    if not _in_rounded_rect(x, y, BADGE_RECT, BADGE_RADIUS):
        return None

    if _in_polygon(x, y, ARROW_POLYGON):
        return (*ARROW, 255)
    if _in_polygon(x, y, _scale_polygon(ARROW_POLYGON, ARROW_HALO_SCALE)):
        return (*ARROW_HALO, 255)
    for line in TEXT_LINES:
        if _in_rounded_rect(x, y, line, TEXT_LINE_RADIUS):
            return (*TEXT_LINE, 255)
    if _in_polygon(x, y, _fold_polygon()):
        return (*FOLD, 255)
    if _in_polygon(x, y, _sheet_polygon()):
        return (*SHEET, 255)
    return (*_badge_color(y), 255)


def _render_rgba(size: int) -> bytearray:
    """指定サイズのRGBAピクセル列を返す。"""

    step = DESIGN_SIZE / (size * SUPERSAMPLE)
    samples = SUPERSAMPLE * SUPERSAMPLE
    pixels = bytearray(size * size * 4)
    offset = 0
    for row in range(size):
        for column in range(size):
            red = green = blue = alpha = 0
            for sub_row in range(SUPERSAMPLE):
                y = (row * SUPERSAMPLE + sub_row + 0.5) * step
                for sub_column in range(SUPERSAMPLE):
                    x = (column * SUPERSAMPLE + sub_column + 0.5) * step
                    color = _sample(x, y)
                    if color is None:
                        continue
                    red += color[0]
                    green += color[1]
                    blue += color[2]
                    alpha += 255
            if alpha:
                # 平均は不透明だった副標本だけで取り、縁の色が暗くなるのを防ぐ。
                opaque = alpha // 255
                pixels[offset] = red // opaque
                pixels[offset + 1] = green // opaque
                pixels[offset + 2] = blue // opaque
                pixels[offset + 3] = alpha // samples
            offset += 4
    return pixels


def _png_chunk(tag: bytes, payload: bytes) -> bytes:
    return (
        struct.pack(">I", len(payload))
        + tag
        + payload
        + struct.pack(">I", zlib.crc32(tag + payload) & 0xFFFFFFFF)
    )


def _encode_png(size: int, pixels: bytearray) -> bytes:
    raw = bytearray()
    stride = size * 4
    for row in range(size):
        raw.append(0)  # フィルタータイプ: None
        raw += pixels[row * stride : (row + 1) * stride]
    header = struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)
    return (
        b"\x89PNG\r\n\x1a\n"
        + _png_chunk(b"IHDR", header)
        + _png_chunk(b"IDAT", zlib.compress(bytes(raw), 9))
        + _png_chunk(b"IEND", b"")
    )


def _write_base64(path: Path, data: bytes) -> None:
    """1行76文字で折り返したBase64テキストとして保存する。"""

    encoded = base64.b64encode(data).decode("ascii")
    lines = [encoded[index : index + 76] for index in range(0, len(encoded), 76)]
    path.write_text("\n".join(lines) + "\n", encoding="ascii")


def main() -> None:
    repository_root = Path(__file__).resolve().parent.parent
    assets_dir = repository_root / "assets"
    assets_dir.mkdir(exist_ok=True)

    for size in ICON_SIZES:
        png = _encode_png(size, _render_rgba(size))
        destination = assets_dir / f"app_icon_{size}.png.b64"
        _write_base64(destination, png)
        print(f"wrote {destination} ({size}x{size}, {len(png)} bytes of PNG)")


if __name__ == "__main__":
    main()
