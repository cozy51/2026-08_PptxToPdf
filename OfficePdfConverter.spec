# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller configuration for the Windows one-file distribution."""

import base64
import struct
from pathlib import Path

from PyInstaller.utils.hooks import collect_all


project_root = Path(SPECPATH).resolve()
build_directory = project_root / "build"
build_directory.mkdir(parents=True, exist_ok=True)


def create_windows_icon(destination: Path) -> None:
    """Create an ICO containing the bundled PNG sizes for Explorer/the taskbar."""

    images = []
    for size in (16, 32, 48, 256):
        encoded = project_root / "assets" / f"app_icon_{size}.png.b64"
        images.append((size, base64.b64decode(encoded.read_text(encoding="ascii"))))

    header_size = 6 + 16 * len(images)
    entries = []
    payload = bytearray()
    for size, image in images:
        # ICO represents 256 pixels with a zero width/height byte.
        dimension = 0 if size == 256 else size
        entries.append(
            struct.pack(
                "<BBBBHHII",
                dimension,
                dimension,
                0,
                0,
                1,
                32,
                len(image),
                header_size + len(payload),
            )
        )
        payload.extend(image)

    destination.write_bytes(
        struct.pack("<HHH", 0, 1, len(images)) + b"".join(entries) + payload
    )


icon_path = build_directory / "app_icon.ico"
create_windows_icon(icon_path)

# tkinterdnd2 ships the native tkdnd DLL and Tcl scripts as package data.  Using
# collect_all keeps those platform-specific files beside the frozen Tcl runtime.
tkdnd_datas, tkdnd_binaries, tkdnd_hiddenimports = collect_all("tkinterdnd2")

a = Analysis(
    [str(project_root / "main.py")],
    pathex=[str(project_root)],
    binaries=tkdnd_binaries,
    datas=tkdnd_datas + [(str(project_root / "assets"), "assets")],
    hiddenimports=tkdnd_hiddenimports
    + ["pythoncom", "pywintypes", "win32com", "win32com.client"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="OfficePdfConverter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(icon_path),
)
