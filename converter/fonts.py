from contextlib import contextmanager
import ctypes
from functools import lru_cache
import os
import re
from pathlib import Path
import sys

from fontTools.ttLib import TTFont

from .errors import ConversionError

ROOT = Path(__file__).resolve().parent.parent
FAMILIES = ("Bell MT", "LM Roman 10")
LM_VERSION = "1.106"


def family_key(name: str) -> str | None:
    name = "".join(c for c in name.lower() if c.isalnum())
    if name.startswith("bellmt"):
        return "Bell MT"
    if name.startswith(("lmroman10", "latinmodernroman10")):
        return "LM Roman 10"
    return None


def font_directories() -> list[Path]:
    directories = [ROOT / "fonts" / "latin-modern" / "v1.106", ROOT / "fonts" / "custom"]
    if sys.platform == "win32":
        directories += [
            Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts",
            Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft/Windows/Fonts",
        ]
    return directories


@lru_cache(maxsize=1)
def inventory() -> dict:
    result = {family: {"name": family, "variants": {}, "available": False} for family in FAMILIES}
    for directory in font_directories():
        if not directory.is_dir():
            continue
        for path in directory.iterdir():
            # Avoid parsing every system font, especially large CJK collections.
            if path.suffix.lower() not in {".ttf", ".otf"}:
                continue
            if directory.name == "Fonts" and not path.name.lower().startswith(("bell", "lmroman10")):
                continue
            try:
                with TTFont(path, lazy=True) as font:
                    names = font["name"]
                    family = next((key for name_id in (1, 16, 6) if (key := family_key(names.getDebugName(name_id) or ""))), None)
                    if not family:
                        continue
                    style = (names.getDebugName(17) or names.getDebugName(2) or "Regular").lower()
                    variant = "bolditalic" if "bold" in style and ("italic" in style or "oblique" in style) else (
                        "bold" if "bold" in style else "italic" if "italic" in style or "oblique" in style else "regular"
                    )
                    fs_type = font["OS/2"].fsType if "OS/2" in font else 0
                    embeddable = not bool(fs_type & (0x0002 | 0x0200))
                    version_match = re.search(r"Version\s+([0-9.]+)", names.getDebugName(5) or "")
                    version = version_match.group(1) if version_match else None
                    if family == "LM Roman 10" and version != LM_VERSION:
                        continue
                    result[family]["variants"].setdefault(variant, {
                        "path": str(path.resolve()), "embeddable": embeddable,
                        "postscript": names.getDebugName(6),
                        "version": version,
                    })
            except (OSError, ValueError, KeyError):
                continue
    for family in result.values():
        regular = family["variants"].get("regular")
        family["available"] = bool(regular and regular["embeddable"])
    return result


def public_inventory() -> list[dict]:
    return [{"name": family["name"], "available": family["available"],
             "version": family["variants"].get("regular", {}).get("version"),
             "variants": [v for v in ('regular', 'bold', 'italic', 'bolditalic') if v in family["variants"]],
             "embeddable": family["available"]} for family in inventory().values()]


def require_fonts(families: set[str]) -> None:
    for name in families:
        if not inventory()[name]["available"]:
            raise ConversionError(
                f"Font {name} regular belum tersedia atau tidak mengizinkan embedding. "
                "Pasang font berlisensi di Windows atau letakkan di fonts/custom, lalu restart aplikasi."
            )


@contextmanager
def registered_fonts(families: set[str]):
    """Register project fonts for this Windows session, then remove our references."""
    if sys.platform != "win32":
        raise ConversionError("Converter ini memerlukan Windows dan Microsoft Word desktop.", 503)
    gdi = ctypes.WinDLL("gdi32", use_last_error=True)
    gdi.AddFontResourceExW.argtypes = [ctypes.c_wchar_p, ctypes.c_uint, ctypes.c_void_p]
    gdi.RemoveFontResourceExW.argtypes = [ctypes.c_wchar_p, ctypes.c_uint, ctypes.c_void_p]
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    user32.SendMessageTimeoutW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_ssize_t, ctypes.c_uint, ctypes.c_uint, ctypes.POINTER(ctypes.c_size_t)]

    def notify_font_change():
        result = ctypes.c_size_t()
        user32.SendMessageTimeoutW(0xFFFF, 0x001D, 0, 0, 2, 1000, ctypes.byref(result))
    added = []
    try:
        for name in families:
            for variant in inventory()[name]["variants"].values():
                path = Path(variant["path"])
                if path.is_relative_to(ROOT / "fonts"):
                    # Flags=0 makes the font visible to the child Word process.
                    # No registry entry or permanent font installation is made.
                    if not gdi.AddFontResourceExW(str(path), 0, None):
                        raise ConversionError(f"Windows gagal memuat font {name}.", 503)
                    added.append(str(path))
        if added:
            notify_font_change()
        yield
    finally:
        for path in reversed(added):
            gdi.RemoveFontResourceExW(path, 0, None)
        if added:
            notify_font_change()
