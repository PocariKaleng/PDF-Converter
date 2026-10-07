import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import threading

from .docx import prepare_docx
from .errors import ConversionError
from .fonts import require_fonts, registered_fonts, ROOT
from .pdf import inspect_pdf, embed_supported_fonts

_lock = threading.Lock()
SCRIPT = ROOT / "scripts" / "word-to-pdf.ps1"


def word_available() -> bool:
    if sys.platform != "win32":
        return False
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"):
            return True
    except OSError:
        return False


def convert(data: bytes, font: str = "original") -> tuple[bytes, dict]:
    if not word_available():
        raise ConversionError("Microsoft Word desktop belum terpasang. Converter memerlukan Windows dan Word desktop aktif.", 503)
    if not _lock.acquire(blocking=False):
        raise ConversionError("Ada dokumen yang sedang dikonversi. Tunggu hingga selesai lalu coba lagi.", 409)
    try:
        runtime = ROOT / ".runtime"
        runtime.mkdir(exist_ok=True)
        with tempfile.TemporaryDirectory(prefix="job-", dir=runtime) as directory:
            work = Path(directory)
            source, target = work / "document.docx", work / "document.pdf"
            expected = prepare_docx(data, source, font)
            require_fonts(expected)
            with registered_fonts(expected):
                command = ["powershell.exe", "-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(SCRIPT), "-InputPath", str(source), "-OutputPath", str(target), "-PidPath", str(work / "word-pid.json")]
                try:
                    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=subprocess.CREATE_NO_WINDOW)
                    stdout, stderr = process.communicate(timeout=120)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.communicate()
                    # The worker records only a newly created Word process.
                    pid_path = work / "word-pid.json"
                    if pid_path.exists():
                        pid = json.loads(pid_path.read_text(encoding="utf-8-sig"))
                        if isinstance(pid, int) and pid > 0:
                            subprocess.run(["taskkill.exe", "/PID", str(pid), "/T", "/F"], capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
                    raise ConversionError("Konversi melewati batas 120 detik. Pastikan Word aktif dan tidak menampilkan dialog aktivasi, lalu coba lagi.", 504)
                except OSError as exc:
                    raise ConversionError("PowerShell atau Microsoft Word tidak dapat dijalankan.", 503) from exc
                if process.returncode or not target.exists():
                    detail = (stderr or stdout).decode("utf-8", errors="replace").strip()
                    raise ConversionError("Word gagal mengonversi dokumen. Pastikan Word telah diaktivasi dan file dapat dibuka. " + detail[:400], 500)
                embed_supported_fonts(target)
                report = inspect_pdf(target, expected)
                return target.read_bytes(), report
    finally:
        _lock.release()
