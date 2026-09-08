"""Punto de entrada portátil: no depende del directorio de trabajo."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

if __name__ == "__main__":
    try:
        from jornada.ui import run
        raise SystemExit(run(ROOT))
    except Exception:
        import traceback
        error = traceback.format_exc()
        try:
            (ROOT / "Error.log").write_text(error, encoding="utf-8")
        except OSError:
            pass
        if sys.platform == "win32":
            import ctypes
            ctypes.windll.user32.MessageBoxW(None, "No se ha podido iniciar Jornada. Extraiga toda la carpeta del ZIP y consulte Error.log.", "Jornada", 0x10)
        else:
            print(error, file=sys.stderr)
        raise SystemExit(1)
