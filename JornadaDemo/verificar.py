"""Comprobación local del paquete, sin conexiones de red."""
from pathlib import Path
import tempfile
import traceback

ROOT=Path(__file__).resolve().parent

try:
    from PySide6.QtCore import qVersion
    from PySide6.QtWidgets import QApplication
    from jornada.core import load_config, generate
    from jornada.pdf import export_reports
    from datetime import date, timedelta
    print('Qt:',qVersion())
    app=QApplication([])
    config=load_config(ROOT)
    if not config.years:
        raise ValueError('No hay años revisados en Festivos.xlsx.')
    year=max(config.years)
    first=date(year,1,1)
    records=generate(config,first,date(year,12,31))
    day=next((r.day for r in records if r.total),records[0].day)
    with tempfile.TemporaryDirectory(prefix='jornada_verificacion_') as tmp:
        _,files=export_reports(config,generate(config,day,day),day,day,output_root=tmp)
        assert all(f.read_bytes().startswith(b'%PDF-') for f in files)
    print('CORRECTO: interfaz, Excel, cálculo y PDF disponibles.')
    print('No se han modificado los archivos de configuración.')
except Exception:
    traceback.print_exc()
    raise SystemExit(1)
