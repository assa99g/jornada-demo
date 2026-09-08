"""Build the offline Windows folder from already downloaded official artifacts."""
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

ROOT=Path(__file__).resolve().parents[1]
APP=ROOT/'JornadaDemo'
OUT=ROOT/'JornadaDemo-Windows-x64.zip'

for name in ('Iniciar.cmd','iniciar.py','runtime/python.exe','runtime/pythonw.exe',
             'runtime/python313.dll','runtime/python313.zip','runtime/python313._pth',
             'runtime/Lib/site-packages/PySide6/QtWidgets.pyd',
             'runtime/Lib/site-packages/PySide6/plugins/platforms/qwindows.dll',
             'Empresa.xlsx','Empleados.xlsx','Horarios.xlsx','Festivos.xlsx','Incidencias.xlsx'):
    assert (APP/name).is_file(),name
pth=(APP/'runtime/python313._pth').read_text()
assert 'Lib/site-packages' in pth and 'import site' in pth

manifest={}
with ZipFile(OUT,'w',ZIP_DEFLATED,compresslevel=6) as z:
    for f in sorted(APP.rglob('*')):
        if not f.is_file() or '__pycache__' in f.parts or f.suffix in ('.pyc','.ndjson') or f.name in ('Error.log','.DS_Store'):
            continue
        relative=f.relative_to(APP)
        z.write(f,Path('JornadaDemo')/relative)
        manifest[str(relative)]=hashlib.sha256(f.read_bytes()).hexdigest()
    z.writestr('JornadaDemo/CONTENIDO-SHA256.json',json.dumps(manifest,indent=2))
with ZipFile(OUT) as z:
    assert z.testzip() is None
sha=hashlib.sha256(OUT.read_bytes()).hexdigest()
(ROOT/'JornadaDemo-Windows-x64.sha256').write_text(f'{sha}  {OUT.name}\n')
print(json.dumps({'archivo':str(OUT),'MB':round(OUT.stat().st_size/1024**2,1),'archivos':len(manifest),'sha256':sha},ensure_ascii=False))
