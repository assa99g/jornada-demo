"""Exercise the actual Qt controls and background jobs using the offscreen platform."""
import shutil
import tempfile
import time
from pathlib import Path
from unittest.mock import patch

from PySide6.QtCore import QDate, QLocale
from PySide6.QtWidgets import QApplication
from jornada.ui import MainWindow
from test_core import change_cell

root=Path(__file__).resolve().parents[1]
(root/'.build/qa').mkdir(parents=True,exist_ok=True)
QLocale.setDefault(QLocale(QLocale.Spanish,QLocale.Spain))
app=QApplication([])
app.setStyle('Fusion')


def wait(window):
    deadline=time.monotonic()+15
    while window.busy and time.monotonic()<deadline:
        app.processEvents()
        time.sleep(.01)
    app.processEvents()
    assert not window.busy,'El trabajo de fondo no terminó'


with tempfile.TemporaryDirectory() as tmp:
    folder=Path(tmp)
    for p in (root/'JornadaDemo').glob('*.xlsx'):
        shutil.copy2(p,folder/p.name)
    w=MainWindow(folder,autoload=False)
    w.start.setDate(QDate(2026,9,1))
    w.end.setDate(QDate(2026,9,30))
    w.show()
    w.preview_button.click()
    wait(w)
    assert w.preview_table.rowCount()==90
    assert w.export_button.isEnabled()
    w.grab().save(str(root/'.build/qa/interfaz.png'))
    for i,name in ((1,'empleados'),(2,'configuracion'),(3,'ayuda')):
        w.nav[i].click()
        app.processEvents()
        assert w.pages.currentIndex()==i
        w.grab().save(str(root/f'.build/qa/interfaz-{name}.png'))
    w.nav[0].click()
    w.employee.setCurrentIndex(w.employee.findData('EMP002'))
    assert not w.export_button.isEnabled()
    w.start.setDate(QDate(2026,9,8))
    w.single_day()
    w.preview_button.click()
    wait(w)
    assert len(w.records)==1 and w.records[0].extra==60
    w.export_button.click()
    wait(w)
    assert len(list(w.result_folder.glob('*.pdf')))==2
    change_cell(folder/'Empresa.xlsx','B6','Empresa modificada')
    with patch('jornada.ui.QMessageBox.warning') as alert:
        w.export_button.click()
        wait(w)
        assert alert.called and 'han cambiado' in alert.call_args.args[2]
    assert not w.export_button.isEnabled()
    (folder/'Extras.xlsx').unlink()
    w.preview_button.click()
    wait(w)
    assert w.records[0].extra==0
    assert 'Sin archivo' in w.message.text()
    w.close()
print('Qt: navegación, filtros, vista previa, exportación, cambio de Excel y ausencia de Extras.xlsx comprobados.')
