from __future__ import annotations

import traceback
from datetime import date
from pathlib import Path

from PySide6.QtCore import Qt, QDate, QLocale, QThread, Signal, QUrl
from PySide6.QtGui import QColor, QDesktopServices, QFont
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFrame, QComboBox, QDateEdit, QTableWidget, QTableWidgetItem,
    QHeaderView, QAbstractItemView, QStackedWidget, QMessageBox, QListWidget, QTextBrowser)

from .core import load_config, generate, totals, duration, clock, ConfigError, REQUIRED, DAYS
from .pdf import export_reports


STYLE = """
QWidget { font-family: 'Segoe UI', 'Arial'; font-size: 13px; color: #24374d; }
QMainWindow, QWidget#main { background: #f4f6f9; }
QFrame#sidebar { background: #203449; }
QFrame#sidebar QLabel { color: #dce5ee; background: transparent; }
QFrame#sidebar QPushButton { text-align: left; border: 0; color: #dce5ee; background: transparent; padding: 13px 17px; border-radius: 7px; }
QFrame#sidebar QPushButton:checked { background: #34516d; color: white; font-weight: bold; }
QFrame#sidebar QPushButton:hover { background: #2b455e; }
QLabel#brand { color: white; font-size: 26px; font-weight: bold; }
QLabel#title { font-size: 26px; font-weight: bold; }
QLabel#subtitle { color: #6a798b; }
QLabel#eyebrow { color: #8497aa; font-size: 10px; font-weight: bold; }
QLabel#demo { color: #8f5d28; background: #fbefd9; padding: 7px 12px; border-radius: 5px; font-weight: bold; font-size: 11px; }
QFrame#panel { background: white; border: 1px solid #e1e6ed; border-radius: 9px; }
QLabel#metric { font-size: 25px; font-weight: bold; }
QLabel#metricLabel { color: #6a798b; font-size: 12px; }
QPushButton { background: white; border: 1px solid #d4dde7; border-radius: 6px; padding: 9px 15px; font-weight: bold; }
QPushButton:hover { background: #edf2f7; }
QPushButton#primary { background: #2463a5; color: white; border: 1px solid #2463a5; }
QPushButton#primary:hover { background: #1b528b; }
QPushButton:disabled { background: #e7ecf1; color: #91a0af; border-color: #e1e6ed; }
QDateEdit, QComboBox { background: white; border: 1px solid #cfd8e2; border-radius: 5px; padding: 7px; min-height: 20px; }
QComboBox QAbstractItemView { background: white; color: #24374d; selection-background-color: #dce9f6; }
QTableWidget { background: white; alternate-background-color: #f7f9fb; border: 0; gridline-color: #edf0f4; selection-background-color: #dfebf8; selection-color: #183956; }
QHeaderView::section { background: #edf2f7; color: #596c80; border: 0; border-bottom: 1px solid #dbe3ec; padding: 10px 7px; font-size: 11px; font-weight: bold; }
QListWidget, QTextBrowser { background: white; border: 1px solid #e0e6ed; border-radius: 6px; padding: 9px; }
QStatusBar { color: #617489; background: #edf2f7; }
QLabel#message { color: #556c83; padding: 7px 0; }
"""


def label(text, name=None):
    w = QLabel(text)
    w.setTextFormat(Qt.PlainText)
    if name:
        w.setObjectName(name)
    return w


def button(text, callback, primary=False):
    b = QPushButton(text)
    if primary:
        b.setObjectName("primary")
    b.clicked.connect(callback)
    return b


def panel():
    frame = QFrame()
    frame.setObjectName("panel")
    return frame


def table(headers):
    t = QTableWidget(0, len(headers))
    t.setHorizontalHeaderLabels(headers)
    t.setEditTriggers(QAbstractItemView.NoEditTriggers)
    t.setSelectionBehavior(QAbstractItemView.SelectRows)
    t.setAlternatingRowColors(True)
    t.setShowGrid(False)
    t.verticalHeader().hide()
    t.verticalHeader().setDefaultSectionSize(34)
    t.horizontalHeader().setStretchLastSection(True)
    return t


def fill_table(widget, values):
    widget.setRowCount(len(values))
    for i, row in enumerate(values):
        for j, value in enumerate(row):
            item = QTableWidgetItem(str(value))
            item.setToolTip(str(value))
            widget.setItem(i, j, item)


class Worker(QThread):
    result = Signal(object)
    error = Signal(str)

    def __init__(self, fn, root):
        super().__init__()
        self.fn, self.root = fn, root

    def run(self):
        try:
            self.result.emit(self.fn())
        except (ConfigError, PermissionError) as exc:
            self.error.emit(str(exc))
        except Exception as exc:
            try:
                (self.root / "Error.log").write_text(traceback.format_exc(), encoding="utf-8")
            except OSError:
                pass
            self.error.emit(f"No se ha completado la operación: {exc}\nConsulte Error.log.")


class MainWindow(QMainWindow):
    def __init__(self, root, autoload=True):
        super().__init__()
        self.root = Path(root).resolve()
        self.config = None
        self.records = []
        self.worker = None
        self.result_folder = None
        self.busy = False
        self.setWindowTitle("Jornada · Simulador de registro horario")
        self.resize(1280, 850)
        self.setMinimumSize(1080, 710)
        self.setStyleSheet(STYLE)
        container = QWidget()
        container.setObjectName("main")
        self.setCentralWidget(container)
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(214)
        side = QVBoxLayout(sidebar)
        side.setContentsMargins(22, 30, 22, 24)
        side.addWidget(label("Jornada", "brand"))
        side.addWidget(label("ALMERÍA · ESPAÑA", "eyebrow"))
        side.addSpacing(38)
        self.nav = []
        self.pages = QStackedWidget()
        for i, name in enumerate(("Informes", "Empleados", "Configuración", "Ayuda")):
            b = button(name, lambda checked=False, index=i: self.navigate(index))
            b.setCheckable(True)
            side.addWidget(b)
            self.nav.append(b)
        side.addStretch()
        side.addWidget(label("FUNCIONAMIENTO LOCAL", "eyebrow"))
        note = label("Sin conexión a internet\nArchivos en esta carpeta")
        note.setWordWrap(True)
        side.addWidget(note)
        side.addSpacing(16)
        side.addWidget(label("Versión 1.0 · Demostración", "eyebrow"))
        layout.addWidget(sidebar)
        layout.addWidget(self.pages, 1)
        self.create_reports_page()
        self.create_employees_page()
        self.create_config_page()
        self.create_help_page()
        self.navigate(0)
        self.statusBar().showMessage("Preparado para leer los archivos Excel.")
        if autoload:
            self.preview()

    def page(self, title, subtitle):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(28, 25, 28, 22)
        layout.setSpacing(15)
        top = QHBoxLayout()
        textcol = QVBoxLayout()
        textcol.addWidget(label(title, "title"))
        textcol.addWidget(label(subtitle, "subtitle"))
        top.addLayout(textcol, 1)
        top.addWidget(label("DATOS SIMULADOS", "demo"), 0, Qt.AlignTop)
        layout.addLayout(top)
        self.pages.addWidget(page)
        return layout

    def create_reports_page(self):
        layout = self.page("Registro de jornada", "Genera y revisa las jornadas de tu plantilla antes de crear los PDF.")
        self.company_label = label("Cargando empresa…", "subtitle")
        layout.addWidget(self.company_label)
        filters = panel()
        form = QHBoxLayout(filters)
        form.setContentsMargins(17, 16, 17, 16)
        today = QDate.currentDate()
        self.start = QDateEdit(QDate(today.year(), today.month(), 1))
        self.end = QDateEdit(QDate(today.year(), today.month(), today.daysInMonth()))
        self.employee = QComboBox()
        self.employee.addItem("Todos los empleados", None)
        self.employee.setMinimumWidth(220)
        for title, widget in (("Desde", self.start), ("Hasta", self.end), ("Empleado", self.employee)):
            col = QVBoxLayout()
            col.addWidget(label(title, "metricLabel"))
            col.addWidget(widget)
            form.addLayout(col, 2 if widget is self.employee else 1)
        for w in (self.start, self.end):
            w.setCalendarPopup(True)
            w.setDisplayFormat("dd/MM/yyyy")
            w.setMinimumDate(QDate(2000, 1, 1))
            w.setMaximumDate(QDate(2100, 12, 31))
            w.dateChanged.connect(self.invalidate)
        self.employee.currentIndexChanged.connect(self.invalidate)
        layout.addWidget(filters)
        actions = QHBoxLayout()
        actions.addWidget(button("Un solo día", self.single_day))
        actions.addWidget(button("Mes completo", self.whole_month))
        actions.addStretch()
        self.preview_button = button("Actualizar vista previa", self.preview)
        actions.addWidget(self.preview_button)
        layout.addLayout(actions)
        metrics = QHBoxLayout()
        self.metrics = {}
        for title, key in (("Empleados", "employees"), ("Ordinarias simuladas", "ordinary"), ("Horas extra", "extra"), ("Total de horas", "total")):
            box = panel()
            content = QVBoxLayout(box)
            content.setContentsMargins(17, 13, 17, 13)
            content.addWidget(label(title, "metricLabel"))
            value = label("—", "metric")
            self.metrics[key] = value
            content.addWidget(value)
            metrics.addWidget(box)
        layout.addLayout(metrics)
        self.preview_table = table(["FECHA", "EMPLEADO", "ENTRADA / SALIDA", "PAUSA", "ORDIN.", "EXTRA", "TOTAL", "SITUACIÓN"])
        for i, width in enumerate((88,169,168,58,64,58,64,132)):
            self.preview_table.setColumnWidth(i,width)
        layout.addWidget(self.preview_table, 1)
        self.message = label("", "message")
        self.message.setWordWrap(True)
        layout.addWidget(self.message)
        bottom = QHBoxLayout()
        self.mode = QComboBox()
        self.mode.addItems(["Individuales y general", "Solo individuales", "Solo general"])
        bottom.addWidget(self.mode)
        bottom.addStretch()
        self.open_button = button("Abrir carpeta de informes", self.open_reports)
        bottom.addWidget(self.open_button)
        self.export_button = button("Generar PDF", self.export, True)
        self.export_button.setEnabled(False)
        bottom.addWidget(self.export_button)
        layout.addLayout(bottom)

    def create_employees_page(self):
        layout = self.page("Empleados", "Plantilla cargada desde Empleados.xlsx. Las fechas de alta y baja delimitan los registros.")
        actions = QHBoxLayout()
        actions.addWidget(button("Editar Empleados.xlsx", lambda: self.open_path(self.root / "Empleados.xlsx")))
        actions.addWidget(button("Recargar archivos", self.preview))
        actions.addStretch()
        layout.addLayout(actions)
        self.employees_table = table(["ID", "NOMBRE", "DNI / NIE", "PUESTO", "ALTA", "BAJA", "HORARIO"])
        for i,w in enumerate((90,200,155,155,100,100,120)):
            self.employees_table.setColumnWidth(i,w)
        layout.addWidget(self.employees_table,1)

    def create_config_page(self):
        layout = self.page("Configuración", "Edita los archivos Excel junto al programa y pulsa Recargar archivos.")
        descriptions = {
            "Empresa.xlsx":"Razón social, CIF/NIF, dirección y variación de las marcas.",
            "Empleados.xlsx":"Identificación, puesto, alta, baja y horario por defecto.",
            "Horarios.xlsx":"Turnos semanales, pausas y asignaciones por mes completo.",
            "Festivos.xlsx":"Fiestas laborales y años revisados para Almería.",
            "Incidencias.xlsx":"Vacaciones, bajas, ausencias, permisos y reducciones.",
            "Extras.xlsx":"Opcional. Intervalos exactos de horas extraordinarias."}
        self.file_states = {}
        for filename, description in descriptions.items():
            box = panel()
            row = QHBoxLayout(box)
            row.setContentsMargins(16, 11, 16, 11)
            col = QVBoxLayout()
            col.addWidget(label(filename))
            col.addWidget(label(description, "subtitle"))
            row.addLayout(col,1)
            state = label("", "subtitle")
            self.file_states[filename]=state
            row.addWidget(state)
            row.addWidget(button("Abrir", lambda checked=False,f=filename:self.open_path(self.root/f)))
            layout.addWidget(box)
        layout.addWidget(button("Recargar archivos", self.preview))
        layout.addStretch()

    def create_help_page(self):
        layout = self.page("Ayuda", "Guía de uso de Jornada · Demostración")
        help = QTextBrowser()
        help.setOpenExternalLinks(False)
        help.setHtml("""<h2>De Excel a PDF en tres pasos</h2>
        <p>1. Edita y guarda los archivos Excel junto a <b>Iniciar.cmd</b>.</p>
        <p>2. Selecciona fechas y empleado; pulsa <b>Actualizar vista previa</b>.</p>
        <p>3. Elige informes individuales, general o ambos y pulsa <b>Generar PDF</b>.</p>
        <h3>Marcas de demostración</h3><p>Los horarios generan entradas y salidas con una variación de hasta diez minutos.
        La semilla de Empresa.xlsx permite repetir las mismas marcas para el mismo empleado y día.
        Los PDF indican que sus datos son simulados.</p>
        <h3>Pausas, ausencias y horas extra</h3><p>Solo se descuentan las pausas marcadas como no pagadas.
        Las ausencias completas no generan trabajo ordinario. Reducción_min acorta el último tramo.
        Extras.xlsx es opcional: si no existe, no se añaden extras. Las desviaciones ordinarias se muestran aparte.</p>
        <h3>Calendario y turnos</h3><p>Se incluye Almería 2026. Para otro año, completa Festivos.xlsx y marca su cobertura como revisada.
        En Horarios.xlsx, Día va de 1 (lunes) a 7 (domingo). Una salida menor que la entrada termina al día siguiente.
        Las horas se calculan según el reloj civil local; no se ajusta la duración por el cambio estacional de hora.</p>
        <h3>Archivos y conservación</h3><p>Cada exportación crea una carpeta nueva en Informes con PDF y una copia de la configuración.
        Los encabezados están en la fila 5; los datos empiezan en la 6. No uses fórmulas.
        Consulta LEEME.txt para los detalles.</p>""")
        layout.addWidget(help,1)
        layout.addWidget(button("Abrir guía completa",lambda:self.open_path(self.root/"LEEME.txt")))

    def navigate(self,index):
        self.pages.setCurrentIndex(index)
        for n,b in enumerate(self.nav):
            b.setChecked(n==index)

    def single_day(self):
        if self.busy:
            return
        self.end.setDate(self.start.date())

    def whole_month(self):
        if self.busy:
            return
        d=self.start.date()
        self.start.setDate(QDate(d.year(),d.month(),1))
        self.end.setDate(QDate(d.year(),d.month(),d.daysInMonth()))

    def invalidate(self,*args):
        self.records=[]
        self.export_button.setEnabled(False)
        self.preview_table.setRowCount(0)
        for metric in self.metrics.values():
            metric.setText("—")
        self.message.setText("Fechas o selección modificadas. Actualiza la vista previa.")

    def set_busy(self,busy):
        self.busy=busy
        for w in (self.preview_button,self.start,self.end,self.employee,self.mode):
            w.setEnabled(not busy)
        self.export_button.setEnabled(not busy and bool(self.records))

    def run_job(self,fn,callback,message):
        if self.busy:
            return
        self.set_busy(True)
        self.statusBar().showMessage(message)
        self.worker=Worker(fn,self.root)
        self.worker.result.connect(callback)
        self.worker.error.connect(self.on_error)
        self.worker.finished.connect(lambda:self.set_busy(False))
        self.worker.start()

    def on_error(self,message):
        self.invalidate()
        self.message.setText(message)
        self.statusBar().showMessage("No se ha completado la operación.")
        QMessageBox.warning(self,"Revisa los datos",message)

    def preview(self):
        if self.busy:
            return
        start,end=self.start.date().toPython(),self.end.date().toPython()
        eid=self.employee.currentData()
        self.invalidate()
        def job():
            config=load_config(self.root)
            return config,generate(config,start,end,eid)
        self.run_job(job,self.show_preview,"Leyendo configuración y generando vista previa…")

    def show_preview(self,result):
        self.config,self.records=result
        self.company_label.setText(f"{self.config.company['Nombre']} · {self.config.company['Municipio']}")
        selected=self.employee.currentData()
        self.employee.blockSignals(True)
        self.employee.clear()
        self.employee.addItem("Todos los empleados",None)
        for e in self.config.employees:
            self.employee.addItem(e.name,e.id)
        self.employee.setCurrentIndex(max(0,self.employee.findData(selected)))
        self.employee.blockSignals(False)
        rows=[]
        for r in self.records:
            spans=" / ".join(f"{clock(a)}–{clock(b)}" for a,b in r.spans) or "—"
            if r.extra_spans:
                spans += " | Extra: " + " / ".join(f"{clock(a)}–{clock(b)}" for a,b in r.extra_spans)
            rows.append([r.day.strftime('%d/%m/%Y'),r.employee.name,spans,duration(r.pause),duration(r.ordinary),duration(r.extra),duration(r.total),r.status])
        fill_table(self.preview_table,rows)
        self.preview_table.resizeRowsToContents()
        t=totals(self.records)
        for key,w in self.metrics.items():
            w.setText(str(len({r.employee.id for r in self.records})) if key=='employees' else duration(t[key]))
        fill_table(self.employees_table,[[e.id,e.name,e.document,e.role,e.start.strftime('%d/%m/%Y'),e.end.strftime('%d/%m/%Y') if e.end else '—',e.schedule] for e in self.config.employees])
        for filename,w in self.file_states.items():
            w.setText("Disponible" if (self.root/filename).is_file() else "No incluido")
        extra = "Extras.xlsx incluido" if self.config.extras_present else "Sin archivo de horas extra"
        self.message.setText(f"{len(self.records)} registros · {extra} · Desviación ordinaria: {duration(t['deviation'])}. Horas en HH:MM.")
        self.export_button.setEnabled(True)
        self.statusBar().showMessage("Vista previa actualizada. Datos simulados.")

    def export(self):
        if not self.records or self.busy:
            return
        start,end=self.start.date().toPython(),self.end.date().toPython()
        mode=("Ambos","Individuales","General")[self.mode.currentIndex()]
        config,records=self.config,list(self.records)
        def job():
            fresh=load_config(self.root)
            if fresh.fingerprint != config.fingerprint:
                raise ConfigError("Los archivos Excel han cambiado. Actualiza la vista previa antes de exportar.")
            return export_reports(config,records,start,end,mode)
        self.run_job(job,self.export_done,"Creando los informes PDF…")

    def export_done(self,result):
        self.result_folder,files=result
        self.message.setText(f"Se han guardado {len(files)} PDF. Pulsa Abrir carpeta de informes para verlos.")
        self.statusBar().showMessage(f"Informes guardados en {self.result_folder.name}")

    def open_path(self,path):
        if not path.exists():
            QMessageBox.information(self,"Archivo no disponible",f"No existe {path.name}. Consulta LEEME.txt para preparar el archivo.")
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(path))):
            QMessageBox.information(self,"Abrir archivo",f"Abre manualmente: {path}")

    def open_reports(self):
        folder=self.result_folder or self.root/"Informes"
        folder.mkdir(exist_ok=True)
        self.open_path(folder)

    def closeEvent(self,event):
        if self.busy:
            self.statusBar().showMessage("Espera a que termine la operación antes de cerrar.")
            event.ignore()
        else:
            event.accept()


def run(root):
    QLocale.setDefault(QLocale(QLocale.Spanish,QLocale.Spain))
    app=QApplication.instance() or QApplication([])
    app.setApplicationName("Jornada Demo")
    app.setStyle("Fusion")
    window=MainWindow(root)
    window.show()
    return app.exec()
