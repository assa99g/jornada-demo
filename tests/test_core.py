import copy
import json
import shutil
import tempfile
import unittest
from dataclasses import replace
from datetime import date, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile
from io import BytesIO

from jornada.core import *
from jornada.pdf import export_reports
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[1] / 'JornadaDemo'


def change_cell(path, cell, value, sheet=1):
    """Change a disposable test fixture, preserving production workbooks."""
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    src = BytesIO(path.read_bytes())
    out = BytesIO()
    with ZipFile(src) as original, ZipFile(out, 'w') as target:
        for item in original.infolist():
            data = original.read(item.filename)
            if item.filename == f'xl/worksheets/sheet{sheet}.xml':
                tree = ET.fromstring(data)
                node = tree.find(f'.//{{{ns}}}c[@r="{cell}"]')
                node.clear()
                node.set('r',cell)
                node.set('t','inlineStr')
                ET.SubElement(ET.SubElement(node,f'{{{ns}}}is'),f'{{{ns}}}t').text = value
                data = ET.tostring(tree)
            target.writestr(item, data)
    path.write_bytes(out.getvalue())


class Calculations(unittest.TestCase):
    def setUp(self):
        self.c = load_config(ROOT)

    def test_month_planned_and_totals(self):
        rs = generate(self.c,'2026-09-01','2026-09-30')
        self.assertEqual(len(rs),90)
        # September has 22 weekdays. Ana: five leave days. Lucía: 1 h reduction.
        self.assertEqual(totals(rs)['planned'],17*450 + 22*480 + 22*240-60)
        self.assertEqual(totals(rs)['extra'],60)
        for r in rs:
            self.assertEqual(r.ordinary,sum(b-a for a,b in r.spans)-r.unpaid)
            self.assertEqual(r.total,r.ordinary+r.extra)
            self.assertEqual(r.extra,sum(b-a for a,b in r.extra_spans))

    def test_day_month_and_employee_consistent(self):
        month = generate(self.c,'2026-09-01','2026-09-30')
        for day in (date(2026,9,8),date(2026,9,10),date(2026,9,15)):
            self.assertEqual(generate(self.c,day,day),[r for r in month if r.day==day])
        self.assertEqual(generate(self.c,'2026-09-01','2026-09-30','EMP002'),[r for r in month if r.employee.id=='EMP002'])

    def test_bounds_both_modes_entire_year(self):
        for mode in ('Simétrica','Antes y después'):
            self.c.company['Variación']=mode
            for r in generate(self.c,'2026-01-01','2026-12-31'):
                for (_,a,b),(sa,sb) in zip(planned_shifts(self.c,r.employee,r.day),r.spans):
                    self.assertTrue(a-10<=sa<=a+(10 if mode=='Simétrica' else 0))
                    self.assertTrue(b-(10 if mode=='Simétrica' else 0)<=sb<=b+10)

    def test_holidays_do_not_generate_work(self):
        for day in self.c.holidays:
            rs=generate(self.c,day,day)
            self.assertTrue(all(r.ordinary==0 and r.status=='Festivo' for r in rs))
        self.assertEqual(len(self.c.holidays),14)
        self.assertIn(date(2026,6,24),self.c.holidays)
        self.assertIn(date(2026,8,29),self.c.holidays)

    def test_work_on_holiday_is_configurable(self):
        self.c.shifts=[replace(s,holidays=True) for s in self.c.shifts]
        r=generate(self.c,'2026-06-24','2026-06-24','EMP001')[0]
        self.assertGreater(r.ordinary,0)
        self.assertEqual(r.status,'Festivo trabajado')

    def test_paid_pause_changes_only_net_time(self):
        a=generate(self.c,'2026-09-01','2026-09-01','EMP001')[0]
        self.c.shifts=[replace(s,paid=True) for s in self.c.shifts]
        b=generate(self.c,'2026-09-01','2026-09-01','EMP001')[0]
        self.assertEqual(a.spans,b.spans)
        self.assertEqual(b.ordinary-a.ordinary,30)
        self.assertEqual(b.planned-a.planned,30)

    def test_absence_and_reduction(self):
        leave=generate(self.c,'2026-09-15','2026-09-15','EMP001')[0]
        reduced=generate(self.c,'2026-09-10','2026-09-10','EMP003')[0]
        self.assertEqual((leave.status,leave.total),('Vacaciones',0))
        self.assertEqual(reduced.planned,180)
        self.assertTrue(170<=reduced.spans[0][1]-reduced.spans[0][0]<=200)

    def test_employment_dates_inclusive(self):
        self.c.employees=[replace(self.c.employees[0],start=date(2026,9,8),end=date(2026,9,10))]
        self.assertEqual([r.day.day for r in generate(self.c,'2026-09-01','2026-09-30')],[8,9,10])

    def test_monthly_assignment(self):
        self.c.assignments['EMP001','2026-09']='PARCIAL'
        r=generate(self.c,'2026-09-01','2026-09-01','EMP001')[0]
        self.assertEqual(r.planned,240)
        r=generate(self.c,'2026-08-31','2026-08-31','EMP001')[0]
        self.assertEqual(r.planned,450)

    def test_overnight_and_touching_shifts(self):
        self.c.shifts=[Shift('NOCHE',i,1,22*60,30*60,0,False,True) for i in range(7)]
        self.c.employees=[replace(self.c.employees[0],schedule='NOCHE')]
        self.c.incidents=[]
        self.c.extras=[]
        r=generate(self.c,'2026-09-08','2026-09-08')[0]
        self.assertEqual(r.planned,480)
        self.assertIn('(+1)',clock(r.spans[0][1]))
        self.c.shifts=[Shift('NOCHE',1,1,480,720,0,False,True),Shift('NOCHE',1,2,720,960,0,False,True)]
        r=generate(self.c,'2026-09-08','2026-09-08')[0]
        self.assertLessEqual(r.spans[0][1],r.spans[1][0])

    def test_extra_overlap_rejected_and_touching_allowed(self):
        self.c.extras=[Extra('EMP001',date(2026,9,8),950,1000,'')]
        with self.assertRaisesRegex(ConfigError,'solapadas'):
            generate(self.c,'2026-09-08','2026-09-08')
        self.c.extras=[Extra('EMP001',date(2026,9,8),960,1020,'')]
        r=generate(self.c,'2026-09-08','2026-09-08','EMP001')[0]
        self.assertEqual(r.extra,60)
        self.assertLessEqual(r.spans[-1][1],960)

    def test_night_extras_cannot_overlap_next_day(self):
        self.c.extras=[Extra('EMP001',date(2026,9,7),23*60,33*60,'')]
        with self.assertRaisesRegex(ConfigError,'solapadas'):
            generate(self.c,'2026-09-07','2026-09-07','EMP001')

    def test_invalid_period_and_unknown_calendar(self):
        for a,b in [('2026-09-10','2026-09-01'),('2027-01-01','2027-01-01'),('2026-01-01','2027-12-31')]:
            with self.assertRaises(ConfigError):
                generate(self.c,a,b)

    def test_parse_dates_and_times(self):
        self.assertEqual(parse_date('08/09/2026'),date(2026,9,8))
        self.assertEqual(minutes('08:05','hora'),485)
        for value in ('24:00','12:60','hola'):
            with self.assertRaises(ConfigError): minutes(value,'hora')


class FileIntegration(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.root=Path(self.temp.name)
        for p in ROOT.glob('*.xlsx'):
            shutil.copy2(p,self.root/p.name)

    def tearDown(self):
        self.temp.cleanup()

    def test_optional_extras_absent(self):
        (self.root/'Extras.xlsx').unlink()
        c=load_config(self.root)
        self.assertFalse(c.extras_present)
        self.assertEqual(totals(generate(c,'2026-09-01','2026-09-30'))['extra'],0)

    def test_duplicate_employee_rejected(self):
        change_cell(self.root/'Empleados.xlsx','A7','EMP001')
        with self.assertRaisesRegex(ConfigError,'duplicado'):
            load_config(self.root)

    def test_unknown_schedule_rejected(self):
        change_cell(self.root/'Empleados.xlsx','G6','DESCONOCIDO')
        with self.assertRaisesRegex(ConfigError,'horario desconocido'):
            load_config(self.root)

    def test_missing_file_rejected(self):
        (self.root/'Empresa.xlsx').unlink()
        with self.assertRaisesRegex(ConfigError,'Falta Empresa.xlsx'):
            load_config(self.root)

    def test_invalid_hours_rejected(self):
        change_cell(self.root/'Extras.xlsx','C6','veinte')
        with self.assertRaisesRegex(ConfigError,'hora no válida'):
            load_config(self.root)

    def test_extra_during_leave_rejected(self):
        change_cell(self.root/'Extras.xlsx','A6','EMP001')
        change_cell(self.root/'Extras.xlsx','B6','15/09/2026')
        with self.assertRaisesRegex(ConfigError,'ausencia'):
            load_config(self.root)

    def test_pdf_snapshot_and_reconciliation(self):
        c=load_config(self.root)
        a,b=date(2026,9,1),date(2026,9,30)
        rs=generate(c,a,b)
        folder,files=export_reports(c,rs,a,b)
        self.assertEqual(len(files),4)
        manifest=json.loads((folder/'Resumen.json').read_text())
        self.assertEqual(manifest['totales_minutos'],totals(rs))
        self.assertEqual(load_config(folder/'Configuracion').fingerprint,c.fingerprint)
        for f in files:
            reader=PdfReader(f)
            for page in reader.pages:
                content=page.extract_text()
                self.assertIn('DEMOSTRACIÓN',content)
                self.assertIn('Datos simulados',content)
            if f.name=='Resumen_general.pdf':
                self.assertIn(duration(totals(rs)['total']),''.join(p.extract_text() for p in reader.pages))
        # Re-export never replaces an existing folder.
        folder2,_=export_reports(c,rs,a,b,'General')
        self.assertNotEqual(folder,folder2)


if __name__=='__main__':
    unittest.main()
