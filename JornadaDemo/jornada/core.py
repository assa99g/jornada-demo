from __future__ import annotations

import hashlib
import json
import re
from io import BytesIO
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from pathlib import Path

from openpyxl import load_workbook


class ConfigError(ValueError):
    """Error de configuración legible para el usuario."""


DAYS = ("Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo")
REQUIRED = ("Empresa.xlsx", "Empleados.xlsx", "Horarios.xlsx", "Festivos.xlsx", "Incidencias.xlsx")
ABSENCES = ("Vacaciones", "Baja médica", "Ausencia", "Permiso", "Jornada reducida")


def text(value):
    return "" if value is None else str(value).strip()


def required(value, label):
    result = text(value)
    if not result:
        raise ConfigError(f"{label}: falta un valor obligatorio.")
    return result


def parse_date(value, label="Fecha", optional=False):
    if optional and value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(text(value), fmt).date()
        except ValueError:
            pass
    raise ConfigError(f"{label}: fecha no válida. Use DD/MM/AAAA.")


def integer(value, label, low=0, high=1440):
    try:
        n = float(value)
        if not n.is_integer() or not low <= n <= high:
            raise ValueError()
        return int(n)
    except (ValueError, TypeError, OverflowError):
        raise ConfigError(f"{label}: indique un entero entre {low} y {high}.") from None


def yes(value, label):
    v = text(value).upper()
    if v in ("SÍ", "SI"):
        return True
    if v == "NO":
        return False
    raise ConfigError(f"{label}: use Sí o No.")


def minutes(value, label):
    if isinstance(value, datetime):
        value = value.time()
    if isinstance(value, time):
        if value.second or value.microsecond:
            raise ConfigError(f"{label}: use horas y minutos, sin segundos.")
        return value.hour * 60 + value.minute
    if isinstance(value, (float, int)) and 0 <= value < 1:
        return round(value * 1440) % 1440
    v = text(value)
    if re.fullmatch(r"\d{1,2}:\d{2}", v):
        h, m = map(int, v.split(":"))
        if h < 24 and m < 60:
            return h * 60 + m
    raise ConfigError(f"{label}: hora no válida. Use HH:MM.")


def duration(value):
    sign = "-" if value < 0 else ""
    h, m = divmod(abs(value), 60)
    return f"{sign}{h:02d}:{m:02d}"


def clock(value):
    day, m = divmod(value, 1440)
    s = f"{m // 60:02d}:{m % 60:02d}"
    return s + (f" ({day:+d})" if day else "")


def rows(root, filename, sheet, headers):
    if filename not in root:
        raise ConfigError(f"Falta {filename}. Colóquelo junto a Iniciar.cmd.")
    try:
        wb = load_workbook(BytesIO(root[filename]), read_only=True, data_only=False)
    except Exception as exc:
        raise ConfigError(f"No se puede leer {filename}. Cierre Excel y compruebe el archivo: {exc}") from exc
    try:
        if sheet not in wb.sheetnames:
            raise ConfigError(f"{filename}: falta la hoja «{sheet}».")
        ws = wb[sheet]
        actual = [text(c.value) for c in next(ws.iter_rows(min_row=5, max_row=5))]
        if any(h not in actual for h in headers):
            raise ConfigError(f"{filename} / {sheet}: conserve los encabezados de la fila 5: {', '.join(headers)}.")
        if len([h for h in actual if h]) != len(set(h for h in actual if h)):
            raise ConfigError(f"{filename} / {sheet}: encabezados duplicados.")
        result = []
        for n, cells in enumerate(ws.iter_rows(min_row=6), 6):
            if not any(c.value is not None for c in cells):
                continue
            if any(c.data_type == "f" for c in cells):
                raise ConfigError(f"{filename} / {sheet}, fila {n}: use valores, no fórmulas.")
            row = {h: cells[actual.index(h)].value if actual.index(h) < len(cells) else None for h in headers}
            row["_label"] = f"{filename} / {sheet}, fila {n}"
            result.append(row)
        return result
    finally:
        wb.close()


@dataclass(frozen=True)
class Employee:
    id: str
    name: str
    document: str
    role: str
    start: date
    end: date | None
    schedule: str

    def active(self, day):
        return self.start <= day and (self.end is None or day <= self.end)


@dataclass(frozen=True)
class Shift:
    schedule: str
    weekday: int
    segment: int
    start: int
    end: int
    pause: int
    paid: bool
    holidays: bool


@dataclass(frozen=True)
class Incident:
    employee: str
    start: date
    end: date
    kind: str
    reduction: int
    note: str


@dataclass(frozen=True)
class Extra:
    employee: str
    day: date
    start: int
    end: int
    note: str


@dataclass
class Config:
    root: Path
    company: dict
    employees: list[Employee]
    shifts: list[Shift]
    assignments: dict
    holidays: dict
    years: set
    incidents: list[Incident]
    extras: list[Extra]
    extras_present: bool
    fingerprint: str
    sources: dict = field(default_factory=dict)


def load_config(root):
    root = Path(root).resolve()
    folder = root
    root = {}
    for name in REQUIRED + ("Extras.xlsx",):
        if (folder / name).is_file():
            try:
                root[name] = (folder / name).read_bytes()
            except OSError as exc:
                raise ConfigError(f"No se puede leer {name}. Cierre Excel y compruebe los permisos.") from exc
    company = {}
    for r in rows(root, "Empresa.xlsx", "Empresa", ("Campo", "Valor", "Descripción")):
        key = required(r["Campo"], r["_label"])
        if key in company:
            raise ConfigError(f"Empresa.xlsx: campo duplicado: {key}.")
        company[key] = required(r["Valor"], r["_label"])
    for key in ("Nombre", "CIF_NIF", "Dirección", "Municipio", "Semilla", "Variación"):
        required(company.get(key), f"Empresa.xlsx / {key}")
    if company["Variación"] not in ("Simétrica", "Antes y después"):
        raise ConfigError("Empresa.xlsx / Variación: use Simétrica o Antes y después.")
    if company["Municipio"].casefold() not in ("almería", "almeria"):
        raise ConfigError("Esta versión utiliza el calendario del municipio de Almería.")

    shifts, keys = [], set()
    for r in rows(root, "Horarios.xlsx", "Turnos", ("Horario", "Día", "Tramo", "Entrada", "Salida", "Pausa_min", "Pausa_pagada", "Trabaja_festivos")):
        label = r["_label"]
        sid = required(r["Horario"], label)
        day = integer(r["Día"], label + " / Día", 1, 7) - 1
        segment = integer(r["Tramo"], label + " / Tramo", 1, 4)
        key = (sid, day, segment)
        if key in keys:
            raise ConfigError(f"{label}: tramo duplicado para el mismo horario y día.")
        keys.add(key)
        a, b = minutes(r["Entrada"], label), minutes(r["Salida"], label)
        if a == b:
            raise ConfigError(f"{label}: entrada y salida no pueden coincidir.")
        if b < a:
            b += 1440
        pause = integer(r["Pausa_min"], label + " / Pausa_min")
        if b - a - pause <= 20:
            raise ConfigError(f"{label}: el tramo debe dejar más de 20 minutos de trabajo tras la pausa.")
        shifts.append(Shift(sid, day, segment, a, b, pause, yes(r["Pausa_pagada"], label), yes(r["Trabaja_festivos"], label)))
    schedules = {s.schedule for s in shifts}
    if not schedules:
        raise ConfigError("Horarios.xlsx: añada al menos un horario.")
    for sid in schedules:
        # Weekly intervals include a second copy so Sunday-to-Monday conflicts are caught.
        week = sorted((s.weekday * 1440 + s.start, s.weekday * 1440 + s.end) for s in shifts if s.schedule == sid)
        doubled = week + [(a + 10080, b + 10080) for a, b in week]
        for (_, end), (start, _) in zip(doubled, doubled[1:]):
            if end > start:
                raise ConfigError(f"Horarios.xlsx: hay tramos solapados en {sid}.")

    employees, ids = [], set()
    for r in rows(root, "Empleados.xlsx", "Empleados", ("ID", "Nombre", "DNI_NIE", "Puesto", "Alta", "Baja", "Horario")):
        label = r["_label"]
        eid = required(r["ID"], label)
        if not re.fullmatch(r"[A-Za-z0-9_-]{1,32}", eid) or eid.casefold() in ids:
            raise ConfigError(f"{label}: ID duplicado o no válido. Use letras, números, guion o guion bajo.")
        ids.add(eid.casefold())
        a, b = parse_date(r["Alta"], label), parse_date(r["Baja"], label, True)
        sid = required(r["Horario"], label)
        if sid not in schedules or (b and b < a):
            raise ConfigError(f"{label}: horario desconocido o baja anterior al alta.")
        employees.append(Employee(eid, required(r["Nombre"], label), required(r["DNI_NIE"], label), required(r["Puesto"], label), a, b, sid))
    if not employees:
        raise ConfigError("Empleados.xlsx: añada al menos un empleado.")
    empmap = {e.id: e for e in employees}

    assignments = {}
    for r in rows(root, "Horarios.xlsx", "Asignaciones", ("ID", "Mes", "Horario")):
        eid, sid = text(r["ID"]), text(r["Horario"])
        month = text(r["Mes"])
        if isinstance(r["Mes"], (date, datetime)):
            if r["Mes"].day != 1:
                raise ConfigError(f"{r['_label']}: el mes debe comenzar el día 1.")
            month = r["Mes"].strftime("%Y-%m")
        if not re.fullmatch(r"\d{4}-\d{2}", month):
            raise ConfigError(f"{r['_label']}: Mes debe ser AAAA-MM. No se permiten cambios a mitad de mes.")
        parse_date(month + "-01", r["_label"])
        if eid not in empmap or sid not in schedules or (eid, month) in assignments:
            raise ConfigError(f"{r['_label']}: empleado/horario desconocido o asignación mensual duplicada.")
        assignments[eid, month] = sid

    years = set()
    for r in rows(root, "Festivos.xlsx", "Cobertura", ("Año", "Municipio", "Revisado")):
        year = integer(r["Año"], r["_label"], 2000, 2100)
        if text(r["Municipio"]).casefold() not in ("almería", "almeria"):
            raise ConfigError(f"{r['_label']}: municipio distinto de Almería.")
        if yes(r["Revisado"], r["_label"]):
            years.add(year)
    holidays = {}
    for r in rows(root, "Festivos.xlsx", "Festivos", ("Fecha", "Nombre", "Ámbito", "Fuente")):
        day = parse_date(r["Fecha"], r["_label"])
        if day in holidays:
            raise ConfigError(f"{r['_label']}: fecha festiva duplicada.")
        holidays[day] = required(r["Nombre"], r["_label"])

    incidents = []
    for r in rows(root, "Incidencias.xlsx", "Incidencias", ("ID", "Desde", "Hasta", "Tipo", "Reducción_min", "Observaciones")):
        eid, kind = text(r["ID"]), text(r["Tipo"])
        a, b = parse_date(r["Desde"], r["_label"]), parse_date(r["Hasta"], r["_label"])
        reduction = integer(r["Reducción_min"] or 0, r["_label"])
        if eid not in empmap or kind not in ABSENCES or b < a:
            raise ConfigError(f"{r['_label']}: empleado, tipo de incidencia o intervalo no válido.")
        if (kind == "Jornada reducida") != (reduction > 0):
            raise ConfigError(f"{r['_label']}: indique Reducción_min solo para Jornada reducida (mayor que cero).")
        if any(i.employee == eid and max(a, i.start) <= min(b, i.end) for i in incidents):
            raise ConfigError(f"{r['_label']}: incidencias solapadas para {eid}.")
        incidents.append(Incident(eid, a, b, kind, reduction, text(r["Observaciones"])))

    extras, present = [], "Extras.xlsx" in root
    if present:
        for r in rows(root, "Extras.xlsx", "Extras", ("ID", "Fecha", "Inicio", "Fin", "Motivo")):
            eid, day = text(r["ID"]), parse_date(r["Fecha"], r["_label"])
            a, b = minutes(r["Inicio"], r["_label"]), minutes(r["Fin"], r["_label"])
            if a == b or eid not in empmap or not empmap[eid].active(day):
                raise ConfigError(f"{r['_label']}: empleado fuera de alta o intervalo de horas extra no válido.")
            if b < a:
                b += 1440
            if any(i.employee == eid and i.start <= day <= i.end and i.kind != "Jornada reducida" for i in incidents):
                raise ConfigError(f"{r['_label']}: horas extra durante una ausencia de día completo.")
            extras.append(Extra(eid, day, a, b, text(r["Motivo"])))
        for eid in empmap:
            ranges = sorted((x.day.toordinal() * 1440 + x.start, x.day.toordinal() * 1440 + x.end) for x in extras if x.employee == eid)
            if any(a[1] > b[0] for a, b in zip(ranges, ranges[1:])):
                raise ConfigError(f"Extras.xlsx: horas extra solapadas para {eid}.")

    digest = hashlib.sha256()
    for f in REQUIRED + (("Extras.xlsx",) if present else ()):
        digest.update(f.encode())
        digest.update(root[f])
    return Config(folder, company, employees, shifts, assignments, holidays, years, incidents, extras, present, digest.hexdigest()[:16], root)


@dataclass
class Record:
    employee: Employee
    day: date
    status: str
    planned: int = 0
    ordinary: int = 0
    extra: int = 0
    pause: int = 0
    unpaid: int = 0
    spans: list = field(default_factory=list)
    extra_spans: list = field(default_factory=list)
    notes: str = ""

    @property
    def total(self):
        return self.ordinary + self.extra

    @property
    def deviation(self):
        return self.ordinary - self.planned


def offset(seed, eid, day, segment, endpoint, mode):
    raw = f"jornada-v1|{seed}|{eid}|{day.isoformat()}|{segment}|{endpoint}".encode()
    n = int.from_bytes(hashlib.sha256(raw).digest()[:8], "big")
    if mode == "Simétrica":
        return n % 21 - 10
    return -(n % 11) if endpoint == "entrada" else n % 11


def planned_shifts(config, employee, day):
    if not employee.active(day):
        return []
    inc = next((i for i in config.incidents if i.employee == employee.id and i.start <= day <= i.end), None)
    if inc and inc.kind != "Jornada reducida":
        return []
    sid = config.assignments.get((employee.id, day.strftime("%Y-%m")), employee.schedule)
    shifts = sorted((s for s in config.shifts if s.schedule == sid and s.weekday == day.weekday() and (day not in config.holidays or s.holidays)), key=lambda s: s.start)
    spans = [(s, s.start, s.end) for s in shifts]
    if inc and spans:
        s, a, b = spans[-1]
        b -= inc.reduction
        if b - a - s.pause <= 20:
            raise ConfigError(f"{employee.name}, {day:%d/%m/%Y}: la reducción no deja un último tramo válido. Reduzca Reducción_min.")
        spans[-1] = (s, a, b)
    return spans


def generate(config, start, end, employee_id=None):
    start, end = parse_date(start), parse_date(end)
    if end < start:
        raise ConfigError("La fecha final es anterior a la inicial.")
    if (end - start).days > 365:
        raise ConfigError("Seleccione como máximo 366 días por informe.")
    missing = set(range(start.year, end.year + 1)) - config.years
    if missing:
        raise ConfigError(f"Falta revisar el calendario de {', '.join(map(str, sorted(missing)))} en Festivos.xlsx / Cobertura.")
    employees = [e for e in config.employees if employee_id is None or e.id == employee_id]
    if not employees:
        raise ConfigError("No se encuentra el empleado seleccionado.")
    records = []
    for e in employees:
        # Generate surrounding days too: overlap validation must not depend on the report range.
        for day in (start + timedelta(days=n) for n in range((end - start).days + 1)):
            if not e.active(day):
                continue
            inc = next((i for i in config.incidents if i.employee == e.id and i.start <= day <= i.end), None)
            shifts = planned_shifts(config, e, day)
            holiday = config.holidays.get(day)
            status = inc.kind if inc else ("Festivo trabajado" if shifts and holiday else "Festivo" if holiday else "Trabajo" if shifts else "Descanso")
            r = Record(e, day, status)
            notes = [v for v in (holiday, inc.note if inc else None) if v]
            nearby = []
            for delta in (-1, 0, 1):
                neighbor = day + timedelta(days=delta)
                for s, a, b in planned_shifts(config, e, neighbor):
                    nearby.append((delta, s.segment, a + delta * 1440, b + delta * 1440))
            extra_near = [(x, (x.day - day).days * 1440 + x.start, (x.day - day).days * 1440 + x.end) for x in config.extras if x.employee == e.id and abs((x.day - day).days) <= 1]
            for s, a, b in shifts:
                sa = a + offset(config.company["Semilla"], e.id, day, s.segment, "entrada", config.company["Variación"])
                sb = b + offset(config.company["Semilla"], e.id, day, s.segment, "salida", config.company["Variación"])
                # Deterministic midpoint bounds avoid overlaps with adjacent simulated shifts.
                for delta, segment, na, nb in nearby:
                    if delta == 0 and segment == s.segment:
                        continue
                    if max(a, na) < min(b, nb):
                        raise ConfigError(f"{e.name}, {day:%d/%m/%Y}: los horarios de meses consecutivos se solapan.")
                    if nb <= a:
                        sa = max(sa, (a + nb) // 2)
                    elif na >= b:
                        sb = min(sb, (b + na) // 2)
                for x, xa, xb in extra_near:
                    if max(a, xa) < min(b, xb):
                        raise ConfigError(f"Extras.xlsx: {e.name}, {x.day:%d/%m/%Y}, horas extra solapadas con el horario ordinario.")
                    if xb <= a:
                        sa = max(sa, xb)
                    elif xa >= b:
                        sb = min(sb, xa)
                unpaid = 0 if s.paid else s.pause
                if sb - sa <= s.pause:
                    raise ConfigError(f"{e.name}, {day:%d/%m/%Y}: pausa mayor que el tramo simulado.")
                r.spans.append((sa, sb))
                r.pause += s.pause
                r.unpaid += unpaid
                r.planned += b - a - unpaid
                r.ordinary += sb - sa - unpaid
            for x, _, _ in extra_near:
                if x.day == day:
                    # Also validate extras against the next/previous day's ordinary schedule.
                    for _, _, a, b in nearby:
                        if max(x.start, a) < min(x.end, b):
                            raise ConfigError(f"Extras.xlsx: {e.name}, {day:%d/%m/%Y}, horas extra solapadas con el horario ordinario.")
                    r.extra_spans.append((x.start, x.end))
                    r.extra += x.end - x.start
                    if x.note:
                        notes.append(x.note)
            if r.extra and not shifts:
                r.status += " + extras"
            r.notes = "; ".join(notes)
            records.append(r)
    if not records:
        raise ConfigError("No hay empleados de alta en el período seleccionado.")
    return sorted(records, key=lambda r: (r.day, r.employee.name, r.employee.id))


def totals(records):
    return {k: sum(getattr(r, k) for r in records) for k in ("planned", "ordinary", "extra", "total", "pause", "unpaid", "deviation")}
