from __future__ import annotations

import calendar
import json
import shutil
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape

import reportlab
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, LongTable, Table, TableStyle, PageBreak, KeepTogether

from .core import DAYS, REQUIRED, duration, clock, totals

MONTHS = ("", "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre")
NAVY = colors.HexColor("#24374D")
MUTED = colors.HexColor("#66758A")
LIGHT = colors.HexColor("#EFF3F7")


def fonts():
    if "Jornada" not in pdfmetrics.getRegisteredFontNames():
        folder = Path(reportlab.__file__).parent / "fonts"
        pdfmetrics.registerFont(TTFont("Jornada", str(folder / "Vera.ttf")))
        pdfmetrics.registerFont(TTFont("JornadaBold", str(folder / "VeraBd.ttf")))
        pdfmetrics.registerFontFamily("Jornada", normal="Jornada", bold="JornadaBold", italic="Jornada", boldItalic="JornadaBold")


def p(value, size=8, bold=False, color=NAVY):
    return Paragraph(escape(str(value)).replace("\n", "<br/>"), ParagraphStyle("cell", fontName="JornadaBold" if bold else "Jornada", fontSize=size, leading=size+3, textColor=color, keepWithNext=bool(bold and size >= 9)))


def grid(headers, rows, widths):
    data = [[p(h, 7, True, colors.white) for h in headers]]
    data += [[p(v, 7) for v in row] for row in rows]
    table = LongTable(data, colWidths=widths, repeatRows=1, hAlign="LEFT")
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("TOPPADDING", (0, 1), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 2.5),
        ("LINEBELOW", (0, -1), (-1, -1), .5, colors.HexColor("#CDD6E0")),
    ]))
    return table


def footer(config, label):
    def draw(c, doc):
        w, h = doc.pagesize
        c.saveState()
        c.setFillColor(MUTED)
        c.setFont("Jornada", 7)
        c.drawString(36, h-24, "JORNADA / " + label)
        c.setStrokeColor(colors.HexColor("#CBD5DF"))
        c.line(36, 42, w-36, 42)
        c.setFillColor(colors.HexColor("#9A5924"))
        c.drawString(36, 29, "DEMOSTRACIÓN · Datos simulados · No acredita jornadas efectivamente trabajadas")
        c.setFillColor(MUTED)
        c.drawRightString(w-36, 29, f"{doc.page}")
        c.setFont("Jornada", 6)
        c.drawString(36, 18, f"Configuración {config.fingerprint} · Horas en HH:MM · Hora civil local de Almería")
        c.restoreState()
    return draw


def heading(config, title, period):
    return [p(title, 19, True), Spacer(1, 8), p(config.company["Nombre"], 10, True),
            p(f"CIF/NIF: {config.company['CIF_NIF']} · {config.company['Dirección']}", 8),
            p(period, 9), Spacer(1, 14)]


def total_block(records, width):
    t = totals(records)
    return grid(["Planificadas", "Ordinarias simuladas", "Extra", "Total simulado", "Desviación ordinaria"],
                [[duration(t[k]) for k in ("planned", "ordinary", "extra", "total", "deviation")]], [width/5]*5)


def note_ranges(records):
    groups=[]
    for r in records:
        if not r.notes:
            continue
        if groups and groups[-1][2] == r.notes and (r.day-groups[-1][1]).days == 1:
            groups[-1][1]=r.day
        else:
            groups.append([r.day,r.day,r.notes])
    return [f"{a:%d/%m}" + (f"–{b:%d/%m}" if a!=b else "") + f": {note}" for a,b,note in groups]


def individual_pdf(path, config, employee, records, start, end):
    fonts()
    width = A4[0]-72
    story = []
    months = sorted({(r.day.year, r.day.month) for r in records})
    for index, (year, month) in enumerate(months):
        if index:
            story.append(PageBreak())
        subset = [r for r in records if (r.day.year, r.day.month) == (year, month)]
        story += heading(config, "Registro individual de jornada", f"{MONTHS[month].capitalize()} {year} · Período solicitado: {start:%d/%m/%Y} – {end:%d/%m/%Y}")
        story += [p(employee.name, 12, True), p(f"{employee.id} · DNI/NIE: {employee.document} · {employee.role}", 8), Spacer(1, 10)]
        rows = []
        for r in subset:
            spans = "\n".join(f"{clock(a)} – {clock(b)}" for a, b in r.spans) or "—"
            if r.extra_spans:
                spans += "\nExtra: " + "; ".join(f"{clock(a)} – {clock(b)}" for a,b in r.extra_spans)
            rows.append([f"{r.day:%d/%m} {DAYS[r.day.weekday()][:3]}", spans, duration(r.pause), duration(r.ordinary), duration(r.extra), duration(r.total), r.status])
        story.append(grid(["Fecha", "Entrada – salida", "Pausa", "Ordin.", "Extra", "Total", "Situación"], rows, [58,170,42,46,42,46,width-404]))
        story += [Spacer(1, 8), total_block(subset, width), Spacer(1, 6)]
        story.append(p(f"Pausas: {duration(sum(r.pause for r in subset))}; no retribuidas descontadas: {duration(sum(r.unpaid for r in subset))}. La pausa entre tramos no se computa como trabajo.", 7, color=MUTED))
        story.append(p("Las desviaciones de las marcas simuladas no se clasifican como horas extra. (+1) indica el día siguiente; (-1), el anterior.", 7, color=MUTED))
        notes = note_ranges(subset)
        if notes:
            story += [Spacer(1, 5), p("Observaciones", 9, True)] + [p(n, 7) for n in notes]
        story += [KeepTogether([Spacer(1, 16), p("Firma de la persona trabajadora: ____________________     Firma de la empresa: ____________________", 7), p("Espacios ilustrativos sin firmas incorporadas.", 6, color=MUTED)])]
    doc = SimpleDocTemplate(str(path), pagesize=A4, leftMargin=36, rightMargin=36, topMargin=43, bottomMargin=55, title="Registro individual de jornada · Demostración", author=config.company["Nombre"])
    doc.build(story, onFirstPage=footer(config, employee.id), onLaterPages=footer(config, employee.id))


def general_pdf(path, config, records, start, end):
    fonts()
    size = landscape(A4)
    width = size[0]-72
    story = heading(config, "Resumen general de jornada", f"{start:%d/%m/%Y} – {end:%d/%m/%Y} · Empleados incluidos: {len({r.employee.id for r in records})}")
    employee_rows = []
    for eid in sorted({r.employee.id for r in records}):
        subset = [r for r in records if r.employee.id == eid]
        t = totals(subset)
        employee_rows.append([subset[0].employee.name, eid, sum(r.total > 0 for r in subset)] + [duration(t[k]) for k in ("planned", "ordinary", "extra", "total", "deviation")])
    story += [grid(["Empleado", "ID", "Días con horas", "Planificadas", "Ordinarias", "Extra", "Total", "Desviación"], employee_rows, [width-515,75,65,75,80,65,75,80]), Spacer(1, 13), total_block(records, width), Spacer(1, 10), p("Las horas ordinarias son simuladas y descuentan las pausas no retribuidas. Las horas extra proceden exclusivamente de Extras.xlsx, si existe. Las ausencias completas no generan horas ordinarias.", 8), Spacer(1, 16), p("Detalle diario de la empresa", 12, True), Spacer(1, 8)]
    daily = []
    for day in sorted({r.day for r in records}):
        subset = [r for r in records if r.day == day]
        t = totals(subset)
        daily.append([f"{day:%d/%m/%Y} · {DAYS[day.weekday()]}", sum(r.total > 0 for r in subset)] + [duration(t[k]) for k in ("planned", "ordinary", "extra", "total")] + [config.holidays.get(day, "")])
    story += [grid(["Fecha", "Personas con horas", "Planificadas", "Ordinarias", "Extra", "Total", "Festivo"], daily, [157,80,82,82,62,72,width-535])]
    doc = SimpleDocTemplate(str(path), pagesize=size, leftMargin=36, rightMargin=36, topMargin=43, bottomMargin=55, title="Resumen general de jornada · Demostración", author=config.company["Nombre"])
    doc.build(story, onFirstPage=footer(config, "RESUMEN GENERAL"), onLaterPages=footer(config, "RESUMEN GENERAL"))


def export_reports(config, records, start, end, mode="Ambos", output_root=None):
    output_root = Path(output_root) if output_root else config.root / "Informes"
    output_root.mkdir(parents=True, exist_ok=True)
    folder = output_root / f"{start:%Y-%m-%d}_{end:%Y-%m-%d}_{datetime.now():%Y%m%d_%H%M%S_%f}"
    staging = folder.with_name(".generando_" + folder.name)
    staging.mkdir()
    files = []
    try:
        if mode in ("Ambos", "General"):
            general_pdf(staging / "Resumen_general.pdf", config, records, start, end)
            files.append("Resumen_general.pdf")
        if mode in ("Ambos", "Individuales"):
            for eid in sorted({r.employee.id for r in records}):
                subset = [r for r in records if r.employee.id == eid]
                name = f"Registro_{eid}.pdf"
                individual_pdf(staging / name, config, subset[0].employee, subset, start, end)
                files.append(name)
        if not files:
            raise ValueError("Tipo de informe desconocido.")
        snapshot = staging / "Configuracion"
        snapshot.mkdir()
        for filename in REQUIRED + (("Extras.xlsx",) if config.extras_present else ()):
            (snapshot / filename).write_bytes(config.sources[filename])
        manifest = {"version": "1.0.0", "tipo": "DEMOSTRACIÓN", "desde": start.isoformat(), "hasta": end.isoformat(), "configuracion": config.fingerprint, "semilla": config.company["Semilla"], "archivos": files, "totales_minutos": totals(records)}
        (staging / "Resumen.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        staging.rename(folder)
    except Exception:
        shutil.rmtree(staging, ignore_errors=True)
        raise
    return folder, [folder / name for name in files]
