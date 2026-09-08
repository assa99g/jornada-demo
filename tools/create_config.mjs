import fs from 'node:fs/promises';
import path from 'node:path';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const root = path.resolve('JornadaDemo');
const qa = path.resolve('.build/qa');
await fs.mkdir(qa, {recursive:true});
const regional = 'https://www.juntadeandalucia.es/boja/2025/93/1.html';
const local = 'https://www.juntadeandalucia.es/boja/2025/197/28';
const date = s => new Date(s+'T12:00:00Z');
const time = s => {const [h,m] = s.split(':').map(Number); return (h*60+m)/1440;};
function col(n) {let s=''; for(n++; n; n=Math.floor((n-1)/26)) s=String.fromCharCode(65+(n-1)%26)+s; return s;}

async function make(filename, sheets) {
  const wb = Workbook.create();
  for (const spec of sheets) {
    const sh = wb.worksheets.add(spec.name);
    const end = col(spec.headers.length-1);
    sh.showGridLines = false;
    sh.getRange(`A1:${end}${Math.max(6,spec.rows.length+5)}`).format.font.name='Arial';
    sh.getRange(`A1:${end}${Math.max(6,spec.rows.length+5)}`).format.font.size=11;
    sh.getRange('A1').values=[[spec.title]];
    sh.getRange('A1').format.font.size=17;
    sh.getRange('A1').format.font.bold=true;
    sh.getRange(`A1:${end}1`).format.rowHeight=29;
    sh.getRange('A2').values=[[spec.note]];
    sh.getRange('A2').format.font.color='#64748B';
    sh.getRange(`A2:${end}2`).format.rowHeight=25;
    sh.getRange(`A5:${end}5`).values=[spec.headers];
    sh.getRange(`A5:${end}5`).format={fill:'#283D55',font:{color:'#FFFFFF',bold:true},rowHeight:34,wrapText:true};
    if(spec.rows.length) sh.getRange(`A6:${end}${5+spec.rows.length}`).values=spec.rows;
    sh.getRange(`A6:${end}${Math.max(6,spec.rows.length+5)}`).format.rowHeight=28;
    sh.getRange(`A6:${end}${Math.max(6,spec.rows.length+5)}`).format.font.color='#164B83';
    sh.getRange(`A6:${end}${Math.max(6,spec.rows.length+5)}`).format.fill='#F2F6FA';
    spec.widths.forEach((w,i)=>sh.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);
    for(const i of spec.dates||[]) {sh.getRange(`${col(i)}6:${col(i)}1000`).setNumberFormat('dd/mm/yyyy'); sh.getRange(`${col(i)}6:${col(i)}1000`).format.horizontalAlignment='center';}
    for(const i of spec.times||[]) {sh.getRange(`${col(i)}6:${col(i)}1000`).setNumberFormat('hh:mm'); sh.getRange(`${col(i)}6:${col(i)}1000`).format.horizontalAlignment='center';}
    spec.headers.forEach((h,i)=>{if(['Año','Pausa_min','Reducción_min'].includes(h)) sh.getRange(`${col(i)}6:${col(i)}1000`).format.horizontalAlignment='center';});
    for(const [i,values] of spec.lists||[]) sh.getRange(`${col(i)}6:${col(i)}1000`).dataValidation={rule:{type:'list',values}};
    sh.freezePanes.freezeRows(5);
    const preview = await wb.render({sheetName:spec.name,range:`A1:${end}${Math.min(13,Math.max(7,spec.rows.length+5))}`,scale:1.4});
    await fs.writeFile(path.join(qa,`${filename}-${spec.name}.png`),new Uint8Array(await preview.arrayBuffer()));
    console.log((await wb.inspect({kind:'table',range:`'${spec.name}'!A5:${end}8`,include:'values',tableMaxRows:4,tableMaxCols:8,maxChars:900})).ndjson);
  }
  console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!',options:{useRegex:true,maxResults:5},summary:'Comprobación de errores'})).ndjson);
  await (await SpreadsheetFile.exportXlsx(wb)).save(path.join(root,filename+'.xlsx'));
}

await make('Empresa',[{name:'Empresa',title:'Datos de la empresa',note:'Edite la columna Valor. Datos ficticios para demostración.',headers:['Campo','Valor','Descripción'],widths:[24,44,89],rows:[
 ['Nombre','Alborán Servicios Demo, S.L.','Razón social de la empresa'],
 ['CIF_NIF','DEMO-B00000000','Identificador ficticio. Sustituya por el dato de la demostración.'],
 ['Dirección','Calle de Ejemplo, 12 · 04001 Almería','Domicilio de la empresa'],
 ['Municipio','Almería','Municipio del calendario de fiestas locales'],
 ['Semilla','ALMERIA-DEMO-2026','Mantenga este valor para repetir las mismas marcas por empleado y día.'],
 ['Variación','Simétrica','Simétrica: ±10 min. Antes y después: entrada −10 a 0; salida 0 a +10 min.'],
]}]);

await make('Empleados',[{name:'Empleados',title:'Plantilla de empleados',note:'Alta y baja incluidas. Baja vacía significa que sigue de alta.',headers:['ID','Nombre','DNI_NIE','Puesto','Alta','Baja','Horario'],widths:[14,30,23,25,17,17,19],dates:[4,5],rows:[
 ['EMP001','Ana García López','DEMO-00000001','Administración',date('2026-01-01'),null,'CONTINUA'],
 ['EMP002','Carlos Martín Ruiz','DEMO-00000002','Atención al cliente',date('2026-01-01'),null,'PARTIDA'],
 ['EMP003','Lucía Fernández Díaz','DEMO-00000003','Auxiliar',date('2026-01-01'),null,'PARCIAL'],
]}]);

const turnos=[];
for(let d=1;d<=5;d++) {
 turnos.push(['CONTINUA',d,1,time('08:00'),time('16:00'),30,'No','No']);
 turnos.push(['PARTIDA',d,1,time('09:00'),time('13:00'),0,'No','No']);
 turnos.push(['PARTIDA',d,2,time('16:00'),time('20:00'),0,'No','No']);
 turnos.push(['PARCIAL',d,1,time('09:00'),time('13:00'),0,'No','No']);
}
await make('Horarios',[
 {name:'Turnos',title:'Horarios semanales',note:'Día: 1 lunes … 7 domingo. Salida menor que entrada: día siguiente.',headers:['Horario','Día','Tramo','Entrada','Salida','Pausa_min','Pausa_pagada','Trabaja_festivos'],widths:[20,10,11,15,15,17,20,22],times:[3,4],lists:[[6,['Sí','No']],[7,['Sí','No']]],rows:turnos},
 {name:'Asignaciones',title:'Asignación mensual de horarios',note:'Opcional. Mes: AAAA-MM. Sin fila se utiliza el horario de Empleados.xlsx.',headers:['ID','Mes','Horario'],widths:[23,25,72],rows:[]}
]);

const holidays=[
 ['2026-01-01','Año Nuevo','Nacional'],['2026-01-06','Epifanía del Señor','Nacional'],
 ['2026-02-28','Día de Andalucía','Autonómico'],['2026-04-02','Jueves Santo','Autonómico'],
 ['2026-04-03','Viernes Santo','Nacional'],['2026-05-01','Fiesta del Trabajo','Nacional'],
 ['2026-06-24','Fiesta local de Almería (junio)','Local'],['2026-08-15','Asunción de la Virgen','Nacional'],
 ['2026-08-29','Fiesta local de Almería (agosto)','Local'],['2026-10-12','Fiesta Nacional de España','Nacional'],
 ['2026-11-02','Todos los Santos (traslado)','Autonómico'],['2026-12-07','Constitución Española (traslado)','Autonómico'],
 ['2026-12-08','Inmaculada Concepción','Nacional'],['2026-12-25','Natividad del Señor','Nacional']
].map(([d,n,a])=>[date(d),n,a,a==='Local'?local:regional]);
await make('Festivos',[
 {name:'Festivos',title:'Calendario laboral de Almería · 2026',note:'Fiestas nacionales, autonómicas y locales. Fuentes: BOJA, enlaces por fecha.',headers:['Fecha','Nombre','Ámbito','Fuente'],widths:[18,43,20,79],dates:[0],rows:holidays},
 {name:'Cobertura',title:'Años del calendario',note:'Marque Sí únicamente cuando haya completado los festivos de ese año.',headers:['Año','Municipio','Revisado'],widths:[22,44,48],lists:[[2,['Sí','No']]],rows:[[2026,'Almería','Sí']]}
]);

await make('Incidencias',[{name:'Incidencias',title:'Ausencias y jornadas reducidas',note:'Ejemplos ficticios. Reducción_min acorta el último tramo del día.',headers:['ID','Desde','Hasta','Tipo','Reducción_min','Observaciones'],widths:[15,18,18,24,20,52],dates:[1,2],lists:[[3,['Vacaciones','Baja médica','Ausencia','Permiso','Jornada reducida']]],rows:[
 ['EMP001',date('2026-09-14'),date('2026-09-18'),'Vacaciones',0,'Vacaciones de ejemplo'],
 ['EMP003',date('2026-09-10'),date('2026-09-10'),'Jornada reducida',60,'Reducción de una hora de ejemplo']
]}]);

await make('Extras',[{name:'Extras',title:'Horas extraordinarias',note:'Archivo opcional. Si lo retira, no se añaden horas extra. Horas exactas, sin variación.',headers:['ID','Fecha','Inicio','Fin','Motivo'],widths:[17,20,18,18,64],dates:[1],times:[2,3],rows:[
 ['EMP002',date('2026-09-08'),time('20:15'),time('21:15'),'Cierre de inventario de ejemplo']
]}]);
