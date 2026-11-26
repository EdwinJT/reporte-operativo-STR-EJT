/**
 * Exportador Nativo OpenXML (.xlsx) 100% Offline — PICOTÓN | CONTROL OPERATIVO STRACON
 * Soporta:
 *  1. Libro de Planilla Operativa (`exportProductionWorkbookToExcel`):
 *     - Comparativo Mensual, Control <EQUIPO> (39 filas con fórmulas, Metas, Tendencia, Avance), Resumen por Área y Base de Datos.
 *  2. Libro Gerencial Acumulado (`exportExecutiveWorkbookToExcel`):
 *     - Hoja `Resumen Gerencial` con 4 tablas gerenciales + 2 gráficos nativos de Excel (Horas vs Meta y Material Día/Noche).
 *     - Hoja `Resumen Semanal` (Domingo a Sábado) + 1 gráfico nativo de Excel.
 *     - Hoja `Detalle Diario` completo por equipo y turno.
 *     - Hojas `Control <EQUIPO>` del último período de valorización y `Base de Datos`.
 */

import {
  buildMatrixModel,
  buildMonthlyComparative,
  buildExecutiveDashboardModel,
  getOperativeTargets
} from './productionService.js';
import {
  formatDiaCortoExcel,
  formatFechaDDMMYYYY,
  getFechaHoyPeru,
  getPeriodTitleSpanish,
  getValuationPeriodRange
} from '../utils/datePeru.js';

function escapeXml(str = '') {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, ' ');
}

function colName(col1Based) {
  let num = col1Based;
  let letters = '';
  while (num > 0) {
    const rem = (num - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    num = Math.floor((num - 1) / 26);
  }
  return letters;
}

function sanitizeSheetName(name = '', used = new Set()) {
  let clean = String(name || 'Hoja')
    .replace(/[:\\/?*\[\]]/g, '-')
    .trim()
    .slice(0, 31);
  if (!clean) clean = 'Hoja';
  let candidate = clean;
  let idx = 2;
  while (used.has(candidate)) {
    const suffix = ` ${idx}`;
    candidate = clean.slice(0, 31 - suffix.length) + suffix;
    idx++;
  }
  used.add(candidate);
  return candidate;
}

function strCell(ref, text, styleId = 0) {
  return `<c r="${ref}" t="inlineStr" s="${styleId}"><is><t>${escapeXml(text)}</t></is></c>`;
}

function numCell(ref, val, styleId = 0) {
  if (val === null || val === undefined || val === '' || !Number.isFinite(Number(val))) {
    return `<c r="${ref}" s="${styleId}"/>`;
  }
  return `<c r="${ref}" s="${styleId}"><v>${Number(val)}</v></c>`;
}

function formulaCell(ref, formula, cachedVal, styleId = 0) {
  const fXml = `<f>${escapeXml(formula)}</f>`;
  const vXml =
    cachedVal !== null && cachedVal !== undefined && Number.isFinite(Number(cachedVal))
      ? `<v>${Number(cachedVal)}</v>`
      : '';
  return `<c r="${ref}" s="${styleId}">${fXml}${vXml}</c>`;
}

function buildMatrixSheetXml(matrix, { tituloPersonalizado = '' } = {}) {
  const days = matrix.days || [];
  const totalDayCols = days.length;
  const totalColIdx = totalDayCols + 2;
  const totalColLetter = colName(totalColIdx);
  const lastDayColLetter = totalDayCols > 0 ? colName(totalDayCols + 1) : 'B';

  const periodTitle =
    (tituloPersonalizado || '').trim() ||
    getPeriodTitleSpanish(matrix.fechaDesde, matrix.fechaHasta) ||
    'CONTROL OPERATIVO';

  const colsXml = [
    `<col min="1" max="1" width="32" customWidth="1"/>`,
    ...days.map((_, idx) => `<col min="${idx + 2}" max="${idx + 2}" width="12.5" customWidth="1"/>`),
    `<col min="${totalColIdx}" max="${totalColIdx}" width="15.5" customWidth="1"/>`
  ].join('');

  const rowsXml = [];
  const merges = [`A1:A1`, `B1:${totalColLetter}1`, `A3:A4`];

  {
    const cells = [
      strCell('A1', 'STRACON', 1),
      strCell('B1', `CONTROL FRACTURA DE ROCA - ${periodTitle} | TURNO DÍA + TURNO NOCHE`, 2)
    ];
    for (let c = 3; c <= totalColIdx; c++) cells.push(`<c r="${colName(c)}1" s="2"/>`);
    rowsXml.push(`<row r="1" ht="30" customHeight="1">${cells.join('')}</row>`);
  }
  rowsXml.push(`<row r="2" ht="7" customHeight="1"/>`);

  {
    const cells = [strCell('A3', 'Fecha', 3)];
    days.forEach((d, i) => cells.push(strCell(`${colName(i + 2)}3`, d.diaSemana, 3)));
    cells.push(strCell(`${totalColLetter}3`, 'TOTAL / PROM.', 3));
    rowsXml.push(`<row r="3" ht="18" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [`<c r="A4" s="3"/>`];
    days.forEach((d, i) => cells.push(strCell(`${colName(i + 2)}4`, d.fechaCorta, 3)));
    cells.push(strCell(`${totalColLetter}4`, 'PERÍODO', 3));
    rowsXml.push(`<row r="4" ht="18" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A5', 'EQUIPO', 4)];
    days.forEach((_, i) => cells.push(strCell(`${colName(i + 2)}5`, matrix.codigoEquipo, 4)));
    cells.push(strCell(`${totalColLetter}5`, matrix.codigoEquipo, 4));
    rowsXml.push(`<row r="5" ht="20" customHeight="1">${cells.join('')}</row>`);
  }

  const fullBandRow = (rowNum, label, styleId, height = 19) => {
    const cells = [strCell(`A${rowNum}`, label, styleId)];
    for (let c = 2; c <= totalColIdx; c++) cells.push(`<c r="${colName(c)}${rowNum}" s="${styleId}"/>`);
    merges.push(`A${rowNum}:${totalColLetter}${rowNum}`);
    rowsXml.push(`<row r="${rowNum}" ht="${height}" customHeight="1">${cells.join('')}</row>`);
  };

  fullBandRow(6, 'TURNO DÍA', 5, 20);
  {
    const cells = [strCell('A7', 'Horómetro Inicial', 9)];
    days.forEach((d, i) => cells.push(numCell(`${colName(i + 2)}7`, d.dia.hmInicial, 20)));
    cells.push(strCell(`${totalColLetter}7`, '-', 20));
    rowsXml.push(`<row r="7" ht="18" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A8', 'Horómetro Final', 9)];
    days.forEach((d, i) => cells.push(numCell(`${colName(i + 2)}8`, d.dia.hmFinal, 20)));
    cells.push(
      numCell(`${totalColLetter}8`, matrix.summary.ultimoHorometro !== null ? matrix.summary.ultimoHorometro : '', 20)
    );
    rowsXml.push(`<row r="8" ht="18" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A9', 'Horas Día', 11)];
    days.forEach((d, i) => {
      const col = colName(i + 2);
      cells.push(
        formulaCell(
          `${col}9`,
          `IF(AND(ISNUMBER(${col}7),ISNUMBER(${col}8),${col}8>=${col}7),ROUND(${col}8-${col}7,1),0)`,
          d.dia.horas,
          21
        )
      );
    });
    cells.push(
      formulaCell(`${totalColLetter}9`, totalDayCols > 0 ? `SUM(B9:${lastDayColLetter}9)` : '0', matrix.summary.horasDia, 21)
    );
    rowsXml.push(`<row r="9" ht="19" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A10', 'Material Fracturado (m³)', 12)];
    days.forEach((d, i) => {
      const col = colName(i + 2);
      cells.push(numCell(`${col}10`, d.dia.material !== null && d.dia.material > 0 ? d.dia.material : '', 24));
    });
    cells.push(
      formulaCell(`${totalColLetter}10`, totalDayCols > 0 ? `SUM(B10:${lastDayColLetter}10)` : '0', matrix.summary.materialDia, 24)
    );
    rowsXml.push(`<row r="10" ht="19" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A11', 'Rendimiento (m³/HM)', 10)];
    days.forEach((d, i) => {
      const col = colName(i + 2);
      cells.push(
        formulaCell(`${col}11`, `IF(${col}9>0,ROUND(${col}10/${col}9,2),0)`, d.dia.record ? d.dia.rendimiento : '', 25)
      );
    });
    cells.push(
      formulaCell(`${totalColLetter}11`, `IF(${totalColLetter}9>0,ROUND(${totalColLetter}10/${totalColLetter}9,2),0)`, matrix.summary.rendDia, 25)
    );
    rowsXml.push(`<row r="11" ht="18" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A12', 'Área de Trabajo', 9)];
    days.forEach((d, i) => cells.push(strCell(`${colName(i + 2)}12`, d.dia.area || '', 26)));
    cells.push(strCell(`${totalColLetter}12`, '', 26));
    rowsXml.push(`<row r="12" ht="20" customHeight="1">${cells.join('')}</row>`);
  }

  fullBandRow(13, '', 34, 10);
  fullBandRow(14, 'TURNO NOCHE', 6, 20);
  {
    const cells = [strCell('A15', 'Horómetro Inicial', 9)];
    days.forEach((d, i) => cells.push(numCell(`${colName(i + 2)}15`, d.noche.hmInicial, 22)));
    cells.push(strCell(`${totalColLetter}15`, '-', 22));
    rowsXml.push(`<row r="15" ht="18" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A16', 'Horómetro Final', 9)];
    days.forEach((d, i) => cells.push(numCell(`${colName(i + 2)}16`, d.noche.hmFinal, 22)));
    cells.push(strCell(`${totalColLetter}16`, '-', 22));
    rowsXml.push(`<row r="16" ht="18" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A17', 'Horas Noche', 13)];
    days.forEach((d, i) => {
      const col = colName(i + 2);
      cells.push(
        formulaCell(
          `${col}17`,
          `IF(AND(ISNUMBER(${col}15),ISNUMBER(${col}16),${col}16>=${col}15),ROUND(${col}16-${col}15,1),0)`,
          d.noche.horas,
          23
        )
      );
    });
    cells.push(
      formulaCell(`${totalColLetter}17`, totalDayCols > 0 ? `SUM(B17:${lastDayColLetter}17)` : '0', matrix.summary.horasNoche, 23)
    );
    rowsXml.push(`<row r="17" ht="19" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A18', 'Material Fracturado (m³)', 14)];
    days.forEach((d, i) => {
      const col = colName(i + 2);
      cells.push(numCell(`${col}18`, d.noche.material !== null && d.noche.material > 0 ? d.noche.material : '', 24));
    });
    cells.push(
      formulaCell(`${totalColLetter}18`, totalDayCols > 0 ? `SUM(B18:${lastDayColLetter}18)` : '0', matrix.summary.materialNoche, 24)
    );
    rowsXml.push(`<row r="18" ht="19" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A19', 'Rendimiento (m³/HM)', 10)];
    days.forEach((d, i) => {
      const col = colName(i + 2);
      cells.push(
        formulaCell(`${col}19`, `IF(${col}17>0,ROUND(${col}18/${col}17,2),0)`, d.noche.record ? d.noche.rendimiento : '', 25)
      );
    });
    cells.push(
      formulaCell(`${totalColLetter}19`, `IF(${totalColLetter}17>0,ROUND(${totalColLetter}18/${totalColLetter}17,2),0)`, matrix.summary.rendNoche, 25)
    );
    rowsXml.push(`<row r="19" ht="18" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A20', 'Área de Trabajo', 9)];
    days.forEach((d, i) => cells.push(strCell(`${colName(i + 2)}20`, d.noche.area || '', 26)));
    cells.push(strCell(`${totalColLetter}20`, '', 26));
    rowsXml.push(`<row r="20" ht="20" customHeight="1">${cells.join('')}</row>`);
  }

  rowsXml.push(`<row r="21" ht="8" customHeight="1"/>`);
  fullBandRow(22, 'TOTAL DÍA', 7, 20);
  {
    const cells = [strCell('A23', 'Horas Totales (Día+Noche)', 15)];
    days.forEach((d, i) => {
      const col = colName(i + 2);
      cells.push(formulaCell(`${col}23`, `ROUND(${col}9+${col}17,1)`, d.total.horas, 27));
    });
    cells.push(
      formulaCell(`${totalColLetter}23`, totalDayCols > 0 ? `SUM(B23:${lastDayColLetter}23)` : '0', matrix.summary.horasTotal, 27)
    );
    rowsXml.push(`<row r="23" ht="19" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A24', 'Material Total (m³)', 16)];
    days.forEach((d, i) => {
      const col = colName(i + 2);
      cells.push(formulaCell(`${col}24`, `${col}10+${col}18`, d.total.material, 28));
    });
    cells.push(
      formulaCell(`${totalColLetter}24`, totalDayCols > 0 ? `SUM(B24:${lastDayColLetter}24)` : '0', matrix.summary.materialTotal, 28)
    );
    rowsXml.push(`<row r="24" ht="19" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A25', 'Rend Total (m³/HM)', 10)];
    days.forEach((d, i) => {
      const col = colName(i + 2);
      cells.push(formulaCell(`${col}25`, `IF(${col}23>0,ROUND(${col}24/${col}23,2),0)`, d.total.rendimiento, 29));
    });
    cells.push(
      formulaCell(`${totalColLetter}25`, `IF(${totalColLetter}23>0,ROUND(${totalColLetter}24/${totalColLetter}23,2),0)`, matrix.summary.rendTotal, 29)
    );
    rowsXml.push(`<row r="25" ht="18" customHeight="1">${cells.join('')}</row>`);
  }

  rowsXml.push(`<row r="26" ht="8" customHeight="1"/>`);
  fullBandRow(27, 'ACUMULADOS Y AVANCE', 8, 20);
  {
    const cells = [strCell('A28', 'Acumulado Hr Día', 17)];
    days.forEach((d, i) => {
      const col = colName(i + 2);
      const prevCol = colName(i + 1);
      cells.push(formulaCell(`${col}28`, i === 0 ? `${col}9` : `${prevCol}28+${col}9`, d.acumulado.hrDia, 30));
    });
    cells.push(formulaCell(`${totalColLetter}28`, `${lastDayColLetter}28`, matrix.summary.horasDia, 30));
    rowsXml.push(`<row r="28" ht="18" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A29', 'Acumulado Hr Noche', 17)];
    days.forEach((d, i) => {
      const col = colName(i + 2);
      const prevCol = colName(i + 1);
      cells.push(formulaCell(`${col}29`, i === 0 ? `${col}17` : `${prevCol}29+${col}17`, d.acumulado.hrNoche, 30));
    });
    cells.push(formulaCell(`${totalColLetter}29`, `${lastDayColLetter}29`, matrix.summary.horasNoche, 30));
    rowsXml.push(`<row r="29" ht="18" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A30', 'Acumulado Total (Hr)', 18)];
    days.forEach((d, i) => {
      const col = colName(i + 2);
      const prevCol = colName(i + 1);
      cells.push(formulaCell(`${col}30`, i === 0 ? `${col}23` : `${prevCol}30+${col}23`, d.acumulado.hrTotal, 27));
    });
    cells.push(formulaCell(`${totalColLetter}30`, `${lastDayColLetter}30`, matrix.summary.horasTotal, 27));
    rowsXml.push(`<row r="30" ht="19" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A31', 'Meta Diaria Total (Hr)', 46)];
    days.forEach((d, i) => cells.push(numCell(`${colName(i + 2)}31`, d.meta.metaDiaria, 45)));
    cells.push(numCell(`${totalColLetter}31`, matrix.summary.metaDiaria, 45));
    rowsXml.push(`<row r="31" ht="18" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A32', 'Acumulado Planeado (Hr)', 46)];
    days.forEach((d, i) => {
      const col = colName(i + 2);
      const prevCol = colName(i + 1);
      cells.push(formulaCell(`${col}32`, i === 0 ? `${col}31` : `${prevCol}32+${col}31`, d.meta.acumuladoPlaneado, 45));
    });
    cells.push(formulaCell(`${totalColLetter}32`, `${lastDayColLetter}32`, matrix.summary.acumuladoPlaneadoFinal, 45));
    rowsXml.push(`<row r="32" ht="18" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A33', 'Tendencia', 46)];
    days.forEach((d, i) => {
      const col = colName(i + 2);
      cells.push(formulaCell(`${col}33`, `IF(${col}32>0,${col}30/${col}32,0)`, d.meta.tendencia, d.meta.tendencia >= 1 ? 47 : 48));
    });
    cells.push(
      formulaCell(
        `${totalColLetter}33`,
        `IF(${totalColLetter}32>0,${totalColLetter}30/${totalColLetter}32,0)`,
        matrix.summary.tendenciaFinal,
        matrix.summary.tendenciaFinal >= 1 ? 47 : 48
      )
    );
    rowsXml.push(`<row r="33" ht="19" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A34', '% de Avance', 46)];
    days.forEach((d, i) => {
      const col = colName(i + 2);
      cells.push(formulaCell(`${col}34`, `IF($B$39>0,${col}30/$B$39,0)`, d.meta.avance, 44));
    });
    cells.push(formulaCell(`${totalColLetter}34`, `IF($B$39>0,${totalColLetter}30/$B$39,0)`, matrix.summary.avanceFinal, 44));
    rowsXml.push(`<row r="34" ht="19" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A35', 'Acumulado Material Total (m³)', 19)];
    days.forEach((d, i) => {
      const col = colName(i + 2);
      const prevCol = colName(i + 1);
      cells.push(formulaCell(`${col}35`, i === 0 ? `${col}24` : `${prevCol}35+${col}24`, d.acumulado.matTotal, 32));
    });
    cells.push(formulaCell(`${totalColLetter}35`, `${lastDayColLetter}35`, matrix.summary.materialTotal, 32));
    rowsXml.push(`<row r="35" ht="19" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A36', 'Rendimiento Acumulado (m³/HM)', 18)];
    days.forEach((d, i) => {
      const col = colName(i + 2);
      cells.push(formulaCell(`${col}36`, `IF(${col}30>0,ROUND(${col}35/${col}30,2),0)`, d.acumulado.rendAcum, 33));
    });
    cells.push(
      formulaCell(`${totalColLetter}36`, `IF(${totalColLetter}30>0,ROUND(${totalColLetter}35/${totalColLetter}30,2),0)`, matrix.summary.rendTotal, 33)
    );
    rowsXml.push(`<row r="36" ht="19" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A37', 'Operador de Día', 9)];
    days.forEach((d, i) => cells.push(strCell(`${colName(i + 2)}37`, d.dia.operador || '', 26)));
    cells.push(strCell(`${totalColLetter}37`, '', 26));
    rowsXml.push(`<row r="37" ht="19" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A38', 'Operador de Noche', 9)];
    days.forEach((d, i) => cells.push(strCell(`${colName(i + 2)}38`, d.noche.operador || '', 26)));
    cells.push(strCell(`${totalColLetter}38`, '', 26));
    rowsXml.push(`<row r="38" ht="19" customHeight="1">${cells.join('')}</row>`);
  }
  {
    const cells = [strCell('A39', 'Meta mensual Horas Totales (Hr)', 46)];
    days.forEach((_, i) => {
      cells.push(i === 0 ? numCell('B39', matrix.summary.metaMensual, 45) : `<c r="${colName(i + 2)}39" s="45"/>`);
    });
    cells.push(numCell(`${totalColLetter}39`, matrix.summary.metaMensual, 45));
    rowsXml.push(`<row r="39" ht="20" customHeight="1">${cells.join('')}</row>`);
  }

  const mergeXml = `<mergeCells count="${merges.length}">${merges.map((ref) => `<mergeCell ref="${ref}"/>`).join('')}</mergeCells>`;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane xSplit="1" ySplit="4" topLeftCell="B5" activePane="bottomRight" state="frozen"/>
    </sheetView>
  </sheetViews>
  <cols>${colsXml}</cols>
  <sheetData>${rowsXml.join('')}</sheetData>
  ${mergeXml}
  <pageSetup orientation="landscape" paperSize="9"/>
</worksheet>`;
}

function buildComparativeSheetXml(compModel) {
  const widths = [14, 18, 16, 17, 14, 16, 21, 19, 21];
  const colsXml = widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('');

  const rowsXml = [];
  rowsXml.push(
    `<row r="1" ht="28" customHeight="1">${strCell('A1', 'STRACON', 1)}${strCell(
      'B1',
      'PICOTÓN | CONTROL OPERATIVO — COMPARATIVO MENSUAL (DÍA + NOCHE)',
      2
    )}<c r="C1" s="2"/><c r="D1" s="2"/><c r="E1" s="2"/><c r="F1" s="2"/><c r="G1" s="2"/><c r="H1" s="2"/><c r="I1" s="2"/></row>`
  );

  const headers = [
    'Equipo',
    'Mes',
    'Turno Día (Hr)',
    'Turno Noche (Hr)',
    'Total (Hr)',
    'Material (m³)',
    'Rendimiento (m³/HM)',
    'Prom. Día (Hr/día)',
    'Prom. Noche (Hr/día)'
  ];
  rowsXml.push(
    `<row r="3" ht="24" customHeight="1">${headers.map((h, i) => strCell(`${colName(i + 1)}3`, h, 35)).join('')}</row>`
  );

  const startRow = 4;
  compModel.rows.forEach((item, idx) => {
    const r = startRow + idx;
    rowsXml.push(
      `<row r="${r}" ht="20" customHeight="1">${[
        strCell(`A${r}`, item.codigoEquipo, 36),
        strCell(`B${r}`, item.mesLabel, 36),
        numCell(`C${r}`, item.horasDia, 38),
        numCell(`D${r}`, item.horasNoche, 38),
        formulaCell(`E${r}`, `ROUND(C${r}+D${r},1)`, item.horasTotal, 38),
        numCell(`F${r}`, item.materialTotal, 37),
        formulaCell(`G${r}`, `IF(E${r}>0,ROUND(F${r}/E${r},2),0)`, item.rendimiento, 39),
        numCell(`H${r}`, item.promDiaHr, 38),
        numCell(`I${r}`, item.promNocheHr, 38)
      ].join('')}</row>`
    );
  });

  const totRow = startRow + compModel.rows.length;
  const endRow = Math.max(startRow, totRow - 1);
  const hasRows = compModel.rows.length > 0;
  const t = compModel.totals;
  rowsXml.push(
    `<row r="${totRow}" ht="22" customHeight="1">${[
      strCell(`A${totRow}`, 'TOTAL FLOTA', 40),
      strCell(`B${totRow}`, `${compModel.rows.length} grupo(s)`, 40),
      formulaCell(`C${totRow}`, hasRows ? `SUM(C${startRow}:C${endRow})` : '0', t.horasDia, 41),
      formulaCell(`D${totRow}`, hasRows ? `SUM(D${startRow}:D${endRow})` : '0', t.horasNoche, 41),
      formulaCell(`E${totRow}`, hasRows ? `SUM(E${startRow}:E${endRow})` : '0', t.horasTotal, 41),
      formulaCell(`F${totRow}`, hasRows ? `SUM(F${startRow}:F${endRow})` : '0', t.materialTotal, 42),
      formulaCell(`G${totRow}`, `IF(E${totRow}>0,ROUND(F${totRow}/E${totRow},2),0)`, t.rendimiento, 43),
      numCell(`H${totRow}`, t.promDiaHr, 41),
      numCell(`I${totRow}`, t.promNocheHr, 41)
    ].join('')}</row>`
  );

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>${colsXml}</cols>
  <sheetData>${rowsXml.join('')}</sheetData>
  <mergeCells count="1"><mergeCell ref="B1:I1"/></mergeCells>
  <pageSetup orientation="landscape" paperSize="9"/>
</worksheet>`;
}

/**
 * Generadores nativos de gráficos OpenXML (DrawingML + ChartML) para insertar gráficos editables en Excel.
 */
function buildColumnBarChartXml({
  title,
  grouping = 'clustered',
  categories = [],
  catFormula,
  series = []
}) {
  const catPointsXml = categories
    .map((c, idx) => `<c:pt idx="${idx}"><c:v>${escapeXml(c)}</c:v></c:pt>`)
    .join('');

  const seriesXml = series
    .map((ser, sIdx) => {
      const valPts = ser.values
        .map((v, idx) => `<c:pt idx="${idx}"><c:v>${Number(v) || 0}</c:v></c:pt>`)
        .join('');
      const fillXml = ser.colorHex
        ? `<c:spPr><a:solidFill><a:srgbClr val="${ser.colorHex}"/></a:solidFill></c:spPr>`
        : '';
      return `
      <c:ser>
        <c:idx val="${sIdx}"/>
        <c:order val="${sIdx}"/>
        <c:tx>
          <c:strRef>
            <c:f>${escapeXml(ser.nameFormula)}</c:f>
            <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXml(ser.name)}</c:v></c:pt></c:strCache>
          </c:strRef>
        </c:tx>
        ${fillXml}
        <c:invertIfNegative val="0"/>
        <c:cat>
          <c:strRef>
            <c:f>${escapeXml(catFormula)}</c:f>
            <c:strCache><c:ptCount val="${categories.length}"/>${catPointsXml}</c:strCache>
          </c:strRef>
        </c:cat>
        <c:val>
          <c:numRef>
            <c:f>${escapeXml(ser.valFormula)}</c:f>
            <c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${ser.values.length}"/>${valPts}</c:numCache>
          </c:numRef>
        </c:val>
      </c:ser>`;
    })
    .join('');

  const overlapTag = grouping === 'stacked' ? '<c:overlap val="100"/>' : '';

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:lang val="es-PE"/>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="es-PE" sz="1200" b="1"><a:solidFill><a:srgbClr val="002B5C"/></a:solidFill></a:rPr><a:t>${escapeXml(title)}</a:t></a:r></a:p></c:rich></c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="${grouping}"/>
        <c:varyColors val="0"/>
        ${seriesXml}
        ${overlapTag}
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="2"/>
        <c:crosses val="autoZero"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="1"/>
        <c:crosses val="autoZero"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="b"/>
      <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;
}

function buildDrawingXml(anchors = []) {
  const anchorsXml = anchors
    .map(
      (a, idx) => `
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>${a.fromCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${a.fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>${a.toCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${a.toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="${idx + 2}" name="${escapeXml(a.name)}"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId${idx + 1}"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>`
    )
    .join('');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  ${anchorsXml}
</xdr:wsDr>`;
}

function buildDrawingRelsXml(chartIndices = []) {
  const rels = chartIndices
    .map(
      (chartNum, idx) =>
        `<Relationship Id="rId${idx + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartNum}.xml"/>`
    )
    .join('\n  ');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${rels}
</Relationships>`;
}

/**
 * Construye la hoja `Resumen Gerencial` con las 4 tablas ejecutivas + serie auxiliar + 2 gráficos nativos.
 */
function buildExecutiveSummarySheetXml(execModel, { tituloPersonalizado = '', hasDrawing = false } = {}) {
  const colsWidths = [16, 14, 15, 15, 14, 17, 16, 15, 16];
  const colsXml = colsWidths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('');

  const rowsXml = [];
  const merges = ['B1:I1'];
  const titulo =
    (tituloPersonalizado || '').trim() ||
    'CONTROL FRACTURA DE ROCA — RESUMEN GERENCIAL ACUMULADO | TURNO DÍA + NOCHE';

  rowsXml.push(
    `<row r="1" ht="30" customHeight="1">${strCell('A1', 'STRACON', 1)}${strCell('B1', titulo, 2)}<c r="C1" s="2"/><c r="D1" s="2"/><c r="E1" s="2"/><c r="F1" s="2"/><c r="G1" s="2"/><c r="H1" s="2"/><c r="I1" s="2"/></row>`
  );

  let r = 3;
  // TABLA 1: RESUMEN MENSUAL POR EQUIPO
  merges.push(`A${r}:I${r}`);
  rowsXml.push(
    `<row r="${r}" ht="22" customHeight="1">${strCell(`A${r}`, '1. RESUMEN MENSUAL POR EQUIPO', 5)}${Array.from(
      { length: 8 },
      (_, k) => `<c r="${colName(k + 2)}${r}" s="5"/>`
    ).join('')}</row>`
  );
  r++;

  const h1 = ['Mes', 'Equipo', 'Horas Máquina (Hr)', 'Meta Mensual (Hr)', '% Cumplimiento', 'Material (m³)', 'Rendimiento (m³/HM)', 'Días Operativos', 'Hr / Día Op.'];
  rowsXml.push(`<row r="${r}" ht="24" customHeight="1">${h1.map((h, i) => strCell(`${colName(i + 1)}${r}`, h, 35)).join('')}</row>`);
  r++;

  const t1Start = r;
  const t1EquipoRowNums = [];
  execModel.monthlyTable.forEach((item) => {
    const isSub = item.rowType === 'subtotal';
    const stTxt = isSub ? 40 : 36;
    const stDec1 = isSub ? 41 : 38;
    const stInt = isSub ? 42 : 37;
    const stDec2 = isSub ? 43 : 39;
    const stPct = isSub ? 44 : item.pctCumplimiento >= 1 ? 47 : 48;
    if (!isSub) t1EquipoRowNums.push(r);

    rowsXml.push(
      `<row r="${r}" ht="19" customHeight="1">${[
        strCell(`A${r}`, item.mesNombre, stTxt),
        strCell(`B${r}`, item.codigoEquipo, stTxt),
        numCell(`C${r}`, item.horasTotal, stDec1),
        numCell(`D${r}`, item.metaMensual, stInt),
        formulaCell(`E${r}`, `IF(D${r}>0,C${r}/D${r},0)`, item.pctCumplimiento, stPct),
        numCell(`F${r}`, item.materialTotal, stInt),
        formulaCell(`G${r}`, `IF(C${r}>0,ROUND(F${r}/C${r},2),0)`, item.rendimiento ?? item.rendTotal, stDec2),
        numCell(`H${r}`, item.diasOperativos, stInt),
        formulaCell(`I${r}`, `IF(H${r}>0,ROUND(C${r}/H${r},2),0)`, item.hrPorDiaOperativo, stDec2)
      ].join('')}</row>`
    );
    r++;
  });

  const t1SumC = t1EquipoRowNums.length > 0 ? t1EquipoRowNums.map((n) => `C${n}`).join('+') : '0';
  const t1SumD = t1EquipoRowNums.length > 0 ? t1EquipoRowNums.map((n) => `D${n}`).join('+') : '0';
  const t1SumF = t1EquipoRowNums.length > 0 ? t1EquipoRowNums.map((n) => `F${n}`).join('+') : '0';
  const t1SumH = t1EquipoRowNums.length > 0 ? t1EquipoRowNums.map((n) => `H${n}`).join('+') : '0';
  const t = execModel.totals;
  rowsXml.push(
    `<row r="${r}" ht="22" customHeight="1">${[
      strCell(`A${r}`, 'TOTAL ACUMULADO', 40),
      strCell(`B${r}`, execModel.equipos.join('+'), 40),
      formulaCell(`C${r}`, t1SumC, t.horasTotal, 41),
      formulaCell(`D${r}`, t1SumD, t.metaMensualTotal, 42),
      formulaCell(`E${r}`, `IF(D${r}>0,C${r}/D${r},0)`, t.pctCumplimiento, 44),
      formulaCell(`F${r}`, t1SumF, t.materialTotal, 42),
      formulaCell(`G${r}`, `IF(C${r}>0,ROUND(F${r}/C${r},2),0)`, t.rendTotal, 43),
      formulaCell(`H${r}`, t1SumH, t.diasOperativos, 42),
      formulaCell(`I${r}`, `IF(H${r}>0,ROUND(C${r}/H${r},2),0)`, t.hrPorDiaOperativo, 43)
    ].join('')}</row>`
  );
  r += 2;

  // TABLA 2: DISTRIBUCIÓN DE HORAS POR TURNO Y EQUIPO
  merges.push(`A${r}:F${r}`);
  rowsXml.push(
    `<row r="${r}" ht="22" customHeight="1">${strCell(`A${r}`, '2. DISTRIBUCIÓN DE HORAS MÁQUINA POR TURNO Y EQUIPO', 6)}${Array.from(
      { length: 5 },
      (_, k) => `<c r="${colName(k + 2)}${r}" s="6"/>`
    ).join('')}</row>`
  );
  r++;
  const h2 = ['Mes', 'Equipo', 'Turno Día (Hr)', 'Turno Noche (Hr)', 'Total (Hr)', '% Turno Noche'];
  rowsXml.push(`<row r="${r}" ht="22" customHeight="1">${h2.map((h, i) => strCell(`${colName(i + 1)}${r}`, h, 35)).join('')}</row>`);
  r++;
  const t2EqRows = [];
  execModel.monthlyTable.forEach((item) => {
    const isSub = item.rowType === 'subtotal';
    const stTxt = isSub ? 40 : 36;
    const stDec1 = isSub ? 41 : 38;
    if (!isSub) t2EqRows.push(r);
    rowsXml.push(
      `<row r="${r}" ht="19" customHeight="1">${[
        strCell(`A${r}`, item.mesNombre, stTxt),
        strCell(`B${r}`, item.codigoEquipo, stTxt),
        numCell(`C${r}`, item.horasDia, stDec1),
        numCell(`D${r}`, item.horasNoche, stDec1),
        formulaCell(`E${r}`, `ROUND(C${r}+D${r},1)`, item.horasTotal, stDec1),
        formulaCell(`F${r}`, `IF(E${r}>0,D${r}/E${r},0)`, item.pctNoche, 44)
      ].join('')}</row>`
    );
    r++;
  });
  rowsXml.push(
    `<row r="${r}" ht="22" customHeight="1">${[
      strCell(`A${r}`, 'TOTAL ACUMULADO', 40),
      strCell(`B${r}`, '', 40),
      formulaCell(`C${r}`, t2EqRows.length ? t2EqRows.map((n) => `C${n}`).join('+') : '0', t.horasDia, 41),
      formulaCell(`D${r}`, t2EqRows.length ? t2EqRows.map((n) => `D${n}`).join('+') : '0', t.horasNoche, 41),
      formulaCell(`E${r}`, `ROUND(C${r}+D${r},1)`, t.horasTotal, 41),
      formulaCell(`F${r}`, `IF(E${r}>0,D${r}/E${r},0)`, t.pctNoche, 44)
    ].join('')}</row>`
  );
  r += 2;

  // TABLA 3: MATERIAL FRACTURADO Y RENDIMIENTO POR TURNO
  merges.push(`A${r}:H${r}`);
  rowsXml.push(
    `<row r="${r}" ht="22" customHeight="1">${strCell(`A${r}`, '3. MATERIAL FRACTURADO Y RENDIMIENTO POR TURNO Y EQUIPO', 7)}${Array.from(
      { length: 7 },
      (_, k) => `<c r="${colName(k + 2)}${r}" s="7"/>`
    ).join('')}</row>`
  );
  r++;
  const h3 = ['Mes', 'Equipo', 'Material Día (m³)', 'Material Noche (m³)', 'Material Total (m³)', 'Rend. Día (m³/HM)', 'Rend. Noche (m³/HM)', 'Rend. Total (m³/HM)'];
  rowsXml.push(`<row r="${r}" ht="22" customHeight="1">${h3.map((h, i) => strCell(`${colName(i + 1)}${r}`, h, 35)).join('')}</row>`);
  r++;
  const t3EqRows = [];
  execModel.monthlyTable.forEach((item) => {
    const isSub = item.rowType === 'subtotal';
    const stTxt = isSub ? 40 : 36;
    const stInt = isSub ? 42 : 37;
    const stDec2 = isSub ? 43 : 39;
    if (!isSub) t3EqRows.push(r);
    rowsXml.push(
      `<row r="${r}" ht="19" customHeight="1">${[
        strCell(`A${r}`, item.mesNombre, stTxt),
        strCell(`B${r}`, item.codigoEquipo, stTxt),
        numCell(`C${r}`, item.materialDia, stInt),
        numCell(`D${r}`, item.materialNoche, stInt),
        formulaCell(`E${r}`, `C${r}+D${r}`, item.materialTotal, stInt),
        numCell(`F${r}`, item.rendDia, stDec2),
        numCell(`G${r}`, item.rendNoche, stDec2),
        numCell(`H${r}`, item.rendTotal, stDec2)
      ].join('')}</row>`
    );
    r++;
  });
  rowsXml.push(
    `<row r="${r}" ht="22" customHeight="1">${[
      strCell(`A${r}`, 'TOTAL ACUMULADO', 40),
      strCell(`B${r}`, '', 40),
      formulaCell(`C${r}`, t3EqRows.length ? t3EqRows.map((n) => `C${n}`).join('+') : '0', t.materialDia, 42),
      formulaCell(`D${r}`, t3EqRows.length ? t3EqRows.map((n) => `D${n}`).join('+') : '0', t.materialNoche, 42),
      formulaCell(`E${r}`, `C${r}+D${r}`, t.materialTotal, 42),
      numCell(`F${r}`, t.rendDia, 43),
      numCell(`G${r}`, t.rendNoche, 43),
      numCell(`H${r}`, t.rendTotal, 43)
    ].join('')}</row>`
  );
  r += 2;

  // TABLA 4: PROMEDIO HORAS MÁQUINA POR DÍA OPERATIVO + RANGO AUXILIAR PARA GRÁFICOS
  merges.push(`A${r}:D${r}`);
  rowsXml.push(
    `<row r="${r}" ht="22" customHeight="1">${strCell(`A${r}`, '4. PROMEDIO HORAS MÁQUINA POR DÍA OPERATIVO', 8)}${Array.from(
      { length: 3 },
      (_, k) => `<c r="${colName(k + 2)}${r}" s="8"/>`
    ).join('')}</row>`
  );
  r++;
  const h4 = ['Mes', 'Equipo', 'Prom. Turno Día (Hr/día)', 'Prom. Turno Noche (Hr/día)'];
  rowsXml.push(`<row r="${r}" ht="22" customHeight="1">${h4.map((h, i) => strCell(`${colName(i + 1)}${r}`, h, 35)).join('')}</row>`);
  r++;
  execModel.monthlyTable.forEach((item) => {
    const isSub = item.rowType === 'subtotal';
    rowsXml.push(
      `<row r="${r}" ht="19" customHeight="1">${[
        strCell(`A${r}`, item.mesNombre, isSub ? 40 : 36),
        strCell(`B${r}`, item.codigoEquipo, isSub ? 40 : 36),
        numCell(`C${r}`, item.promDiaHr, isSub ? 43 : 39),
        numCell(`D${r}`, item.promNocheHr, isSub ? 43 : 39)
      ].join('')}</row>`
    );
    r++;
  });
  r += 2;

  // Tabla Auxiliar limpia (sin subtotales) para alimentar los 2 Gráficos Nativos de Excel
  const auxHeaderRow = r;
  const auxHeaders = ['Periodo - Equipo', 'Horas Máquina (Hr)', 'Meta Mensual (Hr)', 'Material Día (m³)', 'Material Noche (m³)'];
  rowsXml.push(`<row r="${r}" ht="20" customHeight="1">${auxHeaders.map((h, i) => strCell(`${colName(i + 1)}${r}`, h, 35)).join('')}</row>`);
  r++;
  const auxStartRow = r;
  execModel.monthlyChartSeries.forEach((item) => {
    rowsXml.push(
      `<row r="${r}" ht="18" customHeight="1">${[
        strCell(`A${r}`, `${item.mesNombre.slice(0, 3)} ${item.codigoEquipo}`, 36),
        numCell(`B${r}`, item.horasTotal, 38),
        numCell(`C${r}`, item.metaMensual, 37),
        numCell(`D${r}`, item.materialDia, 37),
        numCell(`E${r}`, item.materialNoche, 37)
      ].join('')}</row>`
    );
    r++;
  });
  const auxEndRow = Math.max(auxStartRow, r - 1);

  const mergeXml = `<mergeCells count="${merges.length}">${merges.map((ref) => `<mergeCell ref="${ref}"/>`).join('')}</mergeCells>`;
  const drawingXml = hasDrawing ? `<drawing r:id="rId1"/>` : '';

  return {
    xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cols>${colsXml}</cols>
  <sheetData>${rowsXml.join('')}</sheetData>
  ${mergeXml}
  <pageSetup orientation="landscape" paperSize="9"/>
  ${drawingXml}
</worksheet>`,
    auxHeaderRow,
    auxStartRow,
    auxEndRow
  };
}

function buildWeeklySheetXml(execModel, { hasDrawing = false } = {}) {
  const equipos = execModel.equipos;
  // Columnas: Inicio(1), Fin(2), Semana(3), HM eq1..N, HM Total, Mat eq1..N, Mat Total, Rendimiento, DiasOp eq1..N
  const headers = [
    'Inicio (Dom)',
    'Fin (Sáb)',
    'Semana',
    ...equipos.map((e) => `HM ${e} (Hr)`),
    'HM Total (Hr)',
    ...equipos.map((e) => `Material ${e} (m³)`),
    'Material Total (m³)',
    'Rendimiento (m³/HM)',
    ...equipos.map((e) => `Días Op. ${e}`)
  ];
  const totalCols = headers.length;
  const colsXml = headers
    .map((h, i) => `<col min="${i + 1}" max="${i + 1}" width="${i === 2 ? 20 : 14}" customWidth="1"/>`)
    .join('');

  const rowsXml = [];
  rowsXml.push(
    `<row r="1" ht="28" customHeight="1">${strCell('A1', 'STRACON', 1)}${strCell(
      'B1',
      'RESUMEN SEMANAL (DOMINGO A SÁBADO) POR EQUIPO',
      2
    )}${Array.from({ length: totalCols - 2 }, (_, k) => `<c r="${colName(k + 3)}1" s="2"/>`).join('')}</row>`
  );
  rowsXml.push(
    `<row r="3" ht="24" customHeight="1">${headers.map((h, i) => strCell(`${colName(i + 1)}3`, h, 35)).join('')}</row>`
  );

  const startRow = 4;
  const hmEqStartCol = 4;
  const hmTotCol = hmEqStartCol + equipos.length;
  const matEqStartCol = hmTotCol + 1;
  const matTotCol = matEqStartCol + equipos.length;
  const rendCol = matTotCol + 1;

  execModel.weeklyTable.forEach((wk, idx) => {
    const r = startRow + idx;
    const cells = [
      strCell(`A${r}`, wk.inicioDom, 36),
      strCell(`B${r}`, wk.finSab, 36),
      strCell(`C${r}`, wk.semanaLabel, 36)
    ];
    equipos.forEach((eq) => cells.push(numCell(`${colName(cells.length + 1)}${r}`, wk.byEquip[eq]?.horas || 0, 38)));
    const hmSumExpr =
      equipos.length > 0
        ? `ROUND(SUM(${colName(hmEqStartCol)}${r}:${colName(hmEqStartCol + equipos.length - 1)}${r}),1)`
        : '0';
    cells.push(formulaCell(`${colName(hmTotCol)}${r}`, hmSumExpr, wk.horasTotal, 38));

    equipos.forEach((eq) => cells.push(numCell(`${colName(cells.length + 1)}${r}`, wk.byEquip[eq]?.material || 0, 37)));
    const matSumExpr =
      equipos.length > 0
        ? `SUM(${colName(matEqStartCol)}${r}:${colName(matEqStartCol + equipos.length - 1)}${r})`
        : '0';
    cells.push(formulaCell(`${colName(matTotCol)}${r}`, matSumExpr, wk.materialTotal, 37));

    cells.push(
      formulaCell(
        `${colName(rendCol)}${r}`,
        `IF(${colName(hmTotCol)}${r}>0,ROUND(${colName(matTotCol)}${r}/${colName(hmTotCol)}${r},2),0)`,
        wk.rendimiento,
        39
      )
    );
    equipos.forEach((eq) => cells.push(numCell(`${colName(cells.length + 1)}${r}`, wk.byEquip[eq]?.diasOp || 0, 37)));

    rowsXml.push(`<row r="${r}" ht="19" customHeight="1">${cells.join('')}</row>`);
  });

  const totRow = startRow + execModel.weeklyTable.length;
  const endRow = Math.max(startRow, totRow - 1);
  const hasRows = execModel.weeklyTable.length > 0;
  const totCells = [
    strCell(`A${totRow}`, 'TOTAL', 40),
    strCell(`B${totRow}`, '', 40),
    strCell(`C${totRow}`, `${execModel.weeklyTable.length} semana(s)`, 40)
  ];
  equipos.forEach((eq) => {
    const cL = colName(totCells.length + 1);
    totCells.push(
      formulaCell(`${cL}${totRow}`, hasRows ? `ROUND(SUM(${cL}${startRow}:${cL}${endRow}),1)` : '0', execModel.weeklyTotals.byEquip[eq]?.horas || 0, 41)
    );
  });
  {
    const cL = colName(hmTotCol);
    totCells.push(
      formulaCell(`${cL}${totRow}`, hasRows ? `ROUND(SUM(${cL}${startRow}:${cL}${endRow}),1)` : '0', execModel.weeklyTotals.horasTotal, 41)
    );
  }
  equipos.forEach((eq) => {
    const cL = colName(totCells.length + 1);
    totCells.push(
      formulaCell(`${cL}${totRow}`, hasRows ? `SUM(${cL}${startRow}:${cL}${endRow})` : '0', execModel.weeklyTotals.byEquip[eq]?.material || 0, 42)
    );
  });
  {
    const cL = colName(matTotCol);
    totCells.push(
      formulaCell(`${cL}${totRow}`, hasRows ? `SUM(${cL}${startRow}:${cL}${endRow})` : '0', execModel.weeklyTotals.materialTotal, 42)
    );
  }
  {
    const rL = colName(rendCol);
    const hL = colName(hmTotCol);
    const mL = colName(matTotCol);
    totCells.push(formulaCell(`${rL}${totRow}`, `IF(${hL}${totRow}>0,ROUND(${mL}${totRow}/${hL}${totRow},2),0)`, execModel.weeklyTotals.rendimiento, 43));
  }
  equipos.forEach((eq) => {
    const cL = colName(totCells.length + 1);
    totCells.push(
      formulaCell(`${cL}${totRow}`, hasRows ? `SUM(${cL}${startRow}:${cL}${endRow})` : '0', execModel.weeklyTotals.byEquip[eq]?.diasOp || 0, 42)
    );
  });
  rowsXml.push(`<row r="${totRow}" ht="22" customHeight="1">${totCells.join('')}</row>`);

  const drawingXml = hasDrawing ? `<drawing r:id="rId1"/>` : '';
  return {
    xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols>${colsXml}</cols>
  <sheetData>${rowsXml.join('')}</sheetData>
  <mergeCells count="1"><mergeCell ref="B1:${colName(totalCols)}1"/></mergeCells>
  <pageSetup orientation="landscape" paperSize="9"/>
  ${drawingXml}
</worksheet>`,
    startRow,
    endRow,
    matTotColLetter: colName(matTotCol),
    hmTotColLetter: colName(hmTotCol),
    totalCols
  };
}

function buildDailyDetailSheetXml(execModel) {
  const equipos = execModel.equipos;
  const headers = [
    'Fecha',
    'Mes',
    'Día',
    'Inicio Semana (Dom)',
    ...equipos.flatMap((e) => [`${e} Día (Hr)`, `${e} Noche (Hr)`, `${e} Total (Hr)`, `${e} Material (m³)`]),
    'HM Total (Hr)',
    'Material Total (m³)'
  ];
  const colsXml = headers.map((_, i) => `<col min="${i + 1}" max="${i + 1}" width="${i < 4 ? 15 : 13}" customWidth="1"/>`).join('');

  const rowsXml = [];
  rowsXml.push(
    `<row r="1" ht="24" customHeight="1">${headers.map((h, i) => strCell(`${colName(i + 1)}1`, h, 35)).join('')}</row>`
  );

  execModel.dailyDetail.forEach((d, idx) => {
    const r = idx + 2;
    const cells = [
      strCell(`A${r}`, d.fecha, 36),
      strCell(`B${r}`, d.mesNombre, 36),
      strCell(`C${r}`, d.diaSemana, 36),
      strCell(`D${r}`, d.inicioDom, 36)
    ];
    equipos.forEach((eq) => {
      const info = d.byEquip[eq] || { diaHr: 0, nocheHr: 0, totalHr: 0, material: 0 };
      cells.push(numCell(`${colName(cells.length + 1)}${r}`, info.diaHr, 38));
      cells.push(numCell(`${colName(cells.length + 1)}${r}`, info.nocheHr, 38));
      cells.push(numCell(`${colName(cells.length + 1)}${r}`, info.totalHr, 38));
      cells.push(numCell(`${colName(cells.length + 1)}${r}`, info.material, 37));
    });
    cells.push(numCell(`${colName(cells.length + 1)}${r}`, d.hmTotal, 38));
    cells.push(numCell(`${colName(cells.length + 1)}${r}`, d.materialTotal, 37));
    rowsXml.push(`<row r="${r}" ht="18" customHeight="1">${cells.join('')}</row>`);
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols>${colsXml}</cols>
  <sheetData>${rowsXml.join('')}</sheetData>
</worksheet>`;
}

function buildAreaSummarySheetXml(areaRows = [], codigoEquipo, fechaDesde, fechaHasta) {
  const widths = [28, 14, 16, 20, 16, 20];
  const colsXml = widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('');

  const rowsXml = [];
  rowsXml.push(
    `<row r="1" ht="28" customHeight="1">${strCell(
      'A1',
      `RESUMEN POR ÁREA DE TRABAJO — ${codigoEquipo} (${formatFechaDDMMYYYY(fechaDesde)} al ${formatFechaDDMMYYYY(fechaHasta)})`,
      2
    )}<c r="B1" s="2"/><c r="C1" s="2"/><c r="D1" s="2"/><c r="E1" s="2"/><c r="F1" s="2"/></row>`
  );

  const headers = ['Área de Trabajo', 'Turnos', 'Horas (HM)', 'Material Fracturado (m³)', '% Material', 'Rendimiento Prom. (m³/HM)'];
  rowsXml.push(`<row r="3" ht="22" customHeight="1">${headers.map((h, i) => strCell(`${colName(i + 1)}3`, h, 35)).join('')}</row>`);

  const startRow = 4;
  let totalTurnos = 0;
  let totalHoras = 0;
  let totalMat = 0;

  areaRows.forEach((a, idx) => {
    const r = startRow + idx;
    totalTurnos += a.turnos;
    totalHoras += a.horas;
    totalMat += a.material;
    const totalRowRef = startRow + areaRows.length;
    rowsXml.push(
      `<row r="${r}" ht="19" customHeight="1">${[
        strCell(`A${r}`, a.area, 36),
        numCell(`B${r}`, a.turnos, 37),
        numCell(`C${r}`, a.horas, 38),
        numCell(`D${r}`, a.material, 37),
        formulaCell(`E${r}`, `IF($D$${totalRowRef}>0,ROUND(D${r}/$D$${totalRowRef}*100,1),0)`, a.porcentajeMaterial, 38),
        formulaCell(`F${r}`, `IF(C${r}>0,ROUND(D${r}/C${r},2),0)`, a.rendimiento, 39)
      ].join('')}</row>`
    );
  });

  const totRow = startRow + areaRows.length;
  const endRow = Math.max(startRow, totRow - 1);
  const hasData = areaRows.length > 0;
  rowsXml.push(
    `<row r="${totRow}" ht="21" customHeight="1">${[
      strCell(`A${totRow}`, 'TOTAL PERÍODO', 40),
      formulaCell(`B${totRow}`, hasData ? `SUM(B${startRow}:B${endRow})` : '0', totalTurnos, 42),
      formulaCell(`C${totRow}`, hasData ? `SUM(C${startRow}:C${endRow})` : '0', Math.round(totalHoras * 10) / 10, 41),
      formulaCell(`D${totRow}`, hasData ? `SUM(D${startRow}:D${endRow})` : '0', Math.round(totalMat * 100) / 100, 42),
      numCell(`E${totRow}`, hasData ? 100 : 0, 41),
      formulaCell(`F${totRow}`, `IF(C${totRow}>0,ROUND(D${totRow}/C${totRow},2),0)`, totalHoras > 0 ? Math.round((totalMat / totalHoras) * 100) / 100 : 0, 43)
    ].join('')}</row>`
  );

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>${colsXml}</cols>
  <sheetData>${rowsXml.join('')}</sheetData>
  <mergeCells count="1"><mergeCell ref="A1:F1"/></mergeCells>
</worksheet>`;
}

function buildDatabaseSheetXml(records = []) {
  const columns = [
    { header: 'Fecha', width: 13 },
    { header: 'Día', width: 12 },
    { header: 'Equipo', width: 13 },
    { header: 'Turno', width: 10 },
    { header: 'HM Inicial', width: 13 },
    { header: 'HM Final', width: 13 },
    { header: 'Horas (HM)', width: 13 },
    { header: 'Material (m³)', width: 15 },
    { header: 'Rendimiento (m³/HM)', width: 19 },
    { header: 'Área de Trabajo', width: 20 },
    { header: 'Operador', width: 22 },
    { header: 'Observaciones', width: 32 },
    { header: 'Estado Sync', width: 15 },
    { header: 'Hora Lima', width: 12 }
  ];

  const colsXml = columns.map((c, idx) => `<col min="${idx + 1}" max="${idx + 1}" width="${c.width}" customWidth="1"/>`).join('');

  const sorted = [...records].sort((a, b) => {
    if (a.fecha !== b.fecha) return a.fecha.localeCompare(b.fecha);
    if (a.codigo_equipo !== b.codigo_equipo) return a.codigo_equipo.localeCompare(b.codigo_equipo);
    return a.turno === 'DIA' ? -1 : 1;
  });

  const rowsXml = [];
  rowsXml.push(
    `<row r="1" ht="24" customHeight="1">${columns.map((c, i) => strCell(`${colName(i + 1)}1`, c.header, 35)).join('')}</row>`
  );

  sorted.forEach((rec, idx) => {
    const r = idx + 2;
    const { diaSemana } = formatDiaCortoExcel(rec.fecha);
    rowsXml.push(
      `<row r="${r}" ht="18" customHeight="1">${[
        strCell(`A${r}`, rec.fecha, 36),
        strCell(`B${r}`, diaSemana, 36),
        strCell(`C${r}`, rec.codigo_equipo, 36),
        strCell(`D${r}`, rec.turno === 'DIA' ? 'DÍA' : 'NOCHE', 36),
        numCell(`E${r}`, rec.horometro_inicial, 38),
        numCell(`F${r}`, rec.horometro_final, 38),
        formulaCell(`G${r}`, `ROUND(F${r}-E${r},1)`, rec.horas_trabajadas, 38),
        numCell(`H${r}`, rec.material_fracturado, 37),
        formulaCell(`I${r}`, `IF(G${r}>0,ROUND(H${r}/G${r},2),0)`, rec.rendimiento, 39),
        strCell(`J${r}`, rec.area_trabajo || '', 36),
        strCell(`K${r}`, rec.operador || '', 36),
        strCell(`L${r}`, rec.observaciones || '', 36),
        strCell(`M${r}`, String(rec.estado || 'pendiente').toUpperCase(), 36),
        strCell(`N${r}`, rec.hora_peru || '', 36)
      ].join('')}</row>`
    );
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols>${colsXml}</cols>
  <sheetData>${rowsXml.join('')}</sheetData>
</worksheet>`;
}

function buildStylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="4">
    <numFmt numFmtId="164" formatCode="0.0"/>
    <numFmt numFmtId="165" formatCode="#,##0"/>
    <numFmt numFmtId="166" formatCode="0.00"/>
    <numFmt numFmtId="167" formatCode="0.0%"/>
  </numFmts>
  <fonts count="8">
    <font><sz val="10"/><color rgb="FF0F172A"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FF0F172A"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FF000000"/><name val="Calibri"/></font>
    <font><sz val="9"/><color rgb="FF1E293B"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FF002B5C"/><name val="Calibri"/></font>
  </fonts>
  <fills count="21">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE40046"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1B1F4B"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFC000"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFC65911"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF375623"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF404040"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF4B084"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF9BC2E6"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFC6E0B4"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE0F2FE"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFDCFCE7"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFEF3C7"/></patternFill></fill>
  </fills>
  <borders count="3">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FF94A3B8"/></left>
      <right style="thin"><color rgb="FF94A3B8"/></right>
      <top style="thin"><color rgb="FF94A3B8"/></top>
      <bottom style="thin"><color rgb="FF94A3B8"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FF1E3A5F"/></left>
      <right style="thin"><color rgb="FF1E3A5F"/></right>
      <top style="thin"><color rgb="FF1E3A5F"/></top>
      <bottom style="thin"><color rgb="FF1E3A5F"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="49">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="3" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="3" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="5" fillId="4" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="8" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="17" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="17" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="10" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="13" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="12" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="13" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="15" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="14" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="16" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="15" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="13" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="9" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="1" fillId="10" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="11" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="1" fillId="12" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="165" fontId="0" fillId="13" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="166" fontId="0" fillId="14" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="6" fillId="17" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="164" fontId="1" fillId="15" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="165" fontId="1" fillId="14" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="166" fontId="1" fillId="14" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="16" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="165" fontId="0" fillId="16" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="165" fontId="1" fillId="13" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="166" fontId="1" fillId="15" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="4" fillId="3" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="17" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="165" fontId="0" fillId="17" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="17" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="166" fontId="0" fillId="17" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="7" fillId="15" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="164" fontId="7" fillId="15" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="165" fontId="7" fillId="15" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="166" fontId="7" fillId="15" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="167" fontId="1" fillId="18" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="18" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="18" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="167" fontId="1" fillId="19" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="167" fontId="1" fillId="20" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  </cellXfs>
</styleSheet>`;
}

const CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32Bytes(uint8Arr) {
  let crc = 0xffffffff;
  for (let i = 0; i < uint8Arr.length; i++) {
    crc = CRC32_TABLE[(crc ^ uint8Arr[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function createZipBlobFromEntries(entries) {
  const encoder = new TextEncoder();
  const localFileChunks = [];
  const centralDirChunks = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBytes = encoder.encode(entry.name);
    const dataBytes = entry.content instanceof Uint8Array ? entry.content : encoder.encode(entry.content);
    const crc = crc32Bytes(dataBytes);
    const size = dataBytes.length;

    const localHeader = new Uint8Array(30 + nameBytes.length);
    const lv = new DataView(localHeader.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true);
    lv.setUint16(6, 0x0800, true);
    lv.setUint16(8, 0, true);
    lv.setUint16(10, 0, true);
    lv.setUint16(12, 0x5772, true);
    lv.setUint32(14, crc, true);
    lv.setUint32(18, size, true);
    lv.setUint32(22, size, true);
    lv.setUint16(26, nameBytes.length, true);
    lv.setUint16(28, 0, true);
    localHeader.set(nameBytes, 30);
    localFileChunks.push(localHeader, dataBytes);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(centralHeader.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true);
    cv.setUint16(6, 20, true);
    cv.setUint16(8, 0x0800, true);
    cv.setUint16(10, 0, true);
    cv.setUint16(12, 0, true);
    cv.setUint16(14, 0x5772, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, size, true);
    cv.setUint32(24, size, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint16(30, 0, true);
    cv.setUint16(32, 0, true);
    cv.setUint16(34, 0, true);
    cv.setUint16(36, 0, true);
    cv.setUint32(38, 0, true);
    cv.setUint32(42, offset, true);
    centralHeader.set(nameBytes, 46);
    centralDirChunks.push(centralHeader);

    offset += localHeader.length + dataBytes.length;
  }

  const centralDirSize = centralDirChunks.reduce((acc, c) => acc + c.length, 0);
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(4, 0, true);
  ev.setUint16(6, 0, true);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  ev.setUint32(12, centralDirSize, true);
  ev.setUint32(16, offset, true);
  ev.setUint16(20, 0, true);

  return new Blob([...localFileChunks, ...centralDirChunks, eocd], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
}

export function generateMultiSheetWorkbookBlob(sheets = [], { drawings = [], charts = [] } = {}) {
  const sheetOverrides = sheets
    .map(
      (_, i) =>
        `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    )
    .join('\n  ');

  const drawingOverrides = drawings
    .map(
      (_, i) =>
        `<Override PartName="/xl/drawings/drawing${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`
    )
    .join('\n  ');

  const chartOverrides = charts
    .map(
      (_, i) =>
        `<Override PartName="/xl/charts/chart${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`
    )
    .join('\n  ');

  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${sheetOverrides}
  ${drawingOverrides}
  ${chartOverrides}
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;

  const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  const sheetsTagXml = sheets
    .map((s, i) => `<sheet name="${escapeXml(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`)
    .join('\n    ');

  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    ${sheetsTagXml}
  </sheets>
</workbook>`;

  const sheetRelsXml = sheets
    .map(
      (_, i) =>
        `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`
    )
    .join('\n  ');

  const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${sheetRelsXml}
  <Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

  const entries = [
    { name: '[Content_Types].xml', content: contentTypesXml },
    { name: '_rels/.rels', content: rootRelsXml },
    { name: 'xl/workbook.xml', content: workbookXml },
    { name: 'xl/_rels/workbook.xml.rels', content: workbookRelsXml },
    { name: 'xl/styles.xml', content: buildStylesXml() }
  ];

  sheets.forEach((s, i) => {
    entries.push({ name: `xl/worksheets/sheet${i + 1}.xml`, content: s.xml });
    if (s.drawingNum) {
      entries.push({
        name: `xl/worksheets/_rels/sheet${i + 1}.xml.rels`,
        content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${s.drawingNum}.xml"/>
</Relationships>`
      });
    }
  });

  drawings.forEach((d, i) => {
    entries.push({ name: `xl/drawings/drawing${i + 1}.xml`, content: d.xml });
    entries.push({ name: `xl/drawings/_rels/drawing${i + 1}.xml.rels`, content: d.relsXml });
  });

  charts.forEach((cXml, i) => {
    entries.push({ name: `xl/charts/chart${i + 1}.xml`, content: cXml });
  });

  return createZipBlobFromEntries(entries);
}

function triggerDownload(blob, fileName) {
  if (typeof document === 'undefined') return;
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

export function exportProductionWorkbookToExcel({
  records = [],
  codigoEquipo = 'STREX02',
  fechaDesde,
  fechaHasta,
  tituloPersonalizado = '',
  metaDiaria,
  metaMensual,
  includeAllEquipments = false,
  skipBrowserDownload = false
} = {}) {
  const targets = getOperativeTargets();
  const targetDaily = Number(metaDiaria) > 0 ? Number(metaDiaria) : targets.metaDiaria;
  const targetMonthly = Number(metaMensual) > 0 ? Number(metaMensual) : targets.metaMensual;

  const desde = fechaDesde || getFechaHoyPeru();
  const hasta = fechaHasta || getFechaHoyPeru();
  const isConsolidated = includeAllEquipments || !codigoEquipo || codigoEquipo === 'TODOS';

  const rangeRecords = records.filter((r) => !r.eliminado && r.fecha >= desde && r.fecha <= hasta);
  const filteredRecords = isConsolidated ? rangeRecords : rangeRecords.filter((r) => r.codigo_equipo === codigoEquipo);

  const compModel = buildMonthlyComparative(filteredRecords);
  const usedNames = new Set();
  const sheets = [];

  if (isConsolidated) {
    sheets.push({
      name: sanitizeSheetName('Comparativo Mensual', usedNames),
      xml: buildComparativeSheetXml(compModel)
    });
    const equiposPresentes = Array.from(new Set(rangeRecords.map((r) => r.codigo_equipo))).sort();
    const targetList = equiposPresentes.length > 0 ? equiposPresentes : ['STREX01', 'STREX02'];
    for (const eq of targetList) {
      const matrix = buildMatrixModel(rangeRecords, {
        codigoEquipo: eq,
        fechaDesde: desde,
        fechaHasta: hasta,
        metaDiaria: targetDaily,
        metaMensual: targetMonthly
      });
      sheets.push({
        name: sanitizeSheetName(`Control ${eq}`, usedNames),
        xml: buildMatrixSheetXml(matrix, { tituloPersonalizado })
      });
    }
    const allMatrix = buildMatrixModel(rangeRecords, {
      codigoEquipo: '',
      fechaDesde: desde,
      fechaHasta: hasta,
      metaDiaria: targetDaily,
      metaMensual: targetMonthly
    });
    sheets.push({
      name: sanitizeSheetName('Resumen por Área', usedNames),
      xml: buildAreaSummarySheetXml(allMatrix.areaBreakdown, 'TODOS', desde, hasta)
    });
  } else {
    const matrix = buildMatrixModel(filteredRecords, {
      codigoEquipo,
      fechaDesde: desde,
      fechaHasta: hasta,
      metaDiaria: targetDaily,
      metaMensual: targetMonthly
    });
    sheets.push({
      name: sanitizeSheetName(`Control ${codigoEquipo}`, usedNames),
      xml: buildMatrixSheetXml(matrix, { tituloPersonalizado })
    });
    sheets.push({
      name: sanitizeSheetName('Comparativo Mensual', usedNames),
      xml: buildComparativeSheetXml(compModel)
    });
    sheets.push({
      name: sanitizeSheetName('Resumen por Área', usedNames),
      xml: buildAreaSummarySheetXml(matrix.areaBreakdown, codigoEquipo, desde, hasta)
    });
  }

  sheets.push({
    name: sanitizeSheetName('Base de Datos', usedNames),
    xml: buildDatabaseSheetXml(filteredRecords)
  });

  const blob = generateMultiSheetWorkbookBlob(sheets);
  const tagEq = isConsolidated ? 'CONSOLIDADO_FLOTA' : codigoEquipo;
  const fileName = `Control_Fractura_Roca_${tagEq}_${desde}_al_${hasta}.xlsx`;

  if (!skipBrowserDownload) triggerDownload(blob, fileName);

  return {
    fileName,
    sheetNames: sheets.map((s) => s.name),
    totalRecords: filteredRecords.length,
    blobSizeBytes: blob.size,
    blob
  };
}

/**
 * Exportador Gerencial Completo con Gráficos Nativos OpenXML:
 *  - Resumen Gerencial (4 tablas + 2 gráficos nativos)
 *  - Resumen Semanal (Domingo a Sábado + 1 gráfico nativo)
 *  - Detalle Diario
 *  - Control <EQUIPO> (último período activo)
 *  - Base de Datos
 */
export function exportExecutiveWorkbookToExcel({
  records = [],
  fechaDesde = '',
  fechaHasta = '',
  tituloPersonalizado = '',
  skipBrowserDownload = false
} = {}) {
  const filtered = records.filter((r) => {
    if (r.eliminado) return false;
    if (fechaDesde && r.fecha < fechaDesde) return false;
    if (fechaHasta && r.fecha > fechaHasta) return false;
    return true;
  });

  const execModel = buildExecutiveDashboardModel(filtered);
  const usedNames = new Set();
  const sheets = [];
  const drawings = [];
  const charts = [];

  const hasChartSeries = execModel.monthlyChartSeries.length > 0;
  const summarySheetName = sanitizeSheetName('Resumen Gerencial', usedNames);
  const summaryInfo = buildExecutiveSummarySheetXml(execModel, {
    tituloPersonalizado,
    hasDrawing: hasChartSeries
  });

  if (hasChartSeries) {
    const cats = execModel.monthlyChartSeries.map((x) => `${x.mesNombre.slice(0, 3)} ${x.codigoEquipo}`);
    const catRef = `'${summarySheetName}'!$A$${summaryInfo.auxStartRow}:$A$${summaryInfo.auxEndRow}`;
    const c1 = buildColumnBarChartXml({
      title: 'Horas Máquina vs Meta Mensual (Hr)',
      grouping: 'clustered',
      categories: cats,
      catFormula: catRef,
      series: [
        {
          name: 'Horas Máquina (Hr)',
          nameFormula: `'${summarySheetName}'!$B$${summaryInfo.auxHeaderRow}`,
          valFormula: `'${summarySheetName}'!$B$${summaryInfo.auxStartRow}:$B$${summaryInfo.auxEndRow}`,
          values: execModel.monthlyChartSeries.map((x) => x.horasTotal),
          colorHex: '1B1F4B'
        },
        {
          name: 'Meta Mensual (Hr)',
          nameFormula: `'${summarySheetName}'!$C$${summaryInfo.auxHeaderRow}`,
          valFormula: `'${summarySheetName}'!$C$${summaryInfo.auxStartRow}:$C$${summaryInfo.auxEndRow}`,
          values: execModel.monthlyChartSeries.map((x) => x.metaMensual),
          colorHex: 'E40046'
        }
      ]
    });
    const c2 = buildColumnBarChartXml({
      title: 'Material Fracturado por Turno (Día + Noche m³)',
      grouping: 'stacked',
      categories: cats,
      catFormula: catRef,
      series: [
        {
          name: 'Material Día (m³)',
          nameFormula: `'${summarySheetName}'!$D$${summaryInfo.auxHeaderRow}`,
          valFormula: `'${summarySheetName}'!$D$${summaryInfo.auxStartRow}:$D$${summaryInfo.auxEndRow}`,
          values: execModel.monthlyChartSeries.map((x) => x.materialDia),
          colorHex: '1B1F4B'
        },
        {
          name: 'Material Noche (m³)',
          nameFormula: `'${summarySheetName}'!$E$${summaryInfo.auxHeaderRow}`,
          valFormula: `'${summarySheetName}'!$E$${summaryInfo.auxStartRow}:$E$${summaryInfo.auxEndRow}`,
          values: execModel.monthlyChartSeries.map((x) => x.materialNoche),
          colorHex: 'E40046'
        }
      ]
    });
    charts.push(c1, c2);
    drawings.push({
      xml: buildDrawingXml([
        { fromCol: 10, fromRow: 2, toCol: 19, toRow: 18, name: 'Grafico Horas vs Meta' },
        { fromCol: 10, fromRow: 20, toCol: 19, toRow: 36, name: 'Grafico Material Turnos' }
      ]),
      relsXml: buildDrawingRelsXml([1, 2])
    });
    sheets.push({ name: summarySheetName, xml: summaryInfo.xml, drawingNum: 1 });
  } else {
    sheets.push({ name: summarySheetName, xml: summaryInfo.xml });
  }

  // Hoja 2: Resumen Semanal (Domingo a Sábado) + Gráfico 3
  const hasWeeks = execModel.weeklyTable.length > 0;
  const weeklySheetName = sanitizeSheetName('Resumen Semanal', usedNames);
  const weeklyInfo = buildWeeklySheetXml(execModel, { hasDrawing: hasWeeks });
  if (hasWeeks) {
    const chartIdx = charts.length + 1;
    const drawingIdx = drawings.length + 1;
    const wCats = execModel.weeklyTable.map((w) => w.semanaLabel);
    const wCatRef = `'${weeklySheetName}'!$C$${weeklyInfo.startRow}:$C$${weeklyInfo.endRow}`;
    const c3 = buildColumnBarChartXml({
      title: 'Evolución Semanal — Material Fracturado Total (m³)',
      grouping: 'clustered',
      categories: wCats,
      catFormula: wCatRef,
      series: [
        {
          name: 'Material Total (m³)',
          nameFormula: `'${weeklySheetName}'!$${weeklyInfo.matTotColLetter}$3`,
          valFormula: `'${weeklySheetName}'!$${weeklyInfo.matTotColLetter}$${weeklyInfo.startRow}:$${weeklyInfo.matTotColLetter}$${weeklyInfo.endRow}`,
          values: execModel.weeklyTable.map((w) => w.materialTotal),
          colorHex: '1B1F4B'
        }
      ]
    });
    charts.push(c3);
    drawings.push({
      xml: buildDrawingXml([
        {
          fromCol: weeklyInfo.totalCols + 1,
          fromRow: 2,
          toCol: weeklyInfo.totalCols + 11,
          toRow: 22,
          name: 'Grafico Semanal Material'
        }
      ]),
      relsXml: buildDrawingRelsXml([chartIdx])
    });
    sheets.push({ name: weeklySheetName, xml: weeklyInfo.xml, drawingNum: drawingIdx });
  } else {
    sheets.push({ name: weeklySheetName, xml: weeklyInfo.xml });
  }

  // Hoja 3: Detalle Diario
  sheets.push({
    name: sanitizeSheetName('Detalle Diario', usedNames),
    xml: buildDailyDetailSheetXml(execModel)
  });

  // Hojas 4..N: Control último período por cada equipo
  const dates = filtered.map((r) => r.fecha).sort();
  const minDate = fechaDesde || dates[0] || getFechaHoyPeru();
  const maxDate = fechaHasta || dates[dates.length - 1] || getFechaHoyPeru();
  const lastPeriod = getValuationPeriodRange(maxDate);
  const targets = getOperativeTargets();

  for (const eq of execModel.equipos) {
    const matrix = buildMatrixModel(filtered, {
      codigoEquipo: eq,
      fechaDesde: lastPeriod.desde,
      fechaHasta: maxDate > lastPeriod.hasta ? maxDate : lastPeriod.hasta,
      metaDiaria: targets.metaDiaria,
      metaMensual: targets.metaMensual
    });
    sheets.push({
      name: sanitizeSheetName(`Control ${eq}`, usedNames),
      xml: buildMatrixSheetXml(matrix, { tituloPersonalizado })
    });
  }

  // Hoja final: Base de Datos
  sheets.push({
    name: sanitizeSheetName('Base de Datos', usedNames),
    xml: buildDatabaseSheetXml(filtered)
  });

  const blob = generateMultiSheetWorkbookBlob(sheets, { drawings, charts });
  const fileName = `Reporte_Gerencial_Fractura_Roca_${minDate}_al_${maxDate}.xlsx`;

  if (!skipBrowserDownload) triggerDownload(blob, fileName);

  return {
    fileName,
    sheetNames: sheets.map((s) => s.name),
    chartCount: charts.length,
    totalRecords: filtered.length,
    blobSizeBytes: blob.size,
    blob
  };
}
