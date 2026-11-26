/**
 * Vista 4 — Tablero Gerencial Acumulado STRACON (Colores Corporativos #1B1F4B y #E40046)
 * Incluye las 3 gráficas oficiales de Gerencia + Horas vs Meta + Evolución Semanal (Dom-Sáb).
 */

import {
  buildExecutiveDashboardModel,
  getAllMonthlyTargets,
  saveMonthlyTarget,
  getDataDateRange
} from '../services/productionService.js';
import { exportExecutiveWorkbookToExcel } from '../services/excelExporter.js';
import { escapeHtml } from '../utils/codeNormalizer.js';

const STRACON_NAVY = '#1B1F4B';
const STRACON_CRIMSON = '#E40046';

export class GerenciaDashboardView {
  constructor(mountEl, app) {
    this.mountEl = mountEl;
    this.app = app;
    this.fechaDesde = '';
    this.fechaHasta = '';
    this.weeklyMetric = 'material'; // 'material' | 'horas'
    this.showTargetsEditor = false;
  }

  async mount() {
    this.render();
  }

  getEquipMarker(codigoEquipo = '', equipList = []) {
    const idx = equipList.indexOf(codigoEquipo);
    if (idx === 0) return '●';
    if (idx === 1) return '▲';
    if (idx === 2) return '■';
    return '◆';
  }

  /**
   * Dibuja el eje X jerárquico idéntico a las gráficas gerenciales STRACON:
   *  - Fila superior: ● STREX01 / ▲ STREX02
   *  - Fila inferior: Mes centrado (JULIO, AGOSTO, SETIEMBRE, OCTUBRE, NOVIEMBRE)
   *  - Líneas divisorias verticales separando cada mes
   */
  renderTwoTierAxisXml(series = [], equList = [], padL, plotW, plotH, padT) {
    if (series.length === 0) return '';
    const slotW = plotW / series.length;
    const baseY = padT + plotH;
    const parts = [];

    // Borde inferior e izquierdo/derecho del área del eje
    parts.push(`<line x1="${padL}" y1="${baseY}" x2="${padL + plotW}" y2="${baseY}" stroke="#CBD5E1" stroke-width="1.2"/>`);
    parts.push(`<line x1="${padL}" y1="${baseY}" x2="${padL}" y2="${baseY + 46}" stroke="#CBD5E1" stroke-width="1.2"/>`);

    // Agrupar spans contiguos por mes
    const monthSpans = [];
    series.forEach((s, idx) => {
      const last = monthSpans[monthSpans.length - 1];
      if (last && last.mesNombre === s.mesNombre) {
        last.endIdx = idx;
      } else {
        monthSpans.push({ mesNombre: s.mesNombre, startIdx: idx, endIdx: idx });
      }
      const cx = padL + (idx + 0.5) * slotW;
      const mark = this.getEquipMarker(s.codigoEquipo, equList);
      parts.push(
        `<text x="${cx.toFixed(1)}" y="${baseY + 18}" text-anchor="middle" font-size="10.5" font-weight="800" fill="${STRACON_NAVY}">${mark} ${escapeHtml(
          s.codigoEquipo
        )}</text>`
      );
    });

    monthSpans.forEach((span) => {
      const xStart = padL + span.startIdx * slotW;
      const xEnd = padL + (span.endIdx + 1) * slotW;
      const xMid = (xStart + xEnd) / 2;
      parts.push(`<line x1="${xEnd.toFixed(1)}" y1="${baseY}" x2="${xEnd.toFixed(1)}" y2="${baseY + 46}" stroke="#CBD5E1" stroke-width="1.2"/>`);
      parts.push(
        `<text x="${xMid.toFixed(1)}" y="${baseY + 38}" text-anchor="middle" font-size="10.5" font-weight="800" fill="${STRACON_NAVY}">${escapeHtml(
          span.mesNombre
        )}</text>`
      );
    });

    return parts.join('');
  }

  /**
   * Gráfica 1 (Oficial Gerencia): Horas Máquina por Turno (Día vs Noche) — por Mes y Excavadora
   */
  renderChartTurnosHoras(series = [], equList = []) {
    if (series.length === 0) return '<p>Sin datos para graficar.</p>';
    const slotW = 92;
    const padL = 62;
    const padR = 24;
    const padT = 26;
    const padB = 54;
    const plotW = Math.max(520, series.length * slotW);
    const plotH = 210;
    const svgW = padL + plotW + padR;
    const svgH = padT + plotH + padB;

    const rawMax = Math.max(50, ...series.map((s) => s.horasTotal));
    const step = rawMax <= 200 ? 50 : rawMax <= 400 ? 50 : 100;
    const maxVal = Math.ceil((rawMax * 1.1) / step) * step;
    const ticks = [];
    for (let v = 0; v <= maxVal; v += step) ticks.push(v);

    const grid = ticks
      .map((val) => {
        const y = padT + plotH - (val / maxVal) * plotH;
        return `
          <line x1="${padL}" y1="${y.toFixed(1)}" x2="${padL + plotW}" y2="${y.toFixed(1)}" stroke="#D1D5DB" stroke-width="1"/>
          <text x="${padL - 8}" y="${(y + 3.5).toFixed(1)}" text-anchor="end" font-size="10" font-weight="600" fill="#4B5563">${val.toFixed(
          1
        )}</text>`;
      })
      .join('');

    const slot = plotW / series.length;
    const bw = Math.min(54, slot * 0.56);

    const bars = series
      .map((s, idx) => {
        const cx = padL + (idx + 0.5) * slot;
        const x = cx - bw / 2;
        const hDia = (s.horasDia / maxVal) * plotH;
        const hNoche = (s.horasNoche / maxVal) * plotH;
        const yDia = padT + plotH - hDia;
        const yNoche = yDia - hNoche;
        return `
          <g>
            ${
              hDia > 0
                ? `<rect x="${x.toFixed(1)}" y="${yDia.toFixed(1)}" width="${bw.toFixed(1)}" height="${hDia.toFixed(1)}" fill="${STRACON_NAVY}"><title>${s.mesNombre} ${s.codigoEquipo} Día: ${s.horasDia.toFixed(1)} Hr</title></rect>`
                : ''
            }
            ${
              hDia >= 16
                ? `<text x="${cx.toFixed(1)}" y="${(yDia + hDia / 2 + 4).toFixed(1)}" text-anchor="middle" font-size="10.5" font-weight="800" fill="#FFFFFF">${s.horasDia.toFixed(
                    1
                  )}</text>`
                : ''
            }
            ${
              hNoche > 0
                ? `<rect x="${x.toFixed(1)}" y="${yNoche.toFixed(1)}" width="${bw.toFixed(1)}" height="${hNoche.toFixed(1)}" fill="${STRACON_CRIMSON}"><title>${s.mesNombre} ${s.codigoEquipo} Noche: ${s.horasNoche.toFixed(1)} Hr</title></rect>`
                : ''
            }
            ${
              hNoche >= 16
                ? `<text x="${cx.toFixed(1)}" y="${(yNoche + hNoche / 2 + 4).toFixed(1)}" text-anchor="middle" font-size="10.5" font-weight="800" fill="#FFFFFF">${s.horasNoche.toFixed(
                    1
                  )}</text>`
                : ''
            }
            <text x="${cx.toFixed(1)}" y="${Math.max(13, yNoche - 7).toFixed(1)}" text-anchor="middle" font-size="11" font-weight="900" fill="#111827">${s.horasTotal.toFixed(
          1
        )}</text>
          </g>`;
      })
      .join('');

    const axisX = this.renderTwoTierAxisXml(series, equList, padL, plotW, plotH, padT);

    return `
      <div class="stracon-chart-card">
        <h3 class="stracon-chart-title">Horas Máquina por Turno (Día vs Noche) — por Mes y Excavadora</h3>
        <svg viewBox="0 0 ${svgW} ${svgH}" style="width:100%; min-width:380px; height:auto; display:block;">
          <text x="16" y="${(padT + plotH / 2).toFixed(1)}" text-anchor="middle" font-size="10.5" font-weight="700" fill="#4B5563" transform="rotate(-90 16 ${(
      padT +
      plotH / 2
    ).toFixed(1)})">Horas (Hr)</text>
          ${grid}
          ${bars}
          ${axisX}
        </svg>
        <div class="stracon-chart-legend">
          <span><span class="legend-dot" style="background:${STRACON_NAVY};"></span>Turno Día (Hr)</span>
          <span><span class="legend-dot" style="background:${STRACON_CRIMSON};"></span>Turno Noche (Hr)</span>
        </div>
      </div>`;
  }

  /**
   * Gráfica 2 (Oficial Gerencia): Material Fracturado y Rendimiento por Mes y Excavadora
   */
  renderChartMaterialRendimiento(series = [], equList = []) {
    if (series.length === 0) return '<p>Sin datos para graficar.</p>';
    const slotW = 92;
    const padL = 68;
    const padR = 54;
    const padT = 26;
    const padB = 54;
    const plotW = Math.max(520, series.length * slotW);
    const plotH = 210;
    const svgW = padL + plotW + padR;
    const svgH = padT + plotH + padB;

    const rawMaxMat = Math.max(1000, ...series.map((s) => s.materialTotal));
    const stepMat = rawMaxMat <= 4000 ? 1000 : 2000;
    const maxMat = Math.ceil((rawMaxMat * 1.08) / stepMat) * stepMat;
    const rawMaxRend = Math.max(14, ...series.map((s) => s.rendTotal));
    const maxRend = Math.ceil((rawMaxRend * 1.15) / 4) * 4;

    const stepsCount = 4;
    const grid = Array.from({ length: stepsCount + 1 }, (_, i) => {
      const f = i / stepsCount;
      const y = padT + plotH * (1 - f);
      const valMat = maxMat * f;
      const valRend = maxRend * f;
      return `
        <line x1="${padL}" y1="${y.toFixed(1)}" x2="${padL + plotW}" y2="${y.toFixed(1)}" stroke="#D1D5DB" stroke-width="1"/>
        <text x="${padL - 8}" y="${(y + 3.5).toFixed(1)}" text-anchor="end" font-size="10" font-weight="600" fill="#4B5563">${valMat.toLocaleString(
        'en-US',
        { minimumFractionDigits: 1, maximumFractionDigits: 1 }
      )}</text>
        <text x="${(padL + plotW + 8).toFixed(1)}" y="${(y + 3.5).toFixed(1)}" text-anchor="start" font-size="10" font-weight="600" fill="#4B5563">${valRend.toFixed(
        1
      )}</text>`;
    }).join('');

    const slot = plotW / series.length;
    const bw = Math.min(48, slot * 0.5);

    const bars = series
      .map((s, idx) => {
        const cx = padL + (idx + 0.5) * slot;
        const x = cx - bw / 2;
        const h = (s.materialTotal / maxMat) * plotH;
        const y = padT + plotH - h;
        const matLabel = s.materialTotal.toLocaleString('en-US', {
          minimumFractionDigits: 1,
          maximumFractionDigits: 1
        });
        return `
          <g>
            <rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${bw.toFixed(1)}" height="${h.toFixed(1)}" fill="${STRACON_NAVY}">
              <title>${s.mesNombre} ${s.codigoEquipo}: ${matLabel} m³</title>
            </rect>
            <text x="${cx.toFixed(1)}" y="${Math.max(13, y - 7).toFixed(1)}" text-anchor="middle" font-size="10.5" font-weight="900" fill="#111827">${matLabel}</text>
          </g>`;
      })
      .join('');

    const pts = series.map((s, idx) => {
      const cx = padL + (idx + 0.5) * slot;
      const cy = padT + plotH - (s.rendTotal / maxRend) * plotH;
      return { cx, cy, rend: s.rendTotal, label: `${s.mesNombre} ${s.codigoEquipo}` };
    });
    const poly = pts.map((p) => `${p.cx.toFixed(1)},${p.cy.toFixed(1)}`).join(' ');
    const dots = pts
      .map(
        (p) => `
        <g>
          <circle cx="${p.cx.toFixed(1)}" cy="${p.cy.toFixed(1)}" r="5" fill="${STRACON_CRIMSON}" stroke="#FFFFFF" stroke-width="1.2">
            <title>${p.label}: ${p.rend.toFixed(2)} m³/HM</title>
          </circle>
          <text x="${(p.cx + 9).toFixed(1)}" y="${(p.cy + 4).toFixed(1)}" text-anchor="start" font-size="10.5" font-weight="900" fill="${STRACON_CRIMSON}" stroke="#FFFFFF" stroke-width="2.6" paint-order="stroke">${p.rend.toFixed(
          1
        )}</text>
        </g>`
      )
      .join('');

    const axisX = this.renderTwoTierAxisXml(series, equList, padL, plotW, plotH, padT);

    return `
      <div class="stracon-chart-card">
        <h3 class="stracon-chart-title">Material Fracturado y Rendimiento por Mes y Excavadora</h3>
        <svg viewBox="0 0 ${svgW} ${svgH}" style="width:100%; min-width:380px; height:auto; display:block;">
          <text x="15" y="${(padT + plotH / 2).toFixed(1)}" text-anchor="middle" font-size="10.5" font-weight="700" fill="#4B5563" transform="rotate(-90 15 ${(
      padT +
      plotH / 2
    ).toFixed(1)})">Material Fracturado (m³)</text>
          <text x="${svgW - 12}" y="${(padT + plotH / 2).toFixed(1)}" text-anchor="middle" font-size="10.5" font-weight="700" fill="#4B5563" transform="rotate(-90 ${
      svgW - 12
    } ${(padT + plotH / 2).toFixed(1)})">Rendimiento (m³/HM)</text>
          ${grid}
          ${bars}
          <polyline fill="none" stroke="${STRACON_CRIMSON}" stroke-width="3.2" points="${poly}" />
          ${dots}
          ${axisX}
        </svg>
        <div class="stracon-chart-legend">
          <span><span class="legend-dot" style="background:${STRACON_NAVY};"></span>Material Fracturado (m³)</span>
          <span><span style="display:inline-flex; align-items:center; margin-right:0.35rem;"><span style="width:18px; height:3px; background:${STRACON_CRIMSON}; display:inline-block;"></span><span style="width:9px; height:9px; border-radius:999px; background:${STRACON_CRIMSON}; display:inline-block; margin-left:-13px;"></span></span>Rendimiento (m³/HM)</span>
        </div>
      </div>`;
  }

  /**
   * Gráfica 3 (Oficial Gerencia): Promedio de Horas Máquina por Día Operativo — Mes, Excavadora y Turno
   */
  renderChartPromediosDia(series = [], equList = []) {
    if (series.length === 0) return '<p>Sin datos para graficar.</p>';
    const slotW = 92;
    const padL = 54;
    const padR = 20;
    const padT = 26;
    const padB = 54;
    const plotW = Math.max(520, series.length * slotW);
    const plotH = 210;
    const svgW = padL + plotW + padR;
    const svgH = padT + plotH + padB;

    const maxVal = Math.max(12, Math.ceil((Math.max(...series.map((s) => Math.max(s.promDiaHr, s.promNocheHr))) + 1) / 2) * 2);
    const ticks = [];
    for (let v = 0; v <= maxVal; v += 2) ticks.push(v);

    const grid = ticks
      .map((val) => {
        const y = padT + plotH - (val / maxVal) * plotH;
        return `
          <line x1="${padL}" y1="${y.toFixed(1)}" x2="${padL + plotW}" y2="${y.toFixed(1)}" stroke="#D1D5DB" stroke-width="1"/>
          <text x="${padL - 8}" y="${(y + 3.5).toFixed(1)}" text-anchor="end" font-size="10" font-weight="600" fill="#4B5563">${val.toFixed(
          1
        )}</text>`;
      })
      .join('');

    const slot = plotW / series.length;
    const bw = Math.min(26, slot * 0.28);

    const bars = series
      .map((s, idx) => {
        const cx = padL + (idx + 0.5) * slot;
        const hD = (s.promDiaHr / maxVal) * plotH;
        const hN = (s.promNocheHr / maxVal) * plotH;
        const yD = padT + plotH - hD;
        const yN = padT + plotH - hN;
        const xD = cx - bw - 3;
        const xN = cx + 3;
        return `
          <g>
            ${hD > 0 ? `<rect x="${xD.toFixed(1)}" y="${yD.toFixed(1)}" width="${bw}" height="${hD.toFixed(1)}" fill="${STRACON_NAVY}"><title>Prom Día: ${s.promDiaHr.toFixed(2)} Hr/día</title></rect>` : ''}
            ${hN > 0 ? `<rect x="${xN.toFixed(1)}" y="${yN.toFixed(1)}" width="${bw}" height="${hN.toFixed(1)}" fill="${STRACON_CRIMSON}"><title>Prom Noche: ${s.promNocheHr.toFixed(2)} Hr/día</title></rect>` : ''}
            <text x="${(xD + bw / 2).toFixed(1)}" y="${Math.max(13, yD - 7).toFixed(1)}" text-anchor="middle" font-size="10.5" font-weight="900" fill="#111827">${s.promDiaHr.toFixed(
          1
        )}</text>
            <text x="${(xN + bw / 2).toFixed(1)}" y="${Math.max(13, yN - 7).toFixed(1)}" text-anchor="middle" font-size="10.5" font-weight="900" fill="#111827">${s.promNocheHr.toFixed(
          1
        )}</text>
          </g>`;
      })
      .join('');

    const axisX = this.renderTwoTierAxisXml(series, equList, padL, plotW, plotH, padT);

    return `
      <div class="stracon-chart-card">
        <h3 class="stracon-chart-title">Promedio de Horas Máquina por Día Operativo — Mes, Excavadora y Turno</h3>
        <svg viewBox="0 0 ${svgW} ${svgH}" style="width:100%; min-width:380px; height:auto; display:block;">
          <text x="16" y="${(padT + plotH / 2).toFixed(1)}" text-anchor="middle" font-size="10.5" font-weight="700" fill="#4B5563" transform="rotate(-90 16 ${(
      padT +
      plotH / 2
    ).toFixed(1)})">Hr/Día Operativo</text>
          ${grid}
          ${bars}
          ${axisX}
        </svg>
        <div class="stracon-chart-legend">
          <span><span class="legend-dot" style="background:${STRACON_NAVY};"></span>Prom. Turno Día (Hr/día)</span>
          <span><span class="legend-dot" style="background:${STRACON_CRIMSON};"></span>Prom. Turno Noche (Hr/día)</span>
        </div>
      </div>`;
  }

  /**
   * Gráfica 4: Horas Máquina vs Meta Mensual y % Cumplimiento (STRACON #1B1F4B + #E40046)
   */
  renderChartHorasVsMeta(series = [], equList = []) {
    if (series.length === 0) return '<p>Sin datos para graficar.</p>';
    const slotW = 92;
    const padL = 58;
    const padR = 20;
    const padT = 26;
    const padB = 54;
    const plotW = Math.max(520, series.length * slotW);
    const plotH = 210;
    const svgW = padL + plotW + padR;
    const svgH = padT + plotH + padB;
    const rawMax = Math.max(180, ...series.map((s) => Math.max(s.horasTotal, s.metaMensual)));
    const maxVal = Math.ceil((rawMax * 1.12) / 100) * 100;

    const grid = [0, 0.25, 0.5, 0.75, 1]
      .map((f) => {
        const y = padT + plotH * (1 - f);
        return `
          <line x1="${padL}" y1="${y.toFixed(1)}" x2="${padL + plotW}" y2="${y.toFixed(1)}" stroke="#D1D5DB" stroke-width="1"/>
          <text x="${padL - 8}" y="${(y + 3.5).toFixed(1)}" text-anchor="end" font-size="10" font-weight="600" fill="#4B5563">${(
          maxVal * f
        ).toFixed(0)}</text>`;
      })
      .join('');

    const slot = plotW / series.length;
    const bw = Math.min(26, slot * 0.28);

    const bars = series
      .map((s, idx) => {
        const cx = padL + (idx + 0.5) * slot;
        const hReal = (s.horasTotal / maxVal) * plotH;
        const hMeta = (s.metaMensual / maxVal) * plotH;
        const yReal = padT + plotH - hReal;
        const yMeta = padT + plotH - hMeta;
        const xReal = cx - bw - 3;
        const xMeta = cx + 3;
        const pct = Math.round(s.pctCumplimiento * 100);
        return `
          <g>
            <rect x="${xReal.toFixed(1)}" y="${yReal.toFixed(1)}" width="${bw}" height="${hReal.toFixed(1)}" fill="${STRACON_NAVY}">
              <title>${s.mesNombre} ${s.codigoEquipo}: ${s.horasTotal.toFixed(1)} Hr (${pct}%)</title>
            </rect>
            <rect x="${xMeta.toFixed(1)}" y="${yMeta.toFixed(1)}" width="${bw}" height="${hMeta.toFixed(1)}" fill="${STRACON_CRIMSON}" opacity="0.86">
              <title>Meta ${s.mesNombre}: ${s.metaMensual} Hr</title>
            </rect>
            <text x="${(xReal + bw / 2).toFixed(1)}" y="${Math.max(13, yReal - 6).toFixed(1)}" text-anchor="middle" font-size="9.5" font-weight="900" fill="${STRACON_NAVY}">${s.horasTotal.toFixed(
          0
        )}</text>
            <text x="${(xMeta + bw / 2).toFixed(1)}" y="${Math.max(13, yMeta - 6).toFixed(1)}" text-anchor="middle" font-size="9.5" font-weight="900" fill="${STRACON_CRIMSON}">${pct}%</text>
          </g>`;
      })
      .join('');

    const axisX = this.renderTwoTierAxisXml(series, equList, padL, plotW, plotH, padT);

    return `
      <div class="stracon-chart-card">
        <h3 class="stracon-chart-title">Horas Máquina vs Meta Mensual (% Cumplimiento)</h3>
        <svg viewBox="0 0 ${svgW} ${svgH}" style="width:100%; min-width:380px; height:auto; display:block;">
          <text x="16" y="${(padT + plotH / 2).toFixed(1)}" text-anchor="middle" font-size="10.5" font-weight="700" fill="#4B5563" transform="rotate(-90 16 ${(
      padT +
      plotH / 2
    ).toFixed(1)})">Horas (Hr)</text>
          ${grid}
          ${bars}
          ${axisX}
        </svg>
        <div class="stracon-chart-legend">
          <span><span class="legend-dot" style="background:${STRACON_NAVY};"></span>Horas Máquina (Hr)</span>
          <span><span class="legend-dot" style="background:${STRACON_CRIMSON};"></span>Meta Mensual (Hr) / % Cumpl.</span>
        </div>
      </div>`;
  }

  /**
   * Gráfica 5: Evolución Semanal (Domingo a Sábado) en estilo STRACON
   */
  renderChartSemanal(weeklyTable = [], equipos = []) {
    if (weeklyTable.length === 0) return '<p>Sin datos semanales.</p>';
    const isMat = this.weeklyMetric === 'material';
    const slotW = 48;
    const padL = 58;
    const padR = 50;
    const padT = 26;
    const padB = 46;
    const plotW = Math.max(660, weeklyTable.length * slotW);
    const plotH = 205;
    const svgW = padL + plotW + padR;
    const svgH = padT + plotH + padB;

    const palette = [STRACON_NAVY, STRACON_CRIMSON, '#475569', '#0284C7'];
    const totals = weeklyTable.map((w) => (isMat ? w.materialTotal : w.horasTotal));
    const maxMain = Math.max(isMat ? 500 : 50, ...totals) * 1.1;
    const maxRend = Math.max(16, ...weeklyTable.map((w) => w.rendimiento)) * 1.1;

    const grid = [0, 0.25, 0.5, 0.75, 1]
      .map((f) => {
        const y = padT + plotH * (1 - f);
        return `
          <line x1="${padL}" y1="${y.toFixed(1)}" x2="${padL + plotW}" y2="${y.toFixed(1)}" stroke="#D1D5DB" stroke-width="1"/>
          <text x="${padL - 8}" y="${(y + 3.5).toFixed(1)}" text-anchor="end" font-size="10" font-weight="600" fill="#4B5563">${Math.round(
          maxMain * f
        )}</text>
          <text x="${(padL + plotW + 8).toFixed(1)}" y="${(y + 3.5).toFixed(1)}" text-anchor="start" font-size="10" font-weight="600" fill="#4B5563">${(
          maxRend * f
        ).toFixed(1)}</text>`;
      })
      .join('');

    const bars = weeklyTable
      .map((w, idx) => {
        const cx = padL + (idx + 0.5) * (plotW / weeklyTable.length);
        const bw = Math.min(30, (plotW / weeklyTable.length) * 0.64);
        let yCursor = padT + plotH;
        const rects = equipos
          .map((eq, eIdx) => {
            const val = isMat ? w.byEquip[eq]?.material || 0 : w.byEquip[eq]?.horas || 0;
            if (val <= 0) return '';
            const h = (val / maxMain) * plotH;
            yCursor -= h;
            return `<rect x="${(cx - bw / 2).toFixed(1)}" y="${yCursor.toFixed(1)}" width="${bw.toFixed(1)}" height="${h.toFixed(
              1
            )}" fill="${palette[eIdx % palette.length]}"><title>${w.semanaLabel} (${eq}): ${val}</title></rect>`;
          })
          .join('');
        const totVal = isMat ? w.materialTotal : w.horasTotal;
        const shortSem = w.semanaLabel.split(' al ')[0];
        return `
          <g>
            ${rects}
            ${
              totVal > 0
                ? `<text x="${cx.toFixed(1)}" y="${Math.max(13, yCursor - 6).toFixed(1)}" text-anchor="middle" font-size="9" font-weight="900" fill="#111827">${
                    isMat ? Math.round(totVal) : totVal.toFixed(0)
                  }</text>`
                : ''
            }
            <text x="${cx.toFixed(1)}" y="${padT + plotH + 17}" text-anchor="middle" font-size="9" font-weight="700" fill="${STRACON_NAVY}">${escapeHtml(
          shortSem
        )}</text>
          </g>`;
      })
      .join('');

    const pts = weeklyTable.map((w, idx) => {
      const cx = padL + (idx + 0.5) * (plotW / weeklyTable.length);
      const cy = padT + plotH - (w.rendimiento / maxRend) * plotH;
      return { cx, cy, rend: w.rendimiento, label: w.semanaLabel };
    });
    const poly = pts.map((p) => `${p.cx.toFixed(1)},${p.cy.toFixed(1)}`).join(' ');
    const dots = pts
      .map(
        (p) =>
          `<circle cx="${p.cx.toFixed(1)}" cy="${p.cy.toFixed(1)}" r="4" fill="#F59E0B" stroke="${STRACON_NAVY}" stroke-width="1.4"><title>${p.label}: ${p.rend.toFixed(
            2
          )} m³/HM</title></circle>`
      )
      .join('');

    return `
      <div class="stracon-chart-card">
        <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:0.5rem; margin-bottom:0.4rem;">
          <h3 class="stracon-chart-title" style="margin:0; text-align:left;">Evolución Semanal (Domingo a Sábado) — ${weeklyTable.length} Semana(s)</h3>
          <div class="inline-row no-print">
            <button type="button" class="chip-btn ${isMat ? 'is-active' : ''}" data-wk-metric="material">Material (m³)</button>
            <button type="button" class="chip-btn ${!isMat ? 'is-active' : ''}" data-wk-metric="horas">Horas Máquina (Hr)</button>
          </div>
        </div>
        <svg viewBox="0 0 ${svgW} ${svgH}" style="width:100%; min-width:660px; height:auto; display:block;">
          <text x="15" y="${(padT + plotH / 2).toFixed(1)}" text-anchor="middle" font-size="10" font-weight="700" fill="#4B5563" transform="rotate(-90 15 ${(
      padT +
      plotH / 2
    ).toFixed(1)})">${isMat ? 'Material (m³)' : 'Horas (Hr)'}</text>
          <text x="${svgW - 12}" y="${(padT + plotH / 2).toFixed(1)}" text-anchor="middle" font-size="10" font-weight="700" fill="#4B5563" transform="rotate(-90 ${
      svgW - 12
    } ${(padT + plotH / 2).toFixed(1)})">Rendimiento (m³/HM)</text>
          ${grid}
          <line x1="${padL}" y1="${padT + plotH}" x2="${padL + plotW}" y2="${padT + plotH}" stroke="#CBD5E1" stroke-width="1.2"/>
          ${bars}
          <polyline fill="none" stroke="#F59E0B" stroke-width="2.6" points="${poly}" />
          ${dots}
        </svg>
        <div class="stracon-chart-legend">
          ${equipos
            .map(
              (eq, i) =>
                `<span><span class="legend-dot" style="background:${palette[i % palette.length]};"></span>${this.getEquipMarker(
                  eq,
                  equipos
                )} ${escapeHtml(eq)}</span>`
            )
            .join('')}
          <span><span class="legend-dot" style="background:#F59E0B; border-radius:999px;"></span>Rendimiento (m³/HM)</span>
        </div>
      </div>`;
  }

  render() {
    const model = buildExecutiveDashboardModel(this.app.registros, {
      fechaDesde: this.fechaDesde,
      fechaHasta: this.fechaHasta
    });
    const t = model.totals;
    const allTargets = getAllMonthlyTargets();
    const fmtPct = (v) => `${(Number(v || 0) * 100).toFixed(1)}%`;

    this.mountEl.innerHTML = `
      <section class="panel">
        <div class="panel-header">
          <div>
            <h2 class="panel-title">📈 Panel Gerencial Acumulado STRACON — ${model.equipos.join(' + ') || 'Flota'}</h2>
            <p style="margin:0.15rem 0 0; font-size:0.84rem; font-weight:700; color:var(--text-muted);">
              Consolidado histórico por mes de valorización (26 al 25), turno, semanas (Dom–Sáb) y cumplimiento de metas.
            </p>
          </div>
          <div class="inline-row no-print" style="flex-wrap:wrap;">
            <button type="button" class="btn-mine btn-sm btn-secondary" id="btn-toggle-metas-mes">🎯 Metas por Mes</button>
            <button type="button" class="btn-mine btn-sm btn-navy" id="btn-print-gerencia">🖨️ Imprimir / PDF</button>
            <button type="button" class="btn-mine btn-excel" id="btn-export-exec-xlsx">📥 Exportar Reporte Gerencial (.xlsx)</button>
          </div>
        </div>

        <div class="form-grid no-print" style="margin-bottom:0.85rem;">
          <div class="field-group">
            <label class="field-label" for="ger-desde">Filtrar Desde (opcional)</label>
            <input type="date" id="ger-desde" class="input-mine" value="${escapeHtml(this.fechaDesde)}" />
          </div>
          <div class="field-group">
            <label class="field-label" for="ger-hasta">Filtrar Hasta (opcional)</label>
            <input type="date" id="ger-hasta" class="input-mine" value="${escapeHtml(this.fechaHasta)}" />
          </div>
          <div class="field-group" style="justify-content:flex-end;">
            <button type="button" class="btn-mine btn-secondary" id="btn-ger-all-dates">📅 Ver Todo el Historial</button>
          </div>
        </div>

        ${
          this.showTargetsEditor
            ? `<div class="modal-section no-print" style="margin-bottom:0.9rem;">
                <strong>🎯 Metas Mensuales por Equipo (Hr/mes)</strong>
                <div class="form-grid" style="margin-top:0.5rem;">
                  ${model.monthsOrdered
                    .map(
                      (mKey) => `
                    <div class="field-group">
                      <label class="field-label">${escapeHtml(mKey)} (Hr)</label>
                      <input type="number" step="10" min="1" class="input-mine tabular" data-month-target="${escapeHtml(
                        mKey
                      )}" value="${allTargets[mKey] || 480}" />
                    </div>`
                    )
                    .join('')}
                </div>
              </div>`
            : ''
        }

        <div class="kpi-grid">
          <div class="kpi-card navy">
            <div class="kpi-label">Horas Máquina Acumuladas</div>
            <div class="kpi-value tabular">${t.horasTotal.toFixed(1)} Hr</div>
            <div class="kpi-sub">Meta: ${Math.round(t.metaMensualTotal).toLocaleString('es-PE')} Hr | Cumplimiento: <strong>${fmtPct(
      t.pctCumplimiento
    )}</strong></div>
          </div>
          <div class="kpi-card">
            <div class="kpi-label">Material Fracturado Acumulado</div>
            <div class="kpi-value tabular">${Math.round(t.materialTotal).toLocaleString('es-PE')} m³</div>
            <div class="kpi-sub">☀️ ${Math.round(t.materialDia).toLocaleString('es-PE')} m³ | 🌙 ${Math.round(
      t.materialNoche
    ).toLocaleString('es-PE')} m³</div>
          </div>
          <div class="kpi-card navy">
            <div class="kpi-label">Rendimiento Acumulado</div>
            <div class="kpi-value tabular">${t.rendTotal.toFixed(2)} m³/HM</div>
            <div class="kpi-sub">☀️ ${t.rendDia.toFixed(2)} | 🌙 ${t.rendNoche.toFixed(2)} m³/HM</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-label">Utilización Flota</div>
            <div class="kpi-value tabular">${t.hrPorDiaOperativo.toFixed(2)} Hr/día</div>
            <div class="kpi-sub">${t.diasOperativos} días-máq | Turno Noche: <strong>${fmtPct(t.pctNoche)}</strong></div>
          </div>
        </div>
      </section>

      <div class="exec-charts-grid" style="margin-bottom:1rem;">
        ${this.renderChartPromediosDia(model.monthlyChartSeries, model.equipos)}
        ${this.renderChartTurnosHoras(model.monthlyChartSeries, model.equipos)}
        ${this.renderChartMaterialRendimiento(model.monthlyChartSeries, model.equipos)}
        ${this.renderChartHorasVsMeta(model.monthlyChartSeries, model.equipos)}
      </div>

      <div style="margin-bottom:1rem;">
        ${this.renderChartSemanal(model.weeklyTable, model.equipos)}
      </div>

      <section class="panel">
        <div class="panel-header"><h2 class="panel-title">📋 Resumen Mensual por Equipo y Cumplimiento</h2></div>
        <div class="table-scroll">
          <table class="data-table">
            <thead>
              <tr>
                <th>Mes</th>
                <th>Equipo</th>
                <th>Horas Máquina (Hr)</th>
                <th>Meta Mensual (Hr)</th>
                <th>% Cumpl.</th>
                <th>Material (m³)</th>
                <th>Rend. (m³/HM)</th>
                <th>Días Op.</th>
                <th>Hr / Día Op.</th>
              </tr>
            </thead>
            <tbody>
              ${model.monthlyTable
                .map(
                  (r) => `
                <tr class="${r.rowType === 'subtotal' ? 'row-subtotal' : ''}">
                  <td><strong>${escapeHtml(r.mesNombre)}</strong></td>
                  <td><strong>${escapeHtml(r.codigoEquipo)}</strong></td>
                  <td class="tabular">${r.horasTotal.toFixed(1)}</td>
                  <td class="tabular">${Math.round(r.metaMensual)}</td>
                  <td class="tabular"><span class="badge-pill ${
                    r.pctCumplimiento >= 1 ? 'badge-sync-ok' : 'badge-sync-pending'
                  }">${fmtPct(r.pctCumplimiento)}</span></td>
                  <td class="tabular"><strong>${Math.round(r.materialTotal).toLocaleString('es-PE')}</strong></td>
                  <td class="tabular">${r.rendTotal.toFixed(2)}</td>
                  <td class="tabular">${r.diasOperativos}</td>
                  <td class="tabular">${r.hrPorDiaOperativo.toFixed(2)}</td>
                </tr>`
                )
                .join('')}
            </tbody>
            <tfoot>
              <tr>
                <td>TOTAL ACUMULADO</td>
                <td>${escapeHtml(model.equipos.join('+'))}</td>
                <td class="tabular">${t.horasTotal.toFixed(1)}</td>
                <td class="tabular">${Math.round(t.metaMensualTotal).toLocaleString('es-PE')}</td>
                <td class="tabular">${fmtPct(t.pctCumplimiento)}</td>
                <td class="tabular">${Math.round(t.materialTotal).toLocaleString('es-PE')}</td>
                <td class="tabular">${t.rendTotal.toFixed(2)}</td>
                <td class="tabular">${t.diasOperativos}</td>
                <td class="tabular">${t.hrPorDiaOperativo.toFixed(2)}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </section>

      <div class="two-col-grid">
        <section class="panel">
          <div class="panel-header"><h2 class="panel-title">🌓 Distribución Horas por Turno</h2></div>
          <div class="table-scroll">
            <table class="data-table">
              <thead>
                <tr><th>Mes</th><th>Equipo</th><th>Día (Hr)</th><th>Noche (Hr)</th><th>Total (Hr)</th><th>% Noche</th></tr>
              </thead>
              <tbody>
                ${model.monthlyTable
                  .map(
                    (r) => `
                  <tr class="${r.rowType === 'subtotal' ? 'row-subtotal' : ''}">
                    <td><strong>${escapeHtml(r.mesNombre)}</strong></td>
                    <td>${escapeHtml(r.codigoEquipo)}</td>
                    <td class="tabular">${r.horasDia.toFixed(1)}</td>
                    <td class="tabular">${r.horasNoche.toFixed(1)}</td>
                    <td class="tabular"><strong>${r.horasTotal.toFixed(1)}</strong></td>
                    <td class="tabular">${fmtPct(r.pctNoche)}</td>
                  </tr>`
                  )
                  .join('')}
              </tbody>
              <tfoot>
                <tr>
                  <td colspan="2">TOTAL ACUMULADO</td>
                  <td class="tabular">${t.horasDia.toFixed(1)}</td>
                  <td class="tabular">${t.horasNoche.toFixed(1)}</td>
                  <td class="tabular">${t.horasTotal.toFixed(1)}</td>
                  <td class="tabular">${fmtPct(t.pctNoche)}</td>
                </tr>
              </tfoot>
            </table>
          </div>
        </section>

        <section class="panel">
          <div class="panel-header"><h2 class="panel-title">🪨 Material y Rendimiento por Turno</h2></div>
          <div class="table-scroll">
            <table class="data-table">
              <thead>
                <tr><th>Mes</th><th>Equipo</th><th>Mat. Día</th><th>Mat. Noche</th><th>Mat. Total</th><th>Rend. Día</th><th>Rend. Noche</th><th>Prom. Día / Noche</th></tr>
              </thead>
              <tbody>
                ${model.monthlyTable
                  .map(
                    (r) => `
                  <tr class="${r.rowType === 'subtotal' ? 'row-subtotal' : ''}">
                    <td><strong>${escapeHtml(r.mesNombre)}</strong></td>
                    <td>${escapeHtml(r.codigoEquipo)}</td>
                    <td class="tabular">${Math.round(r.materialDia).toLocaleString('es-PE')}</td>
                    <td class="tabular">${Math.round(r.materialNoche).toLocaleString('es-PE')}</td>
                    <td class="tabular"><strong>${Math.round(r.materialTotal).toLocaleString('es-PE')}</strong></td>
                    <td class="tabular">${r.rendDia.toFixed(2)}</td>
                    <td class="tabular">${r.horasNoche > 0 ? r.rendNoche.toFixed(2) : '-'}</td>
                    <td class="tabular">${r.promDiaHr.toFixed(1)} / ${r.promNocheHr.toFixed(1)}</td>
                  </tr>`
                  )
                  .join('')}
              </tbody>
              <tfoot>
                <tr>
                  <td colspan="2">TOTAL ACUMULADO</td>
                  <td class="tabular">${Math.round(t.materialDia).toLocaleString('es-PE')}</td>
                  <td class="tabular">${Math.round(t.materialNoche).toLocaleString('es-PE')}</td>
                  <td class="tabular">${Math.round(t.materialTotal).toLocaleString('es-PE')}</td>
                  <td class="tabular">${t.rendDia.toFixed(2)}</td>
                  <td class="tabular">${t.rendNoche.toFixed(2)}</td>
                  <td class="tabular">${t.promDiaHr.toFixed(1)} / ${t.promNocheHr.toFixed(1)}</td>
                </tr>
              </tfoot>
            </table>
          </div>
        </section>
      </div>

      <section class="panel">
        <div class="panel-header"><h2 class="panel-title">🗓️ Resumen Semanal (Domingo a Sábado) por Equipo</h2></div>
        <div class="table-scroll">
          <table class="data-table">
            <thead>
              <tr>
                <th>Semana</th>
                ${model.equipos.map((e) => `<th>HM ${escapeHtml(e)}</th>`).join('')}
                <th>HM Total</th>
                ${model.equipos.map((e) => `<th>Mat. ${escapeHtml(e)}</th>`).join('')}
                <th>Material Total (m³)</th>
                <th>Rend. (m³/HM)</th>
                ${model.equipos.map((e) => `<th>Días ${escapeHtml(e)}</th>`).join('')}
              </tr>
            </thead>
            <tbody>
              ${model.weeklyTable
                .map(
                  (w) => `
                <tr>
                  <td><strong>${escapeHtml(w.semanaLabel)}</strong></td>
                  ${model.equipos.map((e) => `<td class="tabular">${(w.byEquip[e]?.horas || 0).toFixed(1)}</td>`).join('')}
                  <td class="tabular"><strong>${w.horasTotal.toFixed(1)}</strong></td>
                  ${model.equipos.map((e) => `<td class="tabular">${Math.round(w.byEquip[e]?.material || 0).toLocaleString('es-PE')}</td>`).join('')}
                  <td class="tabular"><strong>${Math.round(w.materialTotal).toLocaleString('es-PE')}</strong></td>
                  <td class="tabular">${w.rendimiento.toFixed(2)}</td>
                  ${model.equipos.map((e) => `<td class="tabular">${w.byEquip[e]?.diasOp || 0}</td>`).join('')}
                </tr>`
                )
                .join('')}
            </tbody>
            <tfoot>
              <tr>
                <td>TOTAL (${model.weeklyTable.length} sem.)</td>
                ${model.equipos.map((e) => `<td class="tabular">${(model.weeklyTotals.byEquip[e]?.horas || 0).toFixed(1)}</td>`).join('')}
                <td class="tabular">${model.weeklyTotals.horasTotal.toFixed(1)}</td>
                ${model.equipos.map((e) => `<td class="tabular">${Math.round(model.weeklyTotals.byEquip[e]?.material || 0).toLocaleString('es-PE')}</td>`).join('')}
                <td class="tabular">${Math.round(model.weeklyTotals.materialTotal).toLocaleString('es-PE')}</td>
                <td class="tabular">${model.weeklyTotals.rendimiento.toFixed(2)}</td>
                ${model.equipos.map((e) => `<td class="tabular">${model.weeklyTotals.byEquip[e]?.diasOp || 0}</td>`).join('')}
              </tr>
            </tfoot>
          </table>
        </div>
      </section>
    `;

    this.bindEvents();
  }

  bindEvents() {
    this.mountEl.querySelector('#ger-desde')?.addEventListener('change', (e) => {
      this.fechaDesde = e.target.value;
      this.render();
    });
    this.mountEl.querySelector('#ger-hasta')?.addEventListener('change', (e) => {
      this.fechaHasta = e.target.value;
      this.render();
    });
    this.mountEl.querySelector('#btn-ger-all-dates')?.addEventListener('click', () => {
      this.fechaDesde = '';
      this.fechaHasta = '';
      this.render();
    });

    this.mountEl.querySelector('#btn-toggle-metas-mes')?.addEventListener('click', () => {
      this.showTargetsEditor = !this.showTargetsEditor;
      this.render();
    });

    this.mountEl.querySelectorAll('[data-month-target]').forEach((inp) => {
      inp.addEventListener('change', () => {
        const mKey = inp.getAttribute('data-month-target');
        saveMonthlyTarget(mKey, inp.value);
        this.render();
      });
    });

    this.mountEl.querySelectorAll('[data-wk-metric]').forEach((btn) => {
      btn.addEventListener('click', () => {
        this.weeklyMetric = btn.getAttribute('data-wk-metric');
        this.render();
      });
    });

    this.mountEl.querySelector('#btn-print-gerencia')?.addEventListener('click', () => {
      window.print();
    });

    this.mountEl.querySelector('#btn-export-exec-xlsx')?.addEventListener('click', () => {
      const range = getDataDateRange(this.app.registros);
      const res = exportExecutiveWorkbookToExcel({
        records: this.app.registros,
        fechaDesde: this.fechaDesde || range.min,
        fechaHasta: this.fechaHasta || range.max
      });
      this.app.showToast(
        'success',
        `Reporte Gerencial Excel exportado (${res.fileName}): ${res.sheetNames.length} hojas y ${res.chartCount} gráficos nativos.`
      );
    });
  }
}
