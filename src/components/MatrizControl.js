/**
 * Vista 2 — Planilla Control STRACON / PICOTÓN (Excel + Metas + Tendencia + Avance + Comparativo Mensual)
 */

import {
  buildMatrixModel,
  buildMonthlyComparative,
  getDataDateRange,
  getOperativeTargets,
  saveOperativeTargets
} from '../services/productionService.js';
import { exportProductionWorkbookToExcel } from '../services/excelExporter.js';
import {
  getFechaHoyPeru,
  getValuationPeriodRange,
  getCalendarMonthRange,
  shiftIsoDate,
  getMonthTitleSpanish,
  getValuationMonthInfo
} from '../utils/datePeru.js';
import { escapeHtml } from '../utils/codeNormalizer.js';

const LAST_EQUIPO_KEY = 'stracon_produccion_last_equipo_v1';

export class MatrizControlView {
  constructor(mountEl, app) {
    this.mountEl = mountEl;
    this.app = app;
    this.codigoEquipo = '';
    this.preset = 'datos';
    this.fechaDesde = '';
    this.fechaHasta = '';
    this.tituloPersonalizado = '';
    this.targets = getOperativeTargets();
    this.initializedRange = false;
  }

  syncDefaultsIfNeeded() {
    const savedEq = localStorage.getItem(LAST_EQUIPO_KEY);
    if (!this.codigoEquipo || !this.app.equipos.some((e) => e.codigo === this.codigoEquipo)) {
      if (savedEq && this.app.equipos.some((e) => e.codigo === savedEq)) {
        this.codigoEquipo = savedEq;
      } else {
        const firstWithData = this.app.equipos.find((e) =>
          this.app.registros.some((r) => r.codigo_equipo === e.codigo)
        );
        this.codigoEquipo = firstWithData?.codigo || this.app.equipos[0]?.codigo || 'STREX02';
      }
    }
    if (!this.initializedRange) {
      this.applyPreset('datos');
      this.initializedRange = true;
    }
  }

  applyPreset(preset) {
    this.preset = preset;
    const eqRecords = this.app.registros.filter((r) => r.codigo_equipo === this.codigoEquipo);
    const rangeData = getDataDateRange(eqRecords.length > 0 ? eqRecords : this.app.registros);

    if (preset === 'datos') {
      const valRange = getValuationPeriodRange(rangeData.max);
      if (rangeData.min >= valRange.desde && rangeData.max <= valRange.hasta) {
        this.fechaDesde = valRange.desde;
        this.fechaHasta = rangeData.max > valRange.hasta ? rangeData.max : valRange.hasta;
      } else {
        this.fechaDesde = rangeData.min;
        this.fechaHasta = rangeData.max;
      }
    } else if (preset === 'valorizacion') {
      const p = getValuationPeriodRange(getFechaHoyPeru());
      this.fechaDesde = p.desde;
      this.fechaHasta = p.hasta;
    } else if (preset === 'mes') {
      const m = getCalendarMonthRange(getFechaHoyPeru());
      this.fechaDesde = m.desde;
      this.fechaHasta = m.hasta;
    } else if (preset === '15dias') {
      const hoy = getFechaHoyPeru();
      this.fechaDesde = shiftIsoDate(hoy, -14);
      this.fechaHasta = hoy;
    }
  }

  async mount() {
    this.targets = getOperativeTargets();
    this.syncDefaultsIfNeeded();
    this.render();
  }

  renderSvgChart(days = []) {
    if (!days || days.length === 0) {
      return '<p style="font-weight:700; color:var(--text-muted);">No hay días en el rango seleccionado.</p>';
    }

    const barSlot = Math.max(30, Math.min(48, Math.floor(880 / Math.max(days.length, 1))));
    const padL = 46;
    const padR = 46;
    const padT = 22;
    const padB = 42;
    const plotW = Math.max(520, days.length * barSlot);
    const plotH = 190;
    const svgW = padL + plotW + padR;
    const svgH = padT + plotH + padB;

    const maxMat = Math.max(100, ...days.map((d) => d.total.material || 0));
    const maxRend = Math.max(10, ...days.map((d) => d.total.rendimiento || 0));
    const stepW = plotW / days.length;
    const barW = Math.max(12, Math.min(28, stepW * 0.62));

    const yRend = (val) => padT + plotH - (val / maxRend) * plotH;

    const gridLines = [0, 0.25, 0.5, 0.75, 1]
      .map((frac) => {
        const y = padT + plotH * (1 - frac);
        const val = Math.round(maxMat * frac);
        return `
          <line x1="${padL}" y1="${y.toFixed(1)}" x2="${padL + plotW}" y2="${y.toFixed(1)}" stroke="#CBD5E1" stroke-dasharray="3,3" />
          <text x="${padL - 6}" y="${(y + 4).toFixed(1)}" text-anchor="end" font-size="10" font-weight="700" fill="#475569">${val}</text>
        `;
      })
      .join('');

    const bars = days
      .map((d, idx) => {
        const cx = padL + idx * stepW + stepW / 2;
        const x = cx - barW / 2;
        const matDia = d.dia.material || 0;
        const matNoche = d.noche.material || 0;
        const hDia = (matDia / maxMat) * plotH;
        const hNoche = (matNoche / maxMat) * plotH;
        const yDia = padT + plotH - hDia;
        const yNoche = yDia - hNoche;
        const showLabel = days.length <= 22 || idx % 2 === 0;
        return `
          <g>
            ${
              hDia > 0
                ? `<rect x="${x.toFixed(1)}" y="${yDia.toFixed(1)}" width="${barW.toFixed(1)}" height="${hDia.toFixed(1)}" fill="#1B1F4B"><title>${d.fechaCorta} Día: ${matDia} m³</title></rect>`
                : ''
            }
            ${
              hNoche > 0
                ? `<rect x="${x.toFixed(1)}" y="${yNoche.toFixed(1)}" width="${barW.toFixed(1)}" height="${hNoche.toFixed(1)}" fill="#E40046"><title>${d.fechaCorta} Noche: ${matNoche} m³</title></rect>`
                : ''
            }
            ${
              d.total.material > 0
                ? `<text x="${cx.toFixed(1)}" y="${Math.max(12, yNoche - 4).toFixed(1)}" text-anchor="middle" font-size="9.5" font-weight="800" fill="#111827">${Math.round(
                    d.total.material
                  )}</text>`
                : ''
            }
            ${
              showLabel
                ? `<text x="${cx.toFixed(1)}" y="${padT + plotH + 16}" text-anchor="middle" font-size="9.5" font-weight="700" fill="#1B1F4B">${escapeHtml(
                    d.fechaCorta
                  )}</text>`
                : ''
            }
          </g>
        `;
      })
      .join('');

    const linePts = days.map((d, idx) => {
      const cx = padL + idx * stepW + stepW / 2;
      const cy = yRend(d.total.rendimiento || 0);
      return { cx, cy, rend: d.total.rendimiento || 0, fechaCorta: d.fechaCorta };
    });

    const polylineStr = linePts.map((p) => `${p.cx.toFixed(1)},${p.cy.toFixed(1)}`).join(' ');
    const dots = linePts
      .map((p) =>
        p.rend > 0
          ? `<circle cx="${p.cx.toFixed(1)}" cy="${p.cy.toFixed(1)}" r="3.6" fill="#F59E0B" stroke="#1B1F4B" stroke-width="1.3"><title>${p.fechaCorta}: ${p.rend.toFixed(
              2
            )} m³/HM</title></circle>`
          : ''
      )
      .join('');

    return `
      <div class="stracon-chart-card">
        <svg viewBox="0 0 ${svgW} ${svgH}" style="width:100%; min-width:${Math.min(640, svgW)}px; height:auto; display:block;">
          ${gridLines}
          <line x1="${padL}" y1="${padT + plotH}" x2="${padL + plotW}" y2="${padT + plotH}" stroke="#1B1F4B" stroke-width="1.5"/>
          ${bars}
          <polyline fill="none" stroke="#F59E0B" stroke-width="2.4" points="${polylineStr}" />
          ${dots}
          <text x="${padL + plotW + 6}" y="${padT + 10}" font-size="10" font-weight="800" fill="#B45309">${maxRend.toFixed(0)} m³/HM</text>
          <text x="${padL + plotW + 6}" y="${padT + plotH}" font-size="10" font-weight="800" fill="#4B5563">0</text>
        </svg>
        <div class="stracon-chart-legend">
          <span><span class="legend-dot" style="background:#1B1F4B;"></span>Material Día (m³)</span>
          <span><span class="legend-dot" style="background:#E40046;"></span>Material Noche (m³)</span>
          <span><span class="legend-dot" style="background:#F59E0B; border-radius:999px;"></span>Rendimiento Total (m³/HM)</span>
        </div>
      </div>
    `;
  }

  render() {
    this.syncDefaultsIfNeeded();
    const matrix = buildMatrixModel(this.app.registros, {
      codigoEquipo: this.codigoEquipo,
      fechaDesde: this.fechaDesde,
      fechaHasta: this.fechaHasta,
      metaDiaria: this.targets.metaDiaria,
      metaMensual: this.targets.metaMensual
    });
    const compModel = buildMonthlyComparative(
      this.app.registros.filter((r) => r.fecha >= this.fechaDesde && r.fecha <= this.fechaHasta)
    );
    const days = matrix.days;
    const s = matrix.summary;
    const valInfo = getValuationMonthInfo(this.fechaHasta || this.fechaDesde);
    const periodTitle =
      this.tituloPersonalizado.trim() ||
      valInfo.label ||
      getMonthTitleSpanish(this.fechaDesde, this.fechaHasta) ||
      'CONTROL OPERATIVO';

    const equipOpts = this.app.equipos
      .map(
        (eq) =>
          `<option value="${escapeHtml(eq.codigo)}" ${
            eq.codigo === this.codigoEquipo ? 'selected' : ''
          }>${escapeHtml(eq.codigo)}</option>`
      )
      .join('');

    const fmt1 = (v) => (v === null || v === undefined ? '' : Number(v).toFixed(1));
    const fmt2 = (v) => (v === null || v === undefined ? '' : Number(v).toFixed(2));
    const fmtPct = (v) => `${(Number(v || 0) * 100).toFixed(1)}%`;
    const fmtMat = (v) => (v === null || v === undefined || v === 0 ? '' : Math.round(v).toLocaleString('es-PE'));

    this.mountEl.innerHTML = `
      <section class="panel">
        <div class="panel-header">
          <h2 class="panel-title">📊 Planilla Control Operativo PICOTÓN — ${escapeHtml(this.codigoEquipo)}</h2>
          <div class="inline-row" style="flex-wrap:wrap;">
            <button type="button" class="btn-mine btn-excel" id="btn-export-excel-main">
              📥 Exportar Excel (${escapeHtml(this.codigoEquipo)})
            </button>
            <button type="button" class="btn-mine btn-navy" id="btn-export-excel-all">
              📚 Exportar Consolidado (Todos)
            </button>
          </div>
        </div>

        <div class="form-grid" style="margin-bottom:0.85rem;">
          <div class="field-group">
            <label class="field-label" for="mat-equipo">Equipo</label>
            <select id="mat-equipo" class="input-mine">${equipOpts}</select>
          </div>

          <div class="field-group">
            <label class="field-label" for="mat-preset">Período Rápido</label>
            <select id="mat-preset" class="input-mine">
              <option value="datos" ${this.preset === 'datos' ? 'selected' : ''}>Período activo con datos (26–25)</option>
              <option value="valorizacion" ${this.preset === 'valorizacion' ? 'selected' : ''}>Valorización actual (26 al 25)</option>
              <option value="mes" ${this.preset === 'mes' ? 'selected' : ''}>Mes Calendario Actual</option>
              <option value="15dias" ${this.preset === '15dias' ? 'selected' : ''}>Últimos 15 días</option>
              <option value="custom" ${this.preset === 'custom' ? 'selected' : ''}>Personalizado</option>
            </select>
          </div>

          <div class="field-group">
            <label class="field-label" for="mat-desde">Desde</label>
            <input type="date" id="mat-desde" class="input-mine" value="${escapeHtml(this.fechaDesde)}" />
          </div>

          <div class="field-group">
            <label class="field-label" for="mat-hasta">Hasta</label>
            <input type="date" id="mat-hasta" class="input-mine" value="${escapeHtml(this.fechaHasta)}" />
          </div>

          <div class="field-group">
            <label class="field-label" for="mat-meta-dia">Meta Diaria (Hr/día)</label>
            <input type="number" step="0.5" min="1" id="mat-meta-dia" class="input-mine tabular" value="${s.metaDiaria}" />
          </div>

          <div class="field-group">
            <label class="field-label" for="mat-meta-mes">Meta Mensual (Hr)</label>
            <input type="number" step="10" min="1" id="mat-meta-mes" class="input-mine tabular" value="${s.metaMensual}" />
          </div>

          <div class="field-group full-span">
            <label class="field-label" for="mat-titulo">Título de Planilla (opcional — aparece en el banner del Excel)</label>
            <input
              type="text"
              id="mat-titulo"
              class="input-mine"
              placeholder="Automático: ${escapeHtml(valInfo.label)}"
              value="${escapeHtml(this.tituloPersonalizado)}"
            />
          </div>
        </div>

        <p style="margin:0 0 0.5rem; font-size:0.84rem; font-weight:700; color:var(--text-muted);">
          💡 Toca cualquier celda de Turno Día o Turno Noche en la planilla para registrarla o editarla directamente.
        </p>

        <div class="excel-matrix-wrap">
          <table class="excel-matrix-table">
            <thead>
              <tr>
                <th class="col-a-sticky excel-banner-logo">STRACON</th>
                <th class="excel-banner-title" colspan="${days.length + 1}">
                  CONTROL FRACTURA DE ROCA - ${escapeHtml(periodTitle)} | TURNO DÍA + TURNO NOCHE
                </th>
              </tr>
              <tr>
                <th class="col-a-sticky excel-date-hdr" rowspan="2">Fecha</th>
                ${days.map((d) => `<th class="excel-date-hdr">${escapeHtml(d.diaSemana)}</th>`).join('')}
                <th class="excel-date-hdr col-total-sticky">TOTAL / PROM.</th>
              </tr>
              <tr>
                ${days.map((d) => `<th class="excel-date-hdr">${escapeHtml(d.fechaCorta)}</th>`).join('')}
                <th class="excel-date-hdr col-total-sticky">PERÍODO</th>
              </tr>
            </thead>
            <tbody>
              <tr class="excel-equipo-row">
                <td class="col-a-sticky">EQUIPO</td>
                ${days.map(() => `<td>${escapeHtml(matrix.codigoEquipo)}</td>`).join('')}
                <td class="col-total-sticky">${escapeHtml(matrix.codigoEquipo)}</td>
              </tr>

              <tr class="excel-band-dia">
                <td class="col-a-sticky">TURNO DÍA</td>
                <td colspan="${days.length + 1}"></td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-label-white">Horómetro Inicial</td>
                ${days
                  .map(
                    (d) =>
                      `<td class="bg-dia-hm matrix-clickable-cell" data-slot-fecha="${d.fecha}" data-slot-turno="DIA">${fmt1(
                        d.dia.hmInicial
                      )}</td>`
                  )
                  .join('')}
                <td class="bg-dia-hm col-total-sticky">-</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-label-white">Horómetro Final</td>
                ${days
                  .map(
                    (d) =>
                      `<td class="bg-dia-hm matrix-clickable-cell" data-slot-fecha="${d.fecha}" data-slot-turno="DIA">${fmt1(
                        d.dia.hmFinal
                      )}</td>`
                  )
                  .join('')}
                <td class="bg-dia-hm col-total-sticky">${s.ultimoHorometro !== null ? s.ultimoHorometro.toFixed(1) : '-'}</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-dia-hr">Horas Día</td>
                ${days
                  .map(
                    (d) =>
                      `<td class="bg-dia-hr matrix-clickable-cell" data-slot-fecha="${d.fecha}" data-slot-turno="DIA">${d.dia.horas.toFixed(
                        1
                      )}</td>`
                  )
                  .join('')}
                <td class="bg-dia-hr col-total-sticky">${s.horasDia.toFixed(1)}</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-mat-green">Material Fracturado (m³)</td>
                ${days
                  .map(
                    (d) =>
                      `<td class="bg-mat-green matrix-clickable-cell" data-slot-fecha="${d.fecha}" data-slot-turno="DIA">${fmtMat(
                        d.dia.material
                      )}</td>`
                  )
                  .join('')}
                <td class="bg-mat-green col-total-sticky">${Math.round(s.materialDia).toLocaleString('es-PE')}</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-label-white">Rendimiento (m³/HM)</td>
                ${days
                  .map(
                    (d) =>
                      `<td class="bg-rend-green">${d.dia.record ? fmt2(d.dia.rendimiento) : ''}</td>`
                  )
                  .join('')}
                <td class="bg-rend-green col-total-sticky">${s.rendDia.toFixed(2)}</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-label-white">Área de Trabajo</td>
                ${days
                  .map(
                    (d) =>
                      `<td class="bg-area-white matrix-clickable-cell" data-slot-fecha="${d.fecha}" data-slot-turno="DIA">${escapeHtml(
                        d.dia.area
                      )}</td>`
                  )
                  .join('')}
                <td class="bg-area-white col-total-sticky"></td>
              </tr>

              <tr class="excel-sep-blue">
                <td class="col-a-sticky"></td>
                <td colspan="${days.length + 1}"></td>
              </tr>

              <tr class="excel-band-noche">
                <td class="col-a-sticky">TURNO NOCHE</td>
                <td colspan="${days.length + 1}"></td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-label-white">Horómetro Inicial</td>
                ${days
                  .map(
                    (d) =>
                      `<td class="bg-noche-hm matrix-clickable-cell" data-slot-fecha="${d.fecha}" data-slot-turno="NOCHE">${fmt1(
                        d.noche.hmInicial
                      )}</td>`
                  )
                  .join('')}
                <td class="bg-noche-hm col-total-sticky">-</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-label-white">Horómetro Final</td>
                ${days
                  .map(
                    (d) =>
                      `<td class="bg-noche-hm matrix-clickable-cell" data-slot-fecha="${d.fecha}" data-slot-turno="NOCHE">${fmt1(
                        d.noche.hmFinal
                      )}</td>`
                  )
                  .join('')}
                <td class="bg-noche-hm col-total-sticky">-</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-noche-hr">Horas Noche</td>
                ${days
                  .map(
                    (d) =>
                      `<td class="bg-noche-hr matrix-clickable-cell" data-slot-fecha="${d.fecha}" data-slot-turno="NOCHE">${d.noche.horas.toFixed(
                        1
                      )}</td>`
                  )
                  .join('')}
                <td class="bg-noche-hr col-total-sticky">${s.horasNoche.toFixed(1)}</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-mat-green">Material Fracturado (m³)</td>
                ${days
                  .map(
                    (d) =>
                      `<td class="bg-mat-green matrix-clickable-cell" data-slot-fecha="${d.fecha}" data-slot-turno="NOCHE">${fmtMat(
                        d.noche.material
                      )}</td>`
                  )
                  .join('')}
                <td class="bg-mat-green col-total-sticky">${Math.round(s.materialNoche).toLocaleString('es-PE')}</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-label-white">Rendimiento (m³/HM)</td>
                ${days
                  .map(
                    (d) =>
                      `<td class="bg-rend-green">${d.noche.record ? fmt2(d.noche.rendimiento) : ''}</td>`
                  )
                  .join('')}
                <td class="bg-rend-green col-total-sticky">${s.rendNoche.toFixed(2)}</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-label-white">Área de Trabajo</td>
                ${days
                  .map(
                    (d) =>
                      `<td class="bg-area-white matrix-clickable-cell" data-slot-fecha="${d.fecha}" data-slot-turno="NOCHE">${escapeHtml(
                        d.noche.area
                      )}</td>`
                  )
                  .join('')}
                <td class="bg-area-white col-total-sticky"></td>
              </tr>

              <tr class="excel-band-total">
                <td class="col-a-sticky">TOTAL DÍA</td>
                <td colspan="${days.length + 1}"></td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-tot-yellow">Horas Totales (Día+Noche)</td>
                ${days.map((d) => `<td class="bg-tot-yellow">${d.total.horas.toFixed(1)}</td>`).join('')}
                <td class="bg-tot-yellow col-total-sticky">${s.horasTotal.toFixed(1)}</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-rend-green">Material Total (m³)</td>
                ${days
                  .map((d) => `<td class="bg-rend-green">${d.total.material > 0 ? Math.round(d.total.material) : '-'}</td>`)
                  .join('')}
                <td class="bg-rend-green col-total-sticky">${Math.round(s.materialTotal).toLocaleString('es-PE')}</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-label-white">Rend Total (m³/HM)</td>
                ${days.map((d) => `<td class="bg-rend-green">${d.total.rendimiento.toFixed(2)}</td>`).join('')}
                <td class="bg-rend-green col-total-sticky">${s.rendTotal.toFixed(2)}</td>
              </tr>

              <tr class="excel-band-acum">
                <td class="col-a-sticky">ACUMULADOS Y AVANCE</td>
                <td colspan="${days.length + 1}"></td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-acum-gray">Acumulado Hr Día</td>
                ${days.map((d) => `<td class="bg-acum-gray">${d.acumulado.hrDia.toFixed(1)}</td>`).join('')}
                <td class="bg-acum-gray col-total-sticky">${s.horasDia.toFixed(1)}</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-acum-gray">Acumulado Hr Noche</td>
                ${days.map((d) => `<td class="bg-acum-gray">${d.acumulado.hrNoche.toFixed(1)}</td>`).join('')}
                <td class="bg-acum-gray col-total-sticky">${s.horasNoche.toFixed(1)}</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-tot-yellow">Acumulado Total (Hr)</td>
                ${days.map((d) => `<td class="bg-tot-yellow">${d.acumulado.hrTotal.toFixed(1)}</td>`).join('')}
                <td class="bg-tot-yellow col-total-sticky">${s.horasTotal.toFixed(1)}</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-meta-blue">Meta Diaria Total (Hr)</td>
                ${days.map((d) => `<td class="bg-meta-blue">${d.meta.metaDiaria.toFixed(1)}</td>`).join('')}
                <td class="bg-meta-blue col-total-sticky">${s.metaDiaria.toFixed(1)}</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-meta-blue">Acumulado Planeado (Hr)</td>
                ${days.map((d) => `<td class="bg-meta-blue">${d.meta.acumuladoPlaneado.toFixed(1)}</td>`).join('')}
                <td class="bg-meta-blue col-total-sticky">${s.acumuladoPlaneadoFinal.toFixed(1)}</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-meta-blue">Tendencia</td>
                ${days
                  .map(
                    (d) =>
                      `<td class="${d.meta.tendencia >= 1 ? 'bg-trend-ok' : 'bg-trend-warn'}">${fmtPct(
                        d.meta.tendencia
                      )}</td>`
                  )
                  .join('')}
                <td class="${s.tendenciaFinal >= 1 ? 'bg-trend-ok' : 'bg-trend-warn'} col-total-sticky">${fmtPct(
                  s.tendenciaFinal
                )}</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-meta-blue">% de Avance</td>
                ${days.map((d) => `<td class="bg-meta-blue">${fmtPct(d.meta.avance)}</td>`).join('')}
                <td class="bg-meta-blue col-total-sticky">${fmtPct(s.avanceFinal)}</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-mat-green">Acumulado Material Total (m³)</td>
                ${days
                  .map(
                    (d) =>
                      `<td class="bg-mat-green"><strong>${Math.round(d.acumulado.matTotal).toLocaleString('es-PE')}</strong></td>`
                  )
                  .join('')}
                <td class="bg-mat-green col-total-sticky"><strong>${Math.round(s.materialTotal).toLocaleString('es-PE')}</strong></td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-tot-yellow">Rendimiento Acumulado (m³/HM)</td>
                ${days.map((d) => `<td class="bg-tot-yellow">${d.acumulado.rendAcum.toFixed(2)}</td>`).join('')}
                <td class="bg-tot-yellow col-total-sticky">${s.rendTotal.toFixed(2)}</td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-label-white">Operador de Día</td>
                ${days.map((d) => `<td class="bg-area-white">${escapeHtml(d.dia.operador)}</td>`).join('')}
                <td class="bg-area-white col-total-sticky"></td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-label-white">Operador de Noche</td>
                ${days.map((d) => `<td class="bg-area-white">${escapeHtml(d.noche.operador)}</td>`).join('')}
                <td class="bg-area-white col-total-sticky"></td>
              </tr>
              <tr>
                <td class="col-a-sticky bg-meta-blue">Meta mensual Horas Totales (Hr)</td>
                ${days
                  .map((_, i) => `<td class="bg-meta-blue">${i === 0 ? s.metaMensual : ''}</td>`)
                  .join('')}
                <td class="bg-meta-blue col-total-sticky">${s.metaMensual}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </section>

      <section class="panel">
        <div class="panel-header">
          <h2 class="panel-title">🏆 Comparativo Mensual por Equipo (Turnos Diurno y Nocturno)</h2>
        </div>
        ${
          compModel.rows.length === 0
            ? `<p style="font-weight:700; color:var(--text-muted);">No hay datos en este rango para el comparativo.</p>`
            : `<div class="table-scroll">
                <table class="data-table">
                  <thead>
                    <tr>
                      <th>Equipo</th>
                      <th>Mes (Valorización)</th>
                      <th>Turno Día (Hr)</th>
                      <th>Turno Noche (Hr)</th>
                      <th>Total (Hr)</th>
                      <th>Material (m³)</th>
                      <th>Rendimiento (m³/HM)</th>
                      <th>Prom. Día (Hr/día)</th>
                      <th>Prom. Noche (Hr/día)</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${compModel.rows
                      .map(
                        (r) => `
                        <tr>
                          <td><strong>${escapeHtml(r.codigoEquipo)}</strong></td>
                          <td>${escapeHtml(r.mesLabel)}</td>
                          <td class="tabular">${r.horasDia.toFixed(1)}</td>
                          <td class="tabular">${r.horasNoche.toFixed(1)}</td>
                          <td class="tabular"><strong>${r.horasTotal.toFixed(1)}</strong></td>
                          <td class="tabular"><strong>${Math.round(r.materialTotal).toLocaleString('es-PE')}</strong></td>
                          <td class="tabular"><strong>${r.rendimiento.toFixed(2)}</strong></td>
                          <td class="tabular">${r.promDiaHr.toFixed(2)}</td>
                          <td class="tabular">${r.promNocheHr.toFixed(2)}</td>
                        </tr>`
                      )
                      .join('')}
                  </tbody>
                  <tfoot>
                    <tr>
                      <td>TOTAL FLOTA</td>
                      <td>${compModel.rows.length} grupo(s)</td>
                      <td class="tabular">${compModel.totals.horasDia.toFixed(1)}</td>
                      <td class="tabular">${compModel.totals.horasNoche.toFixed(1)}</td>
                      <td class="tabular">${compModel.totals.horasTotal.toFixed(1)}</td>
                      <td class="tabular">${Math.round(compModel.totals.materialTotal).toLocaleString('es-PE')}</td>
                      <td class="tabular">${compModel.totals.rendimiento.toFixed(2)}</td>
                      <td class="tabular">${compModel.totals.promDiaHr.toFixed(2)}</td>
                      <td class="tabular">${compModel.totals.promNocheHr.toFixed(2)}</td>
                    </tr>
                  </tfoot>
                </table>
              </div>`
        }
      </section>

      <section class="panel">
        <div class="panel-header">
          <h2 class="panel-title">📈 Evolución Diaria de Producción (m³) y Rendimiento (m³/HM)</h2>
        </div>
        ${this.renderSvgChart(days)}
      </section>

      <section class="panel">
        <div class="panel-header">
          <h2 class="panel-title">📍 Resumen por Área de Trabajo (${escapeHtml(this.codigoEquipo)})</h2>
        </div>
        ${
          matrix.areaBreakdown.length === 0
            ? `<p style="font-weight:700; color:var(--text-muted);">No hay horas ni material registrado por área en este rango.</p>`
            : `<div class="table-scroll">
                <table class="data-table">
                  <thead>
                    <tr>
                      <th>Área de Trabajo</th>
                      <th>Turnos</th>
                      <th>Horas (HM)</th>
                      <th>Material Fracturado (m³)</th>
                      <th>% Participación</th>
                      <th>Rendimiento Prom. (m³/HM)</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${matrix.areaBreakdown
                      .map(
                        (a) => `
                        <tr>
                          <td><strong>${escapeHtml(a.area)}</strong></td>
                          <td class="tabular">${a.turnos}</td>
                          <td class="tabular">${a.horas.toFixed(1)}</td>
                          <td class="tabular"><strong>${Math.round(a.material).toLocaleString('es-PE')}</strong></td>
                          <td class="tabular">${a.porcentajeMaterial.toFixed(1)}%</td>
                          <td class="tabular"><strong>${a.rendimiento.toFixed(2)}</strong></td>
                        </tr>`
                      )
                      .join('')}
                  </tbody>
                  <tfoot>
                    <tr>
                      <td>TOTAL PERÍODO</td>
                      <td class="tabular">${matrix.areaBreakdown.reduce((acc, a) => acc + a.turnos, 0)}</td>
                      <td class="tabular">${s.horasTotal.toFixed(1)}</td>
                      <td class="tabular">${Math.round(s.materialTotal).toLocaleString('es-PE')}</td>
                      <td class="tabular">100.0%</td>
                      <td class="tabular">${s.rendTotal.toFixed(2)}</td>
                    </tr>
                  </tfoot>
                </table>
              </div>`
        }
      </section>
    `;

    this.bindEvents();
  }

  bindEvents() {
    this.mountEl.querySelector('#mat-equipo')?.addEventListener('change', (e) => {
      this.codigoEquipo = e.target.value;
      localStorage.setItem(LAST_EQUIPO_KEY, this.codigoEquipo);
      if (this.preset !== 'custom') this.applyPreset(this.preset);
      this.render();
    });

    this.mountEl.querySelector('#mat-preset')?.addEventListener('change', (e) => {
      this.applyPreset(e.target.value);
      this.render();
    });

    this.mountEl.querySelector('#mat-desde')?.addEventListener('change', (e) => {
      this.fechaDesde = e.target.value;
      this.preset = 'custom';
      this.render();
    });

    this.mountEl.querySelector('#mat-hasta')?.addEventListener('change', (e) => {
      this.fechaHasta = e.target.value;
      this.preset = 'custom';
      this.render();
    });

    const updateTargets = () => {
      const md = this.mountEl.querySelector('#mat-meta-dia')?.value;
      const mm = this.mountEl.querySelector('#mat-meta-mes')?.value;
      this.targets = saveOperativeTargets({ metaDiaria: md, metaMensual: mm });
      this.render();
    };
    this.mountEl.querySelector('#mat-meta-dia')?.addEventListener('change', updateTargets);
    this.mountEl.querySelector('#mat-meta-mes')?.addEventListener('change', updateTargets);

    this.mountEl.querySelector('#mat-titulo')?.addEventListener('input', (e) => {
      this.tituloPersonalizado = e.target.value;
    });

    this.mountEl.querySelector('#btn-export-excel-main')?.addEventListener('click', () => {
      const res = exportProductionWorkbookToExcel({
        records: this.app.registros,
        codigoEquipo: this.codigoEquipo,
        fechaDesde: this.fechaDesde,
        fechaHasta: this.fechaHasta,
        tituloPersonalizado: this.tituloPersonalizado,
        metaDiaria: this.targets.metaDiaria,
        metaMensual: this.targets.metaMensual,
        includeAllEquipments: false
      });
      this.app.showToast(
        'success',
        `Excel generado (${res.fileName}): ${res.sheetNames.length} hojas (${res.sheetNames.join(', ')}).`
      );
    });

    this.mountEl.querySelector('#btn-export-excel-all')?.addEventListener('click', () => {
      const res = exportProductionWorkbookToExcel({
        records: this.app.registros,
        codigoEquipo: 'TODOS',
        fechaDesde: this.fechaDesde,
        fechaHasta: this.fechaHasta,
        tituloPersonalizado: this.tituloPersonalizado,
        metaDiaria: this.targets.metaDiaria,
        metaMensual: this.targets.metaMensual,
        includeAllEquipments: true
      });
      this.app.showToast(
        'success',
        `Consolidado Excel generado (${res.fileName}): ${res.sheetNames.length} hojas (${res.sheetNames.join(', ')}).`
      );
    });

    this.mountEl.querySelectorAll('[data-slot-fecha]').forEach((cell) => {
      cell.addEventListener('click', () => {
        const fecha = cell.getAttribute('data-slot-fecha');
        const turno = cell.getAttribute('data-slot-turno');
        this.app.openCapturaForSlot({
          codigoEquipo: this.codigoEquipo,
          fecha,
          turno
        });
      });
    });
  }
}
