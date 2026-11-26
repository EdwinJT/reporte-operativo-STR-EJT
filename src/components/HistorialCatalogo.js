/**
 * Vista 3 — Historial Filtrado + Administración de Catálogo de Equipos y Áreas
 */

import { deleteShiftRecord, getDataDateRange } from '../services/productionService.js';
import {
  createOrTouchEquipment,
  deleteEquipment,
  createOrTouchArea,
  deleteArea
} from '../services/catalogService.js';
import { exportProductionWorkbookToExcel } from '../services/excelExporter.js';
import { formatFechaDDMMYYYY } from '../utils/datePeru.js';
import { escapeHtml, round1, round2 } from '../utils/codeNormalizer.js';

export class HistorialCatalogoView {
  constructor(mountEl, app) {
    this.mountEl = mountEl;
    this.app = app;
    this.filters = {
      equipo: 'TODOS',
      turno: 'TODOS',
      area: 'TODAS',
      desde: '',
      hasta: '',
      texto: ''
    };
  }

  async mount() {
    this.render();
  }

  getFilteredRecords() {
    const q = this.filters.texto.trim().toLowerCase();
    return this.app.registros.filter((r) => {
      if (this.filters.equipo !== 'TODOS' && r.codigo_equipo !== this.filters.equipo) return false;
      if (this.filters.turno !== 'TODOS' && r.turno !== this.filters.turno) return false;
      if (this.filters.area !== 'TODAS' && (r.area_trabajo || '') !== this.filters.area) return false;
      if (this.filters.desde && r.fecha < this.filters.desde) return false;
      if (this.filters.hasta && r.fecha > this.filters.hasta) return false;
      if (q) {
        const haystack = `${r.codigo_equipo} ${r.area_trabajo} ${r.operador} ${r.observaciones} ${r.fecha}`.toLowerCase();
        if (!haystack.includes(q)) return false;
      }
      return true;
    });
  }

  render() {
    const rows = this.getFilteredRecords();
    const totHoras = round1(rows.reduce((acc, r) => acc + (Number(r.horas_trabajadas) || 0), 0));
    const totMat = round2(rows.reduce((acc, r) => acc + (Number(r.material_fracturado) || 0), 0));
    const totRend = totHoras > 0 ? round2(totMat / totHoras) : 0;

    const eqOpts = [
      `<option value="TODOS" ${this.filters.equipo === 'TODOS' ? 'selected' : ''}>Todos los equipos</option>`,
      ...this.app.equipos.map(
        (e) =>
          `<option value="${escapeHtml(e.codigo)}" ${
            this.filters.equipo === e.codigo ? 'selected' : ''
          }>${escapeHtml(e.codigo)}</option>`
      )
    ].join('');

    const areaOpts = [
      `<option value="TODAS" ${this.filters.area === 'TODAS' ? 'selected' : ''}>Todas las áreas</option>`,
      ...this.app.areas.map(
        (a) =>
          `<option value="${escapeHtml(a.nombre)}" ${
            this.filters.area === a.nombre ? 'selected' : ''
          }>${escapeHtml(a.nombre)}</option>`
      )
    ].join('');

    this.mountEl.innerHTML = `
      <section class="panel">
        <div class="panel-header">
          <h2 class="panel-title">🗂️ Historial Completo de Turnos (${rows.length})</h2>
          <div class="inline-row" style="flex-wrap:wrap;">
            <button type="button" class="btn-mine btn-sm btn-secondary" id="btn-clear-filters">Limpiar Filtros</button>
            <button type="button" class="btn-mine btn-sm btn-excel" id="btn-export-filtered">📥 Exportar Excel Filtrado</button>
          </div>
        </div>

        <div class="form-grid" style="margin-bottom:0.85rem;">
          <div class="field-group">
            <label class="field-label" for="flt-equipo">Equipo</label>
            <select id="flt-equipo" class="input-mine">${eqOpts}</select>
          </div>
          <div class="field-group">
            <label class="field-label" for="flt-turno">Turno</label>
            <select id="flt-turno" class="input-mine">
              <option value="TODOS" ${this.filters.turno === 'TODOS' ? 'selected' : ''}>Todos</option>
              <option value="DIA" ${this.filters.turno === 'DIA' ? 'selected' : ''}>☀️ TURNO DÍA</option>
              <option value="NOCHE" ${this.filters.turno === 'NOCHE' ? 'selected' : ''}>🌙 TURNO NOCHE</option>
            </select>
          </div>
          <div class="field-group">
            <label class="field-label" for="flt-area">Área de Trabajo</label>
            <select id="flt-area" class="input-mine">${areaOpts}</select>
          </div>
          <div class="field-group">
            <label class="field-label" for="flt-desde">Desde</label>
            <input type="date" id="flt-desde" class="input-mine" value="${escapeHtml(this.filters.desde)}" />
          </div>
          <div class="field-group">
            <label class="field-label" for="flt-hasta">Hasta</label>
            <input type="date" id="flt-hasta" class="input-mine" value="${escapeHtml(this.filters.hasta)}" />
          </div>
          <div class="field-group">
            <label class="field-label" for="flt-texto">Buscar</label>
            <input type="text" id="flt-texto" class="input-mine" placeholder="Operador, área, nota..." value="${escapeHtml(this.filters.texto)}" />
          </div>
        </div>

        <div class="table-scroll">
          <table class="data-table">
            <thead>
              <tr>
                <th>Fecha</th>
                <th>Equipo</th>
                <th>Turno</th>
                <th>HM Ini.</th>
                <th>HM Fin.</th>
                <th>Horas</th>
                <th>Material (m³)</th>
                <th>Rend. (m³/HM)</th>
                <th>Área</th>
                <th>Operador</th>
                <th>Acciones</th>
              </tr>
            </thead>
            <tbody>
              ${
                rows.length === 0
                  ? `<tr><td colspan="11" style="text-align:center; font-weight:700; padding:1rem;">No hay registros que coincidan con los filtros.</td></tr>`
                  : rows
                      .map(
                        (r) => `
                    <tr>
                      <td><strong>${formatFechaDDMMYYYY(r.fecha)}</strong></td>
                      <td><strong>${escapeHtml(r.codigo_equipo)}</strong></td>
                      <td><span class="badge-pill ${
                        r.turno === 'DIA' ? 'badge-dia' : 'badge-noche'
                      }">${r.turno === 'DIA' ? '☀️ DÍA' : '🌙 NOCHE'}</span></td>
                      <td class="tabular">${Number(r.horometro_inicial).toFixed(1)}</td>
                      <td class="tabular">${Number(r.horometro_final).toFixed(1)}</td>
                      <td class="tabular"><strong>${Number(r.horas_trabajadas).toFixed(1)}</strong></td>
                      <td class="tabular"><strong>${Number(r.material_fracturado).toLocaleString('es-PE')}</strong></td>
                      <td class="tabular">${Number(r.rendimiento).toFixed(2)}</td>
                      <td>${escapeHtml(r.area_trabajo || '-')}</td>
                      <td>${escapeHtml(r.operador || '-')}</td>
                      <td>
                        <div class="inline-row">
                          <button type="button" class="btn-mine btn-sm btn-secondary" data-hist-edit="${escapeHtml(r.id)}">✏️</button>
                          <button type="button" class="btn-mine btn-sm btn-danger" data-hist-del="${escapeHtml(r.id)}">🗑️</button>
                        </div>
                      </td>
                    </tr>`
                      )
                      .join('')
              }
            </tbody>
            <tfoot>
              <tr>
                <td colspan="5">TOTALES FILTRADOS (${rows.length} turnos)</td>
                <td class="tabular">${totHoras.toFixed(1)} HM</td>
                <td class="tabular">${Math.round(totMat).toLocaleString('es-PE')} m³</td>
                <td class="tabular">${totRend.toFixed(2)} m³/HM</td>
                <td colspan="3"></td>
              </tr>
            </tfoot>
          </table>
        </div>
      </section>

      <div class="two-col-grid">
        <section class="panel">
          <div class="panel-header">
            <h2 class="panel-title">🚜 Catálogo de Equipos (${this.app.equipos.length})</h2>
          </div>
          <div class="inline-row" style="margin-bottom:0.75rem;">
            <input type="text" id="cat-new-eq-code" class="input-mine" placeholder="Código (ej. STREX04)" />
            <input type="text" id="cat-new-eq-desc" class="input-mine" placeholder="Descripción (opcional)" />
            <button type="button" class="btn-mine btn-primary" id="btn-cat-add-eq">+ Agregar</button>
          </div>
          <div class="table-scroll">
            <table class="data-table">
              <thead>
                <tr><th>Código</th><th>Descripción</th><th></th></tr>
              </thead>
              <tbody>
                ${this.app.equipos
                  .map(
                    (e) => `
                  <tr>
                    <td><strong>${escapeHtml(e.codigo)}</strong></td>
                    <td>${escapeHtml(e.categoria || '')}</td>
                    <td style="text-align:right;">
                      <button type="button" class="btn-mine btn-sm btn-danger" data-del-eq="${escapeHtml(e.codigo)}">🗑️</button>
                    </td>
                  </tr>`
                  )
                  .join('')}
              </tbody>
            </table>
          </div>
        </section>

        <section class="panel">
          <div class="panel-header">
            <h2 class="panel-title">📍 Catálogo de Áreas de Trabajo (${this.app.areas.length})</h2>
          </div>
          <div class="inline-row" style="margin-bottom:0.75rem;">
            <input type="text" id="cat-new-area" class="input-mine" placeholder="Área (ej. SULF 810, Rampa Norte)" />
            <button type="button" class="btn-mine btn-primary" id="btn-cat-add-area">+ Agregar</button>
          </div>
          <div class="table-scroll">
            <table class="data-table">
              <thead>
                <tr><th>Área de Trabajo</th><th></th></tr>
              </thead>
              <tbody>
                ${this.app.areas
                  .map(
                    (a) => `
                  <tr>
                    <td><strong>${escapeHtml(a.nombre)}</strong></td>
                    <td style="text-align:right;">
                      <button type="button" class="btn-mine btn-sm btn-danger" data-del-area="${escapeHtml(a.nombre)}">🗑️</button>
                    </td>
                  </tr>`
                  )
                  .join('')}
              </tbody>
            </table>
          </div>
        </section>
      </div>
    `;

    this.bindEvents(rows);
  }

  bindEvents(filteredRows) {
    const bindSelect = (sel, key) => {
      this.mountEl.querySelector(sel)?.addEventListener('change', (e) => {
        this.filters[key] = e.target.value;
        this.render();
      });
    };
    bindSelect('#flt-equipo', 'equipo');
    bindSelect('#flt-turno', 'turno');
    bindSelect('#flt-area', 'area');
    bindSelect('#flt-desde', 'desde');
    bindSelect('#flt-hasta', 'hasta');

    this.mountEl.querySelector('#flt-texto')?.addEventListener('input', (e) => {
      this.filters.texto = e.target.value;
      this.render();
    });

    this.mountEl.querySelector('#btn-clear-filters')?.addEventListener('click', () => {
      this.filters = { equipo: 'TODOS', turno: 'TODOS', area: 'TODAS', desde: '', hasta: '', texto: '' };
      this.render();
    });

    this.mountEl.querySelector('#btn-export-filtered')?.addEventListener('click', () => {
      const range = getDataDateRange(filteredRows);
      const res = exportProductionWorkbookToExcel({
        records: filteredRows,
        codigoEquipo: this.filters.equipo,
        fechaDesde: this.filters.desde || range.min,
        fechaHasta: this.filters.hasta || range.max
      });
      this.app.showToast('success', `Excel exportado (${res.fileName}) con ${res.totalRecords} turno(s).`);
    });

    this.mountEl.querySelectorAll('[data-hist-edit]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = btn.getAttribute('data-hist-edit');
        const rec = this.app.registros.find((r) => r.id === id);
        if (rec) this.app.editRecordInCaptura(rec);
      });
    });

    this.mountEl.querySelectorAll('[data-hist-del]').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const id = btn.getAttribute('data-hist-del');
        if (!confirm('¿Eliminar este registro de turno?')) return;
        await deleteShiftRecord(id);
        await this.app.reloadAllData();
        this.app.showToast('info', 'Registro eliminado.');
        this.render();
        this.app.triggerSilentSync();
      });
    });

    this.mountEl.querySelector('#btn-cat-add-eq')?.addEventListener('click', async () => {
      const code = this.mountEl.querySelector('#cat-new-eq-code')?.value || '';
      const desc = this.mountEl.querySelector('#cat-new-eq-desc')?.value || '';
      try {
        const rec = await createOrTouchEquipment(code, desc);
        await this.app.reloadAllData();
        this.app.showToast('success', `Equipo ${rec.codigo} guardado.`);
        this.render();
      } catch (err) {
        this.app.showToast('error', err.message);
      }
    });

    this.mountEl.querySelectorAll('[data-del-eq]').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const code = btn.getAttribute('data-del-eq');
        if (!confirm(`¿Quitar ${code} del catálogo de equipos?`)) return;
        await deleteEquipment(code);
        await this.app.reloadAllData();
        this.render();
      });
    });

    this.mountEl.querySelector('#btn-cat-add-area')?.addEventListener('click', async () => {
      const name = this.mountEl.querySelector('#cat-new-area')?.value || '';
      if (!name.trim()) return;
      const rec = await createOrTouchArea(name);
      await this.app.reloadAllData();
      this.app.showToast('success', `Área "${rec.nombre}" agregada.`);
      this.render();
    });

    this.mountEl.querySelectorAll('[data-del-area]').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const name = btn.getAttribute('data-del-area');
        if (!confirm(`¿Quitar "${name}" del catálogo de áreas?`)) return;
        await deleteArea(name);
        await this.app.reloadAllData();
        this.render();
      });
    });
  }
}
