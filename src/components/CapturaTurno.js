/**
 * Vista 1 — Captura Rápida de Turno (Horómetros, Material Fracturado y Área de Trabajo)
 */

import {
  computeShiftMetrics,
  findExistingShift,
  getSuggestedInitialHorometer,
  saveShiftRecord,
  deleteShiftRecord
} from '../services/productionService.js';
import { createOrTouchEquipment } from '../services/catalogService.js';
import {
  getFechaHoyPeru,
  getFechaAyerPeru,
  getTurnoActualPeru,
  formatFechaDDMMYYYY
} from '../utils/datePeru.js';
import { escapeHtml } from '../utils/codeNormalizer.js';
import { formatShiftWhatsAppReport, copyTextToClipboard } from '../utils/whatsappReport.js';

const LAST_OPERATOR_KEY = 'stracon_produccion_last_operator_v1';
const LAST_EQUIPO_KEY = 'stracon_produccion_last_equipo_v1';

export class CapturaTurnoView {
  constructor(mountEl, app) {
    this.mountEl = mountEl;
    this.app = app;
    this.editingId = null;
    this.form = this.createDefaultForm();
    this.suggestion = null;
    this.existingSlotRecord = null;
    this.showNewEquipInput = false;
    this.lastSavedRecord = null;
  }

  createDefaultForm(overrides = {}) {
    const savedEq = localStorage.getItem(LAST_EQUIPO_KEY);
    const preferredEq =
      savedEq && this.app.equipos.some((e) => e.codigo === savedEq)
        ? savedEq
        : this.app.equipos.find((e) => this.app.registros.some((r) => r.codigo_equipo === e.codigo))?.codigo ||
          this.app.equipos[0]?.codigo ||
          'STREX02';
    return {
      codigo_equipo: overrides.codigo_equipo || preferredEq,
      fecha: overrides.fecha || getFechaHoyPeru(),
      turno: overrides.turno || getTurnoActualPeru(),
      horometro_inicial: overrides.horometro_inicial ?? '',
      horometro_final: overrides.horometro_final ?? '',
      material_fracturado: overrides.material_fracturado ?? '',
      area_trabajo: overrides.area_trabajo ?? '',
      operador: overrides.operador ?? (localStorage.getItem(LAST_OPERATOR_KEY) || ''),
      observaciones: overrides.observaciones ?? ''
    };
  }

  async openPrefilledSlot({ codigoEquipo, fecha, turno }) {
    const existing = await findExistingShift(codigoEquipo, fecha, turno);
    if (existing) {
      this.loadRecordIntoForm(existing);
    } else {
      this.editingId = null;
      this.form = this.createDefaultForm({ codigo_equipo: codigoEquipo, fecha, turno });
      await this.refreshHints();
      if (this.suggestion) {
        this.form.horometro_inicial = String(this.suggestion.suggestedValue);
      }
    }
    this.render();
  }

  loadRecordIntoForm(rec) {
    this.editingId = rec.id;
    this.form = {
      codigo_equipo: rec.codigo_equipo,
      fecha: rec.fecha,
      turno: rec.turno,
      horometro_inicial: String(rec.horometro_inicial),
      horometro_final: String(rec.horometro_final),
      material_fracturado: String(rec.material_fracturado),
      area_trabajo: rec.area_trabajo || '',
      operador: rec.operador || '',
      observaciones: rec.observaciones || ''
    };
    this.existingSlotRecord = null;
    this.suggestion = null;
  }

  async refreshHints() {
    this.existingSlotRecord = !this.editingId
      ? await findExistingShift(this.form.codigo_equipo, this.form.fecha, this.form.turno)
      : null;

    this.suggestion = await getSuggestedInitialHorometer(
      this.form.codigo_equipo,
      this.form.fecha,
      this.form.turno,
      this.editingId
    );
  }

  async mount() {
    await this.refreshHints();
    if (!this.editingId && !this.form.horometro_inicial && this.suggestion) {
      this.form.horometro_inicial = String(this.suggestion.suggestedValue);
    }
    this.render();
  }

  render() {
    const metrics = computeShiftMetrics(
      this.form.horometro_inicial,
      this.form.horometro_final,
      this.form.material_fracturado
    );
    const isNegativeDelta =
      this.form.horometro_inicial !== '' &&
      this.form.horometro_final !== '' &&
      Number(this.form.horometro_final) < Number(this.form.horometro_inicial);

    const equipOptionsHtml = this.app.equipos
      .map(
        (eq) =>
          `<option value="${escapeHtml(eq.codigo)}" ${
            eq.codigo === this.form.codigo_equipo ? 'selected' : ''
          }>${escapeHtml(eq.codigo)}${eq.categoria ? ` — ${escapeHtml(eq.categoria)}` : ''}</option>`
      )
      .join('');

    const areaChipsHtml = this.app.areas
      .map(
        (a) =>
          `<button type="button" class="chip-btn ${
            this.form.area_trabajo === a.nombre ? 'is-active' : ''
          }" data-pick-area="${escapeHtml(a.nombre)}">${escapeHtml(a.nombre)}</button>`
      )
      .join('');

    const recentRows = this.app.registros.slice(0, 12);

    this.mountEl.innerHTML = `
      <section class="panel">
        <div class="panel-header">
          <h2 class="panel-title">
            ${this.editingId ? '✏️ Editando Registro de Turno' : '📝 Registrar Horómetros y Producción'}
          </h2>
          ${
            this.editingId
              ? `<button type="button" class="btn-mine btn-sm btn-secondary" id="btn-cancel-edit">✖️ Cancelar Edición</button>`
              : `<span class="badge-pill badge-sync-ok">💾 Guardado 100% Offline (Funciona sin señal)</span>`
          }
        </div>

        ${
          this.existingSlotRecord
            ? `<div class="duplicate-banner">
                <span>
                  ⚠️ Ya existe un registro para <strong>${escapeHtml(this.form.codigo_equipo)}</strong> el
                  <strong>${formatFechaDDMMYYYY(this.form.fecha)}</strong> (${
                this.form.turno === 'DIA' ? 'TURNO DÍA' : 'TURNO NOCHE'
              }):
                  ${this.existingSlotRecord.horometro_inicial} → ${this.existingSlotRecord.horometro_final} HM
                  (${this.existingSlotRecord.material_fracturado} m³).
                </span>
                <button type="button" class="btn-mine btn-sm btn-navy" id="btn-load-existing-slot">
                  ✏️ Cargar para editar
                </button>
              </div>`
            : ''
        }

        <form id="form-captura-turno" novalidate>
          <div class="form-grid">
            <div class="field-group">
              <label class="field-label" for="inp-equipo">1. Equipo</label>
              <div class="inline-row">
                <select id="inp-equipo" class="input-mine">${equipOptionsHtml}</select>
                <button type="button" class="btn-mine btn-sm btn-secondary" id="btn-toggle-new-equip" title="Agregar nuevo equipo">+ Nuevo</button>
              </div>
              ${
                this.showNewEquipInput
                  ? `<div class="inline-row" style="margin-top:0.35rem;">
                      <input type="text" id="inp-new-equip-code" class="input-mine" placeholder="Ej. STREX04" />
                      <button type="button" class="btn-mine btn-sm btn-primary" id="btn-save-new-equip">Agregar</button>
                    </div>`
                  : ''
              }
            </div>

            <div class="field-group">
              <label class="field-label" for="inp-fecha">2. Fecha (Hora Perú)</label>
              <input type="date" id="inp-fecha" class="input-mine" value="${escapeHtml(this.form.fecha)}" required />
              <div class="chip-row">
                <button type="button" class="chip-btn" id="btn-date-hoy">Hoy</button>
                <button type="button" class="chip-btn" id="btn-date-ayer">Ayer</button>
              </div>
            </div>

            <div class="field-group">
              <span class="field-label">3. Turno</span>
              <div class="turno-toggle-group">
                <button type="button" class="turno-btn is-dia ${this.form.turno === 'DIA' ? 'is-active' : ''}" data-set-turno="DIA">
                  ☀️ TURNO DÍA
                </button>
                <button type="button" class="turno-btn is-noche ${this.form.turno === 'NOCHE' ? 'is-active' : ''}" data-set-turno="NOCHE">
                  🌙 TURNO NOCHE
                </button>
              </div>
            </div>

            <div class="field-group">
              <label class="field-label" for="inp-hm-inicial">4. Horómetro Inicial</label>
              <input
                type="number"
                step="0.1"
                min="0"
                inputmode="decimal"
                id="inp-hm-inicial"
                class="input-mine input-numeric-big"
                placeholder="Ej. 2997.2"
                value="${escapeHtml(this.form.horometro_inicial)}"
                required
              />
              ${
                this.suggestion
                  ? `<div class="suggest-banner">
                      <span>💡 ${escapeHtml(this.suggestion.sourceLabel)} cerró en <strong>${this.suggestion.suggestedValue.toFixed(1)}</strong></span>
                      <button type="button" class="chip-btn" id="btn-apply-hm-suggest">Usar ${this.suggestion.suggestedValue.toFixed(1)}</button>
                    </div>`
                  : ''
              }
            </div>

            <div class="field-group">
              <label class="field-label" for="inp-hm-final">5. Horómetro Final</label>
              <input
                type="number"
                step="0.1"
                min="0"
                inputmode="decimal"
                id="inp-hm-final"
                class="input-mine input-numeric-big"
                placeholder="Ej. 3007.2"
                value="${escapeHtml(this.form.horometro_final)}"
                required
              />
            </div>

            <div class="field-group">
              <label class="field-label" for="inp-material">6. Material Fracturado (m³)</label>
              <input
                type="number"
                step="1"
                min="0"
                inputmode="decimal"
                id="inp-material"
                class="input-mine input-numeric-big"
                placeholder="Ej. 180"
                value="${escapeHtml(this.form.material_fracturado)}"
              />
            </div>

            <div class="field-group full-span">
              <label class="field-label" for="inp-area">7. Área de Trabajo</label>
              <input
                type="text"
                id="inp-area"
                class="input-mine"
                placeholder="Ej. SULF 797, SULF 780, Stock CH SULF..."
                value="${escapeHtml(this.form.area_trabajo)}"
              />
              <div class="chip-row">${areaChipsHtml}</div>
            </div>

            <div class="field-group">
              <label class="field-label" for="inp-operador">Operador / Responsable (opcional)</label>
              <input
                type="text"
                id="inp-operador"
                class="input-mine"
                placeholder="Nombre del operador"
                value="${escapeHtml(this.form.operador)}"
              />
            </div>

            <div class="field-group">
              <label class="field-label" for="inp-obs">Observaciones (opcional)</label>
              <input
                type="text"
                id="inp-obs"
                class="input-mine"
                placeholder="Ej. Cambio de punta, espera voladura..."
                value="${escapeHtml(this.form.observaciones)}"
              />
            </div>
          </div>

          <div class="live-calc-strip" id="live-calc-strip">
            <div>
              <div class="live-calc-item-label">Horas Trabajadas Calculadas</div>
              <div class="live-calc-item-val tabular ${isNegativeDelta ? 'is-error' : ''}" id="live-horas-val">
                ${isNegativeDelta ? '⚠️ HM Final < Inicial' : `${metrics.horas.toFixed(1)} HM`}
              </div>
            </div>
            <div>
              <div class="live-calc-item-label">Rendimiento Calculado</div>
              <div class="live-calc-item-val tabular" id="live-rend-val">
                ${metrics.rendimiento.toFixed(2)} m³/HM
              </div>
            </div>
            <div>
              <div class="live-calc-item-label">Estado de Guardado</div>
              <div style="font-weight:800; font-size:0.95rem; margin-top:0.25rem;">
                ${
                  metrics.horas > 13
                    ? '⚠️ Más de 13 horas en un turno (verifica los horómetros)'
                    : '✅ Se guarda al instante en el dispositivo'
                }
              </div>
            </div>
          </div>

          <div class="inline-row" style="flex-wrap:wrap;">
            <button type="submit" class="btn-mine btn-primary" style="flex:1 1 280px; min-height:54px; font-size:1.04rem;">
              ${this.editingId ? '💾 ACTUALIZAR TURNO' : '💾 GUARDAR TURNO (Offline Ready)'}
            </button>
            <button type="button" class="btn-mine btn-whatsapp" id="btn-copy-whatsapp-form" style="flex:0 1 260px; min-height:54px; font-size:1rem;">
              💬 Copiar para WhatsApp
            </button>
          </div>
        </form>
      </section>

      <section class="panel">
        <div class="panel-header">
          <h2 class="panel-title">📋 Últimos Turnos Registrados (${recentRows.length})</h2>
          <button type="button" class="btn-mine btn-sm btn-navy" id="btn-goto-matrix">
            📊 Ver Planilla Excel Completa
          </button>
        </div>
        ${
          recentRows.length === 0
            ? `<p style="font-weight:700; color:var(--text-muted);">Aún no hay registros. Completa el formulario superior para agregar el primero.</p>`
            : `<div class="table-scroll">
                <table class="data-table">
                  <thead>
                    <tr>
                      <th>Fecha</th>
                      <th>Equipo</th>
                      <th>Turno</th>
                      <th>HM Inicial</th>
                      <th>HM Final</th>
                      <th>Horas</th>
                      <th>Material (m³)</th>
                      <th>Rend. (m³/HM)</th>
                      <th>Área</th>
                      <th>Estado</th>
                      <th>Acciones</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${recentRows
                      .map((r) => {
                        const syncBadge =
                          r.estado === 'sincronizado'
                            ? `<span class="badge-pill badge-sync-ok">☁️ Nube</span>`
                            : r.estado === 'demo'
                              ? `<span class="badge-pill badge-sync-demo">🧪 Ejemplo</span>`
                              : r.estado === 'error'
                                ? `<span class="badge-pill badge-sync-err" title="${escapeHtml(r.error_sync || 'Error al sincronizar')}">⚠️ Reintento</span>`
                                : `<span class="badge-pill badge-sync-pending">📱 Local</span>`;
                        return `
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
                          <td>${syncBadge}</td>
                          <td>
                            <div class="inline-row">
                              <button type="button" class="btn-mine btn-sm btn-whatsapp" data-wa-id="${escapeHtml(r.id)}" title="Copiar reporte para WhatsApp">💬</button>
                              <button type="button" class="btn-mine btn-sm btn-secondary" data-edit-id="${escapeHtml(r.id)}" title="Editar">✏️</button>
                              <button type="button" class="btn-mine btn-sm btn-danger" data-delete-id="${escapeHtml(r.id)}" title="Eliminar">🗑️</button>
                            </div>
                          </td>
                        </tr>`;
                      })
                      .join('')}
                  </tbody>
                </table>
              </div>`
        }
      </section>
    `;

    this.bindEvents();
  }

  updateLiveMetricsOnly() {
    const metrics = computeShiftMetrics(
      this.form.horometro_inicial,
      this.form.horometro_final,
      this.form.material_fracturado
    );
    const isNeg =
      this.form.horometro_inicial !== '' &&
      this.form.horometro_final !== '' &&
      Number(this.form.horometro_final) < Number(this.form.horometro_inicial);

    const hEl = this.mountEl.querySelector('#live-horas-val');
    const rEl = this.mountEl.querySelector('#live-rend-val');
    if (hEl) {
      hEl.textContent = isNeg ? '⚠️ HM Final < Inicial' : `${metrics.horas.toFixed(1)} HM`;
      hEl.classList.toggle('is-error', isNeg);
    }
    if (rEl) {
      rEl.textContent = `${metrics.rendimiento.toFixed(2)} m³/HM`;
    }
  }

  bindEvents() {
    const inpEquipo = this.mountEl.querySelector('#inp-equipo');
    const inpFecha = this.mountEl.querySelector('#inp-fecha');
    const inpHi = this.mountEl.querySelector('#inp-hm-inicial');
    const inpHf = this.mountEl.querySelector('#inp-hm-final');
    const inpMat = this.mountEl.querySelector('#inp-material');
    const inpArea = this.mountEl.querySelector('#inp-area');
    const inpOp = this.mountEl.querySelector('#inp-operador');
    const inpObs = this.mountEl.querySelector('#inp-obs');

    inpEquipo?.addEventListener('change', async () => {
      this.form.codigo_equipo = inpEquipo.value;
      localStorage.setItem(LAST_EQUIPO_KEY, this.form.codigo_equipo);
      await this.refreshHints();
      if (!this.editingId && this.suggestion) {
        this.form.horometro_inicial = String(this.suggestion.suggestedValue);
      }
      this.render();
    });

    inpFecha?.addEventListener('change', async () => {
      this.form.fecha = inpFecha.value;
      await this.refreshHints();
      if (!this.editingId && this.suggestion) {
        this.form.horometro_inicial = String(this.suggestion.suggestedValue);
      }
      this.render();
    });

    this.mountEl.querySelector('#btn-date-hoy')?.addEventListener('click', async () => {
      this.form.fecha = getFechaHoyPeru();
      await this.refreshHints();
      this.render();
    });

    this.mountEl.querySelector('#btn-date-ayer')?.addEventListener('click', async () => {
      this.form.fecha = getFechaAyerPeru();
      await this.refreshHints();
      this.render();
    });

    this.mountEl.querySelectorAll('[data-set-turno]').forEach((btn) => {
      btn.addEventListener('click', async () => {
        this.form.turno = btn.getAttribute('data-set-turno');
        await this.refreshHints();
        if (!this.editingId && this.suggestion) {
          this.form.horometro_inicial = String(this.suggestion.suggestedValue);
        }
        this.render();
      });
    });

    inpHi?.addEventListener('input', () => {
      this.form.horometro_inicial = inpHi.value;
      this.updateLiveMetricsOnly();
    });

    inpHf?.addEventListener('input', () => {
      this.form.horometro_final = inpHf.value;
      this.updateLiveMetricsOnly();
    });

    inpMat?.addEventListener('input', () => {
      this.form.material_fracturado = inpMat.value;
      this.updateLiveMetricsOnly();
    });

    inpArea?.addEventListener('input', () => {
      this.form.area_trabajo = inpArea.value;
    });

    inpOp?.addEventListener('input', () => {
      this.form.operador = inpOp.value;
    });

    inpObs?.addEventListener('input', () => {
      this.form.observaciones = inpObs.value;
    });

    this.mountEl.querySelectorAll('[data-pick-area]').forEach((btn) => {
      btn.addEventListener('click', () => {
        this.form.area_trabajo = btn.getAttribute('data-pick-area');
        if (inpArea) inpArea.value = this.form.area_trabajo;
        this.mountEl.querySelectorAll('[data-pick-area]').forEach((b) => {
          b.classList.toggle('is-active', b.getAttribute('data-pick-area') === this.form.area_trabajo);
        });
      });
    });

    this.mountEl.querySelector('#btn-apply-hm-suggest')?.addEventListener('click', () => {
      if (this.suggestion) {
        this.form.horometro_inicial = String(this.suggestion.suggestedValue);
        if (inpHi) inpHi.value = this.form.horometro_inicial;
        this.updateLiveMetricsOnly();
      }
    });

    this.mountEl.querySelector('#btn-load-existing-slot')?.addEventListener('click', () => {
      if (this.existingSlotRecord) {
        this.loadRecordIntoForm(this.existingSlotRecord);
        this.render();
      }
    });

    this.mountEl.querySelector('#btn-cancel-edit')?.addEventListener('click', async () => {
      this.editingId = null;
      this.form = this.createDefaultForm({ codigo_equipo: this.form.codigo_equipo });
      await this.refreshHints();
      this.render();
    });

    this.mountEl.querySelector('#btn-toggle-new-equip')?.addEventListener('click', () => {
      this.showNewEquipInput = !this.showNewEquipInput;
      this.render();
    });

    this.mountEl.querySelector('#btn-save-new-equip')?.addEventListener('click', async () => {
      const codeInput = this.mountEl.querySelector('#inp-new-equip-code');
      try {
        const rec = await createOrTouchEquipment(codeInput?.value || '');
        this.showNewEquipInput = false;
        this.form.codigo_equipo = rec.codigo;
        await this.app.reloadAllData();
        await this.refreshHints();
        this.app.showToast('success', `Equipo ${rec.codigo} agregado al catálogo.`);
        this.render();
      } catch (err) {
        this.app.showToast('error', err.message);
      }
    });

    this.mountEl.querySelector('#form-captura-turno')?.addEventListener('submit', async (e) => {
      e.preventDefault();
      try {
        if (this.form.operador.trim()) {
          localStorage.setItem(LAST_OPERATOR_KEY, this.form.operador.trim());
        }
        const saved = await saveShiftRecord({
          ...this.form,
          id: this.editingId || undefined
        });
        const wasEdit = Boolean(this.editingId);
        this.lastSavedRecord = saved;
        this.editingId = null;
        this.form = this.createDefaultForm({
          codigo_equipo: saved.codigo_equipo,
          fecha: saved.fecha,
          turno: saved.turno === 'DIA' ? 'NOCHE' : 'DIA',
          horometro_inicial: String(saved.horometro_final),
          area_trabajo: saved.area_trabajo
        });
        await this.app.reloadAllData();
        await this.refreshHints();
        this.app.showToast(
          'success',
          `${wasEdit ? 'Turno actualizado' : 'Turno guardado'}: ${saved.codigo_equipo} (${saved.fecha} ${
            saved.turno
          }) — ${saved.horas_trabajadas.toFixed(1)} HM | ${saved.material_fracturado} m³`
        );
        this.render();
        this.app.triggerSilentSync();
      } catch (err) {
        this.app.showToast('error', err.message);
      }
    });

    this.mountEl.querySelector('#btn-copy-whatsapp-form')?.addEventListener('click', async () => {
      const hasFormValues =
        this.form.horometro_inicial !== '' &&
        this.form.horometro_final !== '' &&
        Number(this.form.horometro_final) >= Number(this.form.horometro_inicial);

      const targetRecord = hasFormValues
        ? {
            ...this.form,
            horas_trabajadas: computeShiftMetrics(
              this.form.horometro_inicial,
              this.form.horometro_final,
              this.form.material_fracturado
            ).horas
          }
        : this.lastSavedRecord || this.app.registros[0] || null;

      if (!targetRecord) {
        this.app.showToast('error', 'Completa Horómetro Inicial y Final o guarda un turno para copiarlo a WhatsApp.');
        return;
      }

      const text = formatShiftWhatsAppReport(targetRecord);
      const copied = await copyTextToClipboard(text);
      const turnoTxt = targetRecord.turno === 'NOCHE' ? 'Noche' : 'Día';
      this.app.showToast(
        copied ? 'success' : 'error',
        copied
          ? `💬 Reporte copiado para WhatsApp: ${targetRecord.codigo_equipo} (${formatFechaDDMMYYYY(
              targetRecord.fecha
            )} — Turno ${turnoTxt}). ¡Listo para pegar!`
          : 'No se pudo copiar al portapapeles. Intenta nuevamente.'
      );
    });

    this.mountEl.querySelectorAll('[data-wa-id]').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const id = btn.getAttribute('data-wa-id');
        const rec = this.app.registros.find((r) => r.id === id);
        if (!rec) return;
        const text = formatShiftWhatsAppReport(rec);
        const copied = await copyTextToClipboard(text);
        const turnoTxt = rec.turno === 'NOCHE' ? 'Noche' : 'Día';
        this.app.showToast(
          copied ? 'success' : 'error',
          copied
            ? `💬 Reporte copiado para WhatsApp: ${rec.codigo_equipo} (${formatFechaDDMMYYYY(
                rec.fecha
              )} — Turno ${turnoTxt}).`
            : 'No se pudo copiar al portapapeles.'
        );
      });
    });

    this.mountEl.querySelectorAll('[data-edit-id]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = btn.getAttribute('data-edit-id');
        const rec = this.app.registros.find((r) => r.id === id);
        if (rec) {
          this.loadRecordIntoForm(rec);
          this.render();
          window.scrollTo({ top: 0, behavior: 'smooth' });
        }
      });
    });

    this.mountEl.querySelectorAll('[data-delete-id]').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const id = btn.getAttribute('data-delete-id');
        const rec = this.app.registros.find((r) => r.id === id);
        if (!rec) return;
        if (!confirm(`¿Eliminar el registro de ${rec.codigo_equipo} (${rec.fecha} ${rec.turno})?`)) return;
        await deleteShiftRecord(id);
        await this.app.reloadAllData();
        await this.refreshHints();
        this.app.showToast('info', 'Registro eliminado.');
        this.render();
        this.app.triggerSilentSync();
      });
    });

    this.mountEl.querySelector('#btn-goto-matrix')?.addEventListener('click', () => {
      this.app.switchTab('matriz');
    });
  }
}
