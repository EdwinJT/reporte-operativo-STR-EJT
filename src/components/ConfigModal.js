/**
 * Modal de Configuración: Conexión Supabase Cloud, Modo Offline / PWA y Datos de Ejemplo.
 */

import { supabase } from '../config/supabaseClient.js';
import { pwaManager } from '../services/pwaManager.js';
import { syncQueue } from '../services/syncQueue.js';
import { loadDemoData, clearDemoData } from '../services/productionService.js';
import { getStorageUsageEstimate } from '../db/storageManager.js';
import { escapeHtml } from '../utils/codeNormalizer.js';

export class ConfigModal {
  constructor(mountEl, app) {
    this.mountEl = mountEl;
    this.app = app;
    this.testFeedback = null;
  }

  async render() {
    const cfg = supabase.getConfig();
    const pwa = pwaManager.getState();
    const syncState = syncQueue.getState();
    const storage = await getStorageUsageEstimate();
    const demoCount = this.app.registros.filter((r) => r.es_demo || r.estado === 'demo').length;

    this.mountEl.innerHTML = `
      <div class="modal-backdrop" id="cfg-backdrop">
        <div class="modal-card" role="dialog" aria-modal="true">
          <div class="panel-header">
            <h2 class="panel-title">⚙️ Nube Supabase, Modo Offline (PWA) y Datos</h2>
            <button type="button" class="btn-mine btn-sm btn-secondary" id="btn-close-cfg">✖️ Cerrar</button>
          </div>

          <div class="modal-section">
            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:0.5rem; margin-bottom:0.5rem;">
              <strong>☁️ 1. Conexión a Supabase Cloud</strong>
              <span class="badge-pill ${cfg.isConfigured ? 'badge-sync-ok' : 'badge-sync-pending'}">
                ${cfg.isConfigured ? '✅ Conectado a Supabase' : '📱 Modo Local (Sin Supabase aún)'}
              </span>
            </div>
            <p style="font-size:0.85rem; margin:0 0 0.65rem; color:var(--text-muted);">
              La app funciona 100% sin internet guardando en IndexedDB. Cuando conectes tu proyecto Supabase
              (y ejecutes <code>supabase_setup.sql</code> en el SQL Editor), todos los dispositivos sincronizarán automáticamente.
            </p>
            <div class="form-grid">
              <div class="field-group full-span">
                <label class="field-label" for="cfg-supa-url">SUPABASE_URL</label>
                <input type="url" id="cfg-supa-url" class="input-mine" placeholder="https://tu-proyecto.supabase.co" value="${
                  cfg.isPlaceholder ? '' : escapeHtml(cfg.url)
                }" />
              </div>
              <div class="field-group full-span">
                <label class="field-label" for="cfg-supa-key">SUPABASE_ANON_KEY (solo clave pública anon)</label>
                <input type="text" id="cfg-supa-key" class="input-mine" placeholder="eyJhbGci..." value="${
                  cfg.isPlaceholder ? '' : escapeHtml(cfg.anonKey)
                }" />
              </div>
            </div>
            ${
              this.testFeedback
                ? `<div class="toast-banner ${this.testFeedback.ok ? 'success' : 'error'}" style="margin-top:0.65rem;">
                    <span>${escapeHtml(this.testFeedback.message)}</span>
                  </div>`
                : ''
            }
            <div class="inline-row" style="margin-top:0.75rem; flex-wrap:wrap;">
              <button type="button" class="btn-mine btn-sm btn-primary" id="btn-save-supa">🔗 Probar y Guardar Conexión</button>
              <button type="button" class="btn-mine btn-sm btn-navy" id="btn-sync-now-modal" ${
                !cfg.isConfigured ? 'disabled' : ''
              }>🔄 Sincronizar Ahora</button>
              ${
                cfg.isConfigured
                  ? `<button type="button" class="btn-mine btn-sm btn-danger" id="btn-clear-supa">Desconectar</button>`
                  : ''
              }
            </div>
            <p style="font-size:0.8rem; margin:0.5rem 0 0; color:var(--text-muted);">
              Estado sync: ${escapeHtml(syncState.lastStatusMessage)}
            </p>
          </div>

          <div class="modal-section">
            <strong>📡 2. Estado Offline (PWA) y Memoria del Dispositivo</strong>
            <div style="font-size:0.86rem; margin-top:0.45rem; display:grid; gap:0.3rem;">
              <div>• Service Worker: <strong>${escapeHtml(pwa.swStatus)}</strong></div>
              <div>• Almacenamiento persistente: <strong>${storage.isPersisted ? 'Activado ✅' : 'Estándar'}</strong> (${storage.usageMB} MB usados)</div>
              <div>• Registros en memoria local: <strong>${this.app.registros.length}</strong></div>
            </div>
            <div class="inline-row" style="margin-top:0.7rem; flex-wrap:wrap;">
              <button type="button" class="btn-mine btn-sm ${
                pwa.simulatedAirplaneMode ? 'btn-danger' : 'btn-secondary'
              }" id="btn-toggle-airplane">
                ${pwa.simulatedAirplaneMode ? '✈️ Modo Avión Prueba: ON (Desactivar)' : '✈️ Probar Modo Avión Offline'}
              </button>
              ${
                pwa.canInstallPrompt
                  ? `<button type="button" class="btn-mine btn-sm btn-excel" id="btn-install-pwa">📲 Instalar App en Dispositivo</button>`
                  : ''
              }
            </div>
          </div>

          <div class="modal-section" style="margin-bottom:0;">
            <strong>🧪 3. Datos de Ejemplo STRACON (STREX02)</strong>
            <p style="font-size:0.84rem; margin:0.35rem 0 0.65rem; color:var(--text-muted);">
              Actualmente tienes <strong>${demoCount}</strong> registro(s) de ejemplo cargados (basados en tu Excel Oct–Nov).
              Los ejemplos nunca se suben a Supabase; puedes limpiarlos cuando empieces tu operación real.
            </p>
            <div class="inline-row" style="flex-wrap:wrap;">
              <button type="button" class="btn-mine btn-sm btn-navy" id="btn-reload-demo">📥 Cargar Datos de Ejemplo</button>
              <button type="button" class="btn-mine btn-sm btn-danger" id="btn-clear-demo" ${
                demoCount === 0 ? 'disabled' : ''
              }>🗑️ Limpiar Solo Ejemplos</button>
            </div>
          </div>
        </div>
      </div>
    `;

    this.bindEvents();
  }

  bindEvents() {
    const close = () => this.app.closeConfigModal();

    this.mountEl.querySelector('#btn-close-cfg')?.addEventListener('click', close);
    this.mountEl.querySelector('#cfg-backdrop')?.addEventListener('click', (e) => {
      if (e.target.id === 'cfg-backdrop') close();
    });

    this.mountEl.querySelector('#btn-save-supa')?.addEventListener('click', async () => {
      const url = this.mountEl.querySelector('#cfg-supa-url')?.value || '';
      const key = this.mountEl.querySelector('#cfg-supa-key')?.value || '';
      try {
        supabase.setConfig(url, key);
        this.testFeedback = { ok: true, message: 'Probando conexión con Supabase...' };
        await this.render();
        const res = await supabase.testConnection();
        this.testFeedback = res;
        this.app.updateTopbarStatus();
        if (res.ok) {
          await syncQueue.processQueue({ silent: true });
          await this.app.reloadAllData();
          this.app.renderActiveView();
        }
        await this.render();
      } catch (err) {
        this.testFeedback = { ok: false, message: err.message };
        await this.render();
      }
    });

    this.mountEl.querySelector('#btn-sync-now-modal')?.addEventListener('click', async () => {
      const res = await syncQueue.processQueue({ silent: true });
      await this.app.reloadAllData();
      this.app.renderActiveView();
      this.testFeedback = {
        ok: Boolean(res?.ok),
        message: syncQueue.getState().lastStatusMessage
      };
      await this.render();
    });

    this.mountEl.querySelector('#btn-clear-supa')?.addEventListener('click', () => {
      supabase.clearConfig();
      this.testFeedback = { ok: true, message: 'Desconectado de Supabase. Trabajando en modo local.' };
      this.app.updateTopbarStatus();
      this.render();
    });

    this.mountEl.querySelector('#btn-toggle-airplane')?.addEventListener('click', () => {
      const cur = pwaManager.getState().simulatedAirplaneMode;
      pwaManager.setSimulatedAirplaneMode(!cur);
      this.app.updateTopbarStatus();
      this.render();
    });

    this.mountEl.querySelector('#btn-install-pwa')?.addEventListener('click', async () => {
      await pwaManager.triggerInstallPrompt();
      this.render();
    });

    this.mountEl.querySelector('#btn-reload-demo')?.addEventListener('click', async () => {
      const n = await loadDemoData();
      await this.app.reloadAllData();
      this.app.renderActiveView();
      this.app.showToast('success', `Se cargaron ${n} turnos de ejemplo STRACON (STREX02).`);
      this.render();
    });

    this.mountEl.querySelector('#btn-clear-demo')?.addEventListener('click', async () => {
      const n = await clearDemoData();
      await this.app.reloadAllData();
      this.app.renderActiveView();
      this.app.showToast('info', `Se eliminaron ${n} registros de ejemplo. Tus registros reales siguen intactos.`);
      this.render();
    });
  }
}
