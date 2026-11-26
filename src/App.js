/**
 * Controlador Principal — Control de Horómetros y Producción STRACON (PWA Offline-First)
 */

import { requestPersistentStorage } from './db/storageManager.js';
import { ensureInitialCatalogs, getAllEquipments, getAllAreas } from './services/catalogService.js';
import { ensureInitialDemoIfFresh, getAllRegistros } from './services/productionService.js';
import { pwaManager } from './services/pwaManager.js';
import { syncQueue } from './services/syncQueue.js';
import { supabase } from './config/supabaseClient.js';
import { getFechaHoyPeru, formatFechaDDMMYYYY } from './utils/datePeru.js';
import { escapeHtml } from './utils/codeNormalizer.js';
import { CapturaTurnoView } from './components/CapturaTurno.js';
import { MatrizControlView } from './components/MatrizControl.js';
import { HistorialCatalogoView } from './components/HistorialCatalogo.js';
import { GerenciaDashboardView } from './components/GerenciaDashboard.js';
import { ConfigModal } from './components/ConfigModal.js';

const THEME_STORAGE_KEY = 'stracon_produccion_theme_v1';

export class AppController {
  constructor(mountEl) {
    this.mountEl = mountEl;
    this.activeTab = 'captura'; // 'captura' | 'matriz' | 'historial'
    this.isNightTheme = localStorage.getItem(THEME_STORAGE_KEY) === 'night';
    this.zoomLevel = this.getDefaultZoom();
    this.isTopbarCollapsed = typeof window !== 'undefined' ? window.innerWidth <= 900 : false;
    this.registros = [];
    this.equipos = [];
    this.areas = [];
    this.toast = null;
    this.toastTimer = null;
    this.showConfig = false;
    this.views = {};
    this.configModal = null;
  }

  getDefaultZoom() {
    if (typeof window === 'undefined') return 0.92;
    if (window.innerWidth <= 480) return 0.84;
    if (window.innerWidth <= 900) return 0.88;
    return 0.94;
  }

  async start() {
    document.body.classList.toggle('theme-night', this.isNightTheme);
    document.documentElement.style.setProperty('--ui-zoom', String(this.zoomLevel));

    await requestPersistentStorage();
    await ensureInitialCatalogs();
    await ensureInitialDemoIfFresh();
    await this.reloadAllData();

    await syncQueue.init();
    try {
      await Promise.race([pwaManager.init(), new Promise((r) => setTimeout(r, 1000))]);
    } catch {
      // Continuar si SW restringido
    }

    this.renderShell();
    this.views = {
      captura: new CapturaTurnoView(this.mountEl.querySelector('#view-mount'), this),
      matriz: new MatrizControlView(this.mountEl.querySelector('#view-mount'), this),
      historial: new HistorialCatalogoView(this.mountEl.querySelector('#view-mount'), this),
      gerencia: new GerenciaDashboardView(this.mountEl.querySelector('#view-mount'), this)
    };
    this.configModal = new ConfigModal(this.mountEl.querySelector('#modal-mount'), this);

    pwaManager.subscribe(() => {
      this.updateTopbarStatus();
      this.renderGlobalBanners();
    });
    syncQueue.subscribe(() => this.updateTopbarStatus());

    await this.switchTab(this.activeTab);
    this.triggerSilentSync();
  }

  async reloadAllData() {
    this.registros = await getAllRegistros();
    this.equipos = await getAllEquipments();
    this.areas = await getAllAreas();
    this.updateTopbarStatus();
  }

  triggerSilentSync() {
    if (supabase.getConfig().isConfigured && pwaManager.getState().isOnline) {
      syncQueue.processQueue({ silent: true }).then(async (res) => {
        if (res && !res.skipped) {
          await this.reloadAllData();
          this.renderActiveView();
        }
      });
    }
  }

  showToast(type, message) {
    this.toast = { type, message };
    this.renderToast();
    clearTimeout(this.toastTimer);
    this.toastTimer = setTimeout(() => {
      this.toast = null;
      this.renderToast();
    }, 5500);
  }

  renderToast() {
    const el = this.mountEl.querySelector('#toast-mount');
    if (!el) return;
    if (!this.toast) {
      el.innerHTML = '';
      return;
    }
    el.innerHTML = `
      <div class="toast-banner ${this.toast.type}">
        <span>${escapeHtml(this.toast.message)}</span>
        <button type="button" class="btn-mine btn-sm btn-secondary" id="btn-close-toast">✖️</button>
      </div>
    `;
    el.querySelector('#btn-close-toast')?.addEventListener('click', () => {
      this.toast = null;
      this.renderToast();
    });
  }

  getSyncButtonMeta() {
    const pwa = pwaManager.getState();
    const cfg = supabase.getConfig();
    const sync = syncQueue.getState();
    const pendingCount = this.registros.filter(
      (r) => !r.es_demo && r.estado !== 'demo' && (r.estado === 'pendiente' || r.estado === 'error')
    ).length;

    if (sync.isSyncing) {
      return { cls: 'is-local', text: '🔄 Sincronizando...', disabled: true };
    }
    if (!pwa.isOnline) {
      return {
        cls: 'is-offline',
        text: pendingCount > 0 ? `🔴 Sin señal (${pendingCount} local)` : '🔴 Modo Sin Señal',
        disabled: false
      };
    }
    if (!cfg.isConfigured) {
      return {
        cls: 'is-local',
        text: pendingCount > 0 ? `☁️ Conectar Nube (${pendingCount})` : '☁️ Conectar Nube',
        disabled: false
      };
    }
    if (pendingCount > 0) {
      return { cls: 'is-pending', text: `⬆️ Subir (${pendingCount})`, disabled: false };
    }
    return { cls: 'is-synced', text: '☁️ Nube al día', disabled: false };
  }

  renderShell() {
    const fechaHoy = formatFechaDDMMYYYY(getFechaHoyPeru());
    const pwa = pwaManager.getState();
    const syncMeta = this.getSyncButtonMeta();

    this.mountEl.innerHTML = `
      <header class="topbar ${this.isTopbarCollapsed ? 'is-collapsed' : ''}" id="main-topbar">
        <div class="topbar-main-row">
          <div class="brand-block">
            <div class="brand-badge">STRACON</div>
            <div>
              <h1 class="brand-title">CONTROL DE HORÓMETROS Y PRODUCCIÓN</h1>
              <p class="brand-sub">Fractura de Roca • Operación Offline Día / Noche</p>
            </div>
          </div>

          <div class="topbar-actions">
            <button type="button" class="btn-mine btn-topbar-sync ${syncMeta.cls}" id="btn-topbar-sync" ${
              syncMeta.disabled ? 'disabled' : ''
            }>
              ${syncMeta.text}
            </button>
            <button type="button" class="btn-mine btn-sm btn-secondary" id="btn-toggle-theme" title="Cambiar tema Día/Noche">
              ${this.isNightTheme ? '☀️ Día' : '🌙 Noche'}
            </button>
            <button type="button" class="btn-mine btn-sm btn-secondary" id="btn-open-cfg" title="Configuración Supabase y Offline">
              ⚙️ Opciones
            </button>
            <button type="button" class="btn-mine btn-sm btn-navy" id="btn-collapse-topbar">
              ${this.isTopbarCollapsed ? '▾ Más' : '▴ Menos'}
            </button>
          </div>
        </div>

        <div class="topbar-collapsible">
          <div class="topbar-pills">
            <span class="status-pill">📅 Perú: <strong>${fechaHoy}</strong></span>
            <span class="status-pill ${pwa.isOnline ? 'online' : 'offline'}" id="pill-net-status">
              ${pwa.isOnline ? '🟢 Con Señal' : '🔴 Sin Señal (Guardado Local Activo)'}
            </span>
            <span class="status-pill" id="pill-reg-count">📋 ${this.registros.length} Turno(s)</span>
          </div>
          <div class="topbar-pills">
            <span class="status-pill">
              🔍 Zoom:
              <button type="button" class="btn-mine btn-sm btn-secondary" id="btn-zoom-out" style="min-height:26px; padding:0 0.45rem;">−</button>
              <strong id="label-zoom">${Math.round(this.zoomLevel * 100)}%</strong>
              <button type="button" class="btn-mine btn-sm btn-secondary" id="btn-zoom-in" style="min-height:26px; padding:0 0.45rem;">+</button>
            </span>
          </div>
        </div>
      </header>

      <main class="app-shell">
        <div id="global-banners-mount"></div>
        <div id="toast-mount"></div>

        <nav class="main-tabs" aria-label="Navegación Principal">
          <button type="button" class="tab-btn ${this.activeTab === 'captura' ? 'is-active' : ''}" data-tab="captura">
            📝 1. Registrar Turno
          </button>
          <button type="button" class="tab-btn ${this.activeTab === 'matriz' ? 'is-active' : ''}" data-tab="matriz">
            📊 2. Planilla Control (Excel)
          </button>
          <button type="button" class="tab-btn ${this.activeTab === 'historial' ? 'is-active' : ''}" data-tab="historial">
            🗂️ 3. Historial y Equipos
          </button>
          <button type="button" class="tab-btn ${this.activeTab === 'gerencia' ? 'is-active' : ''}" data-tab="gerencia">
            📈 4. Gerencia (Acumulado)
          </button>
        </nav>

        <div id="view-mount"></div>
      </main>

      <div id="modal-mount"></div>
    `;

    this.renderGlobalBanners();
    this.bindShellEvents();
  }

  renderGlobalBanners() {
    const mount = this.mountEl.querySelector('#global-banners-mount');
    if (!mount) return;
    const pwa = pwaManager.getState();
    const parts = [];

    if (!pwa.isOnline) {
      parts.push(`
        <div class="global-offline-banner">
          <span>🔴 MODO SIN SEÑAL ACTIVO — Puedes seguir registrando horómetros, material y exportando Excel sin internet.</span>
        </div>
      `);
    }
    if (pwa.updateAvailable) {
      parts.push(`
        <div class="global-update-banner">
          <span>✨ Hay una nueva versión de la aplicación disponible.</span>
          <button type="button" class="btn-mine btn-sm btn-primary" id="btn-apply-pwa-update">Actualizar ahora</button>
        </div>
      `);
    }
    mount.innerHTML = parts.join('');
    mount.querySelector('#btn-apply-pwa-update')?.addEventListener('click', () => pwaManager.applyUpdate());
  }

  updateTopbarStatus() {
    const syncBtn = this.mountEl.querySelector('#btn-topbar-sync');
    if (syncBtn) {
      const m = this.getSyncButtonMeta();
      syncBtn.className = `btn-mine btn-topbar-sync ${m.cls}`;
      syncBtn.textContent = m.text;
      syncBtn.disabled = m.disabled;
    }
    const netPill = this.mountEl.querySelector('#pill-net-status');
    const pwa = pwaManager.getState();
    if (netPill) {
      netPill.className = `status-pill ${pwa.isOnline ? 'online' : 'offline'}`;
      netPill.textContent = pwa.isOnline ? '🟢 Con Señal' : '🔴 Sin Señal (Guardado Local Activo)';
    }
    const countPill = this.mountEl.querySelector('#pill-reg-count');
    if (countPill) {
      countPill.textContent = `📋 ${this.registros.length} Turno(s)`;
    }
  }

  bindShellEvents() {
    this.mountEl.querySelectorAll('[data-tab]').forEach((btn) => {
      btn.addEventListener('click', () => this.switchTab(btn.getAttribute('data-tab')));
    });

    this.mountEl.querySelector('#btn-open-cfg')?.addEventListener('click', () => this.openConfigModal());

    this.mountEl.querySelector('#btn-topbar-sync')?.addEventListener('click', async () => {
      const cfg = supabase.getConfig();
      if (!cfg.isConfigured) {
        this.openConfigModal();
        return;
      }
      const res = await syncQueue.processQueue({ silent: true });
      await this.reloadAllData();
      this.renderActiveView();
      this.showToast(res?.ok ? 'success' : 'info', syncQueue.getState().lastStatusMessage);
    });

    this.mountEl.querySelector('#btn-toggle-theme')?.addEventListener('click', () => {
      this.isNightTheme = !this.isNightTheme;
      document.body.classList.toggle('theme-night', this.isNightTheme);
      localStorage.setItem(THEME_STORAGE_KEY, this.isNightTheme ? 'night' : 'day');
      const btn = this.mountEl.querySelector('#btn-toggle-theme');
      if (btn) btn.textContent = this.isNightTheme ? '☀️ Día' : '🌙 Noche';
    });

    this.mountEl.querySelector('#btn-collapse-topbar')?.addEventListener('click', () => {
      this.isTopbarCollapsed = !this.isTopbarCollapsed;
      const topbar = this.mountEl.querySelector('#main-topbar');
      topbar?.classList.toggle('is-collapsed', this.isTopbarCollapsed);
      const btn = this.mountEl.querySelector('#btn-collapse-topbar');
      if (btn) btn.textContent = this.isTopbarCollapsed ? '▾ Más' : '▴ Menos';
    });

    const applyZoom = (delta) => {
      this.zoomLevel = Math.max(0.65, Math.min(1.2, Math.round((this.zoomLevel + delta) * 100) / 100));
      document.documentElement.style.setProperty('--ui-zoom', String(this.zoomLevel));
      const lbl = this.mountEl.querySelector('#label-zoom');
      if (lbl) lbl.textContent = `${Math.round(this.zoomLevel * 100)}%`;
    };
    this.mountEl.querySelector('#btn-zoom-out')?.addEventListener('click', () => applyZoom(-0.06));
    this.mountEl.querySelector('#btn-zoom-in')?.addEventListener('click', () => applyZoom(0.06));
  }

  async switchTab(tabKey) {
    this.activeTab = tabKey;
    this.mountEl.querySelectorAll('[data-tab]').forEach((btn) => {
      btn.classList.toggle('is-active', btn.getAttribute('data-tab') === tabKey);
    });
    const view = this.views[tabKey];
    if (view) await view.mount();
  }

  renderActiveView() {
    const view = this.views[this.activeTab];
    if (view && typeof view.render === 'function') {
      view.render();
    }
  }

  async openCapturaForSlot(slot) {
    this.activeTab = 'captura';
    this.mountEl.querySelectorAll('[data-tab]').forEach((btn) => {
      btn.classList.toggle('is-active', btn.getAttribute('data-tab') === 'captura');
    });
    await this.views.captura.openPrefilledSlot(slot);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  async editRecordInCaptura(rec) {
    this.activeTab = 'captura';
    this.mountEl.querySelectorAll('[data-tab]').forEach((btn) => {
      btn.classList.toggle('is-active', btn.getAttribute('data-tab') === 'captura');
    });
    this.views.captura.loadRecordIntoForm(rec);
    this.views.captura.render();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  openConfigModal() {
    this.showConfig = true;
    this.configModal.render();
  }

  closeConfigModal() {
    this.showConfig = false;
    const m = this.mountEl.querySelector('#modal-mount');
    if (m) m.innerHTML = '';
  }
}
