/**
 * Gestor de PWA Instalable, Service Worker y Modo Avión (Offline-First).
 */

export const APP_SHELL_VERSION = 'v1.3.3-supabase-live';

class PwaManagerService {
  constructor() {
    this.deferredInstallPrompt = null;
    this.swRegistration = null;
    this.waitingWorker = null;
    this.updateAvailable = false;
    this.simulatedAirplaneMode = false;
    this.isInstalledStandalone = this.checkIsStandalone();
    this.swStatus = 'Verificando...';
    this.listeners = new Set();
    this.initialized = false;
  }

  checkIsStandalone() {
    if (typeof window === 'undefined') return false;
    const isStandalone = window.matchMedia && window.matchMedia('(display-mode: standalone)').matches;
    const isIos = window.navigator && window.navigator.standalone === true;
    return Boolean(isStandalone || isIos);
  }

  isEffectiveOnline() {
    if (this.simulatedAirplaneMode) return false;
    return typeof navigator !== 'undefined' ? navigator.onLine : true;
  }

  subscribe(listener) {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  notify() {
    const state = this.getState();
    for (const cb of this.listeners) {
      try {
        cb(state);
      } catch {
        // ignore
      }
    }
  }

  getState() {
    return {
      isOnline: this.isEffectiveOnline(),
      browserOnline: typeof navigator !== 'undefined' ? navigator.onLine : true,
      simulatedAirplaneMode: this.simulatedAirplaneMode,
      canInstallPrompt: Boolean(this.deferredInstallPrompt),
      isInstalledStandalone: this.checkIsStandalone(),
      updateAvailable: this.updateAvailable,
      swStatus: this.swStatus,
      cacheVersion: APP_SHELL_VERSION
    };
  }

  async init() {
    if (this.initialized) return this.getState();
    this.initialized = true;

    window.addEventListener('beforeinstallprompt', (e) => {
      e.preventDefault();
      this.deferredInstallPrompt = e;
      this.notify();
    });

    window.addEventListener('appinstalled', () => {
      this.deferredInstallPrompt = null;
      this.isInstalledStandalone = true;
      this.notify();
    });

    window.addEventListener('online', () => this.notify());
    window.addEventListener('offline', () => this.notify());

    await this.registerServiceWorker();
    return this.getState();
  }

  async registerServiceWorker() {
    if (!('serviceWorker' in navigator)) {
      this.swStatus = 'Service Worker no disponible en este navegador';
      this.notify();
      return;
    }
    try {
      const reg = await navigator.serviceWorker.register('./sw.js', { scope: './' });
      this.swRegistration = reg;
      this.swStatus = reg.active ? 'Activo (Listo para trabajar sin señal)' : 'Instalando caché offline...';
      if (reg.waiting) {
        this.waitingWorker = reg.waiting;
        this.updateAvailable = true;
      }
      reg.addEventListener('updatefound', () => {
        const nw = reg.installing;
        if (!nw) return;
        nw.addEventListener('statechange', () => {
          if (nw.state === 'installed' && navigator.serviceWorker.controller) {
            this.waitingWorker = nw;
            this.updateAvailable = true;
            this.swStatus = 'Nueva versión disponible';
            this.notify();
          } else if (nw.state === 'activated') {
            this.swStatus = 'Activo (Listo para trabajar sin señal)';
            this.notify();
          }
        });
      });
      this.notify();
    } catch (err) {
      this.swStatus = `Modo local (${err?.message || 'SW pendiente'})`;
      this.notify();
    }
  }

  setSimulatedAirplaneMode(enabled) {
    this.simulatedAirplaneMode = Boolean(enabled);
    this.notify();
  }

  async triggerInstallPrompt() {
    if (!this.deferredInstallPrompt) return { outcome: 'unavailable' };
    this.deferredInstallPrompt.prompt();
    const res = await this.deferredInstallPrompt.userChoice;
    if (res.outcome === 'accepted') this.deferredInstallPrompt = null;
    this.notify();
    return res;
  }

  applyUpdate() {
    if (this.waitingWorker) {
      this.waitingWorker.postMessage({ type: 'SKIP_WAITING' });
    }
    this.updateAvailable = false;
    this.notify();
    window.location.reload();
  }
}

export const pwaManager = new PwaManagerService();
