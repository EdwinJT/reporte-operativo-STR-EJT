/**
 * Motor de Sincronización Secuencial e Idempotente con Supabase Cloud.
 * - Funciona en segundo plano sin bloquear nunca la captura offline.
 * - Omite los datos de ejemplo (es_demo === true) para no ensuciar la base real.
 * - Soporta Exponential Backoff (30s -> 2m -> 10m) y mutex anti-colisión.
 */

import { localDb } from '../db/localDb.js';
import { supabase } from '../config/supabaseClient.js';
import { pwaManager } from './pwaManager.js';
import { verifyRealInternetConnection } from './connectivity.js';
import { syncCatalogsFromRecords } from './catalogService.js';

const PERIODIC_SYNC_MS = 2 * 60 * 1000;
const BACKOFF_MS = [30000, 120000, 600000];

class SyncQueueService {
  constructor() {
    this.isSyncing = false;
    this.lastRunAt = null;
    this.lastStatusMessage = 'Modo offline local listo';
    this.lastError = null;
    this.periodicTimer = null;
    this.retryTimer = null;
    this.attemptCount = 0;
    this.initialized = false;
    this.listeners = new Set();
  }

  subscribe(cb) {
    this.listeners.add(cb);
    return () => this.listeners.delete(cb);
  }

  notify() {
    const st = this.getState();
    for (const cb of this.listeners) {
      try {
        cb(st);
      } catch {
        // ignore
      }
    }
  }

  getState() {
    const cfg = supabase.getConfig();
    return {
      isSyncing: this.isSyncing,
      isCloudConfigured: cfg.isConfigured,
      lastRunAt: this.lastRunAt,
      lastStatusMessage: this.lastStatusMessage,
      lastError: this.lastError
    };
  }

  async init() {
    if (this.initialized) return;
    this.initialized = true;

    // Recuperar registros que quedaron en 'sincronizando' tras cierre abrupto
    const all = await localDb.registros.toArray();
    for (const r of all) {
      if (r.estado === 'sincronizando') {
        await localDb.registros.put({ ...r, estado: 'pendiente' });
      }
    }

    if (typeof document !== 'undefined') {
      document.addEventListener('visibilitychange', () => {
        if (document.visibilityState === 'visible') {
          this.processQueue({ silent: true });
        }
      });
    }

    if (typeof window !== 'undefined') {
      window.addEventListener('online', () => {
        this.processQueue({ silent: true });
      });
    }

    this.periodicTimer = setInterval(() => {
      this.processQueue({ silent: true });
    }, PERIODIC_SYNC_MS);
  }

  async processQueue({ silent = false } = {}) {
    if (this.isSyncing) return { skipped: true, reason: 'already_syncing' };

    const cfg = supabase.getConfig();
    if (!cfg.isConfigured) {
      this.lastStatusMessage = 'Trabajando 100% en memoria local (Supabase aún sin conectar)';
      this.notify();
      return { skipped: true, reason: 'not_configured' };
    }

    const net = await verifyRealInternetConnection();
    if (!net.online) {
      this.lastStatusMessage = net.reason;
      this.notify();
      return { skipped: true, reason: 'offline', detail: net.reason };
    }

    this.isSyncing = true;
    this.lastError = null;
    this.lastStatusMessage = 'Sincronizando turnos con Supabase...';
    this.notify();

    try {
      const allLocal = await localDb.registros.toArray();
      const pending = allLocal
        .filter((r) => !r.es_demo && r.estado !== 'demo' && (r.estado === 'pendiente' || r.estado === 'error'))
        .sort((a, b) => (a.creado_en || '').localeCompare(b.creado_en || ''));

      let uploaded = 0;
      let failed = 0;
      let firstErrorMessage = '';

      for (const item of pending) {
        await localDb.registros.put({ ...item, estado: 'sincronizando' });
        const { data: syncedRows, error } = await supabase.upsertRegistro(item);
        if (error) {
          failed++;
          if (!firstErrorMessage) firstErrorMessage = error.message;
          await localDb.registros.put({
            ...item,
            estado: 'error',
            error_sync: error.message
          });
        } else {
          uploaded++;
          const serverRow = Array.isArray(syncedRows) && syncedRows.length > 0 ? syncedRows[0] : null;
          if (item.eliminado) {
            await localDb.registros.delete(item.id);
          } else if (serverRow && serverRow.id && serverRow.id !== item.id) {
            await localDb.registros.delete(item.id);
            await localDb.registros.put({
              ...serverRow,
              estado: 'sincronizado',
              es_demo: false,
              eliminado: Boolean(serverRow.eliminado),
              error_sync: ''
            });
          } else {
            await localDb.registros.put({
              ...item,
              ...(serverRow || {}),
              estado: 'sincronizado',
              es_demo: false,
              eliminado: Boolean(serverRow?.eliminado ?? item.eliminado),
              error_sync: ''
            });
          }
        }
      }

      // Descargar registros remotos y fusionar sin sobrescribir cambios pendientes locales
      const { data: remoteRows, error: pullError } = await supabase.selectRegistros();
      let pulled = 0;
      if (!pullError && Array.isArray(remoteRows)) {
        const refreshedLocal = await localDb.registros.toArray();
        const localById = new Map(refreshedLocal.map((r) => [r.id, r]));
        const localBySlot = new Map(
          refreshedLocal.filter((r) => !r.eliminado).map((r) => [`${r.codigo_equipo}__${r.fecha}__${r.turno}`, r])
        );

        for (const remote of remoteRows) {
          const existing = localById.get(remote.id) || localBySlot.get(`${remote.codigo_equipo}__${remote.fecha}__${remote.turno}`);
          if (existing && (existing.estado === 'pendiente' || existing.estado === 'sincronizando')) {
            continue;
          }
          if (!existing || existing.id !== remote.id || (remote.actualizado_en || '') > (existing.actualizado_en || '')) {
            if (existing && existing.id !== remote.id) {
              await localDb.registros.delete(existing.id);
            }
            await localDb.registros.put({
              ...remote,
              estado: 'sincronizado',
              es_demo: false,
              eliminado: Boolean(remote.eliminado)
            });
            pulled++;
          }
        }
        await syncCatalogsFromRecords(remoteRows);
      }

      this.lastRunAt = new Date().toISOString();
      if (failed > 0) {
        this.attemptCount++;
        this.lastError = firstErrorMessage || 'Error al sincronizar';
        const delayMs = BACKOFF_MS[Math.min(this.attemptCount - 1, BACKOFF_MS.length - 1)];
        this.lastStatusMessage = `Subidos ${uploaded}, ${failed} con error (${this.lastError}). Reintento en ${Math.round(delayMs / 1000)} s.`;
        clearTimeout(this.retryTimer);
        this.retryTimer = setTimeout(() => this.processQueue({ silent: true }), delayMs);
      } else {
        this.attemptCount = 0;
        this.lastStatusMessage = `Sincronización al día (${uploaded} subidos, ${pulled} actualizados de nube).`;
      }

      return { ok: failed === 0, uploaded, pulled, failed };
    } catch (err) {
      this.lastError = err?.message || 'Error inesperado de sincronización';
      this.lastStatusMessage = this.lastError;
      if (!silent) throw err;
      return { ok: false, error: this.lastError };
    } finally {
      this.isSyncing = false;
      this.notify();
    }
  }
}

export const syncQueue = new SyncQueueService();
