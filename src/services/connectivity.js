/**
 * Verificación real de salida a internet y Supabase mediante AbortController (8 s).
 */

import { supabase } from '../config/supabaseClient.js';
import { pwaManager } from './pwaManager.js';

export const DEFAULT_PING_TIMEOUT_MS = 8000;

export async function verifyRealInternetConnection(timeoutMs = DEFAULT_PING_TIMEOUT_MS) {
  const pwaState = pwaManager.getState();
  if (!pwaState.isOnline) {
    return {
      online: false,
      reason: pwaState.simulatedAirplaneMode
        ? 'Modo Avión de prueba activado en la app.'
        : 'Sin conexión de red en el dispositivo.'
    };
  }

  const cfg = supabase.getConfig();
  if (!cfg.isConfigured) {
    return {
      online: true,
      cloudConfigured: false,
      reason: 'Dispositivo con señal, trabajando en modo local (Supabase aún no conectado).'
    };
  }

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetch(`${cfg.url}/auth/v1/settings`, {
      method: 'GET',
      headers: {
        apikey: cfg.anonKey,
        Authorization: `Bearer ${cfg.anonKey}`
      },
      signal: controller.signal,
      cache: 'no-store'
    });
    clearTimeout(timer);
    if (res.ok || res.status === 401) {
      return { online: true, cloudConfigured: true, reason: 'Conexión real confirmada con Supabase.' };
    }
    return { online: false, cloudConfigured: true, reason: `Supabase respondió HTTP ${res.status}.` };
  } catch (err) {
    clearTimeout(timer);
    return {
      online: false,
      cloudConfigured: true,
      reason:
        err?.name === 'AbortError'
          ? 'Red sin salida real a Supabase (tiempo de espera de 8 s agotado).'
          : 'No se pudo alcanzar Supabase.'
    };
  }
}
