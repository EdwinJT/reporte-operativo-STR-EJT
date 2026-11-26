/**
 * Cliente ligero y seguro de Supabase (PostgREST) sin dependencias externas.
 * - Solo utiliza VITE_SUPABASE_URL y VITE_SUPABASE_ANON_KEY (o config guardada en localStorage).
 * - Bloquea estrictamente claves privadas 'service_role'.
 */

const STORAGE_CONFIG_KEY = 'stracon_produccion_supabase_config_v1';
const PLACEHOLDER_URL = 'https://tu-proyecto.supabase.co';
const PLACEHOLDER_KEY = 'tu-clave-publica-anon';
const DEFAULT_SUPABASE_URL = 'https://menggkptiggdywblgbpk.supabase.co';
const DEFAULT_SUPABASE_ANON_KEY = 'sb_publishable_8C2cGY3cFv9UN5wFjAg0Gg_dZT_Zp8z';

export function validateAnonKeySecurity(keyString) {
  if (!keyString || typeof keyString !== 'string' || !keyString.trim()) {
    return {
      valid: false,
      isServiceRole: false,
      reason: 'La clave pública ANON está vacía.'
    };
  }

  const trimmed = keyString.trim();
  const lower = trimmed.toLowerCase();
  if (lower.includes('service_role') || lower.includes('service-role') || lower.includes('secret_key')) {
    return {
      valid: false,
      isServiceRole: true,
      reason: 'BLOQUEO DE SEGURIDAD: Se detectó una clave privada "service_role". Solo debes usar la clave pública ANON.'
    };
  }

  const parts = trimmed.split('.');
  if (parts.length >= 2) {
    try {
      const base64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
      const payload = JSON.parse(atob(base64));
      if (payload.role === 'service_role') {
        return {
          valid: false,
          isServiceRole: true,
          reason: 'BLOQUEO DE SEGURIDAD: El token JWT pertenece a "service_role". Usa únicamente la clave "anon".'
        };
      }
    } catch {
      // No es JWT estándar (ej. sb_publishable_*), permitir
    }
  }

  return {
    valid: true,
    isServiceRole: false,
    reason: 'Clave pública ANON válida.'
  };
}

function readEnvVars() {
  let envUrl = '';
  let envAnonKey = '';
  try {
    if (typeof import.meta !== 'undefined' && import.meta.env) {
      envUrl = import.meta.env.VITE_SUPABASE_URL || '';
      envAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY || '';
    }
  } catch {
    // Entorno sin bundler
  }
  return { envUrl: envUrl.trim(), envAnonKey: envAnonKey.trim() };
}

class SupabaseClientService {
  getConfig() {
    const { envUrl, envAnonKey } = readEnvVars();
    let saved = null;
    try {
      const raw = localStorage.getItem(STORAGE_CONFIG_KEY);
      if (raw) saved = JSON.parse(raw);
    } catch {
      saved = null;
    }

    const url = (saved?.url || envUrl || DEFAULT_SUPABASE_URL).trim().replace(/\/$/, '');
    const anonKey = (saved?.anonKey || envAnonKey || DEFAULT_SUPABASE_ANON_KEY).trim();
    const isPlaceholder =
      !url ||
      url.includes('tu-proyecto.supabase.co') ||
      !anonKey ||
      anonKey === PLACEHOLDER_KEY;

    const security = isPlaceholder
      ? { valid: false, isServiceRole: false, reason: 'Supabase aún no configurado (Modo Offline Local).' }
      : validateAnonKeySecurity(anonKey);

    return {
      url,
      anonKey,
      isPlaceholder,
      isConfigured: !isPlaceholder && security.valid,
      security,
      source: saved?.url ? 'localStorage' : envUrl ? 'env' : 'default'
    };
  }

  setConfig(url, anonKey) {
    const cleanUrl = String(url || '').trim().replace(/\/$/, '');
    const cleanKey = String(anonKey || '').trim();
    if (!cleanUrl.startsWith('https://')) {
      throw new Error('La URL de Supabase debe empezar con https:// (ej. https://xxxx.supabase.co).');
    }
    const sec = validateAnonKeySecurity(cleanKey);
    if (!sec.valid) {
      throw new Error(sec.reason);
    }
    localStorage.setItem(
      STORAGE_CONFIG_KEY,
      JSON.stringify({ url: cleanUrl, anonKey: cleanKey, updatedAt: new Date().toISOString() })
    );
    return this.getConfig();
  }

  clearConfig() {
    localStorage.removeItem(STORAGE_CONFIG_KEY);
    return this.getConfig();
  }

  _headers(cfg, extra = {}) {
    return {
      apikey: cfg.anonKey,
      Authorization: `Bearer ${cfg.anonKey}`,
      'Content-Type': 'application/json',
      ...extra
    };
  }

  async testConnection(timeoutMs = 8000) {
    const cfg = this.getConfig();
    if (!cfg.isConfigured) {
      return { ok: false, message: 'Configura primero la URL y la clave pública ANON de Supabase.' };
    }
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const endpoint = `${cfg.url}/rest/v1/registros_produccion?select=id&limit=1`;
      const res = await fetch(endpoint, {
        method: 'GET',
        headers: this._headers(cfg),
        signal: controller.signal,
        cache: 'no-store'
      });
      clearTimeout(timer);
      if (res.ok) {
        return { ok: true, message: 'Conexión exitosa con Supabase y tabla public.registros_produccion verificada.' };
      }
      const text = await res.text().catch(() => '');
      if (res.status === 404 || text.includes('does not exist')) {
        return {
          ok: false,
          message: 'Conectó al proyecto, pero falta ejecutar supabase_setup.sql en el SQL Editor de Supabase.'
        };
      }
      return { ok: false, message: `Supabase respondió HTTP ${res.status}: ${text || 'Revisa URL y clave ANON.'}` };
    } catch (err) {
      clearTimeout(timer);
      return {
        ok: false,
        message:
          err?.name === 'AbortError'
            ? 'Tiempo de espera agotado (8 s): no hay salida a internet o la URL es incorrecta.'
            : `No se pudo conectar: ${err?.message || 'error de red'}`
      };
    }
  }

  async upsertRegistro(row, timeoutMs = 10000) {
    const cfg = this.getConfig();
    if (!cfg.isConfigured) return { data: null, error: { message: 'Supabase no configurado.' } };

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const payload = {
        id: row.id,
        codigo_equipo: row.codigo_equipo,
        fecha: row.fecha,
        turno: row.turno,
        horometro_inicial: Number(row.horometro_inicial),
        horometro_final: Number(row.horometro_final),
        horas_trabajadas: Number(row.horas_trabajadas),
        material_fracturado: Number(row.material_fracturado),
        rendimiento: Number(row.rendimiento),
        area_trabajo: row.area_trabajo || '',
        operador: row.operador || '',
        observaciones: row.observaciones || '',
        hora_peru: row.hora_peru || '',
        creado_en: row.creado_en || new Date().toISOString(),
        actualizado_en: row.actualizado_en || new Date().toISOString(),
        dispositivo: row.dispositivo || '',
        eliminado: Boolean(row.eliminado)
      };

      const endpoint = `${cfg.url}/rest/v1/registros_produccion?on_conflict=id`;
      const res = await fetch(endpoint, {
        method: 'POST',
        headers: this._headers(cfg, { Prefer: 'resolution=merge-duplicates,return=representation' }),
        body: JSON.stringify([payload]),
        signal: controller.signal
      });
      clearTimeout(timer);

      if (!res.ok) {
        const errText = await res.text().catch(() => '');
        return { data: null, error: { message: `HTTP ${res.status}: ${errText}` } };
      }
      const data = await res.json().catch(() => [payload]);
      return { data, error: null };
    } catch (err) {
      clearTimeout(timer);
      return { data: null, error: { message: err?.message || 'Error de red' } };
    }
  }

  async selectRegistros(timeoutMs = 10000) {
    const cfg = this.getConfig();
    if (!cfg.isConfigured) return { data: [], error: null };

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const endpoint = `${cfg.url}/rest/v1/registros_produccion?select=*&eliminado=eq.false&order=fecha.desc,turno.asc&limit=2000`;
      const res = await fetch(endpoint, {
        method: 'GET',
        headers: this._headers(cfg),
        signal: controller.signal,
        cache: 'no-store'
      });
      clearTimeout(timer);
      if (!res.ok) {
        const errText = await res.text().catch(() => '');
        return { data: null, error: { message: `HTTP ${res.status}: ${errText}` } };
      }
      const data = await res.json();
      return { data: Array.isArray(data) ? data : [], error: null };
    } catch (err) {
      clearTimeout(timer);
      return { data: null, error: { message: err?.message || 'Error de red' } };
    }
  }

  async upsertEquipo(equipo) {
    const cfg = this.getConfig();
    if (!cfg.isConfigured) return;
    try {
      await fetch(`${cfg.url}/rest/v1/equipos_produccion?on_conflict=codigo`, {
        method: 'POST',
        headers: this._headers(cfg, { Prefer: 'resolution=merge-duplicates' }),
        body: JSON.stringify([{ codigo: equipo.codigo, categoria: equipo.categoria || '', actualizado_en: new Date().toISOString() }])
      });
    } catch {
      // Silencioso en segundo plano
    }
  }

  async upsertArea(nombre) {
    const cfg = this.getConfig();
    if (!cfg.isConfigured || !nombre) return;
    try {
      await fetch(`${cfg.url}/rest/v1/areas_trabajo?on_conflict=nombre`, {
        method: 'POST',
        headers: this._headers(cfg, { Prefer: 'resolution=merge-duplicates' }),
        body: JSON.stringify([{ nombre, actualizado_en: new Date().toISOString() }])
      });
    } catch {
      // Silencioso en segundo plano
    }
  }
}

export const supabase = new SupabaseClientService();
