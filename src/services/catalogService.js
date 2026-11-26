/**
 * Catálogo dinámico de Equipos y Áreas de Trabajo en IndexedDB.
 */

import { localDb } from '../db/localDb.js';
import { INITIAL_EQUIPOS, INITIAL_AREAS } from '../data/demoSeed.js';
import { normalizeEquipmentCode, normalizeAreaName } from '../utils/codeNormalizer.js';
import { supabase } from '../config/supabaseClient.js';

export function sortEquipmentList(list = []) {
  return [...list].sort((a, b) =>
    String(a?.codigo || '').localeCompare(String(b?.codigo || ''), 'es', {
      numeric: true,
      sensitivity: 'base'
    })
  );
}

export function sortAreaList(list = []) {
  return [...list].sort((a, b) =>
    String(a?.nombre || '').localeCompare(String(b?.nombre || ''), 'es', {
      numeric: true,
      sensitivity: 'base'
    })
  );
}

export async function ensureInitialCatalogs() {
  const existingEq = await localDb.equipos.toArray();
  const now = new Date().toISOString();
  if (existingEq.length === 0) {
    for (const eq of INITIAL_EQUIPOS) {
      await localDb.equipos.put({
        codigo: eq.codigo,
        categoria: eq.categoria,
        ultimo_uso: now,
        creado_en: now
      });
    }
  }

  const existingAreas = await localDb.areas.toArray();
  if (existingAreas.length === 0) {
    for (const area of INITIAL_AREAS) {
      await localDb.areas.put({
        nombre: area,
        ultimo_uso: now,
        creado_en: now
      });
    }
  }
}

export async function getAllEquipments() {
  const list = await localDb.equipos.toArray();
  return sortEquipmentList(list);
}

export async function createOrTouchEquipment(rawCode, categoria = '') {
  const codigo = normalizeEquipmentCode(rawCode);
  if (!codigo) throw new Error('Ingresa un código de equipo válido (ej. STREX02).');
  const existing = await localDb.equipos.get(codigo);
  const now = new Date().toISOString();
  const record = {
    codigo,
    categoria: categoria.trim() || existing?.categoria || 'Equipo de Producción / Fractura de Roca',
    ultimo_uso: now,
    creado_en: existing?.creado_en || now
  };
  await localDb.equipos.put(record);
  if (supabase.getConfig().isConfigured && navigator.onLine) {
    supabase.upsertEquipo(record);
  }
  return record;
}

export async function deleteEquipment(rawCode) {
  const codigo = normalizeEquipmentCode(rawCode);
  if (!codigo) return false;
  await localDb.equipos.delete(codigo);
  return true;
}

export async function getAllAreas() {
  const list = await localDb.areas.toArray();
  return sortAreaList(list);
}

export async function createOrTouchArea(rawName) {
  const nombre = normalizeAreaName(rawName);
  if (!nombre) return null;
  const existing = await localDb.areas.get(nombre);
  const now = new Date().toISOString();
  const record = {
    nombre,
    ultimo_uso: now,
    creado_en: existing?.creado_en || now
  };
  await localDb.areas.put(record);
  if (supabase.getConfig().isConfigured && navigator.onLine) {
    supabase.upsertArea(nombre);
  }
  return record;
}

export async function deleteArea(rawName) {
  const nombre = normalizeAreaName(rawName);
  if (!nombre) return false;
  await localDb.areas.delete(nombre);
  return true;
}

export async function syncCatalogsFromRecords(records = []) {
  const eqMap = new Map((await localDb.equipos.toArray()).map((e) => [e.codigo, e]));
  const areaMap = new Map((await localDb.areas.toArray()).map((a) => [a.nombre, a]));
  const now = new Date().toISOString();

  for (const r of records) {
    const eqCode = normalizeEquipmentCode(r.codigo_equipo || '');
    if (eqCode && !eqMap.has(eqCode)) {
      const rec = {
        codigo: eqCode,
        categoria: 'Equipo de Producción / Fractura de Roca',
        ultimo_uso: r.actualizado_en || now,
        creado_en: r.creado_en || now
      };
      await localDb.equipos.put(rec);
      eqMap.set(eqCode, rec);
    }
    const areaName = normalizeAreaName(r.area_trabajo || '');
    if (areaName && !areaMap.has(areaName)) {
      const rec = {
        nombre: areaName,
        ultimo_uso: r.actualizado_en || now,
        creado_en: r.creado_en || now
      };
      await localDb.areas.put(rec);
      areaMap.set(areaName, rec);
    }
  }
}
