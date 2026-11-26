/**
 * Motor de Negocio para Horómetros, Material Fracturado, Metas, Comparativo Mensual y Planilla STRACON.
 */

import { localDb } from '../db/localDb.js';
import { buildDemoRegistros } from '../data/demoSeed.js';
import {
  normalizeEquipmentCode,
  normalizeAreaName,
  parseNumOrNull,
  round1,
  round2
} from '../utils/codeNormalizer.js';
import {
  esFechaValida,
  getFechaHoyPeru,
  getHoraActualPeru,
  generateDateRange,
  formatDiaCortoExcel,
  getValuationMonthInfo,
  getSundayWeekStart,
  formatWeekRangeLabel,
  shiftIsoDate
} from '../utils/datePeru.js';
import { createOrTouchEquipment, createOrTouchArea, syncCatalogsFromRecords } from './catalogService.js';
import { supabase } from '../config/supabaseClient.js';

const DEMO_INIT_FLAG = 'stracon_produccion_demo_seeded_v3';
const TARGETS_STORAGE_KEY = 'stracon_produccion_metas_v1';
const MONTHLY_TARGETS_KEY = 'stracon_produccion_metas_mes_v1';

const DEFAULT_HISTORICAL_MONTH_TARGETS = {
  '2026-07': 180,
  '2026-08': 180,
  '2026-09': 360,
  '2026-10': 400,
  '2026-11': 480
};

export function getAllMonthlyTargets() {
  try {
    const raw = localStorage.getItem(MONTHLY_TARGETS_KEY);
    if (raw) {
      return { ...DEFAULT_HISTORICAL_MONTH_TARGETS, ...JSON.parse(raw) };
    }
  } catch {
    // ignore storage errors
  }
  return { ...DEFAULT_HISTORICAL_MONTH_TARGETS };
}

export function getMonthlyTargetForMonth(mesKey) {
  const map = getAllMonthlyTargets();
  if (map[mesKey] && Number(map[mesKey]) > 0) {
    return Number(map[mesKey]);
  }
  return getOperativeTargets().metaMensual;
}

export function saveMonthlyTarget(mesKey, metaHoras) {
  const cleanVal = Math.max(1, round1(Number(metaHoras) || 480));
  const map = getAllMonthlyTargets();
  map[mesKey] = cleanVal;
  localStorage.setItem(MONTHLY_TARGETS_KEY, JSON.stringify(map));
  return map;
}

export function getOperativeTargets() {
  try {
    const raw = localStorage.getItem(TARGETS_STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      return {
        metaDiaria: Math.max(1, Number(parsed.metaDiaria) || 16),
        metaMensual: Math.max(1, Number(parsed.metaMensual) || 480)
      };
    }
  } catch {
    // ignore storage errors
  }
  return { metaDiaria: 16, metaMensual: 480 };
}

export function saveOperativeTargets({ metaDiaria, metaMensual }) {
  const clean = {
    metaDiaria: Math.max(1, round1(Number(metaDiaria) || 16)),
    metaMensual: Math.max(1, round1(Number(metaMensual) || 480))
  };
  localStorage.setItem(TARGETS_STORAGE_KEY, JSON.stringify(clean));
  return clean;
}

export function computeShiftMetrics(hmInicial, hmFinal, materialM3) {
  const hi = parseNumOrNull(hmInicial);
  const hf = parseNumOrNull(hmFinal);
  const mat = parseNumOrNull(materialM3) ?? 0;

  if (hi === null || hf === null) {
    return { validNumbers: false, horas: 0, rendimiento: 0, deltaRaw: 0 };
  }
  const deltaRaw = hf - hi;
  const horas = deltaRaw >= 0 ? round1(deltaRaw) : 0;
  const rendimiento = horas > 0 && mat >= 0 ? round2(mat / horas) : 0;
  return {
    validNumbers: true,
    horas,
    rendimiento,
    deltaRaw: round1(deltaRaw)
  };
}

export function validateShiftInput({
  codigo_equipo,
  fecha,
  turno,
  horometro_inicial,
  horometro_final,
  material_fracturado
}) {
  const errors = [];
  const equipo = normalizeEquipmentCode(codigo_equipo);
  if (!equipo) errors.push('Selecciona o ingresa el código de equipo (ej. STREX02).');
  if (!esFechaValida(fecha)) errors.push('Ingresa una fecha válida.');
  if (!['DIA', 'NOCHE'].includes(turno)) errors.push('Selecciona el turno (DÍA o NOCHE).');

  const hi = parseNumOrNull(horometro_inicial);
  const hf = parseNumOrNull(horometro_final);
  const mat = parseNumOrNull(material_fracturado);

  if (hi === null || hi < 0) errors.push('Ingresa un Horómetro Inicial válido (>= 0).');
  if (hf === null || hf < 0) errors.push('Ingresa un Horómetro Final válido (>= 0).');
  if (hi !== null && hf !== null && hf < hi) {
    errors.push(`El Horómetro Final (${hf}) no puede ser menor que el Horómetro Inicial (${hi}).`);
  }
  if (mat !== null && mat < 0) {
    errors.push('El Material Fracturado (m³) no puede ser negativo.');
  }

  return {
    valid: errors.length === 0,
    errors,
    cleaned: {
      codigo_equipo: equipo,
      fecha,
      turno,
      horometro_inicial: hi !== null ? round1(hi) : 0,
      horometro_final: hf !== null ? round1(hf) : 0,
      material_fracturado: mat !== null ? round2(mat) : 0
    }
  };
}

export async function getAllRegistros() {
  const all = await localDb.registros.toArray();
  return all
    .filter((r) => !r.eliminado)
    .sort((a, b) => {
      if (a.fecha !== b.fecha) return b.fecha.localeCompare(a.fecha);
      if (a.turno !== b.turno) return a.turno === 'NOCHE' ? -1 : 1;
      return (b.actualizado_en || '').localeCompare(a.actualizado_en || '');
    });
}

export async function getRegistroById(id) {
  if (!id) return null;
  const r = await localDb.registros.get(id);
  return r && !r.eliminado ? r : null;
}

export async function findExistingShift(codigoEquipo, fecha, turno) {
  const eq = normalizeEquipmentCode(codigoEquipo);
  if (!eq || !fecha || !turno) return null;
  const all = await getAllRegistros();
  return all.find((r) => r.codigo_equipo === eq && r.fecha === fecha && r.turno === turno) || null;
}

export async function getSuggestedInitialHorometer(codigoEquipo, fecha, turno, excludeId = null) {
  const eq = normalizeEquipmentCode(codigoEquipo);
  if (!eq || !fecha) return null;
  const all = (await getAllRegistros()).filter((r) => r.codigo_equipo === eq && r.id !== excludeId);
  if (all.length === 0) return null;

  const rank = (r) => `${r.fecha}_${r.turno === 'DIA' ? '1' : '2'}`;
  const targetRank = `${fecha}_${turno === 'DIA' ? '1' : '2'}`;

  const previous = all
    .filter((r) => rank(r) < targetRank)
    .sort((a, b) => rank(b).localeCompare(rank(a)));

  if (previous.length === 0) return null;
  const prev = previous[0];
  return {
    suggestedValue: round1(prev.horometro_final),
    sourceDate: prev.fecha,
    sourceTurno: prev.turno,
    sourceLabel: `${prev.turno === 'DIA' ? 'Turno Día' : 'Turno Noche'} (${prev.fecha})`
  };
}

export async function saveShiftRecord(formData) {
  const validation = validateShiftInput(formData);
  if (!validation.valid) {
    throw new Error(validation.errors[0]);
  }

  const { cleaned } = validation;
  const areaClean = normalizeAreaName(formData.area_trabajo || '');
  const operadorClean = String(formData.operador || '').trim();
  const obsClean = String(formData.observaciones || '').trim();
  const metrics = computeShiftMetrics(
    cleaned.horometro_inicial,
    cleaned.horometro_final,
    cleaned.material_fracturado
  );

  const nowIso = new Date().toISOString();
  const existingBySlot = await findExistingShift(cleaned.codigo_equipo, cleaned.fecha, cleaned.turno);
  let targetId = formData.id || existingBySlot?.id || crypto.randomUUID();
  if (formData.id && existingBySlot && existingBySlot.id !== formData.id) {
    await localDb.registros.delete(formData.id);
    targetId = existingBySlot.id;
  }
  const existingById = await localDb.registros.get(targetId);

  const record = {
    id: targetId,
    codigo_equipo: cleaned.codigo_equipo,
    fecha: cleaned.fecha,
    turno: cleaned.turno,
    horometro_inicial: cleaned.horometro_inicial,
    horometro_final: cleaned.horometro_final,
    horas_trabajadas: metrics.horas,
    material_fracturado: cleaned.material_fracturado,
    rendimiento: metrics.rendimiento,
    area_trabajo: areaClean,
    operador: operadorClean,
    observaciones: obsClean,
    hora_peru: getHoraActualPeru(),
    creado_en: existingById?.creado_en || nowIso,
    actualizado_en: nowIso,
    dispositivo: formData.dispositivo || 'PWA Campo',
    estado: 'pendiente',
    es_demo: false,
    eliminado: false
  };

  await localDb.registros.put(record);
  await createOrTouchEquipment(cleaned.codigo_equipo);
  if (areaClean) {
    await createOrTouchArea(areaClean);
  }

  return record;
}

export async function deleteShiftRecord(id) {
  const existing = await localDb.registros.get(id);
  if (!existing) return false;

  if (existing.es_demo || existing.estado === 'demo' || !supabase.getConfig().isConfigured) {
    await localDb.registros.delete(id);
    return true;
  }

  if (existing.estado === 'sincronizado') {
    await localDb.registros.put({
      ...existing,
      eliminado: true,
      estado: 'pendiente',
      actualizado_en: new Date().toISOString()
    });
  } else {
    await localDb.registros.delete(id);
  }
  return true;
}

export async function ensureInitialDemoIfFresh() {
  if (localStorage.getItem(DEMO_INIT_FLAG) === '1') return false;
  const all = await localDb.registros.toArray();
  const hasRealRecords = all.some((r) => !r.es_demo && r.estado !== 'demo');
  localStorage.setItem(DEMO_INIT_FLAG, '1');

  // Si solo había ejemplos antiguos v1 (sin registros reales del usuario), actualizar al set completo EX01+EX02
  if (!hasRealRecords) {
    await clearDemoData();
    await loadDemoData();
    return true;
  }
  return false;
}

export async function loadDemoData() {
  const demoRows = buildDemoRegistros();
  for (const row of demoRows) {
    await localDb.registros.put(row);
  }
  await syncCatalogsFromRecords(demoRows);
  return demoRows.length;
}

export async function clearDemoData() {
  const all = await localDb.registros.toArray();
  let removed = 0;
  for (const r of all) {
    if (r.es_demo || r.estado === 'demo') {
      await localDb.registros.delete(r.id);
      removed++;
    }
  }
  return removed;
}

/**
 * Construye la matriz horizontal completa (PICOTÓN | STRACON) con metas, tendencia y avance.
 */
export function buildMatrixModel(records = [], { codigoEquipo, fechaDesde, fechaHasta, metaDiaria, metaMensual }) {
  const targets = getOperativeTargets();
  const targetDaily = Number(metaDiaria) > 0 ? Number(metaDiaria) : targets.metaDiaria;
  const targetMonthly = Number(metaMensual) > 0 ? Number(metaMensual) : targets.metaMensual;

  const eq = normalizeEquipmentCode(codigoEquipo || '');
  const eqRecords = records.filter((r) => !r.eliminado && (!eq || r.codigo_equipo === eq));

  const dayList = generateDateRange(fechaDesde, fechaHasta, 366);
  const byDateTurno = new Map();
  for (const r of eqRecords) {
    byDateTurno.set(`${r.fecha}__${r.turno}`, r);
  }

  let acumHrDia = 0;
  let acumHrNoche = 0;
  let acumHrTotal = 0;
  let acumMatDia = 0;
  let acumMatNoche = 0;
  let acumMatTotal = 0;
  let bestDay = null;
  let diasConProduccion = 0;
  let diasActivosDia = 0;
  let diasActivosNoche = 0;
  let ultimoHorometro = null;

  const areaStatsMap = new Map();

  const addAreaStat = (areaRaw, horas, mat) => {
    const area = areaRaw ? areaRaw.trim() : 'Sin área asignada';
    if (horas <= 0 && mat <= 0) return;
    const prev = areaStatsMap.get(area) || { area, turnos: 0, horas: 0, material: 0 };
    prev.turnos += 1;
    prev.horas = round1(prev.horas + horas);
    prev.material = round2(prev.material + mat);
    areaStatsMap.set(area, prev);
  };

  const days = dayList.map((fecha, idx) => {
    const { diaSemana, fechaCorta } = formatDiaCortoExcel(fecha);
    const recDia = byDateTurno.get(`${fecha}__DIA`) || null;
    const recNoche = byDateTurno.get(`${fecha}__NOCHE`) || null;

    const diaHoras = recDia ? round1(recDia.horas_trabajadas) : 0;
    const diaMat = recDia ? round2(recDia.material_fracturado) : 0;
    const diaRend = diaHoras > 0 ? round2(diaMat / diaHoras) : 0;
    if (diaHoras > 0) diasActivosDia += 1;

    const nocheHoras = recNoche ? round1(recNoche.horas_trabajadas) : 0;
    const nocheMat = recNoche ? round2(recNoche.material_fracturado) : 0;
    const nocheRend = nocheHoras > 0 ? round2(nocheMat / nocheHoras) : 0;
    if (nocheHoras > 0) diasActivosNoche += 1;

    const totalHoras = round1(diaHoras + nocheHoras);
    const totalMat = round2(diaMat + nocheMat);
    const totalRend = totalHoras > 0 ? round2(totalMat / totalHoras) : 0;

    if (recDia) {
      ultimoHorometro = round1(recDia.horometro_final);
      addAreaStat(recDia.area_trabajo, diaHoras, diaMat);
    }
    if (recNoche) {
      ultimoHorometro = round1(recNoche.horometro_final);
      addAreaStat(recNoche.area_trabajo, nocheHoras, nocheMat);
    }

    if (totalHoras > 0 || totalMat > 0) {
      diasConProduccion += 1;
      if (!bestDay || totalMat > bestDay.material) {
        bestDay = { fecha, fechaCorta, horas: totalHoras, material: totalMat, rendimiento: totalRend };
      }
    }

    acumHrDia = round1(acumHrDia + diaHoras);
    acumHrNoche = round1(acumHrNoche + nocheHoras);
    acumHrTotal = round1(acumHrTotal + totalHoras);
    acumMatDia = round2(acumMatDia + diaMat);
    acumMatNoche = round2(acumMatNoche + nocheMat);
    acumMatTotal = round2(acumMatTotal + totalMat);
    const acumRend = acumHrTotal > 0 ? round2(acumMatTotal / acumHrTotal) : 0;

    const acumuladoPlaneado = round1(targetDaily * (idx + 1));
    const tendencia = acumuladoPlaneado > 0 ? acumHrTotal / acumuladoPlaneado : 0;
    const avance = targetMonthly > 0 ? acumHrTotal / targetMonthly : 0;

    return {
      fecha,
      diaSemana,
      fechaCorta,
      dia: {
        record: recDia,
        hmInicial: recDia ? round1(recDia.horometro_inicial) : null,
        hmFinal: recDia ? round1(recDia.horometro_final) : null,
        horas: diaHoras,
        material: recDia ? diaMat : null,
        rendimiento: recDia ? diaRend : null,
        area: recDia?.area_trabajo || '',
        operador: recDia?.operador || ''
      },
      noche: {
        record: recNoche,
        hmInicial: recNoche ? round1(recNoche.horometro_inicial) : null,
        hmFinal: recNoche ? round1(recNoche.horometro_final) : null,
        horas: nocheHoras,
        material: recNoche ? nocheMat : null,
        rendimiento: recNoche ? nocheRend : null,
        area: recNoche?.area_trabajo || '',
        operador: recNoche?.operador || ''
      },
      total: {
        horas: totalHoras,
        material: totalMat,
        rendimiento: totalRend
      },
      acumulado: {
        hrDia: acumHrDia,
        hrNoche: acumHrNoche,
        hrTotal: acumHrTotal,
        matDia: acumMatDia,
        matNoche: acumMatNoche,
        matTotal: acumMatTotal,
        rendAcum: acumRend
      },
      meta: {
        metaDiaria: targetDaily,
        acumuladoPlaneado,
        tendencia,
        avance
      }
    };
  });

  const rendDiaGlobal = acumHrDia > 0 ? round2(acumMatDia / acumHrDia) : 0;
  const rendNocheGlobal = acumHrNoche > 0 ? round2(acumMatNoche / acumHrNoche) : 0;
  const rendTotalGlobal = acumHrTotal > 0 ? round2(acumMatTotal / acumHrTotal) : 0;
  const promDiaHr = diasActivosDia > 0 ? round2(acumHrDia / diasActivosDia) : 0;
  const promNocheHr = diasActivosNoche > 0 ? round2(acumHrNoche / diasActivosNoche) : 0;
  const acumuladoPlaneadoFinal = round1(targetDaily * days.length);
  const tendenciaFinal = acumuladoPlaneadoFinal > 0 ? acumHrTotal / acumuladoPlaneadoFinal : 0;
  const avanceFinal = targetMonthly > 0 ? acumHrTotal / targetMonthly : 0;

  const areaBreakdown = Array.from(areaStatsMap.values())
    .map((item) => ({
      ...item,
      rendimiento: item.horas > 0 ? round2(item.material / item.horas) : 0,
      porcentajeMaterial: acumMatTotal > 0 ? round1((item.material / acumMatTotal) * 100) : 0
    }))
    .sort((a, b) => b.material - a.material);

  return {
    codigoEquipo: eq || 'TODOS',
    fechaDesde,
    fechaHasta,
    days,
    summary: {
      totalDiasRango: days.length,
      diasConProduccion,
      diasActivosDia,
      diasActivosNoche,
      horasDia: acumHrDia,
      horasNoche: acumHrNoche,
      horasTotal: acumHrTotal,
      promDiaHr,
      promNocheHr,
      materialDia: acumMatDia,
      materialNoche: acumMatNoche,
      materialTotal: acumMatTotal,
      rendDia: rendDiaGlobal,
      rendNoche: rendNocheGlobal,
      rendTotal: rendTotalGlobal,
      metaDiaria: targetDaily,
      metaMensual: targetMonthly,
      acumuladoPlaneadoFinal,
      tendenciaFinal,
      avanceFinal,
      bestDay,
      ultimoHorometro
    },
    areaBreakdown
  };
}

/**
 * Consolida todos los registros por Mes de Valorización (26–25) y Equipo
 * igual que la tabla "Comparativo Turnos diurno y nocturno" del Excel Picotón.
 */
export function buildMonthlyComparative(records = []) {
  const groups = new Map();

  for (const r of records) {
    if (r.eliminado) continue;
    const eq = normalizeEquipmentCode(r.codigo_equipo || '');
    if (!eq || !esFechaValida(r.fecha)) continue;
    const valMonth = getValuationMonthInfo(r.fecha);
    const key = `${valMonth.key}__${eq}`;

    if (!groups.has(key)) {
      groups.set(key, {
        codigoEquipo: eq,
        mesKey: valMonth.key,
        mesNombre: valMonth.mesNombre,
        mesLabel: valMonth.label,
        horasDia: 0,
        horasNoche: 0,
        materialTotal: 0,
        diasDiaSet: new Set(),
        diasNocheSet: new Set()
      });
    }

    const g = groups.get(key);
    const hrs = Number(r.horas_trabajadas) || 0;
    const mat = Number(r.material_fracturado) || 0;
    g.materialTotal = round2(g.materialTotal + mat);

    if (r.turno === 'DIA') {
      g.horasDia = round1(g.horasDia + hrs);
      if (hrs > 0) g.diasDiaSet.add(r.fecha);
    } else if (r.turno === 'NOCHE') {
      g.horasNoche = round1(g.horasNoche + hrs);
      if (hrs > 0) g.diasNocheSet.add(r.fecha);
    }
  }

  const rows = Array.from(groups.values())
    .sort((a, b) => {
      if (a.mesKey !== b.mesKey) return a.mesKey.localeCompare(b.mesKey);
      return a.codigoEquipo.localeCompare(b.codigoEquipo);
    })
    .map((g) => {
      const horasTotal = round1(g.horasDia + g.horasNoche);
      const diasActivosDia = g.diasDiaSet.size;
      const diasActivosNoche = g.diasNocheSet.size;
      return {
        codigoEquipo: g.codigoEquipo,
        mesKey: g.mesKey,
        mesNombre: g.mesNombre,
        mesLabel: g.mesLabel,
        horasDia: g.horasDia,
        horasNoche: g.horasNoche,
        horasTotal,
        materialTotal: g.materialTotal,
        rendimiento: horasTotal > 0 ? round2(g.materialTotal / horasTotal) : 0,
        diasActivosDia,
        diasActivosNoche,
        promDiaHr: diasActivosDia > 0 ? round2(g.horasDia / diasActivosDia) : 0,
        promNocheHr: diasActivosNoche > 0 ? round2(g.horasNoche / diasActivosNoche) : 0
      };
    });

  const totHorasDia = round1(rows.reduce((a, r) => a + r.horasDia, 0));
  const totHorasNoche = round1(rows.reduce((a, r) => a + r.horasNoche, 0));
  const totHoras = round1(totHorasDia + totHorasNoche);
  const totMat = round2(rows.reduce((a, r) => a + r.materialTotal, 0));
  const totDiasDia = rows.reduce((a, r) => a + r.diasActivosDia, 0);
  const totDiasNoche = rows.reduce((a, r) => a + r.diasActivosNoche, 0);

  return {
    rows,
    totals: {
      horasDia: totHorasDia,
      horasNoche: totHorasNoche,
      horasTotal: totHoras,
      materialTotal: totMat,
      rendimiento: totHoras > 0 ? round2(totMat / totHoras) : 0,
      promDiaHr: totDiasDia > 0 ? round2(totHorasDia / totDiasDia) : 0,
      promNocheHr: totDiasNoche > 0 ? round2(totHorasNoche / totDiasNoche) : 0
    }
  };
}

export function getDataDateRange(records = []) {
  const valid = records.filter((r) => !r.eliminado && esFechaValida(r.fecha));
  if (valid.length === 0) {
    const hoy = getFechaHoyPeru();
    return { min: hoy, max: hoy };
  }
  let min = valid[0].fecha;
  let max = valid[0].fecha;
  for (const r of valid) {
    if (r.fecha < min) min = r.fecha;
    if (r.fecha > max) max = r.fecha;
  }
  return { min, max };
}

/**
 * Construye el modelo completo para el Tablero Gerencial (Acumulado Julio-Noviembre+,
 * Resumen Mensual, Distribución por Turno, Rendimiento, Promedios Hr/Día, Semanal Dom-Sáb y Detalle Diario).
 */
export function buildExecutiveDashboardModel(records = [], { fechaDesde = '', fechaHasta = '' } = {}) {
  const active = records.filter((r) => {
    if (r.eliminado || !esFechaValida(r.fecha)) return false;
    if (fechaDesde && r.fecha < fechaDesde) return false;
    if (fechaHasta && r.fecha > fechaHasta) return false;
    return true;
  });

  const equipos = Array.from(
    new Set(active.map((r) => normalizeEquipmentCode(r.codigo_equipo || '')).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b, 'es', { numeric: true }));

  const monthEqMap = new Map();
  const weekMap = new Map();
  const dailyMap = new Map();

  for (const r of active) {
    const eq = normalizeEquipmentCode(r.codigo_equipo || '');
    if (!eq) continue;

    const valMonth = getValuationMonthInfo(r.fecha);
    const hrs = Number(r.horas_trabajadas) || 0;
    const mat = Number(r.material_fracturado) || 0;
    const isDia = r.turno === 'DIA';
    const isNoche = r.turno === 'NOCHE';

    // 1. Agrupación Mensual por Equipo
    const meKey = `${valMonth.key}__${eq}`;
    if (!monthEqMap.has(meKey)) {
      monthEqMap.set(meKey, {
        mesKey: valMonth.key,
        mesNombre: valMonth.mesNombre,
        mesLabel: valMonth.label,
        codigoEquipo: eq,
        horasDia: 0,
        horasNoche: 0,
        materialDia: 0,
        materialNoche: 0,
        diasOpSet: new Set(),
        diasDiaSet: new Set(),
        diasNocheSet: new Set()
      });
    }
    const me = monthEqMap.get(meKey);
    if (isDia) {
      me.horasDia = round1(me.horasDia + hrs);
      me.materialDia = round2(me.materialDia + mat);
      if (hrs > 0) me.diasDiaSet.add(r.fecha);
    } else if (isNoche) {
      me.horasNoche = round1(me.horasNoche + hrs);
      me.materialNoche = round2(me.materialNoche + mat);
      if (hrs > 0) me.diasNocheSet.add(r.fecha);
    }
    if (hrs > 0) me.diasOpSet.add(r.fecha);

    // 2. Agrupación Semanal (Domingo a Sábado)
    const inicioDom = getSundayWeekStart(r.fecha);
    if (!weekMap.has(inicioDom)) {
      weekMap.set(inicioDom, {
        inicioDom,
        finSab: shiftIsoDate(inicioDom, 6),
        semanaLabel: formatWeekRangeLabel(inicioDom),
        byEquip: {}
      });
    }
    const wk = weekMap.get(inicioDom);
    if (!wk.byEquip[eq]) {
      wk.byEquip[eq] = { horas: 0, material: 0, diasSet: new Set() };
    }
    wk.byEquip[eq].horas = round1(wk.byEquip[eq].horas + hrs);
    wk.byEquip[eq].material = round2(wk.byEquip[eq].material + mat);
    if (hrs > 0) wk.byEquip[eq].diasSet.add(r.fecha);

    // 3. Detalle Diario
    if (!dailyMap.has(r.fecha)) {
      const { diaSemana } = formatDiaCortoExcel(r.fecha);
      dailyMap.set(r.fecha, {
        fecha: r.fecha,
        mesNombre: valMonth.mesNombre,
        diaSemana,
        inicioDom,
        byEquip: {}
      });
    }
    const dy = dailyMap.get(r.fecha);
    if (!dy.byEquip[eq]) {
      dy.byEquip[eq] = { diaHr: 0, nocheHr: 0, totalHr: 0, material: 0 };
    }
    if (isDia) dy.byEquip[eq].diaHr = round1(dy.byEquip[eq].diaHr + hrs);
    if (isNoche) dy.byEquip[eq].nocheHr = round1(dy.byEquip[eq].nocheHr + hrs);
    dy.byEquip[eq].totalHr = round1(dy.byEquip[eq].diaHr + dy.byEquip[eq].nocheHr);
    dy.byEquip[eq].material = round2(dy.byEquip[eq].material + mat);
  }

  // Agrupar filas mensuales ordenadas y construir subtotales
  const monthsOrdered = Array.from(new Set(Array.from(monthEqMap.values()).map((x) => x.mesKey))).sort();
  const monthlyTable = [];
  const monthlyChartSeries = [];

  let totHrDia = 0;
  let totHrNoche = 0;
  let totMatDia = 0;
  let totMatNoche = 0;
  let totMeta = 0;
  let totDiasOp = 0;
  let totDiasDia = 0;
  let totDiasNoche = 0;

  for (const mKey of monthsOrdered) {
    const eqRowsRaw = Array.from(monthEqMap.values())
      .filter((x) => x.mesKey === mKey)
      .sort((a, b) => a.codigoEquipo.localeCompare(b.codigoEquipo, 'es', { numeric: true }));

    const metaMes = getMonthlyTargetForMonth(mKey);
    let mHrDia = 0;
    let mHrNoche = 0;
    let mMatDia = 0;
    let mMatNoche = 0;
    let mDiasOp = 0;
    let mDiasDia = 0;
    let mDiasNoche = 0;
    let mMeta = 0;

    for (const g of eqRowsRaw) {
      const horasTotal = round1(g.horasDia + g.horasNoche);
      const materialTotal = round2(g.materialDia + g.materialNoche);
      const diasOperativos = g.diasOpSet.size;
      const diasDia = g.diasDiaSet.size;
      const diasNoche = g.diasNocheSet.size;

      const row = {
        rowType: 'equipo',
        mesKey: g.mesKey,
        mesNombre: g.mesNombre,
        mesLabel: g.mesLabel,
        codigoEquipo: g.codigoEquipo,
        horasDia: g.horasDia,
        horasNoche: g.horasNoche,
        horasTotal,
        pctNoche: horasTotal > 0 ? g.horasNoche / horasTotal : 0,
        metaMensual: metaMes,
        pctCumplimiento: metaMes > 0 ? horasTotal / metaMes : 0,
        materialDia: g.materialDia,
        materialNoche: g.materialNoche,
        materialTotal,
        rendDia: g.horasDia > 0 ? round2(g.materialDia / g.horasDia) : 0,
        rendNoche: g.horasNoche > 0 ? round2(g.materialNoche / g.horasNoche) : 0,
        rendTotal: horasTotal > 0 ? round2(materialTotal / horasTotal) : 0,
        diasOperativos,
        hrPorDiaOperativo: diasOperativos > 0 ? round2(horasTotal / diasOperativos) : 0,
        promDiaHr: diasDia > 0 ? round2(g.horasDia / diasDia) : 0,
        promNocheHr: diasNoche > 0 ? round2(g.horasNoche / diasNoche) : 0
      };

      monthlyTable.push(row);
      monthlyChartSeries.push(row);

      mHrDia = round1(mHrDia + g.horasDia);
      mHrNoche = round1(mHrNoche + g.horasNoche);
      mMatDia = round2(mMatDia + g.materialDia);
      mMatNoche = round2(mMatNoche + g.materialNoche);
      mDiasOp += diasOperativos;
      mDiasDia += diasDia;
      mDiasNoche += diasNoche;
      mMeta += metaMes;
    }

    totHrDia = round1(totHrDia + mHrDia);
    totHrNoche = round1(totHrNoche + mHrNoche);
    totMatDia = round2(totMatDia + mMatDia);
    totMatNoche = round2(totMatNoche + mMatNoche);
    totDiasOp += mDiasOp;
    totDiasDia += mDiasDia;
    totDiasNoche += mDiasNoche;
    totMeta += mMeta;

    if (eqRowsRaw.length >= 2) {
      const mHrTotal = round1(mHrDia + mHrNoche);
      const mMatTotal = round2(mMatDia + mMatNoche);
      monthlyTable.push({
        rowType: 'subtotal',
        mesKey: mKey,
        mesNombre: `${eqRowsRaw[0].mesNombre} TOTAL`,
        mesLabel: `${eqRowsRaw[0].mesNombre} TOTAL`,
        codigoEquipo: eqRowsRaw.map((x) => x.codigoEquipo.replace(/^STR/, '')).join('+'),
        horasDia: mHrDia,
        horasNoche: mHrNoche,
        horasTotal: mHrTotal,
        pctNoche: mHrTotal > 0 ? mHrNoche / mHrTotal : 0,
        metaMensual: mMeta,
        pctCumplimiento: mMeta > 0 ? mHrTotal / mMeta : 0,
        materialDia: mMatDia,
        materialNoche: mMatNoche,
        materialTotal: mMatTotal,
        rendDia: mHrDia > 0 ? round2(mMatDia / mHrDia) : 0,
        rendNoche: mHrNoche > 0 ? round2(mMatNoche / mHrNoche) : 0,
        rendTotal: mHrTotal > 0 ? round2(mMatTotal / mHrTotal) : 0,
        diasOperativos: mDiasOp,
        hrPorDiaOperativo: mDiasOp > 0 ? round2(mHrTotal / mDiasOp) : 0,
        promDiaHr: mDiasDia > 0 ? round2(mHrDia / mDiasDia) : 0,
        promNocheHr: mDiasNoche > 0 ? round2(mHrNoche / mDiasNoche) : 0
      });
    }
  }

  const totHrTotal = round1(totHrDia + totHrNoche);
  const totMatTotal = round2(totMatDia + totMatNoche);

  const totals = {
    horasDia: totHrDia,
    horasNoche: totHrNoche,
    horasTotal: totHrTotal,
    pctNoche: totHrTotal > 0 ? totHrNoche / totHrTotal : 0,
    metaMensualTotal: totMeta,
    pctCumplimiento: totMeta > 0 ? totHrTotal / totMeta : 0,
    materialDia: totMatDia,
    materialNoche: totMatNoche,
    materialTotal: totMatTotal,
    rendDia: totHrDia > 0 ? round2(totMatDia / totHrDia) : 0,
    rendNoche: totHrNoche > 0 ? round2(totMatNoche / totHrNoche) : 0,
    rendTotal: totHrTotal > 0 ? round2(totMatTotal / totHrTotal) : 0,
    diasOperativos: totDiasOp,
    hrPorDiaOperativo: totDiasOp > 0 ? round2(totHrTotal / totDiasOp) : 0,
    promDiaHr: totDiasDia > 0 ? round2(totHrDia / totDiasDia) : 0,
    promNocheHr: totDiasNoche > 0 ? round2(totHrNoche / totDiasNoche) : 0
  };

  // Semanas ordenadas (Domingo a Sábado)
  const weeklyTable = Array.from(weekMap.values())
    .sort((a, b) => a.inicioDom.localeCompare(b.inicioDom))
    .map((wk) => {
      let horasTotal = 0;
      let materialTotal = 0;
      const cleanByEquip = {};
      for (const eq of equipos) {
        const raw = wk.byEquip[eq];
        const horas = raw ? raw.horas : 0;
        const material = raw ? raw.material : 0;
        const diasOp = raw ? raw.diasSet.size : 0;
        cleanByEquip[eq] = { horas, material, diasOp };
        horasTotal = round1(horasTotal + horas);
        materialTotal = round2(materialTotal + material);
      }
      return {
        inicioDom: wk.inicioDom,
        finSab: wk.finSab,
        semanaLabel: wk.semanaLabel,
        byEquip: cleanByEquip,
        horasTotal,
        materialTotal,
        rendimiento: horasTotal > 0 ? round2(materialTotal / horasTotal) : 0
      };
    });

  const weeklyTotals = {
    byEquip: {},
    horasTotal: totHrTotal,
    materialTotal: totMatTotal,
    rendimiento: totHrTotal > 0 ? round2(totMatTotal / totHrTotal) : 0
  };
  for (const eq of equipos) {
    weeklyTotals.byEquip[eq] = {
      horas: round1(weeklyTable.reduce((acc, w) => acc + (w.byEquip[eq]?.horas || 0), 0)),
      material: round2(weeklyTable.reduce((acc, w) => acc + (w.byEquip[eq]?.material || 0), 0)),
      diasOp: weeklyTable.reduce((acc, w) => acc + (w.byEquip[eq]?.diasOp || 0), 0)
    };
  }

  // Detalle diario ordenado
  const dailyDetail = Array.from(dailyMap.values())
    .sort((a, b) => a.fecha.localeCompare(b.fecha))
    .map((d) => {
      let hmTotal = 0;
      let materialTotal = 0;
      const cleanByEquip = {};
      for (const eq of equipos) {
        const raw = d.byEquip[eq] || { diaHr: 0, nocheHr: 0, totalHr: 0, material: 0 };
        cleanByEquip[eq] = raw;
        hmTotal = round1(hmTotal + raw.totalHr);
        materialTotal = round2(materialTotal + raw.material);
      }
      return {
        ...d,
        byEquip: cleanByEquip,
        hmTotal,
        materialTotal
      };
    });

  return {
    equipos,
    monthsOrdered,
    monthlyTable,
    monthlyChartSeries,
    totals,
    weeklyTable,
    weeklyTotals,
    dailyDetail
  };
}
