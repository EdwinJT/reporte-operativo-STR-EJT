/**
 * Utilidades de fecha y hora estrictas en zona horaria Perú (America/Lima, UTC-5)
 * Evita desfases UTC después de las 19:00 h en Perú.
 */

const DIAS_SEMANA_ES = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
const MESES_CORTOS_ES = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Set', 'Oct', 'Nov', 'Dic'];
const MESES_LARGOS_ES = [
  'ENERO',
  'FEBRERO',
  'MARZO',
  'ABRIL',
  'MAYO',
  'JUNIO',
  'JULIO',
  'AGOSTO',
  'SETIEMBRE',
  'OCTUBRE',
  'NOVIEMBRE',
  'DICIEMBRE'
];

export function getFechaHoyPeru(dateObj = new Date()) {
  const formatter = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'America/Lima',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  });
  const parts = formatter.formatToParts(dateObj);
  const year = parts.find((p) => p.type === 'year')?.value;
  const month = parts.find((p) => p.type === 'month')?.value;
  const day = parts.find((p) => p.type === 'day')?.value;
  return `${year}-${month}-${day}`;
}

export function getHoraActualPeru(dateObj = new Date()) {
  const formatter = new Intl.DateTimeFormat('es-PE', {
    timeZone: 'America/Lima',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false
  });
  return formatter.format(dateObj);
}

export function getTurnoActualPeru(dateObj = new Date()) {
  const formatter = new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/Lima',
    hour: '2-digit',
    hour12: false
  });
  const hour = Number(formatter.format(dateObj));
  return hour >= 7 && hour < 19 ? 'DIA' : 'NOCHE';
}

export function shiftIsoDate(isoDate, deltaDays = 0) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(isoDate || '').trim());
  if (!m) return getFechaHoyPeru();
  const d = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
  d.setUTCDate(d.getUTCDate() + deltaDays);
  const y = d.getUTCFullYear();
  const mo = String(d.getUTCMonth() + 1).padStart(2, '0');
  const da = String(d.getUTCDate()).padStart(2, '0');
  return `${y}-${mo}-${da}`;
}

export function getFechaAyerPeru() {
  return shiftIsoDate(getFechaHoyPeru(), -1);
}

export function esFechaValida(fechaStr) {
  if (!fechaStr || typeof fechaStr !== 'string') return false;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fechaStr)) return false;
  const [y, m, d] = fechaStr.split('-').map(Number);
  if (m < 1 || m > 12 || d < 1 || d > 31 || y < 2000 || y > 2100) return false;
  const test = new Date(Date.UTC(y, m - 1, d));
  return test.getUTCFullYear() === y && test.getUTCMonth() === m - 1 && test.getUTCDate() === d;
}

export function formatFechaDDMMYYYY(fechaStr) {
  if (!fechaStr || typeof fechaStr !== 'string') return '';
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(fechaStr.trim());
  if (m) return `${m[3]}/${m[2]}/${m[1]}`;
  return fechaStr;
}

/**
 * Devuelve { diaSemana: 'martes', fechaCorta: '27-Oct' } igual que la cabecera STRACON.
 */
export function formatDiaCortoExcel(isoDate) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(isoDate || '').trim());
  if (!m) return { diaSemana: '', fechaCorta: String(isoDate || '') };
  const y = Number(m[1]);
  const mo = Number(m[2]) - 1;
  const da = Number(m[3]);
  const utc = new Date(Date.UTC(y, mo, da));
  const diaSemana = DIAS_SEMANA_ES[utc.getUTCDay()];
  const mesCorto = MESES_CORTOS_ES[mo] || '';
  return {
    diaSemana,
    fechaCorta: `${da}-${mesCorto}`
  };
}

export function getMonthTitleSpanish(startIso, endIso = startIso) {
  const m1 = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(startIso || ''));
  const m2 = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(endIso || startIso || ''));
  if (!m1) return '';
  const y1 = Number(m1[1]);
  const mo1 = Number(m1[2]) - 1;
  if (!m2) return `${MESES_LARGOS_ES[mo1]} ${y1}`;
  const y2 = Number(m2[1]);
  const mo2 = Number(m2[2]) - 1;
  if (y1 === y2 && mo1 === mo2) {
    return `${MESES_LARGOS_ES[mo1]} ${y1}`;
  }
  if (y1 === y2) {
    return `${MESES_LARGOS_ES[mo1]} - ${MESES_LARGOS_ES[mo2]} ${y1}`;
  }
  return `${MESES_LARGOS_ES[mo1]} ${y1} - ${MESES_LARGOS_ES[mo2]} ${y2}`;
}

export function generateDateRange(startIso, endIso, maxDays = 366) {
  if (!esFechaValida(startIso) || !esFechaValida(endIso)) return [];
  if (startIso > endIso) return [];
  const list = [];
  let cursor = startIso;
  let guard = 0;
  while (cursor <= endIso && guard < maxDays) {
    list.push(cursor);
    cursor = shiftIsoDate(cursor, 1);
    guard++;
  }
  return list;
}

/**
 * Período de valorización minera en Perú: del 26 del mes anterior al 25 del mes actual
 * (o del 26 del mes actual al 25 del mes siguiente si hoy >= 26).
 */
export function getValuationPeriodRange(referenceIso = getFechaHoyPeru()) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(referenceIso);
  if (!m) return { desde: referenceIso, hasta: referenceIso };
  let y = Number(m[1]);
  let mo = Number(m[2]); // 1..12
  const da = Number(m[3]);

  let startYear = y;
  let startMonth = mo;
  let endYear = y;
  let endMonth = mo;

  if (da >= 26) {
    endMonth = mo + 1;
    if (endMonth > 12) {
      endMonth = 1;
      endYear += 1;
    }
  } else {
    startMonth = mo - 1;
    if (startMonth < 1) {
      startMonth = 12;
      startYear -= 1;
    }
  }

  return {
    desde: `${startYear}-${String(startMonth).padStart(2, '0')}-26`,
    hasta: `${endYear}-${String(endMonth).padStart(2, '0')}-25`
  };
}

export function getCalendarMonthRange(referenceIso = getFechaHoyPeru()) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(referenceIso);
  if (!m) return { desde: referenceIso, hasta: referenceIso };
  const y = Number(m[1]);
  const mo = Number(m[2]);
  const lastDay = new Date(Date.UTC(y, mo, 0)).getUTCDate();
  return {
    desde: `${y}-${String(mo).padStart(2, '0')}-01`,
    hasta: `${y}-${String(mo).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`
  };
}

/**
 * Asigna una fecha YYYY-MM-DD a su mes de valorización minera (26 al 25).
 * Ej.: 2026-10-26..2026-11-25 -> NOVIEMBRE 2026 ('2026-11').
 */
export function getValuationMonthInfo(isoDate) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(isoDate || '').trim());
  if (!m) return { key: '0000-00', mesNombre: 'MES', year: 0, label: 'MES' };
  let y = Number(m[1]);
  let mo = Number(m[2]); // 1..12
  const da = Number(m[3]);
  if (da >= 26) {
    mo += 1;
    if (mo > 12) {
      mo = 1;
      y += 1;
    }
  }
  const mesNombre = MESES_LARGOS_ES[mo - 1] || '';
  return {
    key: `${y}-${String(mo).padStart(2, '0')}`,
    mesNombre,
    year: y,
    label: `${mesNombre} ${y}`
  };
}

/**
 * Devuelve el rango 26–25 para una clave de mes de valorización 'YYYY-MM'.
 * Ej.: '2026-11' -> { desde: '2026-10-26', hasta: '2026-11-25', label: 'NOVIEMBRE 2026' }
 */
export function getValuationRangeByMonthKey(mesKey) {
  const m = /^(\d{4})-(\d{2})$/.exec(String(mesKey || '').trim());
  if (!m) return getValuationPeriodRange(getFechaHoyPeru());
  const y = Number(m[1]);
  const mo = Number(m[2]); // 1..12
  let startYear = y;
  let startMonth = mo - 1;
  if (startMonth < 1) {
    startMonth = 12;
    startYear -= 1;
  }
  const mesNombre = MESES_LARGOS_ES[mo - 1] || '';
  return {
    desde: `${startYear}-${String(startMonth).padStart(2, '0')}-26`,
    hasta: `${y}-${String(mo).padStart(2, '0')}-25`,
    label: `${mesNombre} ${y}`
  };
}

/**
 * Devuelve un título coherente para un rango de fechas:
 * - Si cae dentro de un único mes de valorización (26–25), devuelve ej. 'NOVIEMBRE 2026'.
 * - Si abarca varios meses, devuelve ej. 'JULIO - NOVIEMBRE 2026'.
 */
export function getPeriodTitleSpanish(startIso, endIso = startIso) {
  if (!esFechaValida(startIso)) return getMonthTitleSpanish(startIso, endIso);
  const end = esFechaValida(endIso) ? endIso : startIso;
  const vStart = getValuationMonthInfo(startIso);
  const vEnd = getValuationMonthInfo(end);
  if (vStart.key === vEnd.key) {
    return vEnd.label;
  }
  return getMonthTitleSpanish(startIso, end);
}

/**
 * Devuelve el domingo (YYYY-MM-DD) que inicia la semana (Domingo a Sábado) de isoDate.
 */
export function getSundayWeekStart(isoDate) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(isoDate || '').trim());
  if (!m) return '';
  const d = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
  const dow = d.getUTCDay(); // 0 = Domingo
  d.setUTCDate(d.getUTCDate() - dow);
  const y = d.getUTCFullYear();
  const mo = String(d.getUTCMonth() + 1).padStart(2, '0');
  const da = String(d.getUTCDate()).padStart(2, '0');
  return `${y}-${mo}-${da}`;
}

/**
 * Formatea una semana Domingo-Sábado como '12-Jul al 18-Jul'.
 */
export function formatWeekRangeLabel(sundayIso) {
  if (!sundayIso) return '';
  const satIso = shiftIsoDate(sundayIso, 6);
  const fmtPart = (iso) => {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
    if (!m) return iso;
    const mes = MESES_CORTOS_ES[Number(m[2]) - 1] || '';
    return `${m[3]}-${mes}`;
  };
  return `${fmtPart(sundayIso)} al ${fmtPart(satIso)}`;
}
