/**
 * Generador de Reporte de Turno para WhatsApp (Formato Oficial STRACON — Fractura de Roca).
 */

import { formatFechaDDMMYYYY } from './datePeru.js';
import { parseNumOrNull, round1 } from './codeNormalizer.js';

function formatMaterialValue(val) {
  const num = parseNumOrNull(val);
  if (num === null) return '0';
  return Number.isInteger(num) ? String(num) : String(round1(num));
}

function buildActivitiesList(observaciones = '') {
  const raw = String(observaciones || '').trim();
  const baseActivity = 'Picotoneo de roca';
  if (!raw) {
    return [baseActivity];
  }

  const extras = raw
    .split(/\r?\n|;/)
    .map((line) => line.replace(/^\s*(?:\d+[\.\)\-]|[\-\*\•])\s*/, '').trim())
    .filter(Boolean);

  if (extras.length === 0) {
    return [baseActivity];
  }

  const hasPicotoneo = extras.some((item) => item.toLowerCase().includes('picotoneo'));
  return hasPicotoneo ? extras : [baseActivity, ...extras];
}

/**
 * Construye el texto exacto formateado para WhatsApp del reporte de turno.
 */
export function formatShiftWhatsAppReport(record = {}) {
  const codigoEquipo = String(record.codigo_equipo || 'STREX01').trim().toUpperCase();
  const turnoLabel = record.turno === 'NOCHE' ? 'Noche' : 'Día';
  const fechaLabel = formatFechaDDMMYYYY(record.fecha || '');
  const areaClean = String(record.area_trabajo || '').trim().replace(/:+$/, '');
  const operadorClean = String(record.operador || '').trim() || '-';

  const hi = parseNumOrNull(record.horometro_inicial);
  const hf = parseNumOrNull(record.horometro_final);
  const horasNum =
    parseNumOrNull(record.horas_trabajadas) !== null
      ? round1(parseNumOrNull(record.horas_trabajadas))
      : hi !== null && hf !== null && hf >= hi
        ? round1(hf - hi)
        : 0;

  const hiStr = hi !== null ? hi.toFixed(1) : '-';
  const hfStr = hf !== null ? hf.toFixed(1) : '-';
  const horasStr = horasNum.toFixed(1);
  const prodStr = formatMaterialValue(record.material_fracturado);
  const activities = buildActivitiesList(record.observaciones);

  const lines = [
    '*REPORTE DE FRACTURA DE ROCA*',
    `*Excavadora Komatsu Pc390 - ${codigoEquipo}*`,
    `Turno: ${turnoLabel}`,
    `Fecha: ${fechaLabel}`
  ];

  if (areaClean) {
    lines.push('', `*${areaClean}:*`);
  }

  lines.push(
    '',
    `- Operador: ${operadorClean}`,
    `- Hr I: ${hiStr}`,
    `- Hr F: ${hfStr}`,
    `- Horas trabajadas: ${horasStr} hrs`,
    `- Producción: ${prodStr} m3`,
    '',
    'Actividad:',
    ...activities.map((act, idx) => `${idx + 1}. ${act}`)
  );

  return lines.join('\n');
}

/**
 * Copia texto al portapapeles con fallback compatible para móviles y entornos offline.
 */
export async function copyTextToClipboard(text) {
  if (typeof navigator !== 'undefined' && navigator.clipboard && navigator.clipboard.writeText) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch {
      // Continuar con fallback textarea
    }
  }

  if (typeof document === 'undefined') return false;
  const textarea = document.createElement('textarea');
  textarea.value = text;
  textarea.setAttribute('readonly', '');
  textarea.style.position = 'fixed';
  textarea.style.top = '-9999px';
  textarea.style.left = '-9999px';
  document.body.appendChild(textarea);
  textarea.focus();
  textarea.select();
  let ok = false;
  try {
    ok = document.execCommand('copy');
  } catch {
    ok = false;
  }
  document.body.removeChild(textarea);
  return ok;
}
