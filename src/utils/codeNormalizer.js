/**
 * Normalización de códigos de equipo, áreas de trabajo y redondeos numéricos.
 */

export function normalizeEquipmentCode(rawInput = '') {
  if (!rawInput || typeof rawInput !== 'string') return '';
  let cleaned = rawInput.trim().toUpperCase();
  cleaned = cleaned.replace(/[^A-Z0-9\s\-_]/g, '');
  cleaned = cleaned
    .replace(/[\s_]+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '');
  return cleaned;
}

export function normalizeAreaName(rawInput = '') {
  if (!rawInput || typeof rawInput !== 'string') return '';
  return rawInput.trim().replace(/\s+/g, ' ');
}

export function parseNumOrNull(val) {
  if (val === '' || val === null || val === undefined) return null;
  const num = Number(val);
  return Number.isFinite(num) ? num : null;
}

export function round1(n) {
  return Math.round((Number(n) || 0) * 10) / 10;
}

export function round2(n) {
  return Math.round((Number(n) || 0) * 100) / 100;
}

export function escapeHtml(str = '') {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
