/**
 * Gestión de almacenamiento persistente en el navegador (Storage API).
 */

export async function requestPersistentStorage() {
  try {
    if (navigator.storage && typeof navigator.storage.persist === 'function') {
      const persisted = await navigator.storage.persisted();
      if (persisted) return true;
      return await navigator.storage.persist();
    }
  } catch {
    // ignore storage errors
  }
  return false;
}

export async function getStorageUsageEstimate() {
  try {
    let isPersisted = false;
    if (navigator.storage && typeof navigator.storage.persisted === 'function') {
      isPersisted = await navigator.storage.persisted();
    }
    if (navigator.storage && typeof navigator.storage.estimate === 'function') {
      const est = await navigator.storage.estimate();
      const usage = est.usage || 0;
      const quota = est.quota || 0;
      return {
        isPersisted,
        usageBytes: usage,
        quotaBytes: quota,
        usageMB: (usage / (1024 * 1024)).toFixed(2),
        quotaMB: Math.round(quota / (1024 * 1024))
      };
    }
  } catch {
    // ignore estimate errors
  }
  return { isPersisted: false, usageBytes: 0, quotaBytes: 0, usageMB: '0.00', quotaMB: 0 };
}

export function formatBytes(bytes = 0) {
  const b = Number(bytes) || 0;
  if (b < 1024) return `${b} B`;
  if (b < 1024 * 1024) return `${(b / 1024).toFixed(1)} KB`;
  return `${(b / (1024 * 1024)).toFixed(2)} MB`;
}
