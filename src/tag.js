// addin/src/tag.js
// Pure helpers for the Daybreak compose add-in. No Office.js dependency, so they
// load in the browser AND run under Vitest. The header values match exactly what
// the recipient-side parser (src/scoring/sender-tag.ts) understands.

const INTENTS = new Set(['respond', 'approve', 'review', 'fyi']);

// Returns 'YYYY-MM-DD' for a Date or a 'YYYY-MM-DD' string, or null if invalid.
export function formatByDate(value) {
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return null;
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, '0');
    const d = String(value.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return value;
  }
  return null;
}

// Builds the X-PTO-Triage header value for an intent. respond/approve/review may
// carry an optional ';by=YYYY-MM-DD' deadline; fyi never does.
export function buildTagValue(intent, byDate) {
  if (!INTENTS.has(intent)) {
    throw new Error(`Unknown intent: ${intent}`);
  }
  if (intent !== 'fyi') {
    const by = formatByDate(byDate);
    if (by) return `${intent};by=${by}`;
  }
  return intent;
}

// A pragmatic email check for the optional bcc-routing field.
export function validateBccAddress(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value).trim());
}
