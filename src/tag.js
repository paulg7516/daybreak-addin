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

// Maps a header intent (canonical or legacy) to the canonical lane id.
const INTENT_ALIASES = {
  respond: 'respond', blocked: 'respond',
  approve: 'approve', action: 'approve',
  review: 'review', whenever: 'review',
  fyi: 'fyi',
};

// Parses an X-PTO-Triage header value -> { intent, by } (by is null if absent),
// or null if the value is not a Daybreak tag. The recipient/queue side of buildTagValue.
export function parseTagValue(raw) {
  if (!raw) return null;
  const parts = String(raw).trim().split(';');
  const intent = INTENT_ALIASES[(parts[0] || '').toLowerCase()];
  if (!intent) return null;
  let by = null;
  for (const p of parts.slice(1)) {
    const m = p.trim().match(/^by=(\d{4}-\d{2}-\d{2})$/i);
    if (m) by = m[1];
  }
  return { intent, by: intent === 'fyi' ? null : by };
}
