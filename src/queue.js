// addin/src/queue.js
// The recipient triage queue, rendered inside the Outlook add-in panel (Option A:
// calm flat list). Data comes from Graph (graph.js) when CLIENT_ID is set; otherwise
// from mock items so the panel renders. Rows open the real message in Outlook and can
// be marked Done.
import { CLIENT_ID, fetchTaggedItems } from './graph.js';

// Lane markers are neutral - colour in the panel means urgency (warm chips) only;
// brand shows through the coral logo + top accent bar, matching Daybreak's palette.
const LANES = [
  { id: 'decision', title: 'Needs your decision', dot: '#5f6a7e' },
  { id: 'input', title: 'Needs your input', dot: '#5f6a7e' },
  { id: 'fyi', title: 'FYI', dot: '#5f6a7e' },
];

function daysFromToday(by) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const d = new Date(`${by}T00:00:00`);
  return Math.round((d.getTime() - today.getTime()) / 86400000);
}
function urgency(by) {
  if (!by) return null;
  const n = daysFromToday(by);
  if (Number.isNaN(n)) return null;
  if (n < 0) return { label: 'Overdue', cls: 'overdue' };
  if (n === 0) return { label: 'Today', cls: 'today' };
  if (n <= 7) return { label: 'This week', cls: 'week' };
  return null;
}
function senderName(item) {
  if (item.fromName && item.fromName.trim()) return item.fromName.trim();
  const local = (item.from || '').split('@')[0];
  const tidy = local.split(/[.\-_]+/).filter(Boolean).map((w) => w[0].toUpperCase() + w.slice(1)).join(' ');
  return tidy || item.from || 'Unknown sender';
}
function initials(name) {
  const p = name.split(/\s+/).filter(Boolean);
  return ((p[0]?.[0] ?? '?') + (p[1]?.[0] ?? '')).toUpperCase();
}
function relTime(iso) {
  if (!iso) return '';
  return new Date(iso).toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
}
function el(tag, cls, text) {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text != null) n.textContent = text;
  return n;
}

const REAL_DATA = !!CLIENT_ID;

// Collapse + cleared state persist in localStorage for the prototype; production
// should move these to Office roaming settings so they follow the user across devices.
const COLLAPSE_KEY = 'daybreak.queue.collapsed';
function loadCollapsed() {
  const raw = localStorage.getItem(COLLAPSE_KEY);
  if (raw == null) {
    const initial = ['fyi']; // FYI starts collapsed (no action needed)
    try { localStorage.setItem(COLLAPSE_KEY, JSON.stringify(initial)); } catch { /* ignore */ }
    return new Set(initial);
  }
  try { return new Set(JSON.parse(raw)); } catch { return new Set(); }
}
function saveCollapsed(set) {
  try { localStorage.setItem(COLLAPSE_KEY, JSON.stringify([...set])); } catch { /* ignore */ }
}
const CLEARED_KEY = 'daybreak.queue.cleared';
function loadCleared() {
  try { return new Set(JSON.parse(localStorage.getItem(CLEARED_KEY) || '[]')); } catch { return new Set(); }
}
function saveCleared(set) {
  try { localStorage.setItem(CLEARED_KEY, JSON.stringify([...set])); } catch { /* ignore */ }
}

let currentItems = [];

// Open the tagged message natively in Outlook. Only meaningful against real Graph
// data (mock rows have invented ids with no message behind them).
function openItem(item) {
  if (!REAL_DATA) return;
  try {
    const mb = window.Office && Office.context && Office.context.mailbox;
    if (mb && mb.displayMessageForm) mb.displayMessageForm(item.id);
  } catch { /* ignore */ }
}

function clearItem(item, rowEl) {
  rowEl.classList.add('leaving');
  setTimeout(() => {
    currentItems = currentItems.filter((i) => i.id !== item.id);
    if (REAL_DATA) { const s = loadCleared(); s.add(item.id); saveCleared(s); }
    renderQueue();
  }, 180);
}

function rowNode(item) {
  const row = el('div', 'row');
  row.setAttribute('role', 'button');
  row.tabIndex = 0;
  row.addEventListener('click', () => openItem(item));
  row.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); openItem(item); } });

  const av = el('div', 'av', initials(senderName(item)));
  const bd = el('div', 'bd');

  const r1 = el('div', 'r1');
  r1.appendChild(el('span', 'subj', item.subject));
  const u = urgency(item.by);
  if (u) r1.appendChild(el('span', `chip ${u.cls}`, u.label));
  const done = el('span', 'done', '✓ Done');
  done.setAttribute('role', 'button');
  done.setAttribute('aria-label', 'Mark done');
  done.tabIndex = 0;
  done.addEventListener('click', (e) => { e.stopPropagation(); clearItem(item, row); });
  done.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') { e.stopPropagation(); e.preventDefault(); clearItem(item, row); }
  });
  r1.appendChild(done);
  bd.appendChild(r1);

  const r2 = el('div', 'r2');
  r2.appendChild(document.createTextNode(`${senderName(item)} · ${relTime(item.receivedAt)}`));
  if (item.by) {
    r2.appendChild(document.createTextNode(' · '));
    r2.appendChild(el('span', 'due', `due ${relTime(item.by)}`));
  }
  bd.appendChild(r2);

  row.append(av, bd);
  return row;
}

function caughtUp() {
  const wrap = el('div', 'caughtup');
  wrap.appendChild(el('div', 'cu-badge', '✓'));
  wrap.appendChild(el('h3', null, "You're all caught up"));
  wrap.appendChild(el('p', null, 'Nothing needs you right now.'));
  return wrap;
}

function foot() {
  const f = el('div', 'foot');
  f.append(el('b', null, 'Click a row'), document.createTextNode(' opens that email in Outlook.  '),
    el('b', null, 'Hover'), document.createTextNode(' → ✓ Done to clear.'));
  return f;
}

function render(items) {
  const root = document.getElementById('lanes');
  root.textContent = '';
  if (items.length === 0) { root.appendChild(caughtUp()); return; }

  const collapsed = loadCollapsed();
  for (const lane of LANES) {
    const rows = items.filter((i) => i.intent === lane.id);
    const isCollapsed = collapsed.has(lane.id);
    const section = el('section', isCollapsed ? 'lane collapsed' : 'lane');

    const lh = el('div', 'lh');
    lh.setAttribute('role', 'button');
    lh.tabIndex = 0;
    lh.setAttribute('aria-expanded', String(!isCollapsed));
    const dot = el('span', 'dot');
    dot.style.background = lane.dot;
    lh.appendChild(dot);
    lh.appendChild(el('h2', null, lane.title));
    lh.appendChild(el('span', 'ct', String(rows.length)));
    lh.appendChild(el('span', 'chev', '▾'));

    const body = el('div', 'body');
    if (rows.length === 0) body.appendChild(el('p', 'empty-lane', 'All caught up here.'));
    else rows.forEach((r) => body.appendChild(rowNode(r)));

    const toggle = () => {
      const nowCollapsed = section.classList.toggle('collapsed');
      lh.setAttribute('aria-expanded', String(!nowCollapsed));
      const set = loadCollapsed();
      if (nowCollapsed) set.add(lane.id); else set.delete(lane.id);
      saveCollapsed(set);
    };
    lh.addEventListener('click', toggle);
    lh.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggle(); }
    });

    section.append(lh, body);
    root.appendChild(section);
  }
  root.appendChild(foot());
}

function renderQueue() { render(currentItems); }

function renderSkeleton() {
  const root = document.getElementById('lanes');
  root.textContent = '';
  for (let i = 0; i < 4; i += 1) {
    const s = el('div', 'sk');
    s.appendChild(el('div', 'sk-a'));
    const l = el('div', 'sk-l');
    l.appendChild(el('i'));
    l.appendChild(el('i'));
    s.appendChild(l);
    root.appendChild(s);
  }
}

function setStatus(msg) {
  const s = document.getElementById('status');
  s.textContent = msg || '';
  s.hidden = !msg;
}

function mockItems() {
  const iso = (d) => new Date(Date.now() - d * 86400000).toISOString();
  const day = (d) => { const x = new Date(); x.setDate(x.getDate() + d); return x.toISOString().slice(0, 10); };
  return [
    { id: '1', subject: 'Sign-off needed on the Q3 forecast', from: 'cfo@company.com', fromName: 'Dana Whitfield', receivedAt: iso(2), intent: 'decision', by: day(1) },
    { id: '2', subject: 'Quick question on the staging rollout', from: 'sam@company.com', fromName: 'Sam Okafor', receivedAt: iso(1), intent: 'input', by: null },
    { id: '3', subject: 'Re: incident postmortem - your input?', from: 'lead@company.com', fromName: 'Priya Nair', receivedAt: iso(0), intent: 'input', by: null },
    { id: '4', subject: 'Vendor contract redlines for a look', from: 'legal@company.com', fromName: 'Marcus Cole', receivedAt: iso(3), intent: 'input', by: null },
    { id: '5', subject: 'Heads up: office closed Friday', from: 'facilities@company.com', fromName: 'Facilities', receivedAt: iso(1), intent: 'fyi', by: null },
  ];
}

async function load() {
  if (!REAL_DATA) { currentItems = mockItems(); renderQueue(); return; }
  renderSkeleton();
  try {
    const items = await fetchTaggedItems();
    setStatus('');
    const cleared = loadCleared();
    currentItems = items.filter((i) => !cleared.has(i.id));
    renderQueue();
  } catch (e) {
    setStatus(`Couldn't load your tagged mail: ${e.message}`);
    currentItems = mockItems();
    renderQueue();
  }
}

// Live mode (NAA) needs the Office host bridge, so wait for Office.onReady. Mock mode
// has no such dependency, so it renders immediately (and works in headless preview).
if (REAL_DATA && window.Office && window.Office.onReady) {
  window.Office.onReady(() => { void load(); });
} else {
  void load();
}
