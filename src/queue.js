// addin/src/queue.js
// PROTOTYPE: the recipient triage board, rendered inside the Outlook add-in panel.
// Mirrors the desktop app's stacked board at panel width. Data comes from Graph
// (graph.js) when CLIENT_ID is set; otherwise from mock items so the panel renders.
import { CLIENT_ID, fetchTaggedItems } from './graph.js';

const LANES = [
  { id: 'respond', title: 'Needs your reply', desc: 'People waiting on your answer' },
  { id: 'approve', title: 'Needs your decision', desc: 'Sign-offs and decisions for you' },
  { id: 'review', title: 'Needs your review', desc: 'Worth a look when you can' },
  { id: 'fyi', title: 'FYI', desc: 'No action needed, just so you know' },
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
function hue(s) { let h = 0; for (const c of s) h = (h * 31 + c.charCodeAt(0)) >>> 0; return h % 360; }
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

function rowNode(item) {
  const row = el('div', 'row');
  const av = el('div', 'avatar', initials(senderName(item)));
  av.style.background = `hsl(${hue(item.from || item.subject)} 45% 55%)`;
  const body = el('div', 'body');

  const line1 = el('div', 'line1');
  const u = urgency(item.by);
  if (u) line1.appendChild(el('span', `chip ${u.cls}`, u.label));
  line1.appendChild(el('span', 'subject', item.subject));

  const line2 = el('div', 'line2');
  line2.appendChild(el('span', 'who', senderName(item)));
  line2.appendChild(el('span', 'src', 'Mail'));
  line2.appendChild(el('span', 'dot', '·'));
  line2.appendChild(el('span', 'when', relTime(item.receivedAt)));

  body.append(line1, line2);
  if (item.by) body.appendChild(el('div', 'reason', `by ${relTime(item.by)}`));
  row.append(av, body);
  return row;
}

// Collapse state persists per lane. First run has no stored preference, so FYI
// starts collapsed (it's the "no action needed" lane). localStorage is fine for the
// prototype; production should move this to Office roaming settings.
const COLLAPSE_KEY = 'daybreak.queue.collapsed';
function loadCollapsed() {
  const raw = localStorage.getItem(COLLAPSE_KEY);
  if (raw == null) {
    const initial = ['fyi'];
    try { localStorage.setItem(COLLAPSE_KEY, JSON.stringify(initial)); } catch { /* ignore */ }
    return new Set(initial);
  }
  try { return new Set(JSON.parse(raw)); } catch { return new Set(); }
}
function saveCollapsed(set) {
  try { localStorage.setItem(COLLAPSE_KEY, JSON.stringify([...set])); } catch { /* ignore */ }
}

function render(items) {
  const root = document.getElementById('lanes');
  root.textContent = '';
  const collapsed = loadCollapsed();

  for (const lane of LANES) {
    const rows = items.filter((i) => i.intent === lane.id);
    const isCollapsed = collapsed.has(lane.id);
    const section = el('section', isCollapsed ? 'lane collapsed' : 'lane');
    section.appendChild(el('div', `rail ${lane.id}`));

    const head = el('div', 'lane-head');
    head.setAttribute('role', 'button');
    head.tabIndex = 0;
    head.setAttribute('aria-expanded', String(!isCollapsed));
    head.appendChild(el('span', 'chev'));
    const ht = el('div', 'lane-titles');
    ht.appendChild(el('h2', null, lane.title));
    ht.appendChild(el('p', null, lane.desc));
    head.append(ht, el('span', 'count', String(rows.length)));

    const bodyEl = el('div', 'lane-body');
    if (rows.length === 0) {
      bodyEl.appendChild(el('p', 'empty', 'All caught up here.'));
    } else {
      rows.forEach((r) => bodyEl.appendChild(rowNode(r)));
    }

    const toggle = () => {
      const nowCollapsed = section.classList.toggle('collapsed');
      head.setAttribute('aria-expanded', String(!nowCollapsed));
      const set = loadCollapsed();
      if (nowCollapsed) set.add(lane.id); else set.delete(lane.id);
      saveCollapsed(set);
    };
    head.addEventListener('click', toggle);
    head.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggle(); }
    });

    section.append(head, bodyEl);
    root.appendChild(section);
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
    { id: '1', subject: 'Sign-off needed on the Q3 forecast', from: 'cfo@company.com', fromName: 'Dana Whitfield', receivedAt: iso(2), intent: 'approve', by: day(1) },
    { id: '2', subject: 'Quick question on the staging rollout', from: 'sam@company.com', fromName: 'Sam Okafor', receivedAt: iso(1), intent: 'respond', by: null },
    { id: '3', subject: 'Re: incident postmortem - your input?', from: 'lead@company.com', fromName: 'Priya Nair', receivedAt: iso(0), intent: 'respond', by: day(0) },
    { id: '4', subject: 'Vendor contract redlines for a look', from: 'legal@company.com', fromName: 'Marcus Cole', receivedAt: iso(3), intent: 'review', by: day(5) },
    { id: '5', subject: 'Heads up: office closed Friday', from: 'facilities@company.com', fromName: 'Facilities', receivedAt: iso(1), intent: 'fyi', by: null },
  ];
}

async function load() {
  if (!CLIENT_ID) {
    setStatus('Preview data - set CLIENT_ID in graph.js to read your real tagged mail.');
    render(mockItems());
    return;
  }
  setStatus('Loading your tagged mail...');
  try {
    const items = await fetchTaggedItems();
    setStatus('');
    render(items);
  } catch (e) {
    setStatus(`Could not load from Graph: ${e.message}. Showing preview data.`);
    render(mockItems());
  }
}

// Live mode (NAA) needs the Office host bridge, so wait for Office.onReady. Mock mode
// has no such dependency, so it renders immediately (and works in headless preview).
if (CLIENT_ID && window.Office && window.Office.onReady) {
  window.Office.onReady(() => { void load(); });
} else {
  void load();
}
