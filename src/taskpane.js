// addin/src/taskpane.js
import { buildTagValue, validateBccAddress } from './tag.js';

const HEADER = 'X-PTO-Triage';
const LANE_LABEL = {
  decision: 'Needs your decision',
  input: 'Needs your input',
  fyi: 'FYI',
};

function el(id) { return document.getElementById(id); }
function setStatus(msg, kind) {
  const s = el('status');
  s.textContent = msg;
  s.className = `status${kind ? ' ' + kind : ''}`;
}
function selectedIntent() {
  const checked = document.querySelector('input[name="intent"]:checked');
  return checked ? checked.value : null;
}
function refreshControls() {
  const intent = selectedIntent();
  // Only a decision carries a deadline, and it requires one.
  el('dateRow').hidden = intent !== 'decision';
  el('apply').disabled = intent === null || (intent === 'decision' && !el('byDate').value);
}
function fmtDate(v) {
  if (!v) return '';
  const d = new Date(`${v}T00:00:00`);
  return Number.isNaN(d.getTime()) ? v : d.toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
}

// Swap the form for the success card and spell out exactly what was applied.
function showDone(intent, date, bcc) {
  const summary = el('doneSummary');
  summary.textContent = '';
  const chip = document.createElement('span');
  chip.className = 'chip';
  chip.textContent = LANE_LABEL[intent] || intent;
  summary.appendChild(chip);
  let rest = '';
  if (intent === 'decision' && date) rest += ` by ${fmtDate(date)}`;
  if (bcc) rest += ` · also looped in ${bcc}`;
  if (rest) summary.appendChild(document.createTextNode(rest));

  el('form').hidden = true;
  el('done').hidden = false;
}
function showForm() {
  el('done').hidden = true;
  el('form').hidden = false;
  setStatus('');
  refreshControls();
}

// Promisified Office async call.
function officeCall(fn) {
  return new Promise((resolve, reject) => {
    fn((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
      else reject(new Error(res.error ? res.error.message : 'Office call failed'));
    });
  });
}

// Remember the Bcc we already added so re-applying (after "Change tag") does not
// stack duplicate recipients on the draft.
let addedBcc = null;

async function apply() {
  const intent = selectedIntent();
  if (!intent) return;
  let value;
  try {
    value = buildTagValue(intent, el('byDate').value);
  } catch (e) {
    setStatus(e.message, 'err');
    return;
  }

  const bcc = el('bcc').value.trim();
  if (bcc && !validateBccAddress(bcc)) {
    setStatus('That does not look like an email address.', 'err');
    return;
  }

  el('apply').disabled = true;
  setStatus('Applying...');
  try {
    const item = Office.context.mailbox.item;
    await officeCall((cb) => item.internetHeaders.setAsync({ [HEADER]: value }, cb));
    if (bcc && bcc !== addedBcc) {
      await officeCall((cb) => item.bcc.addAsync([{ emailAddress: bcc }], cb));
      addedBcc = bcc;
    }
    showDone(intent, el('byDate').value, bcc);
  } catch (e) {
    setStatus(`Could not apply: ${e.message}`, 'err');
    el('apply').disabled = false;
  }
}

Office.onReady(() => {
  document.querySelectorAll('input[name="intent"]').forEach((r) => r.addEventListener('change', refreshControls));
  el('apply').addEventListener('click', apply);
  el('change').addEventListener('click', showForm);
  el('byDate').addEventListener('input', refreshControls);
  refreshControls();
});
