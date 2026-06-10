// addin/src/taskpane.js
import { buildTagValue, validateBccAddress } from './tag.js';

const HEADER = 'X-PTO-Triage';

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
  el('dateRow').hidden = intent !== 'action';
  el('apply').disabled = intent === null;
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
    if (bcc) {
      await officeCall((cb) => item.bcc.addAsync([{ emailAddress: bcc }], cb));
    }
    setStatus(bcc ? `Tagged and routing to ${bcc}. Send when ready.` : 'Tagged. Send when ready.', 'ok');
  } catch (e) {
    setStatus(`Could not apply: ${e.message}`, 'err');
  } finally {
    el('apply').disabled = false;
  }
}

Office.onReady(() => {
  document.querySelectorAll('input[name="intent"]').forEach((r) => r.addEventListener('change', refreshControls));
  el('apply').addEventListener('click', apply);
  refreshControls();
});
