// addin/src/graph.js
// Reads the recipient's tagged mail from Microsoft Graph, from inside the add-in,
// using nested-app authentication (NAA) - no backend. PROTOTYPE: the auth + fetch
// path is wired but unproven until run in real Outlook against a registered Azure
// app. Set CLIENT_ID to that app's id to go live; while it is empty, the queue
// falls back to mock data so the panel still renders.
import { parseTagValue } from './tag.js';

// >>> Azure checklist: paste your registered app's Application (client) ID here. <<<
export const CLIENT_ID = '';

let pcaPromise = null;
function getPca() {
  // msal-browser is loaded from CDN in queue.html (window.msal).
  if (!pcaPromise) {
    pcaPromise = window.msal.createNestablePublicClientApplication({
      auth: { clientId: CLIENT_ID, authority: 'https://login.microsoftonline.com/common' },
    });
  }
  return pcaPromise;
}

async function getGraphToken() {
  const pca = await getPca();
  const account = pca.getActiveAccount() || pca.getAllAccounts()[0];
  const request = { scopes: ['Mail.Read'], account };
  try {
    const r = await pca.acquireTokenSilent(request);
    return r.accessToken;
  } catch {
    const r = await pca.acquireTokenPopup(request);
    return r.accessToken;
  }
}

// Graph cannot $filter on a custom header, so we pull recent messages with their
// internet headers and filter for X-PTO-Triage client-side.
function normalize(msg) {
  const headers = msg.internetMessageHeaders || [];
  const tagHeader = headers.find((h) => (h.name || '').toLowerCase() === 'x-pto-triage');
  const tag = parseTagValue(tagHeader && tagHeader.value);
  if (!tag) return null;
  return {
    id: msg.id,
    subject: msg.subject || '(no subject)',
    from: (msg.from && msg.from.emailAddress && msg.from.emailAddress.address) || '',
    fromName: (msg.from && msg.from.emailAddress && msg.from.emailAddress.name) || '',
    receivedAt: msg.receivedDateTime,
    intent: tag.intent,
    by: tag.by,
  };
}

export async function fetchTaggedItems() {
  const token = await getGraphToken();
  const url =
    'https://graph.microsoft.com/v1.0/me/messages' +
    '?$top=50&$select=id,subject,from,receivedDateTime,internetMessageHeaders';
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Graph ${res.status}`);
  const data = await res.json();
  return (data.value || []).map(normalize).filter(Boolean);
}
