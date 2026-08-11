/* ---------------------------------------------------------------------------
 * background.js — service worker (MV3, ES module)
 *
 * Session store: the captured bearer tokens the extension uses to talk to
 * Microsoft Graph on your behalf. It lives exclusively in
 * chrome.storage.session — that is RAM, cleared when the browser closes.
 * Tokens are never written to disk: no storage.local, no storage.sync, no file
 * export. They are never shown or handed out either: the popup receives only
 * the one token it needs for its own Graph calls, and there is no UI that
 * lists, displays or copies a token.
 *
 * chrome.storage.local is used only for non-secret data: user preferences and
 * locally scheduled activation jobs.
 * ------------------------------------------------------------------------- */

import {
  Graph, activationBody, deactivationBody, toRole, FINAL_STATUSES,
} from './lib/graph.js';
import { summarize, classifyToken, vaultKey, isExpired, decodeJwt } from './lib/jwt.js';
import { humanize } from './lib/fmt.js';

const VAULT_KEY = 'vault';
const JOBS_KEY = 'jobs';
const PREFS_KEY = 'prefs';

const REFRESH_ALARM = 'pim:refresh';
const JOB_PREFIX = 'pim:job:';

// theme: 'auto' follows the operating system until the user picks a side.
const DEFAULT_PREFS = { theme: 'auto' };

/**
 * The admin portals the popup can jump to. Keyed by the id the popup sends, so
 * no URL ever travels through a message — an id that is not in this table is
 * refused.
 */
const PORTALS = {
  pim: 'https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadmigratedroles',
  entra: 'https://entra.microsoft.com/',
  azure: 'https://portal.azure.com/',
  m365: 'https://admin.microsoft.com/',
  security: 'https://security.microsoft.com/',
  purview: 'https://purview.microsoft.com/',
  intune: 'https://intune.microsoft.com/',
};

/* ============================== preferences ============================== */

async function getPrefs() {
  const bag = await chrome.storage.local.get(PREFS_KEY);
  return { ...DEFAULT_PREFS, ...(bag[PREFS_KEY] || {}) };
}

async function setPrefs(patch) {
  const merged = { ...(await getPrefs()), ...patch };
  await chrome.storage.local.set({ [PREFS_KEY]: merged });
  return merged;
}

/* ============================= session store ============================= */

/** Read the store. Session storage only — nothing is ever loaded from disk. */
async function getVault() {
  const bag = await chrome.storage.session.get(VAULT_KEY);
  return bag[VAULT_KEY] || {};
}

async function writeVault(vault) {
  await chrome.storage.session.set({ [VAULT_KEY]: vault });
}

/** Store one captured token. Returns the record, or null when it is ignored. */
async function saveToken(rawToken, sourceUrl) {
  const info = summarize(rawToken);
  if (!info) return null;

  const claims = decodeJwt(rawToken);
  if (isExpired(claims)) return null;

  const kind = classifyToken(claims);
  // Only keep tokens for Microsoft resources we understand (Graph or Intune).
  if (kind.resource === 'other') return null;

  const key = vaultKey(info, kind);
  const vault = await getVault();
  const existing = vault[key];

  // Do not overwrite a stored token that is still valid and no older.
  if (existing && existing.expiresAt >= (info.expiresAt || 0) && existing.token === rawToken) {
    return existing;
  }

  vault[key] = {
    key,
    token: rawToken,
    resource: kind.resource,
    scopeClass: kind.scopeClass,
    label: kind.label,
    tenantId: info.tenantId,
    principalId: info.principalId,
    upn: info.upn,
    name: info.name,
    scopes: info.scopes,
    audience: String(info.audience || ''),
    expiresAt: info.expiresAt,
    capturedAt: Date.now(),
    sourceUrl: sourceUrl || '',
  };

  await writeVault(vault);
  await refreshBadges();
  return vault[key];
}

/** Unexpired tokens, PIM-capable first, then newest. */
async function listTokens() {
  const vault = await getVault();
  return Object.values(vault)
    .filter((t) => !isExpired(decodeJwt(t.token)))
    .sort((a, b) => {
      const rank = (t) => (t.scopeClass === 'role' ? 0 : t.resource === 'graph' ? 1 : 2);
      return rank(a) - rank(b) || b.capturedAt - a.capturedAt;
    });
}

/** The Graph token used for PIM: prefer one carrying role-management scopes. */
async function getPimToken() {
  const graph = (await listTokens()).filter((t) => t.resource === 'graph');
  return graph.find((t) => t.scopeClass === 'role') || graph[0] || null;
}

/**
 * What the popup is told about the session: who it belongs to, when it dies,
 * and the token it needs for its own Graph reads. Nothing else — there is no
 * message that hands out the full store.
 */
async function sessionForPopup() {
  const pim = await getPimToken();
  if (!pim) return null;
  return {
    token: pim.token,
    principalId: pim.principalId,
    tenantId: pim.tenantId,
    upn: pim.upn,
    name: pim.name,
    expiresAt: pim.expiresAt,
  };
}

/* ================================ badges ================================= */

async function refreshBadges() {
  const pim = await getPimToken();
  if (!pim) {
    await chrome.action.setBadgeText({ text: '' });
    return;
  }
  try {
    const graph = new Graph(pim.token);
    const rows = await graph.activeInstances(pim.principalId);
    const count = rows.map(toRole).filter((r) => r.assignmentType === 'Activated').length;
    await chrome.action.setBadgeText({ text: count ? String(count) : '' });
    await chrome.action.setBadgeBackgroundColor({ color: '#f5a524' });
    if (chrome.action.setBadgeTextColor) {
      await chrome.action.setBadgeTextColor({ color: '#17130a' });
    }
  } catch {
    /* the badge is a nicety, never a failure path */
  }
}

/* ========================== activate / deactivate ======================== */

async function runActivation({ roles, duration, justification, startAt, ticketNumber, ticketSystem }) {
  const pim = await getPimToken();
  if (!pim) throw new Error('NO_PIM_TOKEN');

  const graph = new Graph(pim.token);
  const settled = await Promise.all(
    roles.map(async (role) => {
      try {
        const created = await graph.createRequest(
          activationBody({
            principalId: pim.principalId,
            role,
            duration,
            // a role may carry its own reason; activationBody applies the fallback
            justification,
            startAt,
            ticketNumber,
            ticketSystem,
          })
        );
        return { role, ok: true, requestId: created.id || '', status: created.status || 'Unknown' };
      } catch (error) {
        return {
          role, ok: false,
          code: error.code || 'error',
          message: error.message || String(error),
          hint: error.hint || '',
        };
      }
    })
  );

  await Promise.all(
    settled
      .filter((r) => r.ok && r.requestId && !FINAL_STATUSES.has(r.status))
      .map(async (result) => {
        for (let i = 0; i < 6; i++) {
          await new Promise((r) => setTimeout(r, 2500));
          try {
            const fresh = await graph.requestStatus(result.requestId);
            result.status = fresh.status || result.status;
            if (FINAL_STATUSES.has(result.status)) break;
          } catch {
            break;
          }
        }
      })
  );

  await refreshBadges();
  return settled;
}

async function runDeactivation({ roles, justification }) {
  const pim = await getPimToken();
  if (!pim) throw new Error('NO_PIM_TOKEN');

  const graph = new Graph(pim.token);
  const results = await Promise.all(
    roles.map(async (role) => {
      try {
        const created = await graph.createRequest(
          deactivationBody({ principalId: pim.principalId, role, justification })
        );
        return { role, ok: true, status: created.status || 'Submitted' };
      } catch (error) {
        return {
          role, ok: false,
          code: error.code || 'error',
          message: error.message || String(error),
          hint: error.hint || '',
        };
      }
    })
  );
  await refreshBadges();
  return results;
}

/* ========================== locally scheduled jobs ======================= */

async function listJobs() {
  const bag = await chrome.storage.local.get(JOBS_KEY);
  return bag[JOBS_KEY] || [];
}

async function writeJobs(jobs) {
  await chrome.storage.local.set({ [JOBS_KEY]: jobs });
}

async function createJob(job) {
  const jobs = await listJobs();
  const record = { id: crypto.randomUUID(), createdAt: Date.now(), lastRun: null, lastStatus: '', ...job };
  jobs.push(record);
  await writeJobs(jobs);
  await chrome.alarms.create(JOB_PREFIX + record.id, {
    when: record.runAt,
    ...(record.repeatDaily ? { periodInMinutes: 1440 } : {}),
  });
  return record;
}

async function deleteJob(id) {
  const jobs = await listJobs();
  await writeJobs(jobs.filter((j) => j.id !== id));
  await chrome.alarms.clear(JOB_PREFIX + id);
}

async function fireJob(id) {
  const jobs = await listJobs();
  const job = jobs.find((j) => j.id === id);
  if (!job) {
    await chrome.alarms.clear(JOB_PREFIX + id);
    return;
  }

  const pim = await getPimToken();
  if (!pim) {
    const fresh = await listJobs();
    const target = fresh.find((j) => j.id === id);
    if (target) { target.lastRun = Date.now(); target.lastStatus = 'No PIM token'; await writeJobs(fresh); }
    await notify('Schedule could not run',
      'No valid PIM token is available. Open the PIM portal to capture a new session, then try again.', true);
    return;
  }

  let summary;
  try {
    const results = await runActivation({
      roles: job.roles, duration: job.duration, justification: job.justification,
      ticketNumber: job.ticketNumber, ticketSystem: job.ticketSystem,
    });
    const ok = results.filter((r) => r.ok).length;
    const failed = results.filter((r) => !r.ok);
    summary = `${ok}/${results.length} roles activated`;
    await notify(
      failed.length ? 'Schedule finished with errors' : 'Roles activated on schedule',
      failed.length
        ? `${summary}. Failed: ${failed.map((f) => f.role.name).join(', ')}`
        : `${summary} for ${humanize(job.durationMs || 0)}`,
      Boolean(failed.length)
    );
  } catch (error) {
    summary = error.message === 'NO_PIM_TOKEN' ? 'No PIM token' : 'Error: ' + error.message;
    await notify('Schedule failed', summary, true);
  }

  const fresh = await listJobs();
  const target = fresh.find((j) => j.id === id);
  if (target) {
    target.lastRun = Date.now();
    target.lastStatus = summary;
    if (!target.repeatDaily) {
      await writeJobs(fresh.filter((j) => j.id !== id));
      await chrome.alarms.clear(JOB_PREFIX + id);
      return;
    }
    target.runAt += 86400000;
    await writeJobs(fresh);
  }
}

async function notify(title, message, isProblem = false) {
  try {
    await chrome.notifications.create({
      type: 'basic',
      iconUrl: chrome.runtime.getURL('icons/icon128.png'),
      title, message, priority: isProblem ? 2 : 1,
    });
  } catch {
    /* the user may have turned notifications off */
  }
}

/* =============================== open portal ============================= */

async function openPortal(id) {
  const url = PORTALS[id];
  if (!url) throw new Error('unknown_portal');

  const host = new URL(url).host;
  const tabs = await chrome.tabs.query({ url: `https://${host}/*` });
  // A tab already sitting on this exact address is just focused. One on the
  // same host but elsewhere is navigated — that is how the Azure portal and the
  // PIM blade, which share a host, stay one click apart.
  const onTarget = tabs.find((t) => (t.url || '').startsWith(url));
  const reuse = onTarget || tabs[0];

  if (reuse) {
    await chrome.tabs.update(reuse.id, onTarget ? { active: true } : { active: true, url });
    await chrome.windows.update(reuse.windowId, { focused: true });
  } else {
    await chrome.tabs.create({ url });
  }
}

/* ================================= events ================================ */

/**
 * Older builds had an opt-in "keep tokens across sessions" switch that copied
 * the vault to disk. That option is gone; wipe anything it may have left behind
 * and drop the dead preference.
 */
async function purgeLegacyDiskTokens() {
  await chrome.storage.local.remove(VAULT_KEY);
  const bag = await chrome.storage.local.get(PREFS_KEY);
  const prefs = bag[PREFS_KEY];
  if (prefs && 'persistTokens' in prefs) {
    delete prefs.persistTokens;
    await chrome.storage.local.set({ [PREFS_KEY]: prefs });
  }
}

function onBoot() {
  chrome.alarms.create(REFRESH_ALARM, { periodInMinutes: 5 });
  purgeLegacyDiskTokens().catch(() => {});
}

chrome.runtime.onInstalled.addListener(onBoot);
chrome.runtime.onStartup.addListener(onBoot);

chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === REFRESH_ALARM) refreshBadges().catch(() => {});
  else if (alarm.name.startsWith(JOB_PREFIX)) fireJob(alarm.name.slice(JOB_PREFIX.length)).catch(() => {});
});

// Fallback path: read the header straight off the network layer on the way to Graph.
try {
  chrome.webRequest.onSendHeaders.addListener(
    (details) => {
      const header = (details.requestHeaders || []).find((h) => h.name.toLowerCase() === 'authorization');
      if (!header || !header.value) return;
      const token = header.value.replace(/^\s*bearer\s+/i, '').trim();
      if (token.split('.').length === 3) saveToken(token, details.url).catch(() => {});
    },
    { urls: ['https://graph.microsoft.com/*'] },
    ['requestHeaders', 'extraHeaders']
  );
} catch {
  /* some Edge builds restrict webRequest — the fetch/XHR hook still works */
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  (async () => {
    switch (message && message.type) {
      case 'token/captured':
        return { ok: Boolean(await saveToken(message.token, message.url)) };

      case 'session/get':
        return { session: await sessionForPopup(), prefs: await getPrefs() };

      case 'portal/open':
        await openPortal(message.portal || 'pim');
        return { ok: true };

      case 'role/activate':
        return { results: await runActivation(message.payload) };
      case 'role/deactivate':
        return { results: await runDeactivation(message.payload) };

      case 'jobs/list':
        return { jobs: await listJobs() };
      case 'jobs/create':
        return { job: await createJob(message.payload) };
      case 'jobs/delete':
        await deleteJob(message.id);
        return { ok: true };

      case 'badge/refresh':
        await refreshBadges();
        return { ok: true };

      case 'prefs/get':
        return { prefs: await getPrefs() };
      case 'prefs/set':
        return { prefs: await setPrefs(message.payload) };

      case 'portal/alive':
        return { ok: true };

      default:
        return { error: 'unknown_message' };
    }
  })()
    .then(sendResponse)
    .catch((error) => sendResponse({ error: error.message || String(error), code: error.code || '' }));
  return true;
});
