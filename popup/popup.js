import {
  Graph, readPolicy, toRole, toHistoryEntry, indexPermissions,
} from '../src/lib/graph.js';
import {
  clock, formatDateTime, formatTime, formatDay, dayKey, humanize, normalizeDuration,
  parseIsoDuration, toIsoDuration, fromDatetimeLocal,
} from '../src/lib/fmt.js';

const $ = (id) => document.getElementById(id);

function esc(v) {
  return String(v ?? '').replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

async function ask(type, extra = {}) {
  const res = await chrome.runtime.sendMessage({ type, ...extra });
  if (res && res.error) throw new Error(res.error);
  return res || {};
}

const keyOf = (r) => `${r.roleDefinitionId}|${r.scopeId || '/'}`;

function toast(title, text = '', tone = '') {
  const node = document.createElement('div');
  node.className = 'toast' + (tone ? ` is-${tone}` : '');
  node.innerHTML = `<div><div class="toast-title">${esc(title)}</div>` +
    (text ? `<div class="toast-text">${esc(text)}</div>` : '') + `</div>`;
  $('toasts').appendChild(node);
  setTimeout(() => {
    node.classList.add('is-out');
    node.addEventListener('animationend', () => node.remove(), { once: true });
  }, tone === 'bad' ? 6500 : 4000);
}

/* ================================= state ================================= */

const state = {
  pim: null,            // the captured session the popup makes Graph calls with
  prefs: { theme: 'auto' },
  graph: null,
  eligible: [], active: [], assigned: [],
  jobs: [],
  policies: new Map(),
  definitions: new Map(),  // roleDefinitionId -> { permissions, description } | null
  selected: new Set(),
  openDrawers: new Map(),  // key -> 'perms' | 'activate'
  permsExpanded: new Set(), // keys whose permission list is fully expanded
  reasons: new Map(),      // key -> per-role justification for a bulk activation
  reasonsOpen: false,
  schedulePrefill: null,

  // activation history
  history: [],
  historyState: 'idle',    // 'idle' | 'loading' | 'ready' | 'error'
  historyError: '',
  historyFilter: 'all',    // 'all' | 'on' | 'off'

  // the aggregated permission index
  permIndex: null,         // Map<action, { grantors: [{ name, kind }] }>
  permState: 'idle',       // 'idle' | 'loading' | 'ready' | 'error'
  permError: '',
  permProgress: [0, 0],
  permScope: 'now',        // 'now' = active + assigned, 'all' = also eligible
  permQuery: '',
  permGroupsOpen: new Set(),
};

/** The admin portals the menu can jump to; the worker owns the actual URLs. */
const PORTALS = [
  {
    id: 'pim', name: 'Privileged Identity Management', host: 'portal.azure.com',
    note: 'role activation blade', mark: 'PIM', tone: 't-brand', session: true,
  },
  {
    id: 'entra', name: 'Microsoft Entra admin center', host: 'entra.microsoft.com',
    note: 'identity, users, groups', mark: 'ID', tone: 't-brand', session: true,
  },
  {
    id: 'azure', name: 'Azure portal', host: 'portal.azure.com',
    note: 'subscriptions and resources', mark: 'AZ', tone: 't-calm', session: true,
  },
  {
    id: 'm365', name: 'Microsoft 365 admin center', host: 'admin.microsoft.com',
    note: 'tenant, licences, users', mark: 'M365', tone: 't-ok',
  },
  {
    id: 'security', name: 'Microsoft Defender', host: 'security.microsoft.com',
    note: 'incidents and threat protection', mark: 'DEF', tone: 't-danger',
  },
  {
    id: 'purview', name: 'Microsoft Purview', host: 'purview.microsoft.com',
    note: 'compliance, DLP, audit', mark: 'PUR', tone: 't-elig',
  },
  {
    id: 'intune', name: 'Microsoft Intune', host: 'intune.microsoft.com',
    note: 'devices and endpoint policy', mark: 'MDM', tone: 't-live', session: true,
  },
];

/** Look a role up in every bucket — permissions are viewable in all states. */
function findRole(key) {
  return state.eligible.find((r) => keyOf(r) === key)
    || state.active.find((r) => keyOf(r) === key)
    || state.assigned.find((r) => keyOf(r) === key)
    || null;
}

/**
 * Roles a schedule can target: everything that goes through PIM, whether it is
 * switched on right now or not. An active role still expires, so scheduling a
 * future activation for it is perfectly normal. Permanently assigned roles are
 * excluded — they never pass through an activation request.
 */
function schedulableRoles() {
  return [...state.eligible, ...state.active]
    .sort((a, b) => a.name.localeCompare(b.name, 'en'));
}

/**
 * Every role this account holds, tagged with the bucket it came from. 'active'
 * and 'assigned' are in force right now; 'eligible' is one activation away.
 */
function taggedRoles() {
  return [
    ...state.active.map((role) => ({ role, kind: 'active' })),
    ...state.assigned.map((role) => ({ role, kind: 'assigned' })),
    ...state.eligible.map((role) => ({ role, kind: 'eligible' })),
  ];
}

const KIND_LABEL = { active: 'active', assigned: 'assigned', eligible: 'eligible' };

/* ================================= theme ================================= */

function resolveTheme(pref) {
  if (pref === 'light' || pref === 'dark') return pref;
  return window.matchMedia('(prefers-color-scheme: light)').matches ? 'light' : 'dark';
}

function applyTheme() {
  const resolved = resolveTheme(state.prefs.theme);
  document.documentElement.dataset.theme = resolved;
  $('btnTheme').title = resolved === 'dark' ? 'Switch to light theme' : 'Switch to dark theme';
  return resolved;
}

async function toggleTheme() {
  const next = resolveTheme(state.prefs.theme) === 'dark' ? 'light' : 'dark';
  state.prefs = { ...state.prefs, theme: next };
  applyTheme();
  try {
    const { prefs } = await ask('prefs/set', { payload: { theme: next } });
    state.prefs = prefs;
  } catch {
    /* the theme is already applied; persisting it is best effort */
  }
}

/* ================================= tick ================================== */

function tick() {
  const now = Date.now();
  for (const n of document.querySelectorAll('[data-tick="clock"]')) {
    const left = Number(n.dataset.end) - now;
    n.textContent = clock(left);
    n.classList.toggle('is-urgent', left < 15 * 60 * 1000);
  }
  for (const n of document.querySelectorAll('[data-tick="drain"]')) {
    const start = Number(n.dataset.start), end = Number(n.dataset.end);
    const span = Math.max(1, end - start);
    n.style.width = `${Math.max(0, Math.min(1, (end - now) / span)) * 100}%`;
    n.classList.toggle('is-urgent', end - now < 15 * 60 * 1000);
  }
  for (const n of document.querySelectorAll('[data-tick="expiry"]')) {
    const left = Number(n.dataset.end) - now;
    n.textContent = humanize(left);
    n.classList.toggle('is-urgent', left < 15 * 60 * 1000);
  }
}

/* =============================== load data =============================== */

async function loadSession() {
  const { session, prefs } = await ask('session/get');
  state.prefs = prefs || state.prefs;
  state.pim = session || null;
  applyTheme();
}

async function loadRoles() {
  const g = state.graph, id = state.pim.principalId;
  const [eligibleRaw, activeRaw] = await Promise.all([
    g.eligibleInstances(id),
    g.activeInstances(id),
  ]);
  const all = activeRaw.map(toRole);
  state.active = all.filter((r) => r.assignmentType === 'Activated')
    .sort((a, b) => (a.end || Infinity) - (b.end || Infinity));
  state.assigned = all.filter((r) => r.assignmentType !== 'Activated')
    .sort((a, b) => a.name.localeCompare(b.name, 'en'));

  // A role that is currently on also stays in the eligible list — drop it there
  // so it is not shown twice (an "Activate" card plus a countdown card) and is
  // not counted twice in the stats.
  const activeKeys = new Set(state.active.map(keyOf));
  state.eligible = eligibleRaw.map(toRole)
    .filter((r) => !activeKeys.has(keyOf(r)))
    .sort((a, b) => a.name.localeCompare(b.name, 'en'));
}

async function hydratePolicies() {
  // Active roles are included: their policy still governs a scheduled re-activation.
  const queue = schedulableRoles().filter((r) => !state.policies.has(keyOf(r)));
  const worker = async () => {
    for (;;) {
      const role = queue.shift();
      if (!role) return;
      try {
        const rules = await state.graph.policyRules(role.roleDefinitionId, role.scopeId);
        state.policies.set(keyOf(role), readPolicy(rules));
      } catch {
        state.policies.set(keyOf(role), null);
      }
    }
  };
  await Promise.all([worker(), worker(), worker(), worker()]);
}

function policyChips(role) {
  const p = state.policies.get(keyOf(role));
  if (p === undefined) return '';
  if (p === null) return `<span class="chip">policy unavailable</span>`;
  const chips = [];
  if (p.maxDuration) chips.push(`<span class="chip is-calm">max ${humanize(p.maxDuration)}</span>`);
  if (p.requireMfa) chips.push(`<span class="chip is-warn">MFA</span>`);
  if (p.requireJustification) chips.push(`<span class="chip">reason required</span>`);
  if (p.requireTicket) chips.push(`<span class="chip is-warn">ticket required</span>`);
  if (p.requireApproval) chips.push(`<span class="chip is-gate">approval required</span>`);
  return chips.join('');
}

/* ================================ render ================================= */

/**
 * Who the session belongs to and how long it lives. The countdown is the one
 * thing worth knowing about the captured token: once it lapses every Graph call
 * fails, and reopening a portal is the fix.
 */
function renderIdentity() {
  const sub = $('whoSub');
  if (!state.pim) {
    sub.innerHTML = `<span class="who-name">No session captured</span>`;
    return;
  }
  const who = state.pim.upn || state.pim.name || state.pim.tenantId;
  const end = state.pim.expiresAt || 0;
  sub.innerHTML = `<span class="who-name" title="${esc(who)}">${esc(who)}</span>` + (end
    ? `<span class="who-exp" data-tick="expiry" data-end="${end}"
         title="The captured session expires then">${esc(humanize(end - Date.now()))}</span>`
    : '');
  tick();
}

function renderStats() {
  const total = state.eligible.length + state.active.length + state.assigned.length;
  $('statTotal').textContent = state.pim ? total : '—';
  $('statEligible').textContent = state.pim ? state.eligible.length : '—';
  $('statActive').textContent = state.pim ? state.active.length : '—';
}

function renderMenuBadges() {
  $('badgeJobs').textContent = state.jobs.length || '';
}

function emptyBlock(title, text, cta = '') {
  return `<div class="empty">
    <div class="empty-mark"><svg viewBox="0 0 24 24" width="24" height="24">
      <path d="M12 2.6 4.8 5.7v5.8c0 4.5 3 8.2 7.2 9.5 4.2-1.3 7.2-5 7.2-9.5V5.7L12 2.6Z"
        fill="none" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/>
      <path d="M12 8.6v4.1M12 15.6v.1" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"/>
    </svg></div>
    <div class="empty-title">${esc(title)}</div>
    <div class="empty-text">${esc(text)}</div>${cta}</div>`;
}

function sectionHead(kind, title, count) {
  return `<div class="section">
    <span class="section-dot ${kind}"></span>
    <span class="section-title">${esc(title)}</span>
    <span class="section-count">${count}</span>
    <span class="section-line"></span>
  </div>`;
}

/** Opens the schedule panel with this role pre-ticked. */
function scheduleButton(key) {
  return `<button class="btn btn-quiet btn-sm btn-icon" data-sched="${esc(key)}" type="button"
    title="Schedule a future activation">
    <svg viewBox="0 0 16 16" width="13" height="13" aria-hidden="true">
      <path d="M3 4h10v9H3zM3 6.5h10M6 2v2M10 2v2" fill="none" stroke="currentColor"
        stroke-width="1.4" stroke-linecap="round"/>
    </svg>
  </button>`;
}

/** The button that opens the permission drawer — present on every card kind. */
function permButton(key) {
  return `<button class="btn btn-quiet btn-sm" data-perms="${esc(key)}" type="button"
    title="View the permissions this role grants">
    <svg viewBox="0 0 16 16" width="13" height="13" aria-hidden="true">
      <path d="M8 2.4 3.6 4.3v3.5c0 2.7 1.8 5 4.4 5.8 2.6-.8 4.4-3.1 4.4-5.8V4.3L8 2.4Z"
        fill="none" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round"/>
    </svg>
    <span>Permissions</span>
  </button>`;
}

function drawerFor(key) {
  return `<div class="drawer" data-drawer-for="${esc(key)}">
    <div class="drawer-inner"><div class="drawer-pad" data-detail="${esc(key)}"></div></div>
  </div>`;
}

function renderRoles() {
  const view = $('roleList');

  if (!state.pim) {
    view.innerHTML = emptyBlock(
      'No session captured',
      'Open the PIM blade in the Azure portal or the Entra admin center and let it finish '
      + 'loading your roles. The extension picks the session up automatically.',
      `<button class="btn btn-primary" id="ctaPortal" type="button">Open PIM portal</button>`
    );
    return;
  }

  const activeHtml = state.active.map((r, i) => {
    const key = keyOf(r);
    const remaining = r.end ? r.end - Date.now() : null;
    const urgent = remaining !== null && remaining < 15 * 60 * 1000;
    return `<article class="rcard k-active" style="--i:${i}">
      <div class="rcard-main">
        <div class="rcard-body">
          <div class="rcard-head">
            <span class="rcard-name">${esc(r.name)}</span>
            <span class="badge badge-active">Active</span>
          </div>
          <div class="rcard-meta"><span>${esc(r.scopeLabel)}</span></div>
          <div class="rcard-meta"><span>Ends ${r.end ? esc(formatDateTime(r.end)) : '—'}</span></div>
        </div>
      </div>
      <div class="rcard-actions">
        ${r.end
          ? `<div class="countdown${urgent ? ' is-urgent' : ''}" data-tick="clock" data-end="${r.end}">${clock(remaining)}</div>`
          : `<span class="chip is-calm">permanent</span>`}
        <div class="rcard-spacer"></div>
        ${scheduleButton(key)}
        ${permButton(key)}
        <button class="btn btn-quiet btn-sm" data-off="${esc(key)}" type="button">Deactivate</button>
      </div>
      ${r.end ? `<div class="drain"><div class="drain-fill${urgent ? ' is-urgent' : ''}"
        data-tick="drain" data-start="${r.start || Date.now()}" data-end="${r.end}"></div></div>` : ''}
      ${drawerFor(key)}
    </article>`;
  }).join('');

  const eligibleHtml = state.eligible.map((r, i) => {
    const key = keyOf(r);
    const picked = state.selected.has(key);
    return `<article class="rcard k-eligible${picked ? ' is-picked' : ''}" style="--i:${i}">
      <div class="rcard-main">
        <div class="rcard-check${picked ? ' is-on' : ''}" data-pick="${esc(key)}" role="checkbox"
             tabindex="0" aria-checked="${picked}" aria-label="Select ${esc(r.name)}">
          <svg viewBox="0 0 18 18"><path d="M4.5 9.2 7.6 12.2 13.5 6" fill="none" stroke="currentColor"
            stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/></svg>
        </div>
        <div class="rcard-body">
          <div class="rcard-head">
            <span class="rcard-name">${esc(r.name)}</span>
            <span class="badge badge-eligible">Eligible</span>
          </div>
          <div class="rcard-meta"><span>Member: <b>${esc(r.memberType || 'Direct')}</b></span>
            <span>${esc(r.scopeLabel)}</span></div>
          <div class="chips" data-chips="${esc(key)}">${policyChips(r)}</div>
        </div>
      </div>
      <div class="rcard-actions">
        ${scheduleButton(key)}
        ${permButton(key)}
        <div class="rcard-spacer"></div>
        <button class="btn-activate" data-activate="${esc(key)}" type="button">
          <svg viewBox="0 0 16 16" width="13" height="13"><path d="M8.5 1 3 9h4l-.5 6L13 7H9z"
            fill="currentColor"/></svg>
          <span class="btn-label">Activate</span><span class="spinner"></span>
        </button>
      </div>
      ${drawerFor(key)}
    </article>`;
  }).join('');

  const assignedHtml = state.assigned.map((r, i) => {
    const key = keyOf(r);
    return `<article class="rcard k-assigned" style="--i:${i}">
      <div class="rcard-main">
        <div class="rcard-body">
          <div class="rcard-head">
            <span class="rcard-name">${esc(r.name)}</span>
            <span class="badge badge-assigned">Assigned</span>
          </div>
          <div class="rcard-meta"><span>Member: <b>${esc(r.memberType || 'Group')}</b></span>
            <span>${esc(r.scopeLabel)}</span></div>
        </div>
      </div>
      <div class="rcard-actions">
        <span class="chip">always on · outside PIM</span>
        <div class="rcard-spacer"></div>
        ${permButton(key)}
      </div>
      ${drawerFor(key)}
    </article>`;
  }).join('');

  if (!activeHtml && !eligibleHtml && !assignedHtml) {
    view.innerHTML = emptyBlock('No roles found',
      'This account has no PIM role assignments, or the captured token cannot read them.');
    return;
  }

  view.innerHTML =
    (activeHtml ? sectionHead('k-active', 'Active now', state.active.length) + activeHtml : '') +
    (eligibleHtml ? sectionHead('k-eligible', 'Eligible', state.eligible.length) + eligibleHtml : '') +
    (assignedHtml ? sectionHead('k-assigned', 'Permanently assigned', state.assigned.length) + assignedHtml : '');
  tick();
}

/* --------------------------- details / activate -------------------------- */

const PERM_PREVIEW = 12;

/**
 * Read one role definition, cached. The expanded roleDefinition Graph returns
 * with an assignment does not carry rolePermissions, so the granular actions
 * always need this second call — and both the per-role drawer and the
 * "My permissions" panel want the same answer.
 */
async function fetchDefinition(roleDefinitionId) {
  if (state.definitions.has(roleDefinitionId)) {
    const cached = state.definitions.get(roleDefinitionId);
    if (cached) return cached;
    throw new Error('Graph would not return this role definition.');
  }
  try {
    const def = await state.graph.roleDefinition(roleDefinitionId);
    const actions = [];
    for (const b of def.rolePermissions || []) actions.push(...(b.allowedResourceActions || []));
    const record = { permissions: actions, description: def.description || '' };
    state.definitions.set(roleDefinitionId, record);
    return record;
  } catch (e) {
    state.definitions.set(roleDefinitionId, null);
    throw e;
  }
}

/** Fill in a role's granular actions in place, if they are not there yet. */
async function applyDefinition(role) {
  if (role.permissions.length) return;
  const record = await fetchDefinition(role.roleDefinitionId);
  role.permissions = record.permissions;
  if (!role.description) role.description = record.description;
}

async function fillPermissions(key) {
  const holder = document.querySelector(`[data-detail="${CSS.escape(key)}"]`);
  if (!holder) return;
  const role = findRole(key);
  if (!role) return;

  if (!role.permissions.length) {
    holder.innerHTML = `<div class="detail-text">Loading permissions…</div>`;
    try {
      await applyDefinition(role);
    } catch (e) {
      holder.innerHTML = `<div class="detail-text">Could not read permissions: ${esc(e.message)}</div>`;
      return;
    }
  }

  const expanded = state.permsExpanded.has(key);
  const groups = new Map();
  for (const perm of role.permissions) {
    const i = perm.indexOf('/');
    const provider = i > 0 ? perm.slice(0, i) : 'other';
    if (!groups.has(provider)) groups.set(provider, []);
    groups.get(provider).push(i > 0 ? perm.slice(i + 1) : perm);
  }

  let hidden = 0;
  const permHtml = [...groups.entries()]
    .sort((a, b) => b[1].length - a[1].length || a[0].localeCompare(b[0]))
    .map(([provider, actions]) => {
      const sorted = actions.sort();
      const shown = expanded ? sorted : sorted.slice(0, PERM_PREVIEW);
      hidden += sorted.length - shown.length;
      return `<div class="perm-group">
        <div class="perm-provider">${esc(provider)}<span class="perm-tally">${sorted.length}</span></div>
        <ul class="perm-list">${shown.map((a) => `<li title="${esc(a)}">${esc(a)}</li>`).join('')}</ul>
      </div>`;
    }).join('');

  const more = hidden > 0
    ? `<button class="perm-more" data-permmore="${esc(key)}" type="button">Show all ${role.permissions.length} permissions (+${hidden} hidden)</button>`
    : expanded && role.permissions.length > PERM_PREVIEW
      ? `<button class="perm-more" data-permless="${esc(key)}" type="button">Show less</button>`
      : '';

  holder.innerHTML =
    (role.description
      ? `<div class="detail-label">Description</div><div class="detail-text">${esc(role.description)}</div>`
      : '') +
    `<div class="detail-label">Permissions (${role.permissions.length})</div>` +
    (permHtml || '<div class="detail-text">This role declares no granular permissions.</div>') +
    more;
}

function fillActivateForm(key) {
  const holder = document.querySelector(`[data-detail="${CSS.escape(key)}"]`);
  if (!holder) return;
  const role = findRole(key);
  const p = state.policies.get(key);
  const cap = p && p.maxDuration ? humanize(p.maxDuration) : null;
  holder.innerHTML = `
    <div class="mini-form">
      <label class="field field-sm"><span>Duration</span>
        <input type="text" data-dur="${esc(key)}" value="8h" spellcheck="false" /></label>
      <label class="field field-grow"><span>Reason</span>
        <input type="text" data-reason="${esc(key)}" placeholder="e.g. INC-4471" /></label>
      <button class="btn-activate" data-confirm="${esc(key)}" type="button">
        <span class="btn-label">Confirm</span><span class="spinner"></span></button>
    </div>
    ${cap ? `<div class="detail-note">Policy allows at most ${esc(cap)}.</div>` : ''}
    <div class="chips">${policyChips(role)}</div>`;
}

function toggleDrawer(key, mode) {
  const drawer = document.querySelector(`[data-drawer-for="${CSS.escape(key)}"]`);
  if (!drawer) return;
  const current = state.openDrawers.get(key);

  if (current === mode) {
    state.openDrawers.delete(key);
    drawer.classList.remove('is-open');
    return;
  }
  state.openDrawers.set(key, mode);
  drawer.classList.add('is-open');
  if (mode === 'perms') fillPermissions(key);
  else fillActivateForm(key);
}

/* ========================== activate / deactivate ======================== */

function clampDuration(raw, roles) {
  const iso = normalizeDuration(raw);
  if (!iso) return { error: 'Invalid duration. Try 8h, 90m or 2h30m.' };
  const caps = roles.map((r) => state.policies.get(keyOf(r)))
    .filter((p) => p && p.maxDuration).map((p) => p.maxDuration);
  if (!caps.length) return { iso };
  const cap = Math.min(...caps);
  if (parseIsoDuration(iso) > cap) return { iso: toIsoDuration(cap), lowered: cap };
  return { iso };
}

function needsReason(roles) {
  return roles.some((r) => requiresReason(r));
}

function requiresReason(role) {
  const p = state.policies.get(keyOf(role));
  return Boolean(p && p.requireJustification);
}

/** The reason a role will actually be sent with: its own override, else the shared one. */
function effectiveReason(key, shared) {
  return (state.reasons.get(key) || '').trim() || shared;
}

/**
 * @param perRole when true, each role carries the justification typed for it in
 *        the per-role sheet, falling back to `reason` where none was given.
 */
async function activate(roles, duration, reason, buttonEl, perRole = false) {
  if (buttonEl) { buttonEl.classList.add('is-busy'); buttonEl.disabled = true; }
  try {
    const { results } = await ask('role/activate', {
      payload: {
        roles: roles.map((r) => {
          const entry = { roleDefinitionId: r.roleDefinitionId, scopeId: r.scopeId, name: r.name };
          const own = perRole ? (state.reasons.get(keyOf(r)) || '').trim() : '';
          if (own) entry.justification = own;
          return entry;
        }),
        duration, justification: reason || 'Activated from Entra PIM Roles',
      },
    });
    const ok = results.filter((r) => r.ok);
    for (const f of results.filter((r) => !r.ok)) {
      toast(`Could not activate ${f.role.name}`, f.hint || f.message, 'bad');
    }
    if (ok.length) {
      const pending = ok.filter((r) => r.status === 'PendingApproval');
      toast(`Activated ${ok.length} role${ok.length > 1 ? 's' : ''}`,
        pending.length
          ? `${pending.length} awaiting approval.`
          : `Duration ${humanize(parseIsoDuration(duration))}.`,
        pending.length ? 'warn' : 'ok');
    }
    state.selected.clear();
    state.reasons.clear();
    state.reasonsOpen = false;
    state.openDrawers = new Map();
    await refresh();
  } catch (e) {
    toast('Activation failed', e.message, 'bad');
  } finally {
    if (buttonEl) { buttonEl.classList.remove('is-busy'); buttonEl.disabled = false; }
  }
}

async function confirmSingle(key) {
  const role = state.eligible.find((r) => keyOf(r) === key);
  if (!role) return;
  const durInput = document.querySelector(`[data-dur="${CSS.escape(key)}"]`);
  const reasonInput = document.querySelector(`[data-reason="${CSS.escape(key)}"]`);
  const reason = (reasonInput?.value || '').trim();
  if (needsReason([role]) && !reason) {
    toast('Reason required', 'The policy on this role requires a justification.', 'warn');
    reasonInput?.focus();
    return;
  }
  const { iso, error, lowered } = clampDuration(durInput?.value || '8h', [role]);
  if (error) { toast('Invalid duration', error, 'bad'); return; }
  if (lowered) toast('Duration shortened', `Policy caps this at ${humanize(lowered)}.`, 'warn');
  const btn = document.querySelector(`[data-confirm="${CSS.escape(key)}"]`);
  await activate([role], iso, reason, btn);
}

async function bulkActivate() {
  const roles = selectedRoles();
  if (!roles.length) return;
  const shared = $('inpReason').value.trim();

  // A role whose policy demands a justification needs one of its own or a shared one.
  const missing = roles.filter((r) => requiresReason(r) && !effectiveReason(keyOf(r), shared));
  if (missing.length) {
    openReasons();
    for (const r of missing) {
      document.querySelector(`[data-rreason="${CSS.escape(keyOf(r))}"]`)
        ?.closest('.reason-row')?.classList.add('is-missing');
    }
    toast('Reason required',
      missing.length === roles.length
        ? 'Every selected role requires a justification.'
        : `Still missing for: ${missing.map((r) => r.name).join(', ')}.`,
      'warn');
    document.querySelector(`[data-rreason="${CSS.escape(keyOf(missing[0]))}"]`)?.focus();
    return;
  }

  const { iso, error, lowered } = clampDuration($('inpDuration').value, roles);
  if (error) { toast('Invalid duration', error, 'bad'); return; }
  if (lowered) toast('Duration shortened', `Policy caps this at ${humanize(lowered)}.`, 'warn');
  await activate(roles, iso, shared, $('btnActivate'), true);
}

async function deactivate(key) {
  const role = state.active.find((r) => keyOf(r) === key);
  if (!role) return;
  try {
    const { results } = await ask('role/deactivate', {
      payload: {
        roles: [{ roleDefinitionId: role.roleDefinitionId, scopeId: role.scopeId, name: role.name }],
        justification: 'Deactivated from Entra PIM Roles',
      },
    });
    const f = results.find((r) => !r.ok);
    if (f) toast(`Could not deactivate ${role.name}`, f.hint || f.message, 'bad');
    else toast(`Deactivated ${role.name}`, '', 'ok');
    await refresh();
  } catch (e) {
    toast('Deactivation failed', e.message, 'bad');
  }
}

const selectedRoles = () => state.eligible.filter((r) => state.selected.has(keyOf(r)));

/* --------------------------- per-role reasons ---------------------------- */

function renderReasonSheet() {
  const shared = $('inpReason').value.trim();
  $('reasonList').innerHTML = selectedRoles().map((r) => {
    const key = keyOf(r);
    const required = requiresReason(r);
    const value = state.reasons.get(key) || '';
    return `<div class="reason-row">
      <span class="reason-name" title="${esc(r.name)}">${esc(r.name)}${required ? '<span class="reason-req">*</span>' : ''}</span>
      <input type="text" data-rreason="${esc(key)}" value="${esc(value)}" autocomplete="off"
        placeholder="${required && !shared ? 'required' : shared ? esc(shared) : 'same as shared reason'}" />
    </div>`;
  }).join('');
}

/** Keeps the toggle label honest about how many overrides are in play. */
function syncReasonButton() {
  const set = selectedRoles().filter((r) => (state.reasons.get(keyOf(r)) || '').trim()).length;
  $('reasonsLabel').textContent = set ? `Reason per role · ${set}` : 'Reason per role';
  $('btnReasons').classList.toggle('is-on', state.reasonsOpen || set > 0);
  $('btnReasons').setAttribute('aria-expanded', String(state.reasonsOpen));
}

function openReasons() {
  state.reasonsOpen = true;
  renderReasonSheet();
  $('reasonSheet').hidden = false;
  syncReasonButton();
}

function toggleReasons() {
  if (state.reasonsOpen) {
    state.reasonsOpen = false;
    $('reasonSheet').hidden = true;
    syncReasonButton();
  } else {
    openReasons();
  }
}

function syncActionbar() {
  const bar = $('actionbar');
  const n = state.selected.size;
  bar.hidden = n === 0;
  $('stage').classList.toggle('has-actionbar', n > 0);
  $('selCount').textContent = String(n);

  // drop overrides for roles that are no longer selected
  for (const key of [...state.reasons.keys()]) {
    if (!state.selected.has(key)) state.reasons.delete(key);
  }
  if (n === 0) {
    state.reasonsOpen = false;
    $('reasonSheet').hidden = true;
  } else if (state.reasonsOpen) {
    renderReasonSheet();
  }
  syncReasonButton();
}

/* ============================ admin portals ============================== */

function renderPortals() {
  $('portalBody').innerHTML = PORTALS.map((p, i) => `
    <button class="portal-row ${p.tone}" data-portal="${esc(p.id)}" type="button" style="--i:${i}">
      <span class="portal-mark">${esc(p.mark)}</span>
      <span class="portal-text">
        <span class="portal-name">${esc(p.name)}${p.session ? '<span class="portal-tag">session</span>' : ''}</span>
        <span class="portal-host">${esc(p.host)} · ${esc(p.note)}</span>
      </span>
      <svg class="portal-go" viewBox="0 0 16 16" width="13" height="13" aria-hidden="true">
        <path d="M6 3.5 10.5 8 6 12.5" fill="none" stroke="currentColor" stroke-width="1.7"
          stroke-linecap="round" stroke-linejoin="round"/>
      </svg>
    </button>`).join('');
}

/* =========================== activation history ========================== */

const HISTORY_TOP = 80;

async function loadHistory(force = false) {
  if (!state.pim) return;
  if (state.historyState === 'loading') return;
  if (state.historyState === 'ready' && !force) return;

  state.historyState = 'loading';
  state.historyError = '';
  renderHistory();
  try {
    const rows = await state.graph.requestHistory(state.pim.principalId, HISTORY_TOP);
    state.history = rows
      .map(toHistoryEntry)
      .sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));
    state.historyState = 'ready';
  } catch (e) {
    state.historyState = 'error';
    state.historyError = e.hint || e.message || String(e);
  }
  renderHistory();
}

function historyRows() {
  if (state.historyFilter === 'all') return state.history;
  return state.history.filter((h) => h.kind === state.historyFilter);
}

function historyRow(h) {
  const statusTone = h.failed
    ? 'is-bad'
    : /^Pending/.test(h.status) ? 'is-warn' : 'is-ok';
  const bits = [];
  if (h.createdAt) bits.push(formatTime(h.createdAt));
  if (h.kind === 'on' && h.duration) bits.push(humanize(h.duration));
  if (h.startAt && h.createdAt && h.startAt - h.createdAt > 120000) {
    bits.push(`starts ${formatDateTime(h.startAt)}`);
  }
  bits.push(h.scopeLabel);

  return `<div class="hrow k-${esc(h.kind)}${h.failed ? ' is-failed' : ''}">
    <span class="hmark">
      <svg viewBox="0 0 16 16" width="12" height="12" aria-hidden="true">${h.kind === 'on'
        ? `<path d="M8.5 1.5 3.5 9h3.4l-.4 5.5L12.5 7H8.9z" fill="currentColor"/>`
        : `<path d="M4 8h8" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/>`}
      </svg>
    </span>
    <span class="hbody">
      <span class="hhead">
        <span class="hname" title="${esc(h.name)}">${esc(h.name)}</span>
        <span class="hstatus ${statusTone}">${esc(h.statusLabel)}</span>
      </span>
      <span class="hmeta">${bits.map((b) => `<span>${esc(b)}</span>`).join('')}</span>
      <span class="haction">${esc(h.actionLabel)}</span>
      ${h.justification ? `<span class="hreason">“${esc(h.justification)}”</span>` : ''}
      ${h.ticketNumber ? `<span class="hreason">ticket ${esc(h.ticketNumber)}</span>` : ''}
    </span>
  </div>`;
}

function renderHistory() {
  const body = $('historyBody');
  for (const btn of document.querySelectorAll('[data-hfilter]')) {
    btn.classList.toggle('is-on', btn.dataset.hfilter === state.historyFilter);
  }

  if (!state.pim) {
    body.innerHTML = emptyBlock('No session captured',
      'Open the PIM portal so the extension can read your request history.');
    return;
  }
  if (state.historyState === 'loading') {
    body.innerHTML = '<div class="skeleton skeleton-sm"></div>'.repeat(5);
    return;
  }
  if (state.historyState === 'error') {
    body.innerHTML = emptyBlock('Could not read the history', state.historyError);
    return;
  }

  const rows = historyRows();
  if (!rows.length) {
    body.innerHTML = emptyBlock(
      state.history.length ? 'Nothing in this view' : 'No requests yet',
      state.history.length
        ? 'Every request in the window is of the other kind. Switch the filter above.'
        : `PIM has no activation request on record for this account. The last ${HISTORY_TOP} `
          + 'requests appear here as soon as you activate or deactivate a role.');
    return;
  }

  // One heading per calendar day, newest first.
  let day = '';
  const chunks = [];
  for (const h of rows) {
    const key = h.createdAt ? dayKey(h.createdAt) : 'unknown';
    if (key !== day) {
      day = key;
      chunks.push(`<div class="hday">${esc(h.createdAt ? formatDay(h.createdAt) : 'Unknown date')}</div>`);
    }
    chunks.push(historyRow(h));
  }
  const capped = state.history.length >= HISTORY_TOP
    ? `<div class="detail-note">Showing the newest ${HISTORY_TOP} requests PIM keeps for this account.</div>`
    : '';
  body.innerHTML = chunks.join('') + capped;
}

/* ============================= my permissions ============================ */

const PERM_GROUP_CAP = 300;

/**
 * Read the definition of every role the account holds, then fold them into one
 * action -> grantors index. This is the only view that needs all definitions at
 * once, so it is loaded on demand rather than at boot.
 */
async function loadPermissionIndex(force = false) {
  if (!state.pim) return;
  if (state.permState === 'loading') return;
  if (state.permState === 'ready' && !force) return;

  const tagged = taggedRoles();
  state.permState = 'loading';
  state.permError = '';
  state.permProgress = [0, tagged.length];
  renderPerms();

  let failures = 0;
  const queue = [...tagged];
  const worker = async () => {
    for (;;) {
      const item = queue.shift();
      if (!item) return;
      try {
        await applyDefinition(item.role);
      } catch {
        failures++;
      }
      state.permProgress = [state.permProgress[0] + 1, tagged.length];
      renderPermProgress();
    }
  };

  try {
    await Promise.all([worker(), worker(), worker(), worker()]);
    state.permIndex = indexPermissions(
      tagged.map(({ role, kind }) => ({ name: role.name, kind, permissions: role.permissions }))
    );
    state.permState = 'ready';
    if (failures) {
      toast('Some roles could not be read',
        `${failures} role definition${failures > 1 ? 's' : ''} were refused by Graph.`, 'warn');
    }
  } catch (e) {
    state.permState = 'error';
    state.permError = e.hint || e.message || String(e);
  }
  renderPerms();
}

function renderPermProgress() {
  const bar = document.querySelector('[data-permbar]');
  if (!bar) return;
  const [done, total] = state.permProgress;
  bar.style.width = `${total ? (done / total) * 100 : 0}%`;
  const label = document.querySelector('[data-permcount]');
  if (label) label.textContent = `${done}/${total} roles read`;
}

/** The buckets the current scope counts as "mine". */
function permScopeKinds() {
  return state.permScope === 'now'
    ? new Set(['active', 'assigned'])
    : new Set(['active', 'assigned', 'eligible']);
}

/**
 * Group the indexed actions by provider for the current scope and search box.
 * `latent` marks an action no role in force right now grants — it only arrives
 * once an eligible role is activated.
 */
function permGroups() {
  const kinds = permScopeKinds();
  const query = state.permQuery.trim().toLowerCase();
  const groups = new Map();
  const roles = new Set();
  let total = 0;

  for (const [action, entry] of state.permIndex) {
    const grantors = entry.grantors.filter((g) => kinds.has(g.kind));
    if (!grantors.length) continue;
    if (query && !action.toLowerCase().includes(query)) continue;

    total++;
    for (const g of grantors) roles.add(g.name);
    const cut = action.indexOf('/');
    const provider = cut > 0 ? action.slice(0, cut) : 'other';
    const tail = cut > 0 ? action.slice(cut + 1) : action;
    if (!groups.has(provider)) groups.set(provider, []);
    groups.get(provider).push({
      action, tail, grantors,
      latent: grantors.every((g) => g.kind === 'eligible'),
    });
  }

  for (const list of groups.values()) list.sort((a, b) => a.tail.localeCompare(b.tail, 'en'));
  return {
    total,
    roles: roles.size,
    groups: [...groups.entries()].sort((a, b) => b[1].length - a[1].length || a[0].localeCompare(b[0])),
  };
}

function renderPerms() {
  const body = $('permsBody');
  for (const btn of document.querySelectorAll('[data-pscope]')) {
    btn.classList.toggle('is-on', btn.dataset.pscope === state.permScope);
  }

  if (!state.pim) {
    body.innerHTML = emptyBlock('No session captured',
      'Open the PIM portal so the extension can read the roles you hold.');
    return;
  }
  if (state.permState === 'loading') {
    const [done, total] = state.permProgress;
    body.innerHTML = `<div class="panel">
      <div class="panel-title">Reading role definitions</div>
      <div class="progress"><div class="progress-fill" data-permbar
        style="width:${total ? (done / total) * 100 : 0}%"></div></div>
      <div class="detail-note" data-permcount>${done}/${total} roles read</div>
    </div>` + '<div class="skeleton skeleton-sm"></div>'.repeat(3);
    return;
  }
  if (state.permState === 'error') {
    body.innerHTML = emptyBlock('Could not build the list', state.permError);
    return;
  }
  if (!state.permIndex) {
    body.innerHTML = emptyBlock('Nothing loaded yet', 'Reopen this panel to read your roles.');
    return;
  }

  // The counters describe what is on screen, so a search narrows all three.
  const { total, roles, groups } = permGroups();
  const query = state.permQuery.trim();

  const head = `<div class="stats stats-flat">
    <div class="stat is-total"><div class="stat-num">${total}</div><div class="stat-label">Actions</div></div>
    <div class="stat is-eligible"><div class="stat-num">${groups.length}</div><div class="stat-label">Providers</div></div>
    <div class="stat is-live"><div class="stat-num">${roles}</div><div class="stat-label">Roles</div></div>
  </div>`;

  if (!groups.length) {
    body.innerHTML = head + emptyBlock(
      query ? 'No action matches' : 'No permissions found',
      query
        ? `Nothing in this scope contains “${query}”.`
        : state.permScope === 'now'
          ? 'No role is in force right now. Switch to “Incl. eligible” to see what you could activate.'
          : 'None of your roles declares a granular action.');
    return;
  }

  // A search is a hunt for one action: open every group that still has a hit.
  const searching = Boolean(query);
  const list = groups.map(([provider, actions]) => {
    const open = searching || state.permGroupsOpen.has(provider);
    const shown = actions.slice(0, PERM_GROUP_CAP);
    const rows = open
      ? `<ul class="perm-list">${shown.map((a) => `
          <li class="${a.latent ? 'is-latent' : ''}" title="${esc(a.action)}&#10;granted by ${
            esc(a.grantors.map((g) => `${g.name} (${KIND_LABEL[g.kind]})`).join(', '))}">
            <span class="perm-action">${esc(a.tail)}</span>
            <span class="perm-count">${a.grantors.length}</span>
          </li>`).join('')}</ul>${actions.length > shown.length
            ? `<div class="detail-note">+${actions.length - shown.length} more; narrow the filter to see them.</div>`
            : ''}`
      : '';
    return `<div class="pgroup${open ? ' is-open' : ''}">
      <button class="pgroup-head" data-pgroup="${esc(provider)}" type="button" aria-expanded="${open}">
        <svg class="pgroup-caret" viewBox="0 0 16 16" width="11" height="11" aria-hidden="true">
          <path d="M6 3.5 10.5 8 6 12.5" fill="none" stroke="currentColor" stroke-width="1.8"
            stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
        <span class="pgroup-name">${esc(provider)}</span>
        <span class="perm-tally">${actions.length}</span>
      </button>
      ${rows}
    </div>`;
  }).join('');

  const foot = `<div class="note">
    <p>${state.permScope === 'now'
      ? 'Actions granted by the roles that are switched on right now, plus your permanent assignments.'
      : 'Includes every eligible role. Actions only an eligible role grants are '
        + '<span class="perm-swatch"></span> dimmed — activate the role first.'}</p>
  </div>`;

  body.innerHTML = head + list + foot;
}

/* =============================== schedules =============================== */

function renderSchedule() {
  const body = $('scheduleBody');
  if (!state.pim) {
    body.innerHTML = emptyBlock('No session captured', 'Open the PIM portal before creating a schedule.');
    return;
  }
  const prefill = state.schedulePrefill;
  const candidates = schedulableRoles();
  const activeKeys = new Set(state.active.map(keyOf));
  const options = candidates.map((r) => {
    const k = keyOf(r);
    const checked = prefill ? prefill === k : state.selected.has(k);
    // Currently-on roles are listed too — the schedule targets the next window.
    const live = activeKeys.has(k)
      ? `<span class="pick-hint">on until ${esc(r.end ? formatTime(r.end) : '—')}</span>`
      : '';
    return `<label><input type="checkbox" value="${esc(k)}" ${checked ? 'checked' : ''} />
      <span>${esc(r.name)}</span>${live}</label>`;
  }).join('');
  state.schedulePrefill = null;

  if (!candidates.length) {
    body.innerHTML = emptyBlock('Nothing to schedule',
      'No PIM role on this account can be activated. Permanently assigned roles are always on '
      + 'and never go through an activation request.');
    return;
  }

  const defaultWhen = new Date(Date.now() + 3600000 - new Date().getTimezoneOffset() * 60000)
    .toISOString().slice(0, 16);

  const jobRows = state.jobs.length ? state.jobs.map((j) => `
    <div class="job-row">
      <div class="job-when">${esc(formatTime(j.runAt))}</div>
      <div class="job-body">
        <div class="job-roles">${esc(j.roles.map((r) => r.name).join(', '))}</div>
        <div class="job-note">${esc(formatDateTime(j.runAt))} · ${esc(humanize(parseIsoDuration(j.duration) || 0))}
          ${j.repeatDaily ? ' · repeats daily' : ''}${j.lastStatus ? ` · last run: ${esc(j.lastStatus)}` : ''}</div>
      </div>
      <button class="btn btn-quiet btn-sm" data-jdel="${esc(j.id)}" type="button">Delete</button>
    </div>`).join('') : `<div class="detail-text">No schedules on this machine yet.</div>`;

  body.innerHTML = `
    <div class="panel">
      <div class="panel-title">New activation schedule</div>
      <div class="detail-label">Pick roles</div>
      <div class="picker" id="jobPicker">${options}</div>
      <div class="stack" style="margin-top:11px">
        <div class="row-2">
          <label class="field"><span>Starts at</span><input type="datetime-local" id="jobWhen" value="${defaultWhen}" /></label>
          <label class="field"><span>Duration</span><input type="text" id="jobDuration" value="8h" spellcheck="false" /></label>
        </div>
        <div class="row-2">
          <label class="field"><span>Runs on</span>
            <select id="jobMode"><option value="server">PIM (server side)</option><option value="local">This browser</option></select></label>
          <label class="field"><span>Reason</span><input type="text" id="jobReason" placeholder="e.g. morning shift" /></label>
        </div>
        <label class="picker picker-flat">
          <span style="display:flex;align-items:center;gap:8px;cursor:pointer">
            <input type="checkbox" id="jobDaily" /> Repeat every day
            <span class="picker-hint">(browser schedules only)</span></span>
        </label>
        <button class="btn btn-primary" id="btnJobCreate" type="button">
          <span class="btn-label">Create schedule</span><span class="spinner"></span></button>
      </div>
      <div class="note">
        <p><b>PIM (server side)</b> submits the request now with a future start time — it still fires
        when your machine is off. <b>This browser</b> waits until the moment arrives, so the browser must
        be running and the captured session must still be valid.</p>
      </div>
      <div class="note">
        <p>Roles marked <b>on until …</b> are active right now. Scheduling one is fine — pick a start
        time after the current window ends, otherwise PIM rejects the request as overlapping.</p>
      </div>
    </div>
    <div class="panel"><div class="panel-title">Schedules on this machine</div>${jobRows}</div>`;
}

async function createJob() {
  const picked = [...document.querySelectorAll('#jobPicker input:checked')].map((i) => i.value);
  if (!picked.length) { toast('No role selected', 'Tick at least one role.', 'warn'); return; }
  const runAt = fromDatetimeLocal($('jobWhen').value);
  if (!runAt) { toast('Invalid start time', 'Pick a date and time again.', 'bad'); return; }

  const roles = schedulableRoles().filter((r) => picked.includes(keyOf(r)));
  const { iso, error, lowered } = clampDuration($('jobDuration').value, roles);
  if (error) { toast('Invalid duration', error, 'bad'); return; }
  if (lowered) toast('Duration shortened', `Policy caps this at ${humanize(lowered)}.`, 'warn');

  const reason = $('jobReason').value.trim() || 'Scheduled activation';
  const mode = $('jobMode').value;
  const daily = $('jobDaily').checked;
  const btn = $('btnJobCreate');
  btn.classList.add('is-busy'); btn.disabled = true;

  const payloadRoles = roles.map((r) => ({ roleDefinitionId: r.roleDefinitionId, scopeId: r.scopeId, name: r.name }));

  try {
    if (mode === 'server') {
      if (runAt <= Date.now()) { toast('Start time already passed', 'Pick a moment in the future.', 'bad'); return; }
      const { results } = await ask('role/activate', {
        payload: { roles: payloadRoles, duration: iso, justification: reason, startAt: runAt },
      });
      for (const f of results.filter((r) => !r.ok)) toast(`Could not schedule ${f.role.name}`, f.hint || f.message, 'bad');
      const ok = results.filter((r) => r.ok).length;
      if (ok) toast(`PIM accepted ${ok} schedule${ok > 1 ? 's' : ''}`, `Starts at ${formatDateTime(runAt)}.`, 'ok');
    } else {
      await ask('jobs/create', {
        payload: { roles: payloadRoles, runAt, duration: iso, durationMs: parseIsoDuration(iso),
          justification: reason, repeatDaily: daily },
      });
      toast('Schedule created on this machine',
        `Runs at ${formatDateTime(runAt)}${daily ? ', every day' : ''}.`, 'ok');
      const { jobs } = await ask('jobs/list');
      state.jobs = jobs;
      renderMenuBadges();
      renderSchedule();
    }
  } catch (e) {
    toast('Could not create schedule', e.message, 'bad');
  } finally {
    btn.classList.remove('is-busy'); btn.disabled = false;
  }
}

/* =============================== overlays ================================ */

function openOverlay(id) {
  $('scrim').hidden = false;
  $(id).hidden = false;
  closeMenu();
}
function closeOverlays() {
  for (const id of ['panelSchedules', 'panelHistory', 'panelPerms', 'panelPortals']) {
    $(id).hidden = true;
  }
  $('scrim').hidden = true;
}

function openMenu() { $('menu').hidden = false; }
function closeMenu() { $('menu').hidden = true; }

/** Hand a portal id to the worker, which owns the URL, then get out of the way. */
async function openPortal(id) {
  try {
    await ask('portal/open', { portal: id });
    window.close();
  } catch (e) {
    toast('Could not open the portal', e.message, 'bad');
  }
}

/* ================================ refresh ================================ */

async function refresh() {
  const btn = $('btnRefresh');
  btn.classList.add('is-spinning');
  try {
    await loadSession();
    renderIdentity();
    renderMenuBadges();

    // the role list is about to change, so both derived views go stale
    state.historyState = 'idle';
    state.permState = 'idle';
    state.permIndex = null;

    if (!state.pim) {
      state.graph = null;
      state.eligible = []; state.active = []; state.assigned = [];
      renderStats();
      renderRoles();
      syncActionbar();
      return;
    }

    state.graph = new Graph(state.pim.token);
    await loadRoles();
    renderStats();
    renderRoles();
    syncActionbar();
    hydratePolicies().then(() => {
      for (const r of state.eligible) {
        const holder = document.querySelector(`[data-chips="${CSS.escape(keyOf(r))}"]`);
        if (holder) holder.innerHTML = policyChips(r);
      }
    });

    const { jobs } = await ask('jobs/list');
    state.jobs = jobs;
    renderMenuBadges();
  } catch (e) {
    toast('Could not load data', e.hint || e.message, 'bad');
  } finally {
    btn.classList.remove('is-spinning');
  }
}

/* ================================= events ================================ */

function wire() {
  $('btnRefresh').addEventListener('click', refresh);
  $('btnTheme').addEventListener('click', toggleTheme);

  // follow the OS while the user has not chosen a theme explicitly
  window.matchMedia('(prefers-color-scheme: light)').addEventListener('change', () => {
    if (state.prefs.theme !== 'light' && state.prefs.theme !== 'dark') applyTheme();
  });

  $('btnMenu').addEventListener('click', (e) => {
    e.stopPropagation();
    $('menu').hidden ? openMenu() : closeMenu();
  });

  document.addEventListener('click', () => closeMenu());
  document.addEventListener('keydown', (e) => {
    if (e.key !== 'Escape') return;
    if (!$('menu').hidden) { closeMenu(); return; }
    closeOverlays();
  });

  $('menu').addEventListener('click', (e) => {
    e.stopPropagation();
    const item = e.target.closest('[data-menu]');
    if (!item) return;
    const action = item.dataset.menu;
    closeMenu();
    if (action === 'schedules') { renderSchedule(); openOverlay('panelSchedules'); }
    else if (action === 'history') { renderHistory(); openOverlay('panelHistory'); loadHistory(); }
    else if (action === 'perms') { renderPerms(); openOverlay('panelPerms'); loadPermissionIndex(); }
    else if (action === 'portals') { renderPortals(); openOverlay('panelPortals'); }
  });

  for (const btn of document.querySelectorAll('[data-close]')) {
    btn.addEventListener('click', closeOverlays);
  }
  $('scrim').addEventListener('click', closeOverlays);

  $('btnActivate').addEventListener('click', bulkActivate);
  $('inpReason').addEventListener('keydown', (e) => { if (e.key === 'Enter') bulkActivate(); });
  // the shared reason doubles as the placeholder for every empty per-role field
  $('inpReason').addEventListener('input', () => { if (state.reasonsOpen) renderReasonSheet(); });

  $('btnReasons').addEventListener('click', toggleReasons);

  $('btnReasonsFill').addEventListener('click', () => {
    const shared = $('inpReason').value.trim();
    if (!shared) { toast('Nothing to copy', 'Type a shared reason first.', 'warn'); return; }
    for (const r of selectedRoles()) state.reasons.set(keyOf(r), shared);
    renderReasonSheet();
    syncReasonButton();
  });

  $('reasonList').addEventListener('input', (e) => {
    const input = e.target.closest('[data-rreason]');
    if (!input) return;
    state.reasons.set(input.dataset.rreason, input.value);
    input.closest('.reason-row')?.classList.remove('is-missing');
    syncReasonButton();
  });

  $('reasonList').addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && e.target.closest('[data-rreason]')) bulkActivate();
  });

  // delegated events for the role list
  $('stage').addEventListener('click', (e) => {
    const pick = e.target.closest('[data-pick]');
    if (pick) {
      const key = pick.dataset.pick;
      state.selected.has(key) ? state.selected.delete(key) : state.selected.add(key);
      const on = state.selected.has(key);
      pick.classList.toggle('is-on', on);
      pick.setAttribute('aria-checked', String(on));
      pick.closest('.rcard').classList.toggle('is-picked', on);
      syncActionbar();
      return;
    }
    const perms = e.target.closest('[data-perms]');
    if (perms) { toggleDrawer(perms.dataset.perms, 'perms'); return; }

    const more = e.target.closest('[data-permmore]');
    if (more) { state.permsExpanded.add(more.dataset.permmore); fillPermissions(more.dataset.permmore); return; }

    const less = e.target.closest('[data-permless]');
    if (less) { state.permsExpanded.delete(less.dataset.permless); fillPermissions(less.dataset.permless); return; }

    const act = e.target.closest('[data-activate]');
    if (act) { toggleDrawer(act.dataset.activate, 'activate'); return; }

    const confirm = e.target.closest('[data-confirm]');
    if (confirm) { confirmSingle(confirm.dataset.confirm); return; }

    const sched = e.target.closest('[data-sched]');
    if (sched) { state.schedulePrefill = sched.dataset.sched; renderSchedule(); openOverlay('panelSchedules'); return; }

    const off = e.target.closest('[data-off]');
    if (off) { deactivate(off.dataset.off); return; }

    if (e.target.closest('#ctaPortal')) { openPortal('pim'); }
  });

  $('stage').addEventListener('keydown', (e) => {
    if (e.key !== ' ' && e.key !== 'Enter') return;
    const pick = e.target.closest('[data-pick]');
    if (!pick) return;
    e.preventDefault();
    pick.click();
  });

  // schedules panel
  $('scheduleBody').addEventListener('click', (e) => {
    if (e.target.closest('#btnJobCreate')) { createJob(); return; }
    const del = e.target.closest('[data-jdel]');
    if (del) {
      ask('jobs/delete', { id: del.dataset.jdel })
        .then(() => ask('jobs/list'))
        .then(({ jobs }) => { state.jobs = jobs; renderMenuBadges(); renderSchedule(); toast('Schedule deleted', '', 'ok'); })
        .catch((err) => toast('Could not delete schedule', err.message, 'bad'));
    }
  });

  // history panel
  $('histFilter').addEventListener('click', (e) => {
    const btn = e.target.closest('[data-hfilter]');
    if (!btn || btn.dataset.hfilter === state.historyFilter) return;
    state.historyFilter = btn.dataset.hfilter;
    renderHistory();
  });
  $('btnHistReload').addEventListener('click', () => loadHistory(true));

  // permissions panel
  $('permScope').addEventListener('click', (e) => {
    const btn = e.target.closest('[data-pscope]');
    if (!btn || btn.dataset.pscope === state.permScope) return;
    state.permScope = btn.dataset.pscope;
    renderPerms();
  });

  $('permSearch').addEventListener('input', (e) => {
    state.permQuery = e.target.value;
    if (state.permState === 'ready') renderPerms();
  });

  $('permsBody').addEventListener('click', (e) => {
    const head = e.target.closest('[data-pgroup]');
    if (!head) return;
    const provider = head.dataset.pgroup;
    // While searching every group is forced open, so a click means "collapse".
    if (state.permQuery.trim()) {
      $('permSearch').value = '';
      state.permQuery = '';
      state.permGroupsOpen.add(provider);
    } else if (state.permGroupsOpen.has(provider)) {
      state.permGroupsOpen.delete(provider);
    } else {
      state.permGroupsOpen.add(provider);
    }
    renderPerms();
  });

  // portals panel
  $('portalBody').addEventListener('click', (e) => {
    const row = e.target.closest('[data-portal]');
    if (row) openPortal(row.dataset.portal);
  });
}

/* ================================== boot ================================= */

(async function boot() {
  wire();
  applyTheme();
  $('roleList').innerHTML = '<div class="skeleton"></div>'.repeat(4);
  state.openDrawers = new Map();
  await refresh();
  setInterval(tick, 1000);
})();
