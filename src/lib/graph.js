/* ---------------------------------------------------------------------------
 * Microsoft Graph client.
 * graph.microsoft.com is the only host this file ever talks to.
 * ------------------------------------------------------------------------- */

import { parseIsoDuration, parseGraphDate } from './fmt.js';

const GRAPH = 'https://graph.microsoft.com';
const BETA = `${GRAPH}/beta`;
const V1 = `${GRAPH}/v1.0`;

/** Plain-language advice for the Graph error codes that actually come up. */
const ERROR_HINTS = {
  InvalidAuthenticationToken: 'The token has expired. Reopen the PIM portal to capture a fresh one.',
  Authorization_RequestDenied:
    'The token lacks permission. RoleAssignmentSchedule.ReadWrite.Directory is required.',
  RoleAssignmentRequestPolicyValidationFailed:
    'Policy violation: the duration, justification or ticket does not satisfy the role policy.',
  RoleNotEligibleException: 'This account is not eligible for that role.',
  RoleAssignmentExists: 'That role is already active.',
  SubjectRequestNotAllowed: 'This role cannot be self-activated.',
  Request_BadRequest: 'The request was malformed.',
};

export class GraphError extends Error {
  constructor(status, code, message, requestId = '') {
    super(message || `HTTP ${status}`);
    this.name = 'GraphError';
    this.status = status;
    this.code = code || 'unknown';
    this.requestId = requestId;
    this.hint = ERROR_HINTS[this.code] || '';
  }
}

export class Graph {
  constructor(token) {
    this.token = token;
  }

  async request(method, url, { params, body, headers } = {}) {
    const target = new URL(url);
    if (params) {
      for (const [key, value] of Object.entries(params)) {
        if (value !== undefined && value !== null) {
          target.searchParams.set(key, value);
        }
      }
    }

    let response;
    for (let attempt = 0; ; attempt++) {
      response = await fetch(target.toString(), {
        method,
        headers: {
          Authorization: `Bearer ${this.token}`,
          'Content-Type': 'application/json',
          Accept: 'application/json',
          ...(headers || {}),
        },
        body: body ? JSON.stringify(body) : undefined,
      });

      // Throttled -> honour Retry-After and retry at most twice.
      if ((response.status === 429 || response.status === 503) && attempt < 2) {
        const wait = Number(response.headers.get('Retry-After') || 2 ** attempt);
        await new Promise((r) => setTimeout(r, Math.min(wait, 10) * 1000));
        continue;
      }
      break;
    }

    if (response.status === 204) return {};

    let payload = null;
    const text = await response.text();
    if (text) {
      try {
        payload = JSON.parse(text);
      } catch {
        payload = { raw: text };
      }
    }

    if (!response.ok) {
      const err = (payload && payload.error) || {};
      throw new GraphError(
        response.status,
        err.code,
        err.message || text.slice(0, 300),
        response.headers.get('request-id') || ''
      );
    }
    return payload || {};
  }

  get(url, params, headers) {
    return this.request('GET', url, { params, headers });
  }

  post(url, body) {
    return this.request('POST', url, { body });
  }

  /** Follow every @odata.nextLink page. */
  async getAll(url, params, headers) {
    const items = [];
    let page = await this.get(url, params, headers);
    for (;;) {
      items.push(...(page.value || []));
      const next = page['@odata.nextLink'];
      if (!next) break;
      page = await this.get(next, undefined, headers);
    }
    return items;
  }

  /* ------------------------------ endpoints ----------------------------- */

  organization() {
    return this.getAll(`${V1}/organization`);
  }

  user(principalId) {
    return this.get(`${V1}/users/${principalId}`, {
      $select:
        'id,displayName,userPrincipalName,mail,jobTitle,accountEnabled,officeLocation',
    });
  }

  /** Exactly the GET the PIM blade itself issues. */
  eligibleInstances(principalId) {
    return this.getAll(
      `${BETA}/roleManagement/directory/roleEligibilityScheduleInstances`,
      {
        $expand: 'principal,roleDefinition,directoryScope',
        $filter: `principalId eq '${principalId}'`,
        $count: 'true',
      },
      // Mandatory alongside $count — without it Graph answers 400.
      { ConsistencyLevel: 'eventual' }
    );
  }

  activeInstances(principalId) {
    return this.getAll(
      `${V1}/roleManagement/directory/roleAssignmentScheduleInstances`,
      {
        $expand: 'roleDefinition,directoryScope',
        $filter: `principalId eq '${principalId}'`,
      },
      { ConsistencyLevel: 'eventual' }
    );
  }

  roleDefinition(roleDefinitionId) {
    return this.get(`${V1}/roleManagement/directory/roleDefinitions/${roleDefinitionId}`);
  }

  async policyRules(roleDefinitionId, scopeId = '/') {
    const data = await this.get(`${V1}/policies/roleManagementPolicyAssignments`, {
      $filter:
        `scopeId eq '${scopeId}' and scopeType eq 'Directory' ` +
        `and roleDefinitionId eq '${roleDefinitionId}'`,
      $expand: 'policy($expand=rules)',
    });
    const first = (data.value || [])[0];
    return (first && first.policy && first.policy.rules) || [];
  }

  createRequest(body) {
    return this.post(`${V1}/roleManagement/directory/roleAssignmentScheduleRequests`, body);
  }

  requestStatus(requestId) {
    return this.get(
      `${V1}/roleManagement/directory/roleAssignmentScheduleRequests/${requestId}`
    );
  }

  /**
   * Every activation request this principal ever made — the activate /
   * deactivate history. One page only: `getAll` would walk months of records
   * behind the scenes, and the panel never shows more than `top` rows.
   */
  async requestHistory(principalId, top = 80) {
    const page = await this.get(
      `${V1}/roleManagement/directory/roleAssignmentScheduleRequests`,
      {
        $filter: `principalId eq '${principalId}'`,
        $expand: 'roleDefinition',
        $top: String(top),
      },
      { ConsistencyLevel: 'eventual' }
    );
    return page.value || [];
  }
}

/* ----------------------------- data shaping ------------------------------ */

/** Reduce the raw PIM policy rules to the handful of flags the UI cares about. */
export function readPolicy(rules) {
  const policy = {
    maxDuration: null,
    requireJustification: false,
    requireMfa: false,
    requireTicket: false,
    requireApproval: false,
  };

  for (const rule of rules || []) {
    const id = rule.id || '';
    if (id === 'Expiration_EndUser_Assignment') {
      policy.maxDuration = parseIsoDuration(rule.maximumDuration);
    } else if (id === 'Enablement_EndUser_Assignment') {
      const enabled = (rule.enabledRules || []).map((x) => String(x).toLowerCase());
      policy.requireJustification = enabled.includes('justification');
      policy.requireMfa = enabled.includes('multifactorauthentication');
      policy.requireTicket = enabled.includes('ticketing');
    } else if (id === 'Approval_EndUser_Assignment') {
      policy.requireApproval = Boolean(
        rule.setting && rule.setting.isApprovalRequired
      );
    }
  }
  return policy;
}

/** Normalise one Graph instance into the shape the UI renders. */
export function toRole(raw) {
  const definition = raw.roleDefinition || {};
  const permissions = [];
  for (const block of definition.rolePermissions || []) {
    permissions.push(...(block.allowedResourceActions || []));
  }
  const scope = raw.directoryScope || {};
  const scopeId = raw.directoryScopeId || '/';

  return {
    instanceId: raw.id || '',
    roleDefinitionId: raw.roleDefinitionId || definition.id || '',
    name: definition.displayName || 'Unnamed role',
    description: definition.description || '',
    scopeId,
    scopeLabel:
      scopeId === '/' || scopeId === ''
        ? 'Tenant-wide'
        : scope.displayName || scope.id || scopeId,
    start: parseGraphDate(raw.startDateTime),
    end: parseGraphDate(raw.endDateTime),
    memberType: raw.memberType || '',
    assignmentType: raw.assignmentType || '',
    isBuiltIn: definition.isBuiltIn !== false,
    permissions,
    policy: null,
  };
}

/**
 * What each PIM request `action` means, and whether it switched a role on or
 * off. `kind` drives the colour of the history row.
 */
export const REQUEST_ACTIONS = {
  selfActivate:    { kind: 'on',  label: 'Activated' },
  selfDeactivate:  { kind: 'off', label: 'Deactivated' },
  selfExtend:      { kind: 'on',  label: 'Extended' },
  selfRenew:       { kind: 'on',  label: 'Renewed' },
  adminAssign:     { kind: 'on',  label: 'Assigned by an admin' },
  adminRemove:     { kind: 'off', label: 'Removed by an admin' },
  adminUpdate:     { kind: 'on',  label: 'Updated by an admin' },
  adminExtend:     { kind: 'on',  label: 'Extended by an admin' },
  adminRenew:      { kind: 'on',  label: 'Renewed by an admin' },
};

/** Statuses that mean the request never took effect. */
export const FAILED_STATUSES = new Set(['Failed', 'Denied', 'Canceled', 'Revoked']);

/**
 * In a log, `Provisioned` means "this request went through" — not "the role is
 * on now", which is what the same word means on a live role card. So the
 * history renames the two success statuses and leaves the rest alone.
 */
const HISTORY_STATUS_LABEL = { Provisioned: 'Succeeded', Granted: 'Succeeded' };

/** Normalise one roleAssignmentScheduleRequest into a history row. */
export function toHistoryEntry(raw) {
  const definition = raw.roleDefinition || {};
  const schedule = raw.scheduleInfo || {};
  const expiration = schedule.expiration || {};
  const action = raw.action || '';
  const meta = REQUEST_ACTIONS[action] || { kind: 'on', label: action || 'Request' };
  const scopeId = raw.directoryScopeId || '/';
  const status = raw.status || '';

  return {
    id: raw.id || '',
    action,
    kind: meta.kind,
    actionLabel: meta.label,
    roleDefinitionId: raw.roleDefinitionId || definition.id || '',
    name: definition.displayName || 'Unnamed role',
    scopeId,
    scopeLabel: scopeId === '/' || scopeId === '' ? 'Tenant-wide' : scopeId,
    status,
    statusLabel: HISTORY_STATUS_LABEL[status] || STATUS_LABEL[status] || status || '—',
    failed: FAILED_STATUSES.has(status),
    createdAt: parseGraphDate(raw.createdDateTime),
    completedAt: parseGraphDate(raw.completedDateTime),
    startAt: parseGraphDate(schedule.startDateTime),
    endAt: parseGraphDate(expiration.endDateTime),
    duration: parseIsoDuration(expiration.duration),
    justification: raw.justification || '',
    ticketNumber: (raw.ticketInfo && raw.ticketInfo.ticketNumber) || '',
  };
}

/**
 * Fold every role a principal holds into one action -> grantors index.
 *
 *   roles   [{ name, permissions, kind }]  kind: 'active' | 'assigned' | 'eligible'
 *   returns Map<action, { providers, grantors: [{ name, kind }] }>
 *
 * `kind` is what separates the two views the panel offers: what is in force
 * right now (active + permanently assigned) from what one activation away.
 */
export function indexPermissions(roles) {
  const index = new Map();
  for (const role of roles) {
    for (const action of role.permissions || []) {
      let entry = index.get(action);
      if (!entry) {
        entry = { grantors: [] };
        index.set(action, entry);
      }
      if (!entry.grantors.some((g) => g.name === role.name && g.kind === role.kind)) {
        entry.grantors.push({ name: role.name, kind: role.kind });
      }
    }
  }
  return index;
}

/** Build the request body for a selfActivate call. */
export function activationBody({
  principalId,
  role,
  duration,
  justification,
  startAt,
  ticketNumber,
  ticketSystem,
}) {
  const schedule = {
    expiration: { type: 'AfterDuration', duration },
  };
  if (startAt) schedule.startDateTime = new Date(startAt).toISOString();

  const body = {
    action: 'selfActivate',
    principalId,
    roleDefinitionId: role.roleDefinitionId,
    directoryScopeId: role.scopeId || '/',
    // A bulk activation may give each role its own reason; the shared one is
    // the fallback for every role that was left blank.
    justification: role.justification || justification,
    scheduleInfo: schedule,
  };
  if (ticketNumber || ticketSystem) {
    body.ticketInfo = {
      ticketNumber: ticketNumber || '',
      ticketSystem: ticketSystem || '',
    };
  }
  return body;
}

export function deactivationBody({ principalId, role, justification }) {
  return {
    action: 'selfDeactivate',
    principalId,
    roleDefinitionId: role.roleDefinitionId,
    directoryScopeId: role.scopeId || '/',
    justification: justification || 'Work finished',
  };
}

export const STATUS_LABEL = {
  Provisioned: 'Active',
  Granted: 'Granted',
  Succeeded: 'Succeeded',
  PendingApproval: 'Awaiting approval',
  PendingApprovalProvisioning: 'Awaiting approval',
  PendingProvisioning: 'Provisioning',
  PendingScheduleCreation: 'Creating schedule',
  ScheduleCreated: 'Scheduled',
  PendingAdminDecision: 'Awaiting admin decision',
  Failed: 'Failed',
  Denied: 'Denied',
  Canceled: 'Cancelled',
  Revoked: 'Revoked',
};

export const FINAL_STATUSES = new Set([
  'Provisioned',
  'Granted',
  'Failed',
  'Denied',
  'Canceled',
  'Revoked',
]);
