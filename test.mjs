/* Unit tests for the pure-logic modules, with a faked fetch layer. */
import assert from 'node:assert/strict';
import {
  Graph, GraphError, readPolicy, toRole, activationBody, deactivationBody,
  toHistoryEntry, indexPermissions,
} from './src/lib/graph.js';
import { decodeJwt, summarize, isExpired, classifyToken, vaultKey } from './src/lib/jwt.js';

// Unsigned sample JWTs; audiences and scopes mirror what the portals really send.
const b64u = (o) => Buffer.from(JSON.stringify(o)).toString('base64url');
const mkjwt = (claims) => `${b64u({ alg: 'none', typ: 'JWT' })}.${b64u(claims)}.sig`;
const FAR = 4102444800; // exp in the year 2100 — valid for a long time
const SAMPLE = {
  graphPim: mkjwt({ aud: 'https://graph.microsoft.com', tid: 't1', oid: 'u1', upn: 'a@x.com',
    scp: 'RoleAssignmentSchedule.ReadWrite.Directory RoleEligibilitySchedule.Read.Directory', exp: FAR }),
  graphIntune: mkjwt({ aud: '00000003-0000-0000-c000-000000000000', tid: 't1', oid: 'u1',
    scp: 'DeviceManagementConfiguration.ReadWrite.All', exp: FAR }),
  intuneApi: mkjwt({ aud: 'https://api.manage.microsoft.com/', tid: 't1', oid: 'u1', exp: FAR }),
  graphPlain: mkjwt({ aud: 'https://graph.microsoft.com', tid: 't1', oid: 'u1', scp: 'User.Read', exp: FAR }),
  other: mkjwt({ aud: 'https://example.com', tid: 't1', exp: FAR }),
};
import {
  parseIsoDuration, toIsoDuration, normalizeDuration, humanize, clock, nextOccurrence,
  dayKey, formatDay,
} from './src/lib/fmt.js';

let passed = 0;
const test = (name, fn) => {
  try {
    fn();
    passed++;
    console.log('  OK   ' + name);
  } catch (e) {
    console.log('  FAIL ' + name + ' -> ' + e.message);
    process.exitCode = 1;
  }
};
const testAsync = async (name, fn) => {
  try {
    await fn();
    passed++;
    console.log('  OK   ' + name);
  } catch (e) {
    console.log('  FAIL ' + name + ' -> ' + e.message);
    process.exitCode = 1;
  }
};

/* ------------------------------- fmt.js ---------------------------------- */

test('parses ISO 8601 durations', () => {
  assert.equal(parseIsoDuration('PT8H'), 8 * 3600000);
  assert.equal(parseIsoDuration('PT1H30M'), 5400000);
  assert.equal(parseIsoDuration('P1DT2H'), 26 * 3600000);
  assert.equal(parseIsoDuration('nonsense'), null);
  assert.equal(parseIsoDuration(''), null);
});

test('converts milliseconds back to ISO 8601', () => {
  assert.equal(toIsoDuration(8 * 3600000), 'PT8H');
  assert.equal(toIsoDuration(5400000), 'PT1H30M');
  assert.equal(toIsoDuration(26 * 3600000), 'P1DT2H');
});

test('accepts the shorthand people actually type', () => {
  assert.equal(normalizeDuration('8h'), 'PT8H');
  assert.equal(normalizeDuration('90m'), 'PT1H30M');
  assert.equal(normalizeDuration('2h30m'), 'PT2H30M');
  assert.equal(normalizeDuration('1d2h'), 'P1DT2H');
  assert.equal(normalizeDuration('PT4H'), 'PT4H');
  assert.equal(normalizeDuration('abc'), null);
  assert.equal(normalizeDuration(''), null);
});

test('renders durations for humans', () => {
  assert.equal(humanize(5400000), '1h 30m');
  assert.equal(humanize(-5000), 'expired');
  assert.equal(humanize(90 * 24 * 3600000), '90d');
  assert.equal(clock(5400000), '01:30:00');
  assert.equal(clock(-100), '00:00:00');
});

test('groups timestamps into calendar days', () => {
  const now = Date.now();
  assert.equal(formatDay(now), 'Today');
  assert.equal(formatDay(now - 86400000), 'Yesterday');
  assert.notEqual(dayKey(now), dayKey(now - 86400000));
  // A key is local-calendar based, so two moments on the same day collapse.
  const noon = new Date(); noon.setHours(12, 0, 0, 0);
  const evening = new Date(); evening.setHours(23, 30, 0, 0);
  assert.equal(dayKey(noon.getTime()), dayKey(evening.getTime()));
  assert.match(formatDay(Date.UTC(2026, 0, 15, 12)), /Jan/);
});

test('finds the next occurrence of a time of day', () => {
  const t = nextOccurrence('09:30');
  assert.ok(t > Date.now(), 'must be in the future');
  assert.equal(new Date(t).getMinutes(), 30);
  assert.equal(nextOccurrence('99:99'), null);
});

/* ------------------------------- jwt.js ---------------------------------- */

const b64 = (obj) => Buffer.from(JSON.stringify(obj)).toString('base64url');
const makeToken = (claims) => `${b64({ alg: 'RS256' })}.${b64(claims)}.signature`;

const PRINCIPAL = '8f3a1c22-1111-2222-3333-444455556666';
const TENANT = 'aaaabbbb-cccc-dddd-eeee-ffff00001111';
const validToken = makeToken({
  aud: 'https://graph.microsoft.com',
  oid: PRINCIPAL,
  tid: TENANT,
  upn: 'admin@contoso.onmicrosoft.com',
  name: 'Lê Thuận',
  scp: 'RoleAssignmentSchedule.ReadWrite.Directory',
  exp: Math.floor(Date.now() / 1000) + 3300,
});

test('decodes a JWT and reads its claims', () => {
  const claims = decodeJwt(validToken);
  assert.equal(claims.oid, PRINCIPAL);
  assert.equal(claims.tid, TENANT);
  assert.equal(decodeJwt('not-a-jwt'), null);
});

test('keeps non-ASCII display names intact', () => {
  assert.equal(decodeJwt(validToken).name, 'Lê Thuận');
  assert.equal(summarize(validToken).name, 'Lê Thuận');
});

test('detects an expired token', () => {
  assert.equal(isExpired(decodeJwt(validToken)), false);
  const old = makeToken({ exp: Math.floor(Date.now() / 1000) - 10 });
  assert.equal(isExpired(decodeJwt(old)), true);
});

/* ------------------------- Graph data conversion ------------------------- */

test('reads the PIM policy rules', () => {
  const policy = readPolicy([
    { id: 'Expiration_EndUser_Assignment', maximumDuration: 'PT6H' },
    { id: 'Enablement_EndUser_Assignment', enabledRules: ['Justification', 'MultiFactorAuthentication'] },
    { id: 'Approval_EndUser_Assignment', setting: { isApprovalRequired: true } },
  ]);
  assert.equal(policy.maxDuration, 6 * 3600000);
  assert.equal(policy.requireJustification, true);
  assert.equal(policy.requireMfa, true);
  assert.equal(policy.requireTicket, false);
  assert.equal(policy.requireApproval, true);
});

test('normalises a Graph instance', () => {
  const role = toRole({
    id: 'elig-1',
    roleDefinitionId: 'f2ef992c-3afb-46b9-b7cf-a126ee74c451',
    directoryScopeId: '/',
    startDateTime: '2026-06-04T04:10:00Z',
    endDateTime: null,
    memberType: 'Direct',
    roleDefinition: {
      displayName: 'Global Reader',
      isBuiltIn: true,
      rolePermissions: [{ allowedResourceActions: ['microsoft.directory/users/standard/read'] }],
    },
    directoryScope: { id: '/' },
  });
  assert.equal(role.name, 'Global Reader');
  assert.equal(role.scopeLabel, 'Tenant-wide');
  assert.equal(role.end, null);
  assert.equal(role.permissions.length, 1);
});

test('permissions survive normalisation for every assignment type', () => {
  // The permission drawer works on active and permanently assigned roles too,
  // so toRole must carry the permission list regardless of assignmentType.
  const definition = {
    displayName: 'User Administrator',
    rolePermissions: [{ allowedResourceActions: ['microsoft.directory/users/create'] }],
  };
  for (const assignmentType of ['Activated', 'Assigned']) {
    const role = toRole({ roleDefinitionId: 'x', assignmentType, roleDefinition: definition });
    assert.equal(role.assignmentType, assignmentType);
    assert.deepEqual(role.permissions, ['microsoft.directory/users/create']);
  }
});

/* ---------------------------- request history ---------------------------- */

test('normalises an activation request into a history row', () => {
  const entry = toHistoryEntry({
    id: 'req-1',
    action: 'selfActivate',
    status: 'Provisioned',
    directoryScopeId: '/',
    createdDateTime: '2026-08-04T09:12:00Z',
    completedDateTime: '2026-08-04T09:12:04Z',
    justification: 'INC-4471',
    ticketInfo: { ticketNumber: 'INC-4471', ticketSystem: 'ITSM' },
    scheduleInfo: {
      startDateTime: '2026-08-04T09:12:00Z',
      expiration: { type: 'AfterDuration', duration: 'PT8H' },
    },
    roleDefinition: { displayName: 'User Administrator' },
  });
  assert.equal(entry.kind, 'on');
  assert.equal(entry.actionLabel, 'Activated');
  assert.equal(entry.name, 'User Administrator');
  // in a log "Provisioned" means the request went through, not "on right now"
  assert.equal(entry.statusLabel, 'Succeeded');
  assert.equal(entry.failed, false);
  assert.equal(entry.duration, 8 * 3600000);
  assert.equal(entry.scopeLabel, 'Tenant-wide');
  assert.equal(entry.ticketNumber, 'INC-4471');
  assert.equal(entry.createdAt, Date.parse('2026-08-04T09:12:00Z'));
});

test('a deactivation is tagged off, a denial is tagged failed', () => {
  const off = toHistoryEntry({ action: 'selfDeactivate', status: 'Revoked', roleDefinition: {} });
  assert.equal(off.kind, 'off');
  assert.equal(off.actionLabel, 'Deactivated');
  assert.equal(off.failed, true);

  const denied = toHistoryEntry({ action: 'selfActivate', status: 'Denied', roleDefinition: {} });
  assert.equal(denied.kind, 'on');
  assert.equal(denied.failed, true);
  assert.equal(denied.statusLabel, 'Denied');

  // An unknown action still renders rather than blowing up.
  const odd = toHistoryEntry({ action: 'adminWhatever', roleDefinition: {} });
  assert.equal(odd.actionLabel, 'adminWhatever');
  assert.equal(odd.kind, 'on');
});

/* -------------------------- permission aggregation ----------------------- */

test('folds every role into one action -> grantors index', () => {
  const index = indexPermissions([
    { name: 'Global Reader', kind: 'assigned', permissions: ['microsoft.directory/users/standard/read'] },
    { name: 'User Administrator', kind: 'active',
      permissions: ['microsoft.directory/users/standard/read', 'microsoft.directory/users/create'] },
    { name: 'Intune Administrator', kind: 'eligible', permissions: ['microsoft.intune/allEntities/allTasks'] },
  ]);

  assert.equal(index.size, 3);
  // one action shared by two roles keeps both grantors, with their buckets
  const shared = index.get('microsoft.directory/users/standard/read');
  assert.equal(shared.grantors.length, 2);
  assert.deepEqual(shared.grantors.map((g) => g.kind).sort(), ['active', 'assigned']);

  // what is in force now excludes anything only an eligible role grants
  const inForce = [...index.entries()]
    .filter(([, e]) => e.grantors.some((g) => g.kind !== 'eligible'))
    .map(([action]) => action);
  assert.equal(inForce.length, 2);
  assert.ok(!inForce.includes('microsoft.intune/allEntities/allTasks'));
});

test('the permission index deduplicates a role listed twice', () => {
  const index = indexPermissions([
    { name: 'Groups Administrator', kind: 'active', permissions: ['a/b', 'a/b'] },
    { name: 'Groups Administrator', kind: 'active', permissions: ['a/b'] },
  ]);
  assert.equal(index.get('a/b').grantors.length, 1);
});

test('roles without permissions do not appear in the index', () => {
  const index = indexPermissions([{ name: 'Empty', kind: 'active' }]);
  assert.equal(index.size, 0);
});

test('builds a selfActivate body against the right schema', () => {
  const body = activationBody({
    principalId: PRINCIPAL,
    role: { roleDefinitionId: 'abc', scopeId: '/' },
    duration: 'PT6H',
    justification: 'INC-4471',
  });
  assert.equal(body.action, 'selfActivate');
  assert.equal(body.principalId, PRINCIPAL);
  assert.equal(body.directoryScopeId, '/');
  assert.equal(body.scheduleInfo.expiration.type, 'AfterDuration');
  assert.equal(body.scheduleInfo.expiration.duration, 'PT6H');
  assert.equal(body.scheduleInfo.startDateTime, undefined);
  assert.equal(body.ticketInfo, undefined);
});

test('a scheduled body carries an ISO startDateTime', () => {
  const when = Date.UTC(2026, 7, 4, 2, 0, 0);
  const body = activationBody({
    principalId: PRINCIPAL,
    role: { roleDefinitionId: 'abc', scopeId: '/' },
    duration: 'PT4H',
    justification: 'on call',
    startAt: when,
    ticketNumber: 'INC-1',
    ticketSystem: 'ITSM',
  });
  assert.equal(body.scheduleInfo.startDateTime, '2026-08-04T02:00:00.000Z');
  assert.deepEqual(body.ticketInfo, { ticketNumber: 'INC-1', ticketSystem: 'ITSM' });
});

test('a per-role reason overrides the shared one', () => {
  // Bulk activation: each role may carry its own justification, and any role
  // left blank falls back to the reason typed once for the whole batch.
  const own = activationBody({
    principalId: PRINCIPAL,
    role: { roleDefinitionId: 'abc', scopeId: '/', justification: 'INC-9001 mailbox' },
    duration: 'PT2H',
    justification: 'shared reason',
  });
  assert.equal(own.justification, 'INC-9001 mailbox');

  const fallback = activationBody({
    principalId: PRINCIPAL,
    role: { roleDefinitionId: 'abc', scopeId: '/' },
    duration: 'PT2H',
    justification: 'shared reason',
  });
  assert.equal(fallback.justification, 'shared reason');
});

test('builds a selfDeactivate body', () => {
  const body = deactivationBody({
    principalId: PRINCIPAL,
    role: { roleDefinitionId: 'abc', scopeId: '/' },
  });
  assert.equal(body.action, 'selfDeactivate');
  assert.equal(body.justification, 'Work finished');
});

/* ---------------------- Graph client against a fake fetch ---------------- */

const calls = [];
function mockFetch(routes) {
  globalThis.fetch = async (url, init) => {
    calls.push({ url, init });
    for (const [pattern, handler] of routes) {
      if (url.includes(pattern)) return handler(url, init);
    }
    return new Response(JSON.stringify({ error: { code: 'notFound', message: url } }), {
      status: 404,
      headers: { 'Content-Type': 'application/json' },
    });
  };
}
const json = (body, status = 200, headers = {}) =>
  new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json', ...headers } });

await testAsync('eligible query uses the right filter and ConsistencyLevel header', async () => {
  let seen = null;
  mockFetch([
    ['roleEligibilityScheduleInstances', (url, init) => {
      seen = { url: new URL(url), headers: init.headers };
      return json({ value: [{ id: 'e1', roleDefinition: { displayName: 'Global Reader' } }] });
    }],
  ]);

  const rows = await new Graph(validToken).eligibleInstances(PRINCIPAL);
  assert.equal(rows.length, 1);
  assert.equal(seen.headers.ConsistencyLevel, 'eventual', 'without it Graph answers 400');
  assert.equal(seen.headers.Authorization, `Bearer ${validToken}`);
  assert.equal(seen.url.searchParams.get('$count'), 'true');
  assert.equal(seen.url.searchParams.get('$filter'), `principalId eq '${PRINCIPAL}'`);
  assert.equal(seen.url.searchParams.get('$expand'), 'principal,roleDefinition,directoryScope');
  assert.equal(seen.url.pathname, '/beta/roleManagement/directory/roleEligibilityScheduleInstances');
});

await testAsync('follows every @odata.nextLink page', async () => {
  let page = 0;
  mockFetch([
    ['roleAssignmentScheduleInstances', () => {
      page++;
      return page === 1
        ? json({
            value: [{ id: 'a1' }],
            '@odata.nextLink':
              'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleInstances?$skiptoken=2',
          })
        : json({ value: [{ id: 'a2' }] });
    }],
  ]);
  const rows = await new Graph(validToken).activeInstances(PRINCIPAL);
  assert.equal(rows.length, 2);
  assert.deepEqual(rows.map((r) => r.id), ['a1', 'a2']);
});

await testAsync('Graph errors are translated into a readable hint', async () => {
  mockFetch([
    ['roleAssignmentScheduleRequests', () =>
      json({ error: { code: 'RoleAssignmentExists', message: 'The Role assignment already exists.' } }, 400)],
  ]);
  await assert.rejects(
    () => new Graph(validToken).createRequest({ action: 'selfActivate' }),
    (err) => {
      assert.ok(err instanceof GraphError);
      assert.equal(err.status, 400);
      assert.equal(err.code, 'RoleAssignmentExists');
      assert.equal(err.hint, 'That role is already active.');
      return true;
    }
  );
});

await testAsync('honours Retry-After when throttled with 429', async () => {
  let attempts = 0;
  mockFetch([
    ['roleDefinitions', () => {
      attempts++;
      return attempts === 1
        ? json({ error: { code: 'throttled', message: 'slow down' } }, 429, { 'Retry-After': '0' })
        : json({ id: 'r1', displayName: 'Global Reader' });
    }],
  ]);
  const definition = await new Graph(validToken).roleDefinition('r1');
  assert.equal(attempts, 2, 'must retry after a 429');
  assert.equal(definition.displayName, 'Global Reader');
});

await testAsync('policy assignments are queried with the right filter', async () => {
  let seen = null;
  mockFetch([
    ['roleManagementPolicyAssignments', (url) => {
      seen = new URL(url);
      return json({
        value: [{ policy: { rules: [{ id: 'Expiration_EndUser_Assignment', maximumDuration: 'PT6H' }] } }],
      });
    }],
  ]);
  const rules = await new Graph(validToken).policyRules('abc', '/');
  assert.equal(readPolicy(rules).maxDuration, 6 * 3600000);
  assert.ok(seen.searchParams.get('$filter').includes("scopeType eq 'Directory'"));
  assert.equal(seen.searchParams.get('$expand'), 'policy($expand=rules)');
});

await testAsync('the history query asks for one page of this principal only', async () => {
  let seen = null;
  let pages = 0;
  mockFetch([
    ['roleAssignmentScheduleRequests', (url, init) => {
      pages++;
      seen = { url: new URL(url), headers: init.headers };
      return json({
        value: [{ id: 'r1', action: 'selfActivate', roleDefinition: { displayName: 'Global Reader' } }],
        // A nextLink must NOT be followed — the panel shows one page.
        '@odata.nextLink':
          'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests?$skiptoken=2',
      });
    }],
  ]);

  const rows = await new Graph(validToken).requestHistory(PRINCIPAL, 80);
  assert.equal(pages, 1, 'the history must not walk every page of records');
  assert.equal(rows.length, 1);
  assert.equal(seen.url.searchParams.get('$filter'), `principalId eq '${PRINCIPAL}'`);
  assert.equal(seen.url.searchParams.get('$top'), '80');
  assert.equal(seen.url.searchParams.get('$expand'), 'roleDefinition');
  assert.equal(seen.headers.ConsistencyLevel, 'eventual');
  assert.equal(toHistoryEntry(rows[0]).name, 'Global Reader');
});

await testAsync('graph.microsoft.com is the only host ever called', async () => {
  const hosts = new Set(calls.map((c) => new URL(c.url).host));
  assert.deepEqual([...hosts], ['graph.microsoft.com']);
});

/* ---------------------- token classification / vault --------------------- */

test('classify: a Graph token with PIM scopes -> role', () => {
  const k = classifyToken(decodeJwt(SAMPLE.graphPim));
  assert.equal(k.resource, 'graph');
  assert.equal(k.scopeClass, 'role');
  assert.match(k.label, /PIM/);
});

test('classify: a Graph token for Intune (aud is an app-id GUID) -> devicemgmt', () => {
  const k = classifyToken(decodeJwt(SAMPLE.graphIntune));
  assert.equal(k.resource, 'graph');
  assert.equal(k.scopeClass, 'devicemgmt');
  assert.match(k.label, /Intune/);
});

test('classify: an api.manage.microsoft.com token -> intune', () => {
  const k = classifyToken(decodeJwt(SAMPLE.intuneApi));
  assert.equal(k.resource, 'intune');
  assert.equal(k.label, 'Intune API');
});

test('classify: a non-Microsoft token -> other (and is dropped)', () => {
  assert.equal(classifyToken(decodeJwt(SAMPLE.other)).resource, 'other');
});

test('vaultKey separates the PIM and Intune-Graph tokens of one tenant', () => {
  const pim = decodeJwt(SAMPLE.graphPim);
  const intune = decodeJwt(SAMPLE.graphIntune);
  const kPim = vaultKey(summarize(SAMPLE.graphPim), classifyToken(pim));
  const kIntune = vaultKey(summarize(SAMPLE.graphIntune), classifyToken(intune));
  assert.notEqual(kPim, kIntune);          // neither may overwrite the other
  assert.match(kPim, /graph\|t1\|role/);
  assert.match(kIntune, /graph\|t1\|devicemgmt/);
});

test('PIM token choice: a role-scoped Graph token beats a plain one', () => {
  // Mirrors getPimToken() from the worker against a stand-in vault.
  const vault = [SAMPLE.graphPlain, SAMPLE.graphPim].map((tok) => {
    const k = classifyToken(decodeJwt(tok));
    return { token: tok, resource: k.resource, scopeClass: k.scopeClass, capturedAt: 1 };
  });
  const graph = vault.filter((t) => t.resource === 'graph');
  const chosen = graph.find((t) => t.scopeClass === 'role') || graph[0];
  assert.equal(chosen.scopeClass, 'role');
});

console.log(`\n${passed} tests passed.`);
