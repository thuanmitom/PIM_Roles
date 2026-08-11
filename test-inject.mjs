/* Run inject.js inside a fake "window" to exercise the token-capture hook. */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const source = fs.readFileSync(new URL('./src/inject.js', import.meta.url), 'utf8');

function buildWindow() {
  const messages = [];
  const fetchCalls = [];
  const xhrSends = [];

  class FakeXHR {
    open(method, url) { this._method = method; this._url = url; this._headers = {}; }
    setRequestHeader(name, value) { this._headers[name] = value; }
    send(body) { xhrSends.push({ url: this._url, headers: this._headers, body }); }
  }

  class FakeHeaders {
    constructor(init) { this._map = new Map(Object.entries(init || {})); }
    get(name) {
      for (const [k, v] of this._map) if (k.toLowerCase() === name.toLowerCase()) return v;
      return null;
    }
  }

  class FakeRequest {
    constructor(url, init) { this.url = url; this.headers = new FakeHeaders((init || {}).headers); }
  }

  const win = {
    location: { origin: 'https://portal.azure.com' },
    postMessage: (data, origin) => messages.push({ data, origin }),
    fetch: async (...args) => { fetchCalls.push(args); return { ok: true }; },
    XMLHttpRequest: FakeXHR,
    Headers: FakeHeaders,
    Request: FakeRequest,
  };

  const context = vm.createContext({
    window: win,
    XMLHttpRequest: FakeXHR,
    Headers: FakeHeaders,
    Request: FakeRequest,
    location: win.location,
    Object,
    String,
    Array,
    Set,
    console,
  });
  vm.runInContext(source, context);

  return { win, messages, fetchCalls, xhrSends, FakeXHR, FakeHeaders, FakeRequest };
}

const b64 = (o) => Buffer.from(JSON.stringify(o)).toString('base64url');
const token = (suffix) =>
  `${b64({ alg: 'RS256' })}.${b64({ aud: 'https://graph.microsoft.com', oid: 'abc' })}.signature${suffix}`;

let passed = 0;
const test = async (name, fn) => {
  try {
    await fn();
    passed++;
    console.log('  OK   ' + name);
  } catch (e) {
    console.log('  FAIL ' + name + ' -> ' + e.message);
    process.exitCode = 1;
  }
};

const GRAPH_URL =
  'https://graph.microsoft.com/beta/roleManagement/directory/roleEligibilityScheduleInstances?$count=true';

await test('captures a token from fetch() with a plain object of headers', async () => {
  const { win, messages } = buildWindow();
  const t = token('1');
  await win.fetch(GRAPH_URL, { headers: { Authorization: `Bearer ${t}` } });

  assert.equal(messages.length, 1);
  assert.equal(messages[0].data.token, t, 'the Bearer prefix must be stripped');
  assert.equal(messages[0].data.kind, 'token');
  assert.equal(messages[0].origin, 'https://portal.azure.com', 'the postMessage origin must be pinned');
  assert.ok(!messages[0].data.url.includes('?'), 'the query string must be dropped');
});

await test('captures a token when headers come as Headers or Request', async () => {
  const { win, messages, FakeHeaders, FakeRequest } = buildWindow();

  await win.fetch(GRAPH_URL, { headers: new FakeHeaders({ authorization: `bearer ${token('2')}` }) });
  assert.equal(messages.length, 1, 'could not read from a Headers instance');

  await win.fetch(new FakeRequest(GRAPH_URL, { headers: { Authorization: `Bearer ${token('3')}` } }));
  assert.equal(messages.length, 2, 'could not read from a Request instance');
});

await test('captures a token from an array of header pairs', async () => {
  const { win, messages } = buildWindow();
  await win.fetch(GRAPH_URL, { headers: [['authorization', `Bearer ${token('4')}`]] });
  assert.equal(messages.length, 1);
});

await test('ignores requests that are not going to Microsoft Graph', async () => {
  const { win, messages } = buildWindow();
  await win.fetch('https://portal.azure.com/api/telemetry', {
    headers: { Authorization: `Bearer ${token('5')}` },
  });
  await win.fetch('https://management.azure.com/subscriptions', {
    headers: { Authorization: `Bearer ${token('6')}` },
  });
  assert.equal(messages.length, 0, 'only Graph traffic may be captured');
});

await test('captures a token sent to the Intune API (manage.microsoft.com)', async () => {
  const { win, messages } = buildWindow();
  await win.fetch('https://api.manage.microsoft.com/devices', {
    headers: { Authorization: `Bearer ${token('int')}` },
  });
  assert.equal(messages.length, 1, 'the Intune token must be captured too');
});

await test('never forwards the same token twice', async () => {
  const { win, messages } = buildWindow();
  const t = token('7');
  for (let i = 0; i < 5; i++) {
    await win.fetch(GRAPH_URL, { headers: { Authorization: `Bearer ${t}` } });
  }
  assert.equal(messages.length, 1, 'deduplication keeps the worker from being spammed');
});

await test('ignores anything that is not a JWT', async () => {
  const { win, messages } = buildWindow();
  await win.fetch(GRAPH_URL, { headers: { Authorization: 'Bearer not-a-jwt' } });
  await win.fetch(GRAPH_URL, { headers: { Authorization: '' } });
  assert.equal(messages.length, 0);
});

await test("the portal's own request still goes through untouched", async () => {
  const { win, fetchCalls } = buildWindow();
  const result = await win.fetch(GRAPH_URL, { headers: { Authorization: `Bearer ${token('8')}` } });
  assert.equal(fetchCalls.length, 1, 'the original fetch must still be called');
  assert.equal(fetchCalls[0][0], GRAPH_URL);
  assert.deepEqual(result, { ok: true }, 'the original fetch result must be returned');
});

await test('a hook failure never breaks fetch', async () => {
  const { win, fetchCalls } = buildWindow();
  // odd input with no .url — the hook must swallow the error, not rethrow it.
  const result = await win.fetch({ odd: 'input' }, { headers: null });
  assert.equal(fetchCalls.length, 1);
  assert.deepEqual(result, { ok: true });
});

await test('captures a token from XMLHttpRequest', async () => {
  const { messages, xhrSends, FakeXHR } = buildWindow();
  const t = token('9');
  const xhr = new FakeXHR();
  xhr.open('GET', GRAPH_URL);
  xhr.setRequestHeader('Authorization', `Bearer ${t}`);
  xhr.setRequestHeader('ConsistencyLevel', 'eventual');
  xhr.send();

  assert.equal(messages.length, 1);
  assert.equal(messages[0].data.token, t);
  assert.equal(xhrSends.length, 1, 'the original send() must still run');
  assert.equal(xhrSends[0].headers.ConsistencyLevel, 'eventual', 'other headers must survive');
});

await test('an XHR to another domain is not captured', async () => {
  const { messages, FakeXHR } = buildWindow();
  const xhr = new FakeXHR();
  xhr.open('POST', 'https://portal.azure.com/api/log');
  xhr.setRequestHeader('Authorization', `Bearer ${token('10')}`);
  xhr.send('{}');
  assert.equal(messages.length, 0);
});

await test('never hooks the same page twice', async () => {
  const { win, messages } = buildWindow();
  const first = win.fetch;
  vm.runInContext(source, vm.createContext({ window: win, XMLHttpRequest: win.XMLHttpRequest, Object, String, Array, Set, console }));
  assert.equal(win.fetch, first, 'the second run must bail out immediately');
  await win.fetch(GRAPH_URL, { headers: { Authorization: `Bearer ${token('11')}` } });
  assert.equal(messages.length, 1, 'stacked hooks must not duplicate messages');
});

console.log(`\n${passed} tests passed.`);
