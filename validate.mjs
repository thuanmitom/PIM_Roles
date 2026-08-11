/* Structural and security checks to run before loading the extension. */
// popup.js calls boot() on import — there is no DOM in Node, so swallow that.
process.on('uncaughtException', () => {});
process.on('unhandledRejection', () => {});

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));
const problems = [];
const notes = [];
const ok = (m) => notes.push('  OK   ' + m);
const bad = (m) => problems.push('  FAIL ' + m);

const read = (p) => fs.readFileSync(path.join(root, p), 'utf8');
const exists = (p) => fs.existsSync(path.join(root, p));

/* ------------------------------ 1. manifest ------------------------------ */
let manifest;
try {
  manifest = JSON.parse(read('manifest.json'));
  ok('manifest.json is valid JSON');
} catch (e) {
  bad('manifest.json is broken: ' + e.message);
  console.log(problems.join('\n'));
  process.exit(1);
}

if (manifest.manifest_version !== 3) bad('manifest_version must be 3');
else ok('manifest_version = 3');

const referenced = [
  manifest.background?.service_worker,
  manifest.action?.default_popup,
  ...Object.values(manifest.icons || {}),
  ...Object.values(manifest.action?.default_icon || {}),
  ...(manifest.content_scripts || []).flatMap((c) => c.js || []),
].filter(Boolean);

for (const file of referenced) {
  if (exists(file)) ok(`exists: ${file}`);
  else bad(`manifest points at a missing file: ${file}`);
}

// every icon the manifest declares must be a real PNG
for (const [size, file] of Object.entries(manifest.icons || {})) {
  if (!exists(file)) continue;
  const head = fs.readFileSync(path.join(root, file)).subarray(0, 8);
  if (head.equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))) {
    ok(`icon ${size} is a real PNG`);
  } else {
    bad(`icon ${size} (${file}) is not a PNG`);
  }
}

// a MAIN-world content script needs an isolated partner to relay through
const worlds = (manifest.content_scripts || []).map((c) => c.world || 'ISOLATED');
if (worlds.includes('MAIN') && worlds.includes('ISOLATED')) {
  ok('both a MAIN and an ISOLATED content script are declared');
} else {
  bad('missing the MAIN + ISOLATED content script pair');
}

for (const need of ['storage', 'alarms']) {
  if ((manifest.permissions || []).includes(need)) ok(`permission present: ${need}`);
  else bad(`missing permission: ${need}`);
}

for (const host of ['https://portal.azure.com/*', 'https://graph.microsoft.com/*']) {
  if ((manifest.host_permissions || []).includes(host)) ok(`host_permission present: ${host}`);
  else bad(`missing host_permission: ${host}`);
}

/* ------------------------ 2. module syntax per file ---------------------- */
const jsFiles = [
  'src/background.js',
  'src/content.js',
  'src/inject.js',
  'src/lib/graph.js',
  'src/lib/jwt.js',
  'src/lib/fmt.js',
  'popup/popup.js',
];

for (const file of jsFiles) {
  if (!exists(file)) {
    bad(`missing file ${file}`);
    continue;
  }
  try {
    // Import for real so genuine syntax errors surface. Runtime errors from the
    // absent chrome.* APIs are expected and ignored.
    await import(pathToFileURL(path.join(root, file)).href).catch((e) => {
      if (e instanceof SyntaxError) throw e;
    });
    ok(`syntax is valid: ${file}`);
  } catch (e) {
    bad(`syntax error in ${file}: ${e.message}`);
  }
}

/* --------------------- 3. cross-check IDs: HTML vs JS -------------------- */
const html = read('popup/popup.html');
const popupJs = read('popup/popup.js');

const htmlIds = new Set([...html.matchAll(/\bid="([^"]+)"/g)].map((m) => m[1]));
const jsStaticIds = new Set([...popupJs.matchAll(/\$\('([A-Za-z0-9_]+)'\)/g)].map((m) => m[1]));
// IDs produced inside popup.js template strings
const jsTemplateIds = new Set([...popupJs.matchAll(/id="([A-Za-z0-9_]+)"/g)].map((m) => m[1]));

let missing = 0;
for (const id of jsStaticIds) {
  if (!htmlIds.has(id) && !jsTemplateIds.has(id)) {
    bad(`popup.js calls $('${id}') but nothing carries that id`);
    missing++;
  }
}
if (!missing) ok(`all ${jsStaticIds.size} ids used by popup.js resolve (static HTML or template)`);

const unused = [...htmlIds].filter((id) => !popupJs.includes(`'${id}'`) && !popupJs.includes(`"${id}"`));
if (unused.length) notes.push(`  INFO ids declared in HTML but unused in JS: ${unused.join(', ')}`);

/* ------------------------ 3b. cross-check CSS classes -------------------- */
const css = read('popup/popup.css');
const definedClasses = new Set([...css.matchAll(/\.(-?[_a-zA-Z][\w-]*)/g)].map((m) => m[1]));

const usedClasses = new Set();
for (const source of [html, popupJs]) {
  for (const m of source.matchAll(/class="([^"$]+)"/g)) {
    for (const cls of m[1].split(/\s+/)) if (cls) usedClasses.add(cls);
  }
  // shapes like: className = 'row' + (x ? ' is-picked' : '') / classList.toggle('is-x')
  for (const m of source.matchAll(/classList\.(?:add|toggle|remove)\('([\w-]+)'/g)) {
    usedClasses.add(m[1]);
  }
}

const undefinedClasses = [...usedClasses].filter((c) => !definedClasses.has(c));
if (undefinedClasses.length) {
  bad(`classes used in markup but absent from popup.css: ${undefinedClasses.join(', ')}`);
} else {
  ok(`all ${usedClasses.size} classes are defined in popup.css`);
}

/* -------------------------- 3c. both themes exist ------------------------ */
if (/\[data-theme="light"\]/.test(css) && /^:root\s*\{/m.test(css)) {
  ok('popup.css defines both a default (dark) and a light theme');
} else {
  bad('popup.css is missing one of the two themes');
}
if (/color-scheme/.test(css)) ok('color-scheme is declared so native controls follow the theme');
else bad('color-scheme is missing — date pickers and selects will not match the theme');

if (/prefers-reduced-motion/.test(css)) ok('animations respect prefers-reduced-motion');
else bad('no prefers-reduced-motion guard around the animations');

/* -------------------- 3d. permissions reachable everywhere --------------- */
// The permission drawer must be offered on all three card kinds — eligible,
// active and permanently assigned — not just on roles you can switch on.
const permButtons = (popupJs.match(/\$\{permButton\(key\)\}/g) || []).length;
if (permButtons >= 3) {
  ok(`the permission button is rendered on all ${permButtons} card kinds`);
} else {
  bad(`the permission button appears on only ${permButtons} card kind(s); expected 3`);
}
const drawers = (popupJs.match(/\$\{drawerFor\(key\)\}/g) || []).length;
if (drawers >= 3) ok('every card kind carries a detail drawer');
else bad(`only ${drawers} card kind(s) carry a detail drawer; expected 3`);

/* --------------------- 3e. the token vault UI is really gone -------------- */
// Tokens are still captured — they are what talks to Graph — but nothing may
// display, list or copy one. That was the whole point of dropping the vault.
if (/panelVault|vaultBody|Token vault|data-copy=/.test(html + popupJs)) {
  bad('remnants of the token vault UI are still in the popup');
} else {
  ok('the popup has no token vault UI left');
}
if (/navigator\.clipboard/.test(popupJs)) {
  bad('the popup still copies something to the clipboard — a token must never be copyable');
} else {
  ok('the popup never touches the clipboard');
}

/* --------------------- 3f. every panel the menu opens exists -------------- */
const panels = [...popupJs.matchAll(/openOverlay\('([A-Za-z0-9_]+)'\)/g)].map((m) => m[1]);
const closed = read('popup/popup.js').match(/for \(const id of \[([^\]]+)\]\)/);
const closedIds = closed ? [...closed[1].matchAll(/'([A-Za-z0-9_]+)'/g)].map((m) => m[1]) : [];
for (const panel of new Set(panels)) {
  if (!htmlIds.has(panel)) bad(`openOverlay('${panel}') but no element carries that id`);
  else if (!closedIds.includes(panel)) bad(`panel '${panel}' is opened but closeOverlays() never hides it`);
  else ok(`panel '${panel}' exists and is closed again by closeOverlays()`);
}

/* --------------------- 4. cross-check messages both ways ----------------- */
const background = read('src/background.js');
const handled = new Set([...background.matchAll(/case '([a-z]+\/[a-z-]+)'/gi)].map((m) => m[1]));
const sentFromPopup = new Set([...popupJs.matchAll(/ask\('([a-z]+\/[a-z-]+)'/gi)].map((m) => m[1]));
const sentFromContent = new Set(
  [...read('src/content.js').matchAll(/type: '([a-z]+\/[a-z-]+)'/gi)].map((m) => m[1])
);

for (const type of [...sentFromPopup, ...sentFromContent]) {
  if (handled.has(type)) ok(`the worker handles message '${type}'`);
  else bad(`nothing handles message '${type}'`);
}

/* ------------------- 4b. cross-check the portal registry ----------------- */
// The worker owns the URLs; the popup only sends an id. Both lists must agree,
// or a menu entry would open nothing.
const workerPortals = new Set(
  [...(background.match(/const PORTALS = \{[\s\S]*?\n\};/) || [''])[0]
    .matchAll(/^\s{2}([a-z0-9]+):\s*'https:/gm)].map((m) => m[1])
);
const popupPortals = [...popupJs.matchAll(/id: '([a-z0-9]+)', name:/g)].map((m) => m[1]);

if (!workerPortals.size) bad('no portal registry found in the worker');
for (const id of popupPortals) {
  if (workerPortals.has(id)) ok(`portal '${id}' resolves to a URL in the worker`);
  else bad(`the popup offers portal '${id}' but the worker has no URL for it`);
}

// The admin centres the popup is expected to reach.
for (const host of [
  'entra.microsoft.com', 'portal.azure.com', 'admin.microsoft.com',
  'security.microsoft.com', 'purview.microsoft.com',
]) {
  if (background.includes(`https://${host}/`)) ok(`portal reachable: ${host}`);
  else bad(`no portal entry for ${host}`);
}

/* ------------------------- 5. security: no leaks ------------------------- */
/** Strip comments so the scans do not trip over the notes themselves. */
const stripComments = (code) =>
  code
    .replace(/\/\*[\s\S]*?\*\//g, '')
    .replace(/(^|[^:'"\\])\/\/.*$/gm, '$1')
    .replace(/<!--[\s\S]*?-->/g, '');

const allSource = stripComments(jsFiles.map(read).join('\n') + html);
const backgroundCode = stripComments(background);

// Every http(s) URL in the source must point at a Microsoft host.
const allowedHosts = [
  'graph.microsoft.com',
  'portal.azure.com',
  'entra.microsoft.com',
  'intune.microsoft.com',
  'endpoint.microsoft.com',
  'manage.microsoft.com',
  'api.manage.microsoft.com',
  'admin.microsoft.com',
  'security.microsoft.com',
  'purview.microsoft.com',
  'www.w3.org',
];
const urls = [...allSource.matchAll(/https?:\/\/([a-z0-9.-]+)/gi)].map((m) => m[1].toLowerCase());
const foreign = [...new Set(urls)].filter((h) => !allowedHosts.includes(h));
if (foreign.length) bad(`URLs pointing at foreign domains: ${foreign.join(', ')}`);
else ok('no endpoints outside Microsoft');

if (/chrome\.storage\.sync/.test(allSource)) bad('storage.sync is used — data would sync to the cloud');
else ok('storage.sync is never used');

// The vault must live in storage.session (RAM) and nowhere else.
if (/storage\.session\.set/.test(backgroundCode)) ok('the vault is written to storage.session (RAM only)');
else bad('the vault is never written to storage.session');

// No code path may put a token on disk, under any preference.
const diskVaultWrites = [...allSource.matchAll(/storage\.local\.set\(\{\s*\[VAULT_KEY\]/g)].length;
if (diskVaultWrites === 0) ok('the vault is never written to storage.local (tokens never touch the disk)');
else bad(`${diskVaultWrites} code path(s) write the vault to storage.local`);

if (/storage\.local\.get\(VAULT_KEY\)/.test(allSource)) {
  bad('the vault is read back from storage.local — tokens must never come off the disk');
} else {
  ok('the vault is never read from storage.local');
}

// The old opt-in switch must be gone entirely, not merely defaulted to false.
if (/persistTokens/.test(allSource.replace(/'persistTokens' in prefs|delete prefs\.persistTokens/g, ''))) {
  bad('persistTokens still appears outside the legacy-cleanup path');
} else {
  ok('the persistTokens option is fully removed');
}

if (/storage\.local\.remove\(VAULT_KEY\)/.test(backgroundCode)) {
  ok('any vault left on disk by an older build is purged on startup');
} else {
  bad('no cleanup of a vault written to disk by an older build');
}

if (/console\.log/.test(allSource)) bad('console.log remains — it could print a token to devtools');
else ok('no console.log anywhere');

/* --------------------------------- result -------------------------------- */
console.log(notes.join('\n'));
console.log('\n' + '='.repeat(70));
if (problems.length) {
  console.log(problems.join('\n'));
  console.log(`\n${problems.length} problem(s) to fix.`);
  process.exit(1);
}
console.log('All checks passed.');
