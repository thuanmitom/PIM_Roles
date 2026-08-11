# Entra PIM Roles

A Chrome/Edge (MV3) extension to **activate, deactivate and schedule Entra ID PIM
roles**, review **what you activated and when**, see **every permission you hold**, and
jump straight to any Microsoft admin portal — all from the toolbar. It talks to
Microsoft Graph directly, so there is no token to copy and paste by hand.

The popup is a small dashboard: status stripes and badges per role, live countdowns on
whatever is switched on, and a full light/dark theme that follows your system by default.

> ⚠️ **This is an administration tool for your own tenant.** It works by reading the
> `Authorization` header the Microsoft portals already send — the same technique
> token-stealing malware uses. So **do not publish it to the Chrome Web Store or Edge
> Add-ons**, and do not install it for anyone else. Load it unpacked, on your own machine.

---

## Install

1. Unpack the **whole** folder to disk (do not load it from inside the Windows zip
   viewer — Chrome will report missing files).
2. Open `chrome://extensions` (or `edge://extensions`) and turn on **Developer mode**.
3. Click **Load unpacked** and pick the folder.
4. Open the PIM blade so it finishes loading your roles:
   `https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadmigratedroles`
   (the same blade in `entra.microsoft.com` works too)
5. Click the extension icon — your account, the counters and the role list appear.

---

## What it does

**Role list** — one card per role, with a status badge:

- `ELIGIBLE` (pink): can be switched on. **Activate** opens a duration + reason form
  inside the card; the calendar button schedules that one role.
- `ACTIVE` (amber): switched on right now. Live countdown, a draining time bar, and
  **Deactivate**.
- `ASSIGNED` (slate): permanent, never passes through PIM.

**Permissions, in every state** — every card, whatever its status, has a
**Permissions** button. It opens a drawer with the role description and the granular
actions it grants, grouped by provider, with a *show all* toggle for long lists. You do
not have to switch a role off to look at what it can do.

**Bulk activation with per-role reasons** — tick several eligible roles and the action
bar appears. Type one **shared reason** for the whole batch, or open **Reason per role**
to give each role its own justification; anything left blank falls back to the shared
one. Roles whose policy demands a justification are marked with `*`, and activation is
blocked until each of them has one.

**Scheduling, including for roles that are already on** — the calendar button is on
active cards too, and the schedule picker lists active roles alongside eligible ones
(marked *on until …*). An active role still expires, so scheduling the next window is
normal; just pick a start time after the current one ends, or PIM rejects the request as
overlapping. Two modes:

- *PIM (server side)* — submitted now with a future start time; it fires even with your
  machine off.
- *This browser* — the extension waits for the moment and can repeat daily; the browser
  must be running and the captured session still valid.

**Stat cards** — total roles · eligible · active.

**Activation history** — the last 80 PIM requests this account made, newest first,
grouped by day (*Today*, *Yesterday*, then the date). Each row shows what happened
(activated, deactivated, extended, or an admin acting on you), how the request ended
(*Succeeded*, *Awaiting approval*, *Denied*, *Revoked* — failures in red), the duration
asked for, the justification and the ticket number. Two filters narrow it to activations
or deactivations only, and the reload button re-reads it from Graph.

**My permissions** — every granular action your roles grant you, folded into one list
and grouped by provider (`microsoft.directory`, `microsoft.intune`, …). Two scopes:

- *In force now* — what the roles switched on right now, plus your permanent
  assignments, actually let you do.
- *Incl. eligible* — the same plus everything an eligible role would add. Actions only
  an eligible role grants are dimmed, so it is obvious what you would have to activate
  first.

Tap a provider to unfold its actions; each one carries a counter of how many of your
roles grant it and names them on hover. The search box filters across every provider at
once, which is the fast way to answer *"can I reset a password right now?"*.

**Admin portals** — one panel with a shortcut to each console, reusing the tab you
already have open on that host:

| Portal | Host |
| --- | --- |
| Privileged Identity Management | `portal.azure.com` (activation blade) |
| Microsoft Entra admin center | `entra.microsoft.com` |
| Azure portal | `portal.azure.com` |
| Microsoft 365 admin center | `admin.microsoft.com` |
| Microsoft Defender | `security.microsoft.com` |
| Microsoft Purview | `purview.microsoft.com` |
| Microsoft Intune | `intune.microsoft.com` |

The ones marked *session* are the three the extension can pick a session up from; the
rest are plain shortcuts.

**Session countdown** — next to your account name in the header, the time the captured
session still has to live. It turns red in the last quarter of an hour: that is the cue
to reopen a portal before Graph starts refusing calls.

**Policy awareness** — reads `roleManagementPolicyAssignments` to show the constraints
per role (maximum duration, MFA, justification, ticket, approval) and automatically
shortens a requested duration that exceeds the cap.

**Theme** — the sun/moon button switches light and dark; until you pick one, the popup
follows the operating system.

---

## Where do the tokens live?

**In memory, and nowhere else.** The captured session is kept in
`chrome.storage.session`, which is RAM: it is gone when the browser closes. There is no
option to persist tokens, no export, and no write to disk — earlier builds had an opt-in
"keep across sessions" switch, and it has been removed. Anything a previous version left
on disk is deleted on startup.

**There is no longer any way to read a token out of the extension.** The old *token
vault* panel — which listed every captured token with a *Copy token* button — is gone,
along with the messages that served it. The worker hands the popup exactly one token, for
its own Graph reads, and nothing renders it. All you see of the session is who it belongs
to and how long it lasts.

Which token is used for PIM? The extension picks the Graph token carrying role-management
scopes (`RoleAssignmentSchedule…`). If you only opened the Intune portal, the captured
token will be a device-management one, and PIM operations will report missing permissions
until you open the PIM blade again.

---

## Security boundary (enforced by `validate.mjs`)

- Scripts are injected only into `portal.azure.com`, `entra.microsoft.com`,
  `intune.microsoft.com` and `endpoint.microsoft.com`. The other portals in the shortcut
  list are only ever *opened* — nothing is injected into them.
- Only tokens sent to `graph.microsoft.com` and `manage.microsoft.com` are captured.
- The extension **only calls out to** `graph.microsoft.com` (CSP `connect-src`) — no
  foreign endpoints, no `storage.sync`, no `console.log` (so a token can never be printed
  to devtools).
- No token is ever displayed, listed or copied: the validator fails the build if the
  popup regrows a vault panel or touches the clipboard.
- The session store is never written to, or read from, `storage.local`.
- The popup only ever sends a portal *id*; the worker owns the URLs, so a message cannot
  navigate a tab anywhere else.
- Tokens never leave your browser.

---

## Tests

```bash
node validate.mjs      # structure + security constraints
node test.mjs          # 34 logic tests (Graph, policies, history, permission index)
node test-inject.mjs   # 12 tests for the token-capture hook
```

Regenerate the shield icons after editing the generator:

```bash
node tools/make-icons.mjs
```

## Layout

```
manifest.json           MV3
src/inject.js           fetch/XHR hook (MAIN world) — reads the Authorization header
src/content.js          relay to the service worker (ISOLATED world)
src/background.js       session store, activate/deactivate, schedules, portals, badge
src/lib/graph.js        Graph client, policy reader, history + permission shaping
src/lib/jwt.js          JWT decoding, token classification, store keys
src/lib/fmt.js          ISO 8601 durations, clocks, day grouping, formatting
popup/                  the UI (html/css/js)
tools/make-icons.mjs    dependency-free PNG generator for the shield icons
icons/                  icons, 16–128
```
