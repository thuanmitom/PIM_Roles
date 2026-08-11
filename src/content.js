/* ---------------------------------------------------------------------------
 * content.js — isolated world; the bridge between inject.js and the worker.
 * ------------------------------------------------------------------------- */

'use strict';

window.addEventListener(
  'message',
  (event) => {
    // Only accept messages this page posted to itself.
    if (event.source !== window) return;

    const data = event.data;
    if (!data || data.__pimConsole !== true || data.kind !== 'token') return;
    if (typeof data.token !== 'string' || data.token.split('.').length !== 3) return;

    try {
      chrome.runtime
        .sendMessage({ type: 'token/captured', token: data.token, url: data.url })
        .catch(() => {
          /* worker asleep or extension just reloaded — nothing to do */
        });
    } catch {
      /* context invalidated after an extension reload */
    }
  },
  false
);

// Let the worker know a portal tab is open (used by the scheduler).
try {
  chrome.runtime.sendMessage({ type: 'portal/alive' }).catch(() => {});
} catch {}
