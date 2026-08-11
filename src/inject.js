/* ---------------------------------------------------------------------------
 * inject.js  —  runs in the MAIN world of the Microsoft portals
 *
 * Its only job: read the Authorization header the portal itself sends when it
 * calls Microsoft Graph, and hand it to the isolated world via postMessage.
 *
 * It sends data NOWHERE else. It does not touch the DOM, cookies or
 * localStorage, and only runs on the hosts listed in the manifest.
 * ------------------------------------------------------------------------- */

(() => {
  'use strict';

  const FLAG = '__pimConsoleHooked__';
  if (window[FLAG]) return;
  try {
    Object.defineProperty(window, FLAG, { value: true, enumerable: false });
  } catch {
    window[FLAG] = true;
  }

  // Capture tokens sent to the Microsoft APIs these portals use.
  //   graph.microsoft.com   -> PIM, Intune (via Graph) and most operations
  //   manage.microsoft.com  -> the dedicated Intune API
  const WATCHED_HOSTS = ['graph.microsoft.com', 'manage.microsoft.com'];

  // Remember what has been forwarded so the worker is not spammed.
  const alreadySent = new Set();

  function isWatched(url) {
    if (typeof url !== 'string') return false;
    return WATCHED_HOSTS.some((host) => url.includes(host));
  }

  function publish(rawAuthorization, url) {
    if (!rawAuthorization) return;

    const token = String(rawAuthorization).replace(/^\s*bearer\s+/i, '').trim();
    // A well-formed JWT has exactly three dot-separated parts.
    if (token.split('.').length !== 3) return;

    const fingerprint = token.slice(-32);
    if (alreadySent.has(fingerprint)) return;
    alreadySent.add(fingerprint);
    if (alreadySent.size > 50) alreadySent.clear();

    window.postMessage(
      { __pimConsole: true, kind: 'token', token, url: String(url).split('?')[0] },
      location.origin
    );
  }

  /** Read one header out of any of the shapes fetch() accepts. */
  function readHeader(headers, name) {
    if (!headers) return '';
    try {
      if (typeof Headers !== 'undefined' && headers instanceof Headers) {
        return headers.get(name) || '';
      }
      if (Array.isArray(headers)) {
        const found = headers.find((pair) => String(pair[0]).toLowerCase() === name);
        return found ? found[1] : '';
      }
      const key = Object.keys(headers).find((k) => k.toLowerCase() === name);
      return key ? headers[key] : '';
    } catch {
      return '';
    }
  }

  /* ---------------------------- hook fetch() ---------------------------- */
  const nativeFetch = window.fetch;
  if (typeof nativeFetch === 'function') {
    window.fetch = function (input, init) {
      try {
        const url =
          typeof input === 'string'
            ? input
            : input && typeof input.url === 'string'
            ? input.url
            : '';

        if (isWatched(url)) {
          let auth = readHeader(init && init.headers, 'authorization');
          if (!auth && typeof Request !== 'undefined' && input instanceof Request) {
            auth = input.headers.get('authorization') || '';
          }
          publish(auth, url);
        }
      } catch {
        /* never break the portal's own request */
      }
      return nativeFetch.apply(this, arguments);
    };
    // Keep the original toString() so the portal does not see a patched fetch.
    try {
      window.fetch.toString = () => nativeFetch.toString();
    } catch {}
  }

  /* ------------------------ hook XMLHttpRequest ------------------------- */
  const XHR = XMLHttpRequest.prototype;
  const nativeOpen = XHR.open;
  const nativeSetHeader = XHR.setRequestHeader;
  const nativeSend = XHR.send;

  XHR.open = function (method, url) {
    try {
      this.__pimUrl = url;
      this.__pimAuth = '';
    } catch {}
    return nativeOpen.apply(this, arguments);
  };

  XHR.setRequestHeader = function (name, value) {
    try {
      if (String(name).toLowerCase() === 'authorization') this.__pimAuth = value;
    } catch {}
    return nativeSetHeader.apply(this, arguments);
  };

  XHR.send = function () {
    try {
      if (isWatched(this.__pimUrl)) publish(this.__pimAuth, this.__pimUrl);
    } catch {}
    return nativeSend.apply(this, arguments);
  };
})();
