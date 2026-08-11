/* Decode the JWT payload. The signature is NOT verified — this only reads claims. */

export function decodeJwt(token) {
  try {
    const parts = String(token || '').split('.');
    if (parts.length !== 3) return null;

    const base64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
    const padded = base64 + '='.repeat((4 - (base64.length % 4)) % 4);
    const binary = atob(padded);

    // Decode as UTF-8 so accented display names survive intact.
    const bytes = Uint8Array.from(binary, (ch) => ch.charCodeAt(0));
    return JSON.parse(new TextDecoder('utf-8').decode(bytes));
  } catch {
    return null;
  }
}

/** Seconds left before the token expires (negative once it has). */
export function secondsUntilExpiry(claims) {
  if (!claims || !claims.exp) return null;
  return claims.exp - Math.floor(Date.now() / 1000);
}

export function isExpired(claims, skewSeconds = 30) {
  const left = secondsUntilExpiry(claims);
  return left === null ? false : left <= skewSeconds;
}

/** Pull out just the claims worth displaying. */
export function summarize(token) {
  const claims = decodeJwt(token);
  if (!claims) return null;
  return {
    principalId: claims.oid || '',
    tenantId: claims.tid || '',
    upn: claims.upn || claims.unique_name || claims.preferred_username || '',
    name: claims.name || '',
    audience: claims.aud || '',
    scopes: claims.scp || (Array.isArray(claims.roles) ? claims.roles.join(' ') : ''),
    expiresAt: claims.exp ? claims.exp * 1000 : null,
  };
}

/* -------------------------- token classification ------------------------- */

// Well-known app IDs, used when `aud` is a GUID rather than a URL.
const APPID = {
  graph: '00000003-0000-0000-c000-000000000000',
  intune: '0000000a-0000-0000-c000-000000000000',
};

const ROLE_SCOPE = /role(?:assignment|eligibility)?schedule|rolemanagement|roleassignment/i;
const INTUNE_SCOPE = /devicemanagement/i;

/**
 * Work out what a token is, so the vault knows whether PIM can use it.
 *   resource:   'graph' | 'intune' | 'other'
 *   scopeClass: 'role' (carries PIM scopes) | 'devicemgmt' | 'general'
 *   label:      the string shown to the user
 */
export function classifyToken(claims) {
  if (!claims) return { resource: 'other', scopeClass: 'general', label: 'Unknown' };

  const aud = String(claims.aud || '').toLowerCase();
  const scopes = String(
    claims.scp || (Array.isArray(claims.roles) ? claims.roles.join(' ') : '')
  ).toLowerCase();

  let resource = 'other';
  if (aud.includes('graph.microsoft.com') || aud.startsWith(APPID.graph)) resource = 'graph';
  else if (aud.includes('manage.microsoft.com') || aud.startsWith(APPID.intune)) resource = 'intune';

  let scopeClass = 'general';
  if (ROLE_SCOPE.test(scopes)) scopeClass = 'role';
  else if (INTUNE_SCOPE.test(scopes)) scopeClass = 'devicemgmt';

  let label;
  if (resource === 'graph') {
    label =
      scopeClass === 'role'
        ? 'Microsoft Graph · PIM'
        : scopeClass === 'devicemgmt'
        ? 'Microsoft Graph · Intune'
        : 'Microsoft Graph';
  } else if (resource === 'intune') {
    label = 'Intune API';
  } else {
    const host = aud.replace(/^https?:\/\//, '').split('/')[0];
    label = host || 'Microsoft';
  }

  return { resource, scopeClass, label };
}

/** Unique vault key per token kind, so a PIM token never overwrites an Intune one. */
export function vaultKey(summary, kind) {
  return `${kind.resource}|${summary.tenantId || 'notenant'}|${kind.scopeClass}`;
}
