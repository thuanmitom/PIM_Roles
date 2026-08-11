/* Duration and time helpers shared by the popup and the service worker. */

const ISO_DURATION =
  /^P(?!$)(?:(\d+)D)?(?:T(?!$)(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/;

/** 'PT8H30M' -> milliseconds. Returns null when the input is not valid. */
export function parseIsoDuration(value) {
  if (!value) return null;
  const m = ISO_DURATION.exec(String(value).trim().toUpperCase());
  if (!m) return null;
  const [, d, h, mi, s] = m.map((x) => (x === undefined ? 0 : x));
  return (
    (Number(d) * 86400 + Number(h) * 3600 + Number(mi) * 60 + Number(s)) * 1000
  );
}

/** Milliseconds -> 'PT8H30M'. */
export function toIsoDuration(ms) {
  let total = Math.max(0, Math.floor(ms / 1000));
  const days = Math.floor(total / 86400);
  total -= days * 86400;
  const hours = Math.floor(total / 3600);
  total -= hours * 3600;
  const minutes = Math.floor(total / 60);
  const seconds = total - minutes * 60;

  let out = 'P';
  if (days) out += `${days}D`;
  let time = '';
  if (hours) time += `${hours}H`;
  if (minutes) time += `${minutes}M`;
  if (seconds) time += `${seconds}S`;
  if (time) out += 'T' + time;
  return out === 'P' ? 'PT0S' : out;
}

/** Accepts '8h', '90m', '2h30m', 'PT4H' -> an ISO 8601 string, or null. */
export function normalizeDuration(raw) {
  const text = String(raw || '').trim().toUpperCase();
  if (!text) return null;
  if (text.startsWith('P')) return parseIsoDuration(text) === null ? null : text;

  const m = /^(?:(\d+)D)?(?:(\d+)H)?(?:(\d+)M)?$/.exec(text);
  if (!m || !m.slice(1).some(Boolean)) return null;
  const ms =
    (Number(m[1] || 0) * 86400 + Number(m[2] || 0) * 3600 + Number(m[3] || 0) * 60) *
    1000;
  return ms > 0 ? toIsoDuration(ms) : null;
}

/** 8100000 -> '2h 15m'. */
export function humanize(ms) {
  if (ms === null || ms === undefined) return '—';
  let total = Math.floor(ms / 1000);
  if (total <= 0) return 'expired';

  const days = Math.floor(total / 86400);
  total -= days * 86400;
  const hours = Math.floor(total / 3600);
  total -= hours * 3600;
  const minutes = Math.floor(total / 60);

  const parts = [];
  if (days) parts.push(`${days}d`);
  if (hours) parts.push(`${hours}h`);
  if (minutes || !parts.length) parts.push(`${minutes}m`);
  return parts.join(' ');
}

/** 8100000 -> '02:15:00'  (used by the countdown clocks). */
export function clock(ms) {
  const total = Math.max(0, Math.floor(ms / 1000));
  const h = String(Math.floor(total / 3600)).padStart(2, '0');
  const m = String(Math.floor((total % 3600) / 60)).padStart(2, '0');
  const s = String(total % 60).padStart(2, '0');
  return `${h}:${m}:${s}`;
}

export function parseGraphDate(value) {
  if (!value) return null;
  const t = Date.parse(value);
  return Number.isNaN(t) ? null : t;
}

const DATE_FMT = new Intl.DateTimeFormat('en-GB', {
  day: '2-digit',
  month: 'short',
  hour: '2-digit',
  minute: '2-digit',
  hour12: false,
});

const TIME_FMT = new Intl.DateTimeFormat('en-GB', {
  hour: '2-digit',
  minute: '2-digit',
  hour12: false,
});

export function formatDateTime(ms) {
  if (!ms) return 'no expiry';
  return DATE_FMT.format(new Date(ms));
}

export function formatTime(ms) {
  if (!ms) return '—';
  return TIME_FMT.format(new Date(ms));
}

const DAY_FMT = new Intl.DateTimeFormat('en-GB', {
  weekday: 'short',
  day: '2-digit',
  month: 'short',
});

/** 'YYYY-M-D' in local time — the grouping key for the history list. */
export function dayKey(ms) {
  const d = new Date(ms);
  return `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}`;
}

/** A day heading: 'Today', 'Yesterday', else 'Tue, 04 Aug'. */
export function formatDay(ms) {
  if (!ms) return '—';
  const key = dayKey(ms);
  const now = Date.now();
  if (key === dayKey(now)) return 'Today';
  if (key === dayKey(now - 86400000)) return 'Yesterday';
  return DAY_FMT.format(new Date(ms));
}

/** Turn 'HH:MM' into the next timestamp at which it occurs (today or tomorrow). */
export function nextOccurrence(hhmm) {
  const m = /^(\d{1,2}):(\d{2})$/.exec(String(hhmm).trim());
  if (!m) return null;
  const hour = Number(m[1]);
  const minute = Number(m[2]);
  if (hour > 23 || minute > 59) return null;

  const target = new Date();
  target.setSeconds(0, 0);
  target.setHours(hour, minute);
  if (target.getTime() <= Date.now()) target.setDate(target.getDate() + 1);
  return target.getTime();
}

/** Turn the value of an <input type="datetime-local"> into a timestamp. */
export function fromDatetimeLocal(value) {
  if (!value) return null;
  const t = new Date(value).getTime();
  return Number.isNaN(t) ? null : t;
}
