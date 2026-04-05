/**
 * Working-days arithmetic.
 *
 * When skipWeekends is true, Saturdays and Sundays are excluded from
 * duration counts and date calculations. If a start date falls on a
 * weekend, it is snapped forward to the next Monday.
 */

function isWeekend(date) {
  const day = date.getDay();
  return day === 0 || day === 6;
}

function toDate(value) {
  if (value instanceof Date) return new Date(value);
  const s = String(value);
  const parts = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (parts) return new Date(Number(parts[1]), Number(parts[2]) - 1, Number(parts[3]), 12);
  return new Date(s);
}

function toIso(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function snapToWorkday(d) {
  while (isWeekend(d)) {
    d.setDate(d.getDate() + 1);
  }
  return d;
}

/**
 * Add `days` working days to `startDate`.
 * Returns an ISO date string (YYYY-MM-DD).
 */
export function addWorkingDays(startDate, days, skipWeekends) {
  if (!startDate || !Number.isFinite(days) || days <= 0) return null;

  const d = toDate(startDate);
  if (isNaN(d.getTime())) return null;

  if (!skipWeekends) {
    d.setDate(d.getDate() + days - 1);
    return toIso(d);
  }

  snapToWorkday(d);
  let remaining = days - 1;

  while (remaining > 0) {
    d.setDate(d.getDate() + 1);
    if (!isWeekend(d)) remaining--;
  }

  return toIso(d);
}

/**
 * Count the number of working days between two dates (inclusive).
 * Returns an integer. If either date falls on a weekend when
 * skipWeekends is true, it is snapped to the nearest working day first.
 */
export function workingDaysBetween(startDate, endDate, skipWeekends) {
  if (!startDate || !endDate) return 0;

  const a = toDate(startDate);
  const b = toDate(endDate);
  if (isNaN(a.getTime()) || isNaN(b.getTime())) return 0;

  if (!skipWeekends) {
    const diff = Math.round((b - a) / (1000 * 60 * 60 * 24)) + 1;
    return Math.max(diff, 0);
  }

  snapToWorkday(a);
  if (b < a) return 0;

  let count = 0;
  const cursor = new Date(a);
  while (cursor <= b) {
    if (!isWeekend(cursor)) count++;
    cursor.setDate(cursor.getDate() + 1);
  }

  return Math.max(count, 0);
}
