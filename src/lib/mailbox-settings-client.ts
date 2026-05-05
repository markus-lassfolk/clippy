import { readFile } from 'node:fs/promises';
import { callGraph, GraphApiError, graphError } from './graph-client.js';
import { graphUserPath } from './graph-user-path.js';

/** Subset of Graph [mailboxSettings](https://learn.microsoft.com/en-us/graph/api/resources/mailboxsettings) for read/patch. */
export interface MailboxSettingsWorkingHours {
  daysOfWeek?: string[];
  startTime?: string;
  endTime?: string;
  timeZone?: { name?: string };
}

export interface MailboxSettingsFull {
  automaticRepliesSetting?: unknown;
  timeZone?: string;
  dateFormat?: string;
  timeFormat?: string;
  workingHours?: MailboxSettingsWorkingHours;
  archiveFolder?: string;
  language?: unknown;
}

export async function getMailboxSettingsFull(
  token: string,
  user?: string
): Promise<{
  ok: boolean;
  data?: MailboxSettingsFull;
  error?: { message: string; code?: string; status?: number };
}> {
  try {
    return await callGraph<MailboxSettingsFull>(token, graphUserPath(user, 'mailboxSettings'));
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status) as {
        ok: boolean;
        data?: MailboxSettingsFull;
        error?: { message: string; code?: string; status?: number };
      };
    }
    return graphError(err instanceof Error ? err.message : 'Failed to get mailbox settings') as {
      ok: boolean;
      data?: MailboxSettingsFull;
      error?: { message: string; code?: string; status?: number };
    };
  }
}

export async function patchMailboxSettings(
  token: string,
  patch: Record<string, unknown>,
  user?: string
): Promise<{ ok: boolean; error?: { message: string; code?: string; status?: number } }> {
  try {
    const result = await callGraph<Record<string, never>>(
      token,
      graphUserPath(user, 'mailboxSettings'),
      {
        method: 'PATCH',
        body: JSON.stringify(patch)
      },
      false
    );
    if (!result.ok) {
      return { ok: false, error: result.error || { message: 'Failed to patch mailbox settings' } };
    }
    return { ok: true };
  } catch (err) {
    if (err instanceof GraphApiError) {
      return { ok: false, error: { message: err.message, code: err.code, status: err.status } };
    }
    return { ok: false, error: { message: err instanceof Error ? err.message : 'Failed to patch mailbox settings' } };
  }
}

export async function readJsonPatchFile(path: string): Promise<Record<string, unknown>> {
  const raw = await readFile(path, 'utf8');
  const parsed = JSON.parse(raw) as unknown;
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    throw new Error('JSON file must contain a single object');
  }
  return parsed as Record<string, unknown>;
}

const DAY_MAP: Record<string, string> = {
  sun: 'sunday',
  sunday: 'sunday',
  mon: 'monday',
  monday: 'monday',
  tue: 'tuesday',
  tuesday: 'tuesday',
  wed: 'wednesday',
  wednesday: 'wednesday',
  thu: 'thursday',
  thursday: 'thursday',
  fri: 'friday',
  friday: 'friday',
  sat: 'saturday',
  saturday: 'saturday'
};

/** Parse "mon,tue,wed" into Graph `daysOfWeek` values. */
export function parseWorkDaysCsv(csv: string): string[] {
  const out: string[] = [];
  for (const part of csv.split(',')) {
    const k = part.trim().toLowerCase();
    if (!k) continue;
    const full = DAY_MAP[k];
    if (full && !out.includes(full)) {
      out.push(full);
    }
  }
  return out;
}

/** Graph expects `HH:mm:ss.0000000` for working hour times. */
export function normalizeWorkingHourTime(hhmm: string): string {
  const t = hhmm.trim();
  const parts = t.split(':');
  if (parts.length < 2) {
    throw new Error(`Invalid time (use HH:mm): ${hhmm}`);
  }
  const h = parseInt(parts[0]!, 10);
  const m = parseInt(parts[1]!, 10);
  if (!Number.isFinite(h) || !Number.isFinite(m) || h < 0 || h > 23 || m < 0 || m > 59) {
    throw new Error(`Invalid time (use HH:mm): ${hhmm}`);
  }
  const secPart = parts[2]?.split('.')[0];
  const s = secPart !== undefined ? parseInt(secPart, 10) : 0;
  if (!Number.isFinite(s) || s < 0 || s > 59) {
    throw new Error(`Invalid time (use HH:mm or HH:mm:ss): ${hhmm}`);
  }
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}.0000000`;
}
