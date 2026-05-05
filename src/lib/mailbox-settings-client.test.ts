import { describe, expect, it } from 'bun:test';
import { normalizeWorkingHourTime, parseWorkDaysCsv } from './mailbox-settings-client.js';

describe('parseWorkDaysCsv', () => {
  it('maps short and long day names', () => {
    expect(parseWorkDaysCsv('mon,wed,Fri')).toEqual(['monday', 'wednesday', 'friday']);
  });
});

describe('normalizeWorkingHourTime', () => {
  it('formats HH:mm with seconds fraction', () => {
    expect(normalizeWorkingHourTime('9:00')).toBe('09:00:00.0000000');
    expect(normalizeWorkingHourTime('17:30')).toBe('17:30:00.0000000');
  });
});
