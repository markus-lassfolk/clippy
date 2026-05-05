import type { FindMeetingTimesRequest } from './graph-schedule.js';

/** Parse `displayName`, `displayName|room@x.com`, or `room@x.com` for findMeetingTimes locationConstraint.locations. */
export function parseMeetingLocationSpecs(
  specs: string[],
  resolveAvailability: boolean
): NonNullable<NonNullable<FindMeetingTimesRequest['locationConstraint']>['locations']> {
  const out: NonNullable<NonNullable<FindMeetingTimesRequest['locationConstraint']>['locations']> = [];
  for (const raw of specs) {
    const s = raw.trim();
    if (!s) continue;
    const pipe = s.indexOf('|');
    let displayName: string | undefined;
    let locationEmailAddress: string | undefined;
    if (pipe >= 0) {
      displayName = s.slice(0, pipe).trim() || undefined;
      locationEmailAddress = s.slice(pipe + 1).trim() || undefined;
    } else if (s.includes('@')) {
      locationEmailAddress = s;
    } else {
      displayName = s;
    }
    out.push({
      ...(resolveAvailability ? { resolveAvailability: true } : {}),
      ...(displayName ? { displayName } : {}),
      ...(locationEmailAddress ? { locationEmailAddress } : {})
    });
  }
  return out;
}

export function buildFindMeetingTimesLocationConstraint(opts: {
  suggestLocations?: boolean;
  locationRequired?: boolean;
  resolveLocationAvailability?: boolean;
  meetingLocation?: string[];
}): FindMeetingTimesRequest['locationConstraint'] | undefined {
  const locations = parseMeetingLocationSpecs(opts.meetingLocation ?? [], opts.resolveLocationAvailability ?? false);
  const hasConstraint = Boolean(opts.suggestLocations) || Boolean(opts.locationRequired) || locations.length > 0;
  if (!hasConstraint) return undefined;
  return {
    ...(opts.locationRequired ? { isRequired: true } : {}),
    ...(opts.suggestLocations ? { suggestLocation: true } : {}),
    ...(locations.length > 0 ? { locations } : {})
  };
}
