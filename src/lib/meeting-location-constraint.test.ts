import { describe, expect, it } from 'bun:test';
import { buildFindMeetingTimesLocationConstraint, parseMeetingLocationSpecs } from './meeting-location-constraint.js';

describe('parseMeetingLocationSpecs', () => {
  it('parses display name, email-only, and pipe form', () => {
    expect(parseMeetingLocationSpecs(['Room A', 'r1@x.com', 'Conf B|r2@x.com'], false)).toEqual([
      { displayName: 'Room A' },
      { locationEmailAddress: 'r1@x.com' },
      { displayName: 'Conf B', locationEmailAddress: 'r2@x.com' }
    ]);
  });

  it('sets resolveAvailability when requested', () => {
    expect(parseMeetingLocationSpecs(['a@b.co'], true)).toEqual([
      { resolveAvailability: true, locationEmailAddress: 'a@b.co' }
    ]);
  });
});

describe('buildFindMeetingTimesLocationConstraint', () => {
  it('returns undefined when no constraint inputs', () => {
    expect(buildFindMeetingTimesLocationConstraint({})).toBeUndefined();
  });

  it('combines flags and locations', () => {
    expect(
      buildFindMeetingTimesLocationConstraint({
        suggestLocations: true,
        locationRequired: true,
        meetingLocation: ['Main|main@x.com']
      })
    ).toEqual({
      isRequired: true,
      suggestLocation: true,
      locations: [{ displayName: 'Main', locationEmailAddress: 'main@x.com' }]
    });
  });
});
