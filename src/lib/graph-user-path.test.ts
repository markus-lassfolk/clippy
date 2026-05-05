import { describe, expect, it } from 'bun:test';
import { graphUserPath } from './graph-user-path.js';

describe('graphUserPath', () => {
  it('uses /me when user is omitted or blank', () => {
    expect(graphUserPath(undefined, 'joinedTeams')).toBe('/me/joinedTeams');
    expect(graphUserPath('', 'joinedTeams')).toBe('/me/joinedTeams');
    expect(graphUserPath('  ', 'manager')).toBe('/me/manager');
  });

  it('encodes UPN in /users/ path for joined teams, manager, directReports', () => {
    expect(graphUserPath('manager@contoso.com', 'joinedTeams')).toBe('/users/manager%40contoso.com/joinedTeams');
    expect(graphUserPath('manager@contoso.com', 'manager')).toBe('/users/manager%40contoso.com/manager');
    expect(graphUserPath('manager@contoso.com', 'directReports')).toBe('/users/manager%40contoso.com/directReports');
  });

  it('strips leading slash from suffix', () => {
    expect(graphUserPath('u@x', '/calendar/events')).toBe('/users/u%40x/calendar/events');
  });
});
