import { describe, expect, test } from 'bun:test';
import { buildTeamsHtmlBodyWithMentions, parseAtSpecs } from './teams-message-compose.js';

describe('teams-message-compose', () => {
  test('parseAtSpecs splits on first colon', () => {
    expect(parseAtSpecs(['aaa-bbb-ccc-ddd-eee-fff:Jane Doe'])).toEqual([
      { userId: 'aaa-bbb-ccc-ddd-eee-fff', displayName: 'Jane Doe' }
    ]);
  });

  test('buildTeamsHtmlBodyWithMentions replaces @displayName and adds mentions', () => {
    const r = buildTeamsHtmlBodyWithMentions('Hello @Jane Doe — please review.', [
      { userId: 'u1', displayName: 'Jane Doe' }
    ]);
    expect(r.body.contentType).toBe('html');
    expect(r.body.content).toContain('<at id="0">Jane Doe</at>');
    expect(r.body.content).toContain('\u2014');
    expect(r.mentions).toHaveLength(1);
    expect((r.mentions[0] as { id: number }).id).toBe(0);
    expect((r.mentions[0] as { mentioned: { user: { id: string } } }).mentioned.user.id).toBe('u1');
  });

  test('buildTeamsHtmlBodyWithMentions finds @displayName when display name has HTML-special chars', () => {
    const r = buildTeamsHtmlBodyWithMentions('Hi @A & B <team> "x"', [
      { userId: 'u1', displayName: 'A & B <team> "x"' }
    ]);
    expect(r.body.content).toContain('<at id="0">A &amp; B &lt;team&gt; &quot;x&quot;</at>');
    expect((r.mentions[0] as { mentioned: { user: { displayName: string } } }).mentioned.user.displayName).toBe(
      'A & B <team> "x"'
    );
  });

  test('throws when @displayName missing from text', () => {
    expect(() => buildTeamsHtmlBodyWithMentions('Hello there', [{ userId: 'u1', displayName: 'Jane' }])).toThrow(
      /@Jane/
    );
  });
});
