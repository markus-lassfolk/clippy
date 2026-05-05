import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  createCalendarGroup,
  createCalendarResource,
  deleteCalendarGroup,
  deleteCalendarResource,
  eventsDeltaPage,
  type GraphCalendarEvent,
  getCalendar,
  getEvent,
  listCalendarGroups,
  listCalendars,
  listCalendarView,
  updateCalendarResource
} from '../lib/graph-calendar-client.js';
import {
  applyDeltaPageToState,
  assertDeltaScopeMatchesState,
  readDeltaStateFile,
  resolveDeltaContinuationUrl,
  writeDeltaStateFile
} from '../lib/graph-delta-state-file.js';
import { acceptEventInvitation, declineEventInvitation, tentativelyAcceptEventInvitation } from '../lib/graph-event.js';
import { checkReadOnly } from '../lib/utils.js';

export const graphCalendarCommand = new Command('graph-calendar').description(
  'Microsoft Graph calendar REST: calendars (list/get/create/update/delete), calendar groups, calendarView, events, deltas, invitation responses (distinct from EWS `calendar` / `respond`)'
);

function formatEventLine(e: GraphCalendarEvent): string {
  const subj = e.subject?.trim() || '(no subject)';
  const start = e.start?.dateTime;
  const end = e.end?.dateTime;
  const tz = e.start?.timeZone || '';
  const when =
    start && end ? `${start} → ${end}${tz ? ` (${tz})` : ''}` : start ? `${start}${tz ? ` (${tz})` : ''}` : '?';
  const allDay = e.isAllDay ? ' [all-day]' : '';
  return `${when}${allDay}\t${subj}\t${e.id}`;
}

graphCalendarCommand
  .command('list-calendars')
  .description('List calendars (Graph GET /calendars)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target mailbox (delegation)')
  .action(async (opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listCalendars(auth.token, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const c of r.data) {
      const label = c.name || '(unnamed)';
      console.log(`${label}\t${c.id}`);
    }
  });

graphCalendarCommand
  .command('get-calendar')
  .description('Get one calendar by id')
  .argument('<calendarId>', 'Calendar id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target mailbox (delegation)')
  .action(async (calendarId: string, opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getCalendar(auth.token, calendarId, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else console.log(JSON.stringify(r.data, null, 2));
  });

graphCalendarCommand
  .command('list-calendar-groups')
  .description('List calendar groups (Graph GET /calendarGroups)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target mailbox (delegation)')
  .action(async (opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listCalendarGroups(auth.token, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const g of r.data) {
      console.log(`${g.name ?? '(unnamed)'}\t${g.id}`);
    }
  });

graphCalendarCommand
  .command('create-calendar-group')
  .description('Create a calendar group (Graph POST /calendarGroups)')
  .requiredOption('--name <name>', 'Display name')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target mailbox (delegation)')
  .action(
    async (opts: { name: string; json?: boolean; token?: string; identity?: string; user?: string }, cmd: any) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await createCalendarGroup(auth.token, opts.name, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      console.log(`Created calendar group: ${r.data.id}`);
    }
  );

graphCalendarCommand
  .command('delete-calendar-group')
  .description('Delete a calendar group (Graph DELETE /calendarGroups/{id})')
  .argument('<groupId>', 'Calendar group id')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target mailbox (delegation)')
  .action(async (groupId: string, opts: { token?: string; identity?: string; user?: string }, cmd: any) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await deleteCalendarGroup(auth.token, groupId, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log('Done.');
  });

graphCalendarCommand
  .command('create-calendar')
  .description('Create a calendar (Graph POST /calendars or POST /calendarGroups/{id}/calendars)')
  .requiredOption('--name <name>', 'Calendar name')
  .option('--color <preset>', 'Color preset (e.g. preset7); see Graph calendar resource')
  .option('--group-id <id>', 'Create inside this calendar group (omit for top-level default group behavior per tenant)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target mailbox (delegation)')
  .action(
    async (
      opts: {
        name: string;
        color?: string;
        groupId?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await createCalendarResource(
        auth.token,
        { name: opts.name, color: opts.color },
        opts.user,
        opts.groupId
      );
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      console.log(`Created calendar: ${r.data.name ?? opts.name}\t${r.data.id}`);
    }
  );

graphCalendarCommand
  .command('update-calendar')
  .description('Update a calendar (Graph PATCH /calendars/{id})')
  .argument('<calendarId>', 'Calendar id')
  .option('--name <name>', 'New display name')
  .option('--color <preset>', 'New color preset')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target mailbox (delegation)')
  .action(
    async (
      calendarId: string,
      opts: { name?: string; color?: string; json?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      if (!opts.name?.trim() && !opts.color?.trim()) {
        console.error('Error: provide --name and/or --color');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await updateCalendarResource(auth.token, calendarId, { name: opts.name, color: opts.color }, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      console.log('Done.');
    }
  );

graphCalendarCommand
  .command('delete-calendar')
  .description('Delete a calendar (Graph DELETE /calendars/{id}; cannot delete default calendar)')
  .argument('<calendarId>', 'Calendar id')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target mailbox (delegation)')
  .action(async (calendarId: string, opts: { token?: string; identity?: string; user?: string }, cmd: any) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await deleteCalendarResource(auth.token, calendarId, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log('Done.');
  });

graphCalendarCommand
  .command('list-view')
  .description('List events in a time window (Graph GET .../calendarView)')
  .requiredOption('--start <iso>', 'Start (ISO 8601, e.g. 2026-04-01T00:00:00Z)')
  .requiredOption('--end <iso>', 'End (ISO 8601, exclusive upper bound in many cases — see Graph docs)')
  .option('-c, --calendar <calendarId>', 'Calendar id (omit for default calendar)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target mailbox (delegation)')
  .action(
    async (opts: {
      start: string;
      end: string;
      calendar?: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await listCalendarView(auth.token, opts.start, opts.end, {
        calendarId: opts.calendar,
        user: opts.user
      });
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      for (const e of r.data) {
        console.log(formatEventLine(e));
      }
    }
  );

graphCalendarCommand
  .command('events-delta')
  .description(
    'One page of events delta sync (use @odata.nextLink as --next for more pages; @odata.deltaLink for baseline). Optional --state-file persists next/delta URLs for unattended loops.'
  )
  .option('-c, --calendar <calendarId>', 'Delta for this calendar only (omit for default calendar /me/events/delta)')
  .option('--next <url>', 'Full @odata.nextLink URL from a previous response (overrides --state-file continuation)')
  .option('--state-file <path>', 'Read/write JSON delta cursor (pending nextLink + stable deltaLink)')
  .option('--json', 'Output raw page JSON (value, nextLink, deltaLink)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target mailbox (delegation)')
  .action(
    async (opts: {
      calendar?: string;
      next?: string;
      stateFile?: string;
      json?: boolean;
      token?: string;
      identity?: string;
      user?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const existingState = opts.stateFile ? await readDeltaStateFile(opts.stateFile) : null;
      if (existingState && existingState.kind !== 'calendarEvents') {
        console.error('Error: state file is not for calendar events-delta (kind must be calendarEvents).');
        process.exit(1);
      }
      try {
        if (existingState) {
          assertDeltaScopeMatchesState(existingState, { calendarId: opts.calendar, user: opts.user });
        }
      } catch (err) {
        console.error(err instanceof Error ? err.message : err);
        process.exit(1);
      }
      const continueUrl = resolveDeltaContinuationUrl({ explicitNext: opts.next, state: existingState });
      const r = await eventsDeltaPage(auth.token, {
        user: opts.user,
        calendarId: opts.calendar,
        nextLink: continueUrl
      });
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.stateFile && r.data) {
        const merged = applyDeltaPageToState(existingState, 'calendarEvents', r.data, {
          calendarId: opts.calendar,
          user: opts.user
        });
        await writeDeltaStateFile(opts.stateFile, merged);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        console.log(`Changes: ${r.data.value?.length ?? 0} item(s)`);
        if (r.data['@odata.nextLink']) console.log(`nextLink: ${r.data['@odata.nextLink']}`);
        if (r.data['@odata.deltaLink']) console.log(`deltaLink: ${r.data['@odata.deltaLink']}`);
        if (opts.stateFile) console.log(`state-file: ${opts.stateFile} (updated)`);
      }
    }
  );

graphCalendarCommand
  .command('get-event')
  .description('Get a single event by id (Graph GET /events/{id})')
  .argument('<eventId>', 'Event id')
  .option('--select <fields>', 'OData $select (comma-separated)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Mailbox that owns the event (delegation)')
  .action(
    async (
      eventId: string,
      opts: {
        select?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getEvent(auth.token, eventId, opts.user, opts.select);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else console.log(JSON.stringify(r.data, null, 2));
    }
  );

function addRespondCommand(
  name: string,
  description: string,
  fn: (o: {
    token: string;
    eventId: string;
    comment?: string;
    sendResponse: boolean;
    user?: string;
  }) => ReturnType<typeof acceptEventInvitation>
) {
  graphCalendarCommand
    .command(name)
    .description(description)
    .argument('<eventId>', 'Event id')
    .option('--comment <text>', 'Optional comment to organizer')
    .option('--no-notify', "Don't send response to organizer")
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity (default: default)')
    .option('--user <email>', 'Mailbox that owns the invitation (delegation)')
    .action(
      async (
        eventId: string,
        opts: {
          comment?: string;
          notify: boolean;
          token?: string;
          identity?: string;
          user?: string;
        },
        cmd: any
      ) => {
        checkReadOnly(cmd);
        const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
        if (!auth.success || !auth.token) {
          console.error(`Auth error: ${auth.error}`);
          process.exit(1);
        }
        const r = await fn({
          token: auth.token,
          eventId,
          comment: opts.comment,
          sendResponse: opts.notify !== false,
          user: opts.user
        });
        if (!r.ok) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        console.log('Done.');
      }
    );
}

addRespondCommand('accept', 'Accept a meeting request (Graph POST .../accept)', acceptEventInvitation);
addRespondCommand('decline', 'Decline a meeting request (Graph POST .../decline)', declineEventInvitation);
addRespondCommand(
  'tentative',
  'Tentatively accept without proposing a new time (Graph POST .../tentativelyAccept)',
  tentativelyAcceptEventInvitation
);
