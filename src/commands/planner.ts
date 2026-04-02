import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  createTask,
  getPlanDetails,
  getTask,
  listGroupPlans,
  listPlanBuckets,
  listPlanTasks,
  listUserPlans,
  listUserTasks,
  normalizeAppliedCategories,
  type PlannerCategorySlot,
  type PlannerPlanDetails,
  type PlannerTask,
  parsePlannerLabelKey,
  updateTask
} from '../lib/planner-client.js';
import { checkReadOnly } from '../lib/utils.js';

const LABEL_SLOTS: PlannerCategorySlot[] = [
  'category1',
  'category2',
  'category3',
  'category4',
  'category5',
  'category6'
];

function formatTaskLabels(task: PlannerTask, descriptions?: PlannerPlanDetails['categoryDescriptions']): string {
  if (!task.appliedCategories) return '';
  const parts: string[] = [];
  for (const slot of LABEL_SLOTS) {
    if (task.appliedCategories[slot]) {
      const name = descriptions?.[slot]?.trim();
      parts.push(name || slot);
    }
  }
  return parts.join(', ');
}

export const plannerCommand = new Command('planner').description('Manage Microsoft Planner tasks and plans');

plannerCommand
  .command('list-my-tasks')
  .description('List tasks assigned to you')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const result = await listUserTasks(auth.token!);
    if (!result.ok || !result.data) {
      console.error(`Error listing tasks: ${result.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(result.data, null, 2));
    } else {
      const planDetailsCache = new Map<string, PlannerPlanDetails['categoryDescriptions']>();
      for (const t of result.data) {
        if (!planDetailsCache.has(t.planId)) {
          const d = await getPlanDetails(auth.token!, t.planId);
          planDetailsCache.set(t.planId, d.ok ? d.data?.categoryDescriptions : undefined);
        }
        const desc = planDetailsCache.get(t.planId);
        const labels = formatTaskLabels(t, desc);
        console.log(`- [${t.percentComplete === 100 ? 'x' : ' '}] ${t.title} (ID: ${t.id})`);
        console.log(`  Plan ID: ${t.planId} | Bucket ID: ${t.bucketId}${labels ? ` | Labels: ${labels}` : ''}`);
      }
    }
  });

plannerCommand
  .command('list-plans')
  .description('List your plans or plans for a group')
  .option('-g, --group <groupId>', 'Group ID to list plans for')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { group?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const result = opts.group ? await listGroupPlans(auth.token!, opts.group) : await listUserPlans(auth.token!);
    if (!result.ok || !result.data) {
      console.error(`Error listing plans: ${result.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(result.data, null, 2));
    } else {
      for (const p of result.data) {
        console.log(`- ${p.title} (ID: ${p.id})`);
      }
    }
  });

plannerCommand
  .command('list-buckets')
  .description('List buckets in a plan')
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { plan: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const result = await listPlanBuckets(auth.token!, opts.plan);
    if (!result.ok || !result.data) {
      console.error(`Error listing buckets: ${result.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(result.data, null, 2));
    } else {
      for (const b of result.data) {
        console.log(`- ${b.name} (ID: ${b.id})`);
      }
    }
  });

plannerCommand
  .command('list-tasks')
  .description('List tasks in a plan')
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { plan: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const result = await listPlanTasks(auth.token!, opts.plan);
    if (!result.ok || !result.data) {
      console.error(`Error listing tasks: ${result.error?.message}`);
      process.exit(1);
    }
    const detailsR = await getPlanDetails(auth.token!, opts.plan);
    const descriptions = detailsR.ok ? detailsR.data?.categoryDescriptions : undefined;
    if (opts.json) {
      console.log(JSON.stringify(result.data, null, 2));
    } else {
      for (const t of result.data) {
        const labels = formatTaskLabels(t, descriptions);
        console.log(
          `- [${t.percentComplete === 100 ? 'x' : ' '}] ${t.title} (ID: ${t.id})${labels ? ` | ${labels}` : ''}`
        );
      }
    }
  });

plannerCommand
  .command('create-task')
  .description('Create a new task in a plan')
  .requiredOption('-p, --plan <planId>', 'Plan ID')
  .requiredOption('-t, --title <title>', 'Task title')
  .option('-b, --bucket <bucketId>', 'Bucket ID')
  .option(
    '--label <slot>',
    'Label slot: 1-6 or category1..category6 (repeatable; names are defined in plan details)',
    (v: string, prev: string[]) => [...prev, v],
    [] as string[]
  )
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: {
        plan: string;
        title: string;
        bucket?: string;
        label?: string[];
        json?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let applied: ReturnType<typeof normalizeAppliedCategories> | undefined;
      if (opts.label?.length) {
        const setTrue: PlannerCategorySlot[] = [];
        for (const raw of opts.label) {
          const slot = parsePlannerLabelKey(raw);
          if (!slot) {
            console.error(`Invalid --label "${raw}". Use 1-6 or category1..category6.`);
            process.exit(1);
          }
          setTrue.push(slot);
        }
        applied = normalizeAppliedCategories(undefined, { setTrue });
      }
      const result = await createTask(auth.token!, opts.plan, opts.title, opts.bucket, undefined, applied);
      if (!result.ok || !result.data) {
        console.error(`Error creating task: ${result.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(result.data, null, 2));
      } else {
        console.log(`Created task: ${result.data.title} (ID: ${result.data.id})`);
      }
    }
  );

plannerCommand
  .command('update-task')
  .description('Update a task')
  .requiredOption('-i, --id <taskId>', 'Task ID')
  .option('--title <title>', 'New title')
  .option('-b, --bucket <bucketId>', 'Move to Bucket ID')
  .option('--percent <percentComplete>', 'Percent complete (0-100)')
  .option('--assign <userId>', 'Assign to user ID')
  .option(
    '--label <slot>',
    'Turn on label slot (1-6 or category1..category6); repeatable',
    (v: string, prev: string[]) => [...prev, v],
    [] as string[]
  )
  .option(
    '--unlabel <slot>',
    'Turn off label slot; repeatable',
    (v: string, prev: string[]) => [...prev, v],
    [] as string[]
  )
  .option('--clear-labels', 'Clear all label slots on the task')
  .option('--json', 'Output JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: {
        id: string;
        title?: string;
        bucket?: string;
        percent?: string;
        assign?: string;
        label?: string[];
        unlabel?: string[];
        clearLabels?: boolean;
        json?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }

      // First, we need to get the task to retrieve its ETag.
      const taskRes = await getTask(auth.token!, opts.id);
      if (!taskRes.ok || !taskRes.data) {
        console.error(`Error fetching task: ${taskRes.error?.message}`);
        process.exit(1);
      }
      const etag = taskRes.data['@odata.etag'];
      if (!etag) {
        console.error('Task does not have an ETag');
        process.exit(1);
      }

      const updates: any = {};
      if (opts.title !== undefined) updates.title = opts.title;
      if (opts.bucket !== undefined) updates.bucketId = opts.bucket;
      if (opts.percent !== undefined) {
        const percentValue = parseInt(opts.percent, 10);
        if (Number.isNaN(percentValue) || percentValue < 0 || percentValue > 100) {
          console.error(`Invalid percent value: ${opts.percent}. Must be a number between 0 and 100.`);
          process.exit(1);
        }
        updates.percentComplete = percentValue;
      }
      if (opts.assign) {
        // Planner API requires a specific structure for assignments.
        updates.assignments = {
          [opts.assign]: {
            '@odata.type': '#microsoft.graph.plannerAssignment',
            orderHint: ' !'
          }
        };
      }

      const labelOps = (opts.label?.length ?? 0) > 0 || (opts.unlabel?.length ?? 0) > 0 || opts.clearLabels;
      if (labelOps) {
        const setTrue: PlannerCategorySlot[] = [];
        const setFalse: PlannerCategorySlot[] = [];
        for (const raw of opts.label ?? []) {
          const slot = parsePlannerLabelKey(raw);
          if (!slot) {
            console.error(`Invalid --label "${raw}". Use 1-6 or category1..category6.`);
            process.exit(1);
          }
          setTrue.push(slot);
        }
        for (const raw of opts.unlabel ?? []) {
          const slot = parsePlannerLabelKey(raw);
          if (!slot) {
            console.error(`Invalid --unlabel "${raw}". Use 1-6 or category1..category6.`);
            process.exit(1);
          }
          setFalse.push(slot);
        }
        if (opts.clearLabels && (setTrue.length > 0 || setFalse.length > 0)) {
          console.error('Error: use --clear-labels alone, or use --label/--unlabel without --clear-labels');
          process.exit(1);
        }
        updates.appliedCategories = normalizeAppliedCategories(taskRes.data.appliedCategories, {
          clearAll: opts.clearLabels,
          setTrue: setTrue.length ? setTrue : undefined,
          setFalse: setFalse.length ? setFalse : undefined
        });
      }

      if (Object.keys(updates).length === 0) {
        console.error(
          'Error: specify at least one of --title, --bucket, --percent, --assign, --label, --unlabel, --clear-labels'
        );
        process.exit(1);
      }

      const result = await updateTask(auth.token!, opts.id, etag, updates);
      if (!result.ok) {
        console.error(`Error updating task: ${result.error?.message}`);
        process.exit(1);
      }

      // Since PATCH returns 204 No Content, get task again to show updated state
      const updatedTaskRes = await getTask(auth.token!, opts.id);
      if (!updatedTaskRes.ok || !updatedTaskRes.data) {
        console.error(`Error fetching updated task: ${updatedTaskRes.error?.message}`);
        process.exit(1);
      }

      if (opts.json) {
        console.log(JSON.stringify(updatedTaskRes.data, null, 2));
      } else {
        console.log(`Updated task: ${opts.id}`);
      }
    }
  );
