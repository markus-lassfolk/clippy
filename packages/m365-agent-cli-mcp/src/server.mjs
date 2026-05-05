#!/usr/bin/env node
/**
 * Thin MCP server: tools spawn `m365-agent-cli` (must be on PATH or set M365_AGENT_CLI_BIN).
 */
import { spawnSync } from 'node:child_process';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';

function cliBinary() {
  return process.env.M365_AGENT_CLI_BIN?.trim() || 'm365-agent-cli';
}

function runM365(args) {
  const bin = cliBinary();
  const r = spawnSync(bin, args, {
    encoding: 'utf8',
    maxBuffer: 50 * 1024 * 1024,
    env: process.env
  });
  if (r.error) {
    throw r.error;
  }
  const err = (r.stderr || '').trim();
  const out = (r.stdout || '').trim();
  if (r.status !== 0) {
    throw new Error(err || out || `m365-agent-cli exited with code ${r.status}`);
  }
  return out;
}

const server = new McpServer({
  name: 'm365-agent-cli-mcp',
  version: '0.1.0'
});

server.registerTool(
  'm365_whoami',
  {
    description: 'JSON summary of the signed-in user (`m365-agent-cli whoami --json`).',
    inputSchema: {}
  },
  async () => {
    const text = runM365(['whoami', '--json']);
    return { content: [{ type: 'text', text }] };
  }
);

server.registerTool(
  'm365_graph_search',
  {
    description: 'Microsoft Search flattened hits (`m365-agent-cli graph-search … --json-hits`).',
    inputSchema: {
      query: z.string().min(1).describe('Search query (KQL-style per Graph docs)'),
      preset: z.enum(['default', 'extended', 'connectors']).optional()
    }
  },
  async ({ query, preset }) => {
    const args = ['graph-search', query, '--json-hits'];
    if (preset) {
      args.push('--preset', preset);
    }
    const text = runM365(args);
    return { content: [{ type: 'text', text }] };
  }
);

server.registerTool(
  'm365_graph_invoke_get',
  {
    description:
      'Read-only Graph GET (`m365-agent-cli --read-only graph invoke -X GET <path>`). Path must start with /.',
    inputSchema: {
      path: z.string().min(1).describe('Example: /v1.0/me/drive/root/children?$top=5'),
      beta: z.boolean().optional().describe('Use Graph beta root URL')
    }
  },
  async ({ path, beta }) => {
    if (!path.startsWith('/')) {
      throw new Error('path must start with /');
    }
    const args = ['--read-only', 'graph', 'invoke', '-X', 'GET'];
    if (beta) {
      args.push('--beta');
    }
    args.push(path);
    const text = runM365(args);
    return { content: [{ type: 'text', text }] };
  }
);

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
