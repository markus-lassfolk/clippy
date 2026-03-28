import { Command } from 'commander';
import { startWebhookServer } from '../lib/webhook-server.js';

export const serveCommand = new Command('serve')
  .description('Start the webhook receiver server')
  .option('-p, --port <port>', 'Port to listen on', '3000')
  .action((options) => {
    startWebhookServer(parseInt(options.port, 10));
  });
