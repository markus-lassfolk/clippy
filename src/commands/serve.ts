import { Command } from 'commander';
import { startWebhookServer } from '../lib/webhook-server.js';

export const serveCommand = new Command('serve')
  .description('Start the webhook receiver server')
  .option('-p, --port <port>', 'Port to listen on', '3000')
  .action((options) => {
    const port = parseInt(options.port, 10);
    if (Number.isNaN(port) || port <= 0 || port > 65535) {
      console.error(`Invalid port "${options.port}". Please provide an integer between 1 and 65535.`);
      process.exit(1);
    }
    startWebhookServer(port);
  });
