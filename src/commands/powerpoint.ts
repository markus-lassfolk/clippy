import { Command } from 'commander';
import { registerOfficeDocumentCommands } from './office-docs-shared.js';

export const powerpointCommand = new Command('powerpoint');
registerOfficeDocumentCommands(powerpointCommand, 'Microsoft PowerPoint');
