import { Command } from 'commander';
import { registerOfficeDocumentCommands } from './office-docs-shared.js';

export const wordCommand = new Command('word');
registerOfficeDocumentCommands(wordCommand, 'Microsoft Word');
