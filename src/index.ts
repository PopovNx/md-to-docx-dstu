#!/usr/bin/env node

import fs from 'node:fs/promises';
import path from 'node:path';

import { Command } from 'commander';
import { Packer } from 'docx';

import { buildDocument } from './docx';
import { loadMarkdownFiles, renderMarkdownFiles } from './markdown';

async function main() {
  const program = new Command();

  program
    .name('md-to-docx')
    .description('Generate a DOCX document directly from Markdown thesis fragments.')
    .option('-i, --input <dir>', 'Markdown directory', path.resolve(process.cwd(), '../pages'))
    .option('-g, --images <dir>', 'Image directory', path.resolve(process.cwd(), '../images'))
    .option('-o, --output <file>', 'Output DOCX file', path.resolve(process.cwd(), './dist/course-project.docx'));

  program.parse();

  const options = program.opts<{ input: string; images: string; output: string }>();
  const markdownFiles = await loadMarkdownFiles(options.input);

  if (markdownFiles.length === 0) {
    throw new Error(`No Markdown files found in ${options.input}`);
  }

  const blocks = await renderMarkdownFiles(markdownFiles, { imageDir: options.images });
  const document = buildDocument(blocks);
  const buffer = await Packer.toBuffer(document);

  await fs.mkdir(path.dirname(options.output), { recursive: true });
  await fs.writeFile(options.output, buffer);

  console.log(`Generated ${options.output}`);
  console.log(`Processed ${markdownFiles.length} Markdown files.`);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.stack ?? error.message : String(error));
  process.exit(1);
});
