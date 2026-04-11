#!/usr/bin/env node

import fs from 'node:fs/promises';
import path from 'node:path';

import { Command } from 'commander';
import { Packer } from 'docx';

import { buildDocument } from './docx';
import { loadInputFiles, renderInputFiles } from './markdown';
import { injectDocxChunks } from './inject';

async function main() {
    const program = new Command();

    program
        .name('md-to-docx')
        .description('Generate a DOCX document directly from Markdown thesis fragments.')
        .option('-i, --input <dir>', 'Input directory (.md and .docx files)', path.resolve(process.cwd(), '../pages'))
        .option('-g, --images <dir>', 'Image directory', path.resolve(process.cwd(), '../images'))
        .option('-o, --output <file>', 'Output DOCX file', path.resolve(process.cwd(), './dist/course-project.docx'));

    program.parse();

    const options = program.opts<{ input: string; images: string; output: string }>();
    const inputFiles = await loadInputFiles(options.input);

    if (inputFiles.length === 0) {
        throw new Error(`No .md or .docx files found in ${options.input}`);
    }

    const { paragraphs, chunks } = await renderInputFiles(inputFiles, { imageDir: options.images });

    const document = buildDocument(paragraphs);
    let buffer = await Packer.toBuffer(document);

    if (chunks.length > 0) {
        buffer = await injectDocxChunks(buffer, chunks);
    }

    await fs.mkdir(path.dirname(options.output), { recursive: true });
    await fs.writeFile(options.output, buffer);

    const mdCount = inputFiles.filter((f) => f.kind === 'markdown').length;
    const docxCount = chunks.length;
    console.log(`Generated ${options.output}`);
    console.log(`Processed ${mdCount} Markdown file(s), injected ${docxCount} DOCX file(s).`);
}

main().catch((error) => {
    console.error(error instanceof Error ? error.stack ?? error.message : String(error));
    process.exit(1);
});
