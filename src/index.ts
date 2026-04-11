#!/usr/bin/env node

import fs from 'node:fs/promises';
import path from 'node:path';

import { Command } from 'commander';
import { Packer } from 'docx';

import { buildDocument } from './docx';
import { loadInputFiles } from './markdown';
import { injectDocxChunks } from './inject';
import { CONFIG_FILE, loadConfig, initConfig } from './config';
import { buildTaskList } from './progress';
import { log } from './logger';

const program = new Command();

program
    .name('md-to-docx')
    .description('Convert Markdown + DOCX fragments into a single DOCX document.')
    .helpOption('-h, --help', 'Show help');

program
    .command('init')
    .description(`Create ${CONFIG_FILE} in the current directory`)
    .action(async () => {
        const configPath = path.resolve(process.cwd(), CONFIG_FILE);
        try {
            await initConfig(configPath);
            console.log(`Created ${CONFIG_FILE}`);
        } catch (err) {
            die(err);
        }
    });

program
    .command('render', { isDefault: true })
    .description(`Render using ${CONFIG_FILE} (default command)`)
    .option('-c, --config <file>', 'Config file path', CONFIG_FILE)
    .action(async (opts: { config: string }) => {
        const configPath = path.resolve(process.cwd(), opts.config);
        const config = await loadConfig(configPath).catch(die);

        const inputFiles = await loadInputFiles(config.dirs.pages).catch(die);
        if (inputFiles.length === 0) die(`No .md or .docx files found in ${config.dirs.pages}`);

        const tasks = buildTaskList(config, inputFiles);

        const { result } = await tasks.run().catch(die);

        log.debug('packing document');
        const document = buildDocument(result.paragraphs);
        let buffer = await Packer.toBuffer(document);

        if (result.chunks.length > 0) {
            log.debug({ count: result.chunks.length }, 'injecting docx chunks');
            buffer = await injectDocxChunks(buffer, result.chunks, config.vars);
        }

        await fs.mkdir(path.dirname(config.dirs.output), { recursive: true });
        await fs.writeFile(config.dirs.output, buffer);

        const kb = (buffer.length / 1024).toFixed(0);
        console.log(`\nWritten ${config.dirs.output}  (${kb} KB)\n`);
    });

program.parse();

function die(err: unknown): never {
    const msg = err instanceof Error ? err.message : String(err);
    log.error(msg);
    console.error(`\nError: ${msg}\n`);
    process.exit(1);
}
