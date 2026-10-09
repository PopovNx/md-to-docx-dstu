import fs from 'node:fs/promises';
import path from 'node:path';

import { Paragraph, Table } from 'docx';

import type { InputFile, RenderOptions } from '../types';
import type { DocxChunk } from '../inject';
import { applyVars } from '../config';
import { createChunkPlaceholder } from '../docx';
import { isPageBreak, renderMarkdown } from './renderer';

export type RenderResult = {
    paragraphs: Array<Paragraph | Table>;
    chunks: DocxChunk[];
};

export async function loadInputFiles(inputDir: string): Promise<InputFile[]> {
    const entries = await fs.readdir(inputDir, { withFileTypes: true });
    return entries
        .filter(e => e.isFile() && /\.(md|docx)$/.test(e.name))
        .sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }))
        .map(e => ({ kind: e.name.endsWith('.docx') ? 'docx' : 'markdown', path: path.join(inputDir, e.name) }));
}

const placeholders = new WeakSet<Paragraph | Table>();

export async function renderInputFile(file: InputFile, opts: RenderOptions, vars: Record<string, string>, result: RenderResult) {
    const { paragraphs, chunks } = result;

    if (file.kind === 'markdown') {
        const blocks = renderMarkdown(applyVars(await fs.readFile(file.path, 'utf8'), vars), opts);
        const last = paragraphs.at(-1);
        if (last && placeholders.has(last) && isPageBreak(blocks[0])) blocks.shift();
        paragraphs.push(...blocks);
        return;
    }

    if (isPageBreak(paragraphs.at(-1))) paragraphs.pop();
    const prev = paragraphs.at(-1);
    const placeholder = createChunkPlaceholder(chunks.length);
    placeholders.add(placeholder);
    chunks.push({ index: chunks.length, path: file.path, breakBefore: prev !== undefined && !placeholders.has(prev) });
    paragraphs.push(placeholder);
}
