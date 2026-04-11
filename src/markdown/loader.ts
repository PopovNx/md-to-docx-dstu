import fs from 'node:fs/promises';
import path from 'node:path';

import { Paragraph } from 'docx';

import type { GeneratorPaths, InputFile } from '../types';
import type { DocxChunk } from '../inject';
import { createChunkPlaceholder } from '../docx';
import { renderMarkdown } from './renderer';

export async function loadInputFiles(inputDir: string): Promise<InputFile[]> {
    const entries = await fs.readdir(inputDir, { withFileTypes: true });

    return entries
        .filter((e) => e.isFile() && (e.name.endsWith('.md') || e.name.endsWith('.docx')))
        .sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }))
        .map((e): InputFile => ({
            kind: e.name.endsWith('.docx') ? 'docx' : 'markdown',
            path: path.join(inputDir, e.name),
        }));
}

export type RenderResult = {
    paragraphs: Paragraph[];
    chunks: DocxChunk[];
};

export async function renderInputFiles(files: InputFile[], paths: GeneratorPaths): Promise<RenderResult> {
    const paragraphs: Paragraph[] = [];
    const chunks: DocxChunk[] = [];

    for (const file of files) {
        if (file.kind === 'markdown') {
            const source = await fs.readFile(file.path, 'utf8');
            paragraphs.push(...renderMarkdown(source, paths));
        } else {
            const index = chunks.length;
            paragraphs.push(createChunkPlaceholder(index));
            chunks.push({ index, path: file.path });
        }
    }

    return { paragraphs, chunks };
}
