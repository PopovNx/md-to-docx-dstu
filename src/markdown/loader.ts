import fs from 'node:fs/promises';
import path from 'node:path';

import { Paragraph } from 'docx';

import type { GeneratorPaths, InputFile } from '../types';
import type { DocxChunk } from '../inject';
import { applyVars } from '../config';
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

export type FileCallback = (name: string, kind: 'markdown' | 'docx') => () => void;

export async function renderInputFiles(
    files: InputFile[],
    paths: GeneratorPaths,
    vars: Record<string, string> = {},
    onFile?: FileCallback,
): Promise<RenderResult> {
    const paragraphs: Paragraph[] = [];
    const chunks: DocxChunk[] = [];

    for (const file of files) {
        const name = path.basename(file.path);
        const done = onFile?.(name, file.kind);

        if (file.kind === 'markdown') {
            const raw = await fs.readFile(file.path, 'utf8');
            const source = Object.keys(vars).length > 0 ? applyVars(raw, vars) : raw;
            paragraphs.push(...renderMarkdown(source, paths));
        } else {
            const index = chunks.length;
            paragraphs.push(createChunkPlaceholder(index));
            chunks.push({ index, path: file.path });
        }

        done?.();
    }

    return { paragraphs, chunks };
}
