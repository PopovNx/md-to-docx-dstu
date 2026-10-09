import fs from 'node:fs/promises';
import path from 'node:path';

import type { InputFile, RenderOptions } from '../types';
import type { DocxChunk } from '../inject';
import { applyVars } from '../config';
import { createChunkPlaceholder } from '../docx';
import { collapseEmptyLines, mark, markOf, type DocxBlock } from './blocks';
import { renderMarkdown } from './renderer';

export type RenderResult = {
    paragraphs: DocxBlock[];
    chunks: Array<Omit<DocxChunk, 'breakBefore'>>;
};

export async function loadInputFiles(inputDir: string): Promise<InputFile[]> {
    const entries = await fs.readdir(inputDir, { withFileTypes: true });
    return entries
        .filter(e => e.isFile() && /\.(md|docx)$/.test(e.name))
        .sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }))
        .map(e => ({ kind: e.name.endsWith('.docx') ? 'docx' : 'markdown', path: path.join(inputDir, e.name) }));
}

export async function renderInputFile(file: InputFile, opts: RenderOptions, vars: Record<string, string>, result: RenderResult) {
    if (file.kind === 'markdown') {
        result.paragraphs.push(...renderMarkdown(applyVars(await fs.readFile(file.path, 'utf8'), vars), opts));
        return;
    }
    const index = result.chunks.length;
    result.paragraphs.push(mark(createChunkPlaceholder(index), 'chunk'));
    result.chunks.push({ index, path: file.path });
}

export function finishPages({ paragraphs, chunks }: RenderResult): { paragraphs: DocxBlock[]; chunks: DocxChunk[] } {
    const collapsed = collapseEmptyLines(paragraphs);
    const blocks = collapsed.filter((block, i) =>
        markOf(block) !== 'pageBreak' || (markOf(collapsed[i - 1]) !== 'chunk' && markOf(collapsed[i + 1]) !== 'chunk'));

    const placeholderAt = blocks.flatMap((block, i) => (markOf(block) === 'chunk' ? [i] : []));
    return {
        paragraphs: blocks,
        chunks: chunks.map((chunk, n) => {
            const prev = blocks[placeholderAt[n]! - 1];
            return { ...chunk, breakBefore: prev !== undefined && markOf(prev) !== 'chunk' };
        }),
    };
}
