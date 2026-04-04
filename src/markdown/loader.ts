import fs from 'node:fs/promises';
import path from 'node:path';

import { Paragraph } from 'docx';

import type { GeneratorPaths } from '../types';
import { renderMarkdown } from './renderer';

export async function loadMarkdownFiles(inputDir: string): Promise<string[]> {
    const entries = await fs.readdir(inputDir, { withFileTypes: true });

    return entries
        .filter((entry) => entry.isFile() && entry.name.endsWith('.md'))
        .map((entry) => path.join(inputDir, entry.name))
        .sort((left, right) => path.basename(left).localeCompare(path.basename(right), undefined, { numeric: true }));
}

export async function renderMarkdownFiles(filePaths: string[], paths: GeneratorPaths): Promise<Paragraph[]> {
    const rendered: Paragraph[] = [];

    for (const filePath of filePaths) {
        const source = await fs.readFile(filePath, 'utf8');
        rendered.push(...renderMarkdown(source, paths));
    }

    return rendered;
}
