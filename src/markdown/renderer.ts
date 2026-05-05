import { type Token } from 'markdown-it';
import MarkdownIt from 'markdown-it';
import markdownItAttrs from 'markdown-it-attrs';
import { Paragraph, Table, TableCell, TableRow } from 'docx';

import type { GeneratorPaths } from '../types';
import {
    createBodyParagraph,
    createCaptionParagraph,
    createCodeParagraph,
    createHeadingParagraph,
    createImageParagraph,
    createListParagraph,
    TABLE_CONTENT_WIDTH,
    createTable,
    createTableCell,
    createTableRow,
    createPageBreak as createDocxPageBreak,
} from '../docx';
import { buildImageRun } from './image';
import { renderInlineTokens } from './inline';

const markdown = new MarkdownIt({ html: true, breaks: false, linkify: true, typographer: false })
    .use(markdownItAttrs);

const LIST_OPEN_TOKENS: Partial<Record<string, 'bullet' | 'number'>> = {
    bullet_list_open: 'bullet',
    ordered_list_open: 'number',
};

type DocxBlock = Paragraph | Table;

const emptySet = new WeakSet<Paragraph>();
const headingSet = new WeakSet<Paragraph>();
const headingLevelSet = new WeakMap<Paragraph, 1 | 2>();
const pageBreakSet = new WeakSet<Paragraph>();

function createEmpty(): Paragraph {
    const p = createBodyParagraph([]);
    emptySet.add(p);
    return p;
}

function isEmpty(p: Paragraph): boolean {
    return emptySet.has(p);
}

function isHeadingParagraph(p: Paragraph): boolean {
    return headingSet.has(p);
}

function getHeadingLevel(p: Paragraph): 1 | 2 | undefined {
    return headingLevelSet.get(p);
}

function isPageBreakParagraph(p: Paragraph): boolean {
    return pageBreakSet.has(p);
}

type ListState = {
    kinds: Array<'bullet' | 'number'>;
    firstBlock: boolean[];
};

function currentListKind(list: ListState): 'bullet' | 'number' | undefined {
    return list.kinds[list.kinds.length - 1];
}

function isFirstBlockInListItem(list: ListState): boolean {
    return list.firstBlock[list.firstBlock.length - 1] ?? false;
}

function markListBlockRendered(list: ListState): void {
    if (list.firstBlock.length > 0) list.firstBlock[list.firstBlock.length - 1] = false;
}

export function renderMarkdown(source: string, paths: GeneratorPaths): DocxBlock[] {
    const tokens = markdown.parse(normalizeSource(source), {});
    const raw: DocxBlock[] = [];
    const list: ListState = { kinds: [], firstBlock: [] };

    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];

        if (token.type === 'html_block' && token.content.includes('DOCX_PAGE_BREAK')) {
            const pageBreak = createDocxPageBreak();
            pageBreakSet.add(pageBreak);
            raw.push(pageBreak);
            continue;
        }

        if (token.type === 'heading_open') {
            const inline = tokens[i + 1];
            if (inline?.type !== 'inline') continue;
            const level = Number(token.tag.replace('h', '')) as 1 | 2;
            const heading = createHeadingParagraph(level, renderInlineTokens(inline.children ?? [], paths));
            headingSet.add(heading);
            headingLevelSet.set(heading, level);

            const prev = raw[raw.length - 1];
            if (level !== 1 || (prev !== undefined && (!isParagraphBlock(prev) || !isPageBreakParagraph(prev)))) {
                raw.push(createEmpty());
            }

            raw.push(heading, createEmpty());
            i += 2;
            continue;
        }

        if (token.type === 'table_open') {
            const { table, nextIndex } = renderTable(tokens, i, paths);
            raw.push(table, createEmpty());
            markListBlockRendered(list);
            i = nextIndex;
            continue;
        }

        const listKind = LIST_OPEN_TOKENS[token.type];
        if (listKind !== undefined) { list.kinds.push(listKind); continue; }
        if (token.type.endsWith('_list_close')) { list.kinds.pop(); continue; }
        if (token.type === 'list_item_open') { list.firstBlock.push(true); continue; }
        if (token.type === 'list_item_close') { list.firstBlock.pop(); continue; }

        if (token.type === 'hr') {
            raw.push(createEmpty());
            continue;
        }

        if (token.type === 'fence' || token.type === 'code_block') {
            raw.push(createCodeParagraph(token.content.replace(/\s+$/, '').split('\n')));
            markListBlockRendered(list);
            continue;
        }

        if (token.type === 'paragraph_open') {
            const inline = tokens[i + 1];
            if (inline?.type !== 'inline') continue;
            const blocks = renderParagraph(inline.children ?? [], paths, list);
            raw.push(...(Array.isArray(blocks) ? blocks : [blocks]));
            i += 2;
            continue;
        }
    }

    return collapseEmpties(raw);
}

function renderParagraph(
    children: Token[],
    paths: GeneratorPaths,
    list: ListState,
): Paragraph | Paragraph[] {
    const imageTokens = extractImageOnlyTokens(children);

    if (imageTokens.length > 0) {
        return imageTokens.flatMap((t) => {
            const caption = t.content.trim();
            const image = createImageParagraph(buildImageRun(t, paths));
            return caption
                ? [createEmpty(), image, createCaptionParagraph(caption), createEmpty()]
                : [createEmpty(), image, createEmpty()];
        });
    }

    const rendered = renderInlineTokens(children, paths);
    const kind = currentListKind(list);
    const isFirstInItem = isFirstBlockInListItem(list);
    markListBlockRendered(list);

    return kind && isFirstInItem
        ? createListParagraph(rendered, kind, list.kinds.length - 1)
        : createBodyParagraph(rendered);
}

function renderTable(tokens: Token[], startIndex: number, paths: GeneratorPaths): { table: Table; nextIndex: number } {
    const rows: TableRow[] = [];
    const weights: number[] = [];
    let i = startIndex + 1;

    while (i < tokens.length && tokens[i].type !== 'table_close') {
        const token = tokens[i];

        if (token.type === 'tr_open') {
            const result = renderTableRow(tokens, i + 1, paths);
            rows.push(result.row);
            for (let col = 0; col < result.widthWeights.length; col++) {
                weights[col] = Math.max(weights[col] ?? 0, result.widthWeights[col]);
            }
            i = result.nextIndex;
            continue;
        }

        i++;
    }

    const columnWidths = normalizeWidths(weights.length > 0 ? weights : [1], TABLE_CONTENT_WIDTH);
    return { table: createTable(rows, columnWidths), nextIndex: i };
}

function renderTableRow(tokens: Token[], startIndex: number, paths: GeneratorPaths): { row: TableRow; nextIndex: number; widthWeights: number[] } {
    const cells: TableCell[] = [];
    const widthWeights: number[] = [];
    let i = startIndex;
    let columnIndex = 0;

    while (i < tokens.length && tokens[i].type !== 'tr_close') {
        const token = tokens[i];

        if (token.type === 'th_open' || token.type === 'td_open') {
            const inline = tokens[i + 1];
            if (inline?.type !== 'inline') {
                i++;
                continue;
            }

            const cellTokens = inline.children ?? [];
            cells.push(createTableCell(renderInlineTokens(cellTokens, paths)));
            widthWeights[columnIndex] = Math.max(widthWeights[columnIndex] ?? 0, estimateInlineWidth(cellTokens));
            columnIndex++;
            i += 3;
            continue;
        }

        i++;
    }

    return { row: createTableRow(cells), nextIndex: i + 1, widthWeights };
}

function estimateInlineWidth(tokens: Token[]): number {
    const text = tokens
        .map((token) => {
            if (token.type === 'text' || token.type === 'code_inline') return token.content;
            if (token.type === 'softbreak' || token.type === 'hardbreak') return ' ';
            if (token.type === 'html_inline' && /<br\s*\/?\s*>/i.test(token.content)) return ' ';
            return '';
        })
        .join(' ')
        .replace(/\s+/g, ' ')
        .trim();

    return Math.max(1, text.length);
}

function normalizeWidths(weights: number[], totalWidth: number): number[] {
    const safeWeights = weights.length > 0 ? weights : [1];
    const normalizedWeights = safeWeights.map((value) => Math.max(1, value));
    const totalWeight = normalizedWeights.reduce((sum, value) => sum + value, 0);

    const exact = normalizedWeights.map((weight) => (totalWidth * weight) / totalWeight);
    const widths = exact.map((value) => Math.max(1, Math.floor(value)));
    let remainder = totalWidth - widths.reduce((sum, value) => sum + value, 0);

    const order = exact
        .map((value, index) => ({ index, fraction: value - Math.floor(value) }))
        .sort((a, b) => b.fraction - a.fraction);

    for (let i = 0; remainder > 0 && order.length > 0; i++) {
        widths[order[i % order.length].index] += 1;
        remainder--;
    }

    while (remainder < 0) {
        const widestIndex = widths.reduce((best, value, index, arr) => (value > arr[best] ? index : best), 0);
        if (widths[widestIndex] <= 1) break;
        widths[widestIndex] -= 1;
        remainder++;
    }

    return widths;
}

function normalizeSource(source: string): string {
    return source.replace(/^\s*\\newpage\s*$/gm, '<!-- DOCX_PAGE_BREAK -->');
}

function extractImageOnlyTokens(children: Token[]): Token[] {
    const images: Token[] = [];
    for (const token of children) {
        if (token.type === 'image') { images.push(token); continue; }
        if (token.type === 'softbreak' || token.type === 'hardbreak') continue;
        if (token.type === 'text' && token.content.trim() === '') continue;
        return [];
    }
    return images;
}

function isParagraphBlock(block: DocxBlock): block is Paragraph {
    return block instanceof Paragraph;
}

function collapseEmpties(paragraphs: DocxBlock[]): DocxBlock[] {
    const out: DocxBlock[] = [];

    for (let i = 0; i < paragraphs.length; i++) {
        const current = paragraphs[i];
        if (!isParagraphBlock(current) || !isEmpty(current)) {
            out.push(current);
            continue;
        }

        let j = i;
        while (j + 1 < paragraphs.length) {
            const nextBlock = paragraphs[j + 1];
            if (!isParagraphBlock(nextBlock) || !isEmpty(nextBlock)) break;
            j++;
        }

        const prev = i > 0 ? paragraphs[i - 1] : undefined;
        const next = j + 1 < paragraphs.length ? paragraphs[j + 1] : undefined;

        const prevNonEmpty = prev !== undefined && (!isParagraphBlock(prev) || !isEmpty(prev));
        const nextNonEmpty = next !== undefined && (!isParagraphBlock(next) || !isEmpty(next));

        if (
            (prevNonEmpty && nextNonEmpty)
            || (prev === undefined && next !== undefined && isParagraphBlock(next) && nextNonEmpty && isHeadingParagraph(next))
            || (next === undefined && prev !== undefined && isParagraphBlock(prev) && prevNonEmpty && isHeadingParagraph(prev))
        ) {
            out.push(current);
        }

        i = j;
    }

    return out;
}
