import MarkdownIt, { type Token } from 'markdown-it';
import markdownItAttrs from 'markdown-it-attrs';
import { Paragraph, Table, TableCell, TableRow } from 'docx';

import type { RenderOptions } from '../types';
import {
    createBodyParagraph,
    createCaptionParagraph,
    createCodeParagraph,
    createHeadingParagraph,
    createImageParagraph,
    createListParagraph,
    createPageBreak,
    createTable,
    createTableCell,
    createTableRow,
} from '../docx';
import { buildImageRun } from './image';
import { inlineText, renderInlineTokens } from './inline';

const markdown = new MarkdownIt({ html: true, breaks: false, linkify: true, typographer: false })
    .use(markdownItAttrs);

const PAGE_BREAK_MARKER = 'DOCX_PAGE_BREAK';

type DocxBlock = Paragraph | Table;
type ListKind = 'bullet' | 'number';

type Mark = 'empty' | 'heading' | 'pageBreak';
const marks = new WeakMap<DocxBlock, Mark>();
const mark = <T extends DocxBlock>(block: T, kind: Mark) => (marks.set(block, kind), block);
const markOf = (block: DocxBlock | undefined) => (block ? marks.get(block) : undefined);
const empty = () => mark(createBodyParagraph([]), 'empty');

export function renderMarkdown(source: string, opts: RenderOptions): DocxBlock[] {
    const tokens = markdown.parse(source.replace(/^\s*\\newpage\s*$/gm, `<!-- ${PAGE_BREAK_MARKER} -->`), {});
    const out: DocxBlock[] = [];

    const lists: ListKind[] = [];
    const freshItems: boolean[] = [];
    const takeFreshItem = () => {
        const fresh = freshItems.at(-1) ?? false;
        if (freshItems.length > 0) freshItems[freshItems.length - 1] = false;
        return fresh;
    };

    let rows: TableRow[] = [];
    let cells: TableCell[] = [];
    let weights: number[] = [];

    tokens.forEach((token, i) => {
        switch (token.type) {
            case 'html_block':
                if (token.content.includes(PAGE_BREAK_MARKER)) out.push(mark(createPageBreak(), 'pageBreak'));
                break;
            case 'bullet_list_open':
                lists.push('bullet');
                break;
            case 'ordered_list_open':
                lists.push('number');
                break;
            case 'bullet_list_close':
            case 'ordered_list_close':
                lists.pop();
                break;
            case 'list_item_open':
                freshItems.push(true);
                break;
            case 'list_item_close':
                freshItems.pop();
                break;
            case 'hr':
                out.push(empty());
                break;
            case 'fence':
            case 'code_block':
                out.push(createCodeParagraph(token.content.trimEnd().split('\n')));
                takeFreshItem();
                break;
            case 'table_open':
                rows = [];
                weights = [];
                break;
            case 'tr_open':
                cells = [];
                break;
            case 'tr_close':
                rows.push(createTableRow(cells));
                break;
            case 'table_close':
                out.push(createTable(rows, splitWidth(weights.length > 0 ? weights : [1], opts.contentWidth)), empty());
                takeFreshItem();
                break;
            case 'inline':
                renderInline(token, tokens[i - 1]);
                break;
        }
    });

    function renderInline(inline: Token, parent: Token | undefined) {
        const children = inline.children ?? [];
        const parentType = parent?.type;

        if (parentType === 'heading_open') {
            const level = Number(parent!.tag.slice(1));
            const prev = out.at(-1);
            if (level !== 1 || (prev && markOf(prev) !== 'pageBreak')) out.push(empty());
            out.push(mark(createHeadingParagraph(level, renderInlineTokens(children, opts)), 'heading'), empty());
            return;
        }

        if (parentType === 'th_open' || parentType === 'td_open') {
            const column = cells.length;
            cells.push(createTableCell(renderInlineTokens(children, opts)));
            weights[column] = Math.max(weights[column] ?? 0, inlineText(children).length, 1);
            return;
        }

        if (parentType !== 'paragraph_open') return;

        const images = imageOnlyTokens(children);
        if (images.length > 0) {
            for (const image of images) {
                const caption = image.content.trim();
                out.push(empty(), createImageParagraph(buildImageRun(image, opts)));
                if (caption) out.push(createCaptionParagraph(caption));
                out.push(empty());
            }
            return;
        }

        const runs = renderInlineTokens(children, opts);
        const kind = lists.at(-1);
        const fresh = takeFreshItem();
        out.push(kind && fresh ? createListParagraph(runs, kind, lists.length - 1) : createBodyParagraph(runs));
    }

    return collapseEmpties(out);
}

function imageOnlyTokens(children: Token[]): Token[] {
    const isBlank = (t: Token) => t.type === 'softbreak' || t.type === 'hardbreak' || (t.type === 'text' && !t.content.trim());
    const images = children.filter(t => t.type === 'image');
    return children.every(t => t.type === 'image' || isBlank(t)) ? images : [];
}

function splitWidth(weights: number[], totalWidth: number): number[] {
    const weightSum = weights.reduce((sum, w) => sum + w, 0);
    const exact = weights.map(w => (totalWidth * w) / weightSum);
    const widths = exact.map(Math.floor);
    const leftover = totalWidth - widths.reduce((sum, w) => sum + w, 0);

    exact
        .map((value, index) => ({ index, fraction: value % 1 }))
        .sort((a, b) => b.fraction - a.fraction)
        .slice(0, leftover)
        .forEach(({ index }) => widths[index]!++);

    return widths;
}

function collapseEmpties(blocks: DocxBlock[]): DocxBlock[] {
    return blocks.filter((block, i) => {
        if (markOf(block) !== 'empty') return true;
        const prev = blocks[i - 1];
        if (markOf(prev) === 'empty') return false;

        let end = i;
        while (markOf(blocks[end + 1]) === 'empty') end++;
        const next = blocks[end + 1];

        if (prev && next) return true;
        return markOf(prev ?? next) === 'heading';
    });
}

export const isPageBreak = (block: DocxBlock | undefined) => markOf(block) === 'pageBreak';
