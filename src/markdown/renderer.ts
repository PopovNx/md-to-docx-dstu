import MarkdownIt, { type Token } from 'markdown-it';
import markdownItAttrs from 'markdown-it-attrs';
import { TableCell, TableRow } from 'docx';

import type { RenderOptions } from '../types';
import {
    type ListRef,
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
import { emptyLine, mark, type DocxBlock } from './blocks';
import { buildImageRun } from './image';
import { inlineText, renderInlineTokens } from './inline';

const markdown = new MarkdownIt({ html: true, breaks: false, linkify: true, typographer: false })
    .use(markdownItAttrs);

const PAGE_BREAK_MARKER = 'DOCX_PAGE_BREAK';

export function renderMarkdown(source: string, opts: RenderOptions): DocxBlock[] {
    const tokens = markdown.parse(source.replace(/^\s*\\newpage\s*$/gm, `<!-- ${PAGE_BREAK_MARKER} -->`), {});
    const out: DocxBlock[] = [];

    const lists: ListRef[] = [];
    let listCount = 0;
    const freshItems: boolean[] = [];
    const takeFreshItem = () => {
        const fresh = freshItems.at(-1) ?? false;
        if (freshItems.length > 0) freshItems[freshItems.length - 1] = false;
        return fresh;
    };

    let rows: TableRow[] = [];
    let cells: TableCell[] = [];
    let headerRow = false;
    let columns: ColumnText[] = [];

    tokens.forEach((token, i) => {
        switch (token.type) {
            case 'html_block':
                if (token.content.includes(PAGE_BREAK_MARKER)) out.push(mark(createPageBreak(), 'pageBreak'));
                break;
            case 'bullet_list_open':
                lists.push({ reference: 'bullets', instance: listCount++ });
                break;
            case 'ordered_list_open':
                lists.push({ reference: token.markup === ')' ? 'numbers-paren' : 'numbers-dot', instance: listCount++ });
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
                out.push(emptyLine());
                break;
            case 'fence':
            case 'code_block':
                out.push(createCodeParagraph(token.content.trimEnd().split('\n')));
                takeFreshItem();
                break;
            case 'table_open':
                rows = [];
                columns = [];
                break;
            case 'tr_open':
                cells = [];
                headerRow = false;
                break;
            case 'tr_close':
                rows.push(createTableRow(cells, headerRow));
                break;
            case 'table_close':
                out.push(createTable(rows, columnWidths(columns.length > 0 ? columns : [{ chars: 1, longestWord: 0 }], opts.contentWidth)), emptyLine());
                takeFreshItem();
                break;
            case 'inline':
                renderInline(token, tokens[i - 1], tokens[i + 2]);
                break;
        }
    });

    function renderInline(inline: Token, parent: Token | undefined, next: Token | undefined) {
        const children = inline.children ?? [];
        const parentType = parent?.type;

        if (parentType === 'heading_open') {
            const level = Number(parent!.tag.slice(1));
            out.push(emptyLine(), mark(createHeadingParagraph(level, renderInlineTokens(children, opts)), 'heading'), emptyLine(true));
            return;
        }

        if (parentType === 'th_open' || parentType === 'td_open') {
            const column = cells.length;
            headerRow ||= parentType === 'th_open';
            cells.push(createTableCell(renderInlineTokens(children, opts), parentType === 'th_open'));
            const text = inlineText(children);
            const stats = (columns[column] ??= { chars: 1, longestWord: 0 });
            stats.chars = Math.max(stats.chars, text.length);
            stats.longestWord = Math.max(stats.longestWord, ...text.split(' ').map(word => word.length));
            return;
        }

        if (parentType !== 'paragraph_open') return;

        const images = imageOnlyTokens(children);
        if (images.length > 0) {
            for (const image of images) {
                const caption = image.content.trim();
                out.push(emptyLine(true), createImageParagraph(buildImageRun(image, opts), { keepNext: caption !== '' }));
                if (caption) out.push(createCaptionParagraph(caption));
                out.push(emptyLine());
            }
            return;
        }

        const runs = renderInlineTokens(children, opts);
        const kind = lists.at(-1);
        const fresh = takeFreshItem();
        const introduces = next?.type === 'table_open';
        out.push(kind && fresh ? createListParagraph(runs, kind, lists.length - 1) : createBodyParagraph(runs, { keepNext: introduces }));
    }

    return out;
}

function imageOnlyTokens(children: Token[]): Token[] {
    const isBlank = (t: Token) => t.type === 'softbreak' || t.type === 'hardbreak' || (t.type === 'text' && !t.content.trim());
    const images = children.filter(t => t.type === 'image');
    return children.every(t => t.type === 'image' || isBlank(t)) ? images : [];
}

type ColumnText = { chars: number; longestWord: number };

const CHAR_TWIPS = 130;
const CELL_PADDING = 216;
const MAX_UNBROKEN_WORD = 16;

const sum = (values: number[]) => values.reduce((total, v) => total + v, 0);

function columnWidths(columns: ColumnText[], totalWidth: number): number[] {
    const minimums = columns.map(c => Math.min(c.longestWord, MAX_UNBROKEN_WORD) * CHAR_TWIPS + CELL_PADDING);
    const spare = totalWidth - sum(minimums);
    const chars = sum(columns.map(c => c.chars));
    const exact = spare <= 0
        ? minimums.map(m => (totalWidth * m) / sum(minimums))
        : minimums.map((m, i) => m + (spare * columns[i]!.chars) / chars);
    const widths = exact.map(Math.floor);
    const leftover = totalWidth - sum(widths);

    exact
        .map((value, index) => ({ index, fraction: value % 1 }))
        .sort((a, b) => b.fraction - a.fraction)
        .slice(0, leftover)
        .forEach(({ index }) => widths[index]!++);

    return widths;
}
