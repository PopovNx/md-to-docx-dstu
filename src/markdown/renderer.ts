import { type Token } from 'markdown-it';
import MarkdownIt from 'markdown-it';
import markdownItAttrs from 'markdown-it-attrs';
import { Paragraph } from 'docx';

import type { GeneratorPaths } from '../types';
import {
    createBodyParagraph,
    createCaptionParagraph,
    createCodeParagraph,
    createHeadingParagraph,
    createImageParagraph,
    createListParagraph,
    createPageBreak,
} from '../docx';
import { buildImageRun } from './image';
import { renderInlineTokens } from './inline';

const markdown = new MarkdownIt({ html: true, breaks: false, linkify: true, typographer: false })
    .use(markdownItAttrs);

const LIST_OPEN_TOKENS: Partial<Record<string, 'bullet' | 'number'>> = {
    bullet_list_open: 'bullet',
    ordered_list_open: 'number',
};

const emptySet = new WeakSet<Paragraph>();
const headingSet = new WeakSet<Paragraph>();

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

export function renderMarkdown(source: string, paths: GeneratorPaths): Paragraph[] {
    const tokens = markdown.parse(normalizeSource(source), {});
    const raw: Paragraph[] = [];
    const list: ListState = { kinds: [], firstBlock: [] };

    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];

        if (token.type === 'html_block' && token.content.includes('DOCX_PAGE_BREAK')) {
            raw.push(createPageBreak());
            continue;
        }

        if (token.type === 'heading_open') {
            const inline = tokens[i + 1];
            if (inline?.type !== 'inline') continue;
            const level = Number(token.tag.replace('h', '')) as 1 | 2;
            const heading = createHeadingParagraph(level, renderInlineTokens(inline.children ?? [], paths));
            headingSet.add(heading);
            raw.push(createEmpty(), heading, createEmpty());
            i += 2;
            continue;
        }

        const listKind = LIST_OPEN_TOKENS[token.type];
        if (listKind !== undefined) { list.kinds.push(listKind); continue; }
        if (token.type.endsWith('_list_close')) { list.kinds.pop(); continue; }
        if (token.type === 'list_item_open') { list.firstBlock.push(true); continue; }
        if (token.type === 'list_item_close') { list.firstBlock.pop(); continue; }

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

function collapseEmpties(paragraphs: Paragraph[]): Paragraph[] {
    const out: Paragraph[] = [];

    for (let i = 0; i < paragraphs.length; i++) {
        const current = paragraphs[i];
        if (!isEmpty(current)) {
            out.push(current);
            continue;
        }

        let j = i;
        while (j + 1 < paragraphs.length && isEmpty(paragraphs[j + 1])) j++;

        const prev = i > 0 ? paragraphs[i - 1] : undefined;
        const next = j + 1 < paragraphs.length ? paragraphs[j + 1] : undefined;

        const prevNonEmpty = prev !== undefined && !isEmpty(prev);
        const nextNonEmpty = next !== undefined && !isEmpty(next);

        if (
            (prevNonEmpty && nextNonEmpty)
            || (prev === undefined && nextNonEmpty && isHeadingParagraph(next))
            || (next === undefined && prevNonEmpty && isHeadingParagraph(prev))
        ) {
            out.push(current);
        }

        i = j;
    }

    return out;
}
