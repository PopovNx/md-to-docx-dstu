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

const markdown = new MarkdownIt({
    html: true,
    breaks: false,
    linkify: true,
    typographer: false,
}).use(markdownItAttrs);

const LIST_OPEN_TOKENS: Partial<Record<string, 'bullet' | 'number'>> = {
    bullet_list_open: 'bullet',
    ordered_list_open: 'number',
};

export function renderMarkdown(source: string, paths: GeneratorPaths): Paragraph[] {
    const prepared = normalizeSource(source);
    const tokens = markdown.parse(prepared, {});
    const blocks: Paragraph[] = [];

    const listStack: Array<'bullet' | 'number'> = [];
    const listItemFirstBlockStack: boolean[] = [];

    for (let index = 0; index < tokens.length; index += 1) {
        const token = tokens[index];

        if (token.type === 'html_block' && token.content.includes('DOCX_PAGE_BREAK')) {
            blocks.push(createPageBreak());
            continue;
        }

        if (token.type === 'heading_open') {
            const inline = tokens[index + 1];
            if (inline?.type !== 'inline') continue;
            const level = Number(token.tag.replace('h', '')) as 1 | 2;
            blocks.push(createHeadingParagraph(level, renderInlineTokens(inline.children ?? [], paths)));
            index += 2;
            continue;
        }

        const listKind = LIST_OPEN_TOKENS[token.type];
        if (listKind !== undefined) { listStack.push(listKind); continue; }
        if (token.type.endsWith('_list_close')) { listStack.pop(); continue; }
        if (token.type === 'list_item_open') { listItemFirstBlockStack.push(true); continue; }
        if (token.type === 'list_item_close') { listItemFirstBlockStack.pop(); continue; }

        if (token.type === 'fence' || token.type === 'code_block') {
            blocks.push(createCodeParagraph(token.content.replace(/\s+$/, '').split('\n')));
            markListItemBlockRendered(listItemFirstBlockStack);
            continue;
        }

        if (token.type === 'paragraph_open') {
            const inline = tokens[index + 1];
            if (inline?.type !== 'inline') continue;
            const rendered = renderParagraphFromInline(inline.children ?? [], paths, listStack, listItemFirstBlockStack);
            if (Array.isArray(rendered)) blocks.push(...rendered);
            else blocks.push(rendered);
            index += 2;
            continue;
        }
    }

    return blocks;
}

function renderParagraphFromInline(
    inlineChildren: Token[],
    paths: GeneratorPaths,
    listStack: Array<'bullet' | 'number'>,
    listItemFirstBlockStack: boolean[],
): Paragraph | Paragraph[] {
    const imageTokens = extractImageOnlyParagraphTokens(inlineChildren);
    let result: Paragraph | Paragraph[];

    if (imageTokens.length > 0) {
        const imageBlocks = imageTokens.flatMap((imageToken) => {
            const caption = imageToken.content.trim();
            return caption
                ? [createImageParagraph(buildImageRun(imageToken, paths)), createCaptionParagraph(caption)]
                : [createImageParagraph(buildImageRun(imageToken, paths))];
        });
        result = [createBodyParagraph([]), ...imageBlocks, createBodyParagraph([])];
    } else {
        const rendered = renderInlineTokens(inlineChildren, paths);
        const currentListKind = listStack[listStack.length - 1];
        const isFirstListBlock = listItemFirstBlockStack[listItemFirstBlockStack.length - 1] ?? false;

        result = currentListKind && isFirstListBlock
            ? createListParagraph(rendered, currentListKind, listStack.length - 1)
            : createBodyParagraph(rendered);
    }

    markListItemBlockRendered(listItemFirstBlockStack);
    return result;
}

function normalizeSource(source: string) {
    return source.replace(/^\s*\\newpage\s*$/gm, '<!-- DOCX_PAGE_BREAK -->');
}

function extractImageOnlyParagraphTokens(inlineChildren: Token[]): Token[] {
    const images: Token[] = [];
    for (const token of inlineChildren) {
        if (token.type === 'image') { images.push(token); continue; }
        if (token.type === 'softbreak' || token.type === 'hardbreak') continue;
        if (token.type === 'text' && token.content.trim() === '') continue;
        return [];
    }
    return images;
}

function markListItemBlockRendered(stack: boolean[]) {
    if (stack.length > 0) stack[stack.length - 1] = false;
}
