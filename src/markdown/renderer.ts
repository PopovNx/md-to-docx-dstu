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

        if (token.type === 'bullet_list_open') { listStack.push('bullet'); continue; }
        if (token.type === 'ordered_list_open') { listStack.push('number'); continue; }
        if (token.type === 'bullet_list_close' || token.type === 'ordered_list_close') { listStack.pop(); continue; }
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
            if (Array.isArray(rendered)) {
                blocks.push(...rendered);
            } else if (rendered) {
                blocks.push(rendered);
            }
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
) {
    const imageTokens = extractImageOnlyParagraphTokens(inlineChildren);
    if (imageTokens.length > 0) {
        const imageBlocks: Paragraph[] = [];
        for (const imageToken of imageTokens) {
            const image = buildImageRun(imageToken, paths);
            const caption = imageToken.content.trim();
            imageBlocks.push(createImageParagraph(image));
            if (caption) imageBlocks.push(createCaptionParagraph(caption));
        }
        markListItemBlockRendered(listItemFirstBlockStack);
        return imageBlocks;
    }

    const rendered = renderInlineTokens(inlineChildren, paths);
    const currentListKind = listStack[listStack.length - 1];
    const currentListItemIsFirstBlock = listItemFirstBlockStack[listItemFirstBlockStack.length - 1] ?? false;

    if (currentListKind && currentListItemIsFirstBlock) {
        markListItemBlockRendered(listItemFirstBlockStack);
        return createListParagraph(rendered, currentListKind, listStack.length - 1);
    }

    markListItemBlockRendered(listItemFirstBlockStack);
    return createBodyParagraph(rendered);
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

function markListItemBlockRendered(listItemFirstBlockStack: boolean[]) {
    if (listItemFirstBlockStack.length > 0) {
        listItemFirstBlockStack[listItemFirstBlockStack.length - 1] = false;
    }
}
