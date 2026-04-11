import { type Token } from 'markdown-it';
import { ParagraphChild, TextRun } from 'docx';

import type { GeneratorPaths } from '../types';
import { createLinkTextRun } from '../docx';
import { buildImageRun } from './image';

type InlineStyleState = {
    bold: boolean;
    italic: boolean;
    strike: boolean;
    linkHref?: string;
};

const STYLE_TOGGLES: Partial<Record<string, Partial<InlineStyleState>>> = {
    strong_open: { bold: true },
    em_open: { italic: true },
    s_open: { strike: true },
};

export function renderInlineTokens(tokens: Token[], paths: GeneratorPaths): ParagraphChild[] {
    const children: ParagraphChild[] = [];
    const styleStack: InlineStyleState[] = [{ bold: false, italic: false, strike: false }];

    for (const token of tokens) {
        const style = styleStack[styleStack.length - 1];

        if (token.type === 'text') {
            children.push(createLinkTextRun(token.content.replace(/\u00a0/g, ' '), style.linkHref, style.bold, style.italic, style.strike));
            continue;
        }

        if (token.type === 'softbreak' || token.type === 'hardbreak') {
            children.push(new TextRun({ break: 1 }));
            continue;
        }

        if (token.type === 'code_inline') {
            children.push(new TextRun({ text: token.content, font: 'Courier New', size: 20 }));
            continue;
        }

        if (token.type === 'html_inline') {
            if (/<br\s*\/?\s*>/i.test(token.content)) children.push(new TextRun({ break: 1 }));
            continue;
        }

        if (token.type === 'image') {
            children.push(buildImageRun(token, paths));
            continue;
        }

        if (token.type === 'link_open') {
            styleStack.push({ ...style, linkHref: token.attrGet('href') ?? '' });
            continue;
        }

        const styleChange = STYLE_TOGGLES[token.type];
        if (styleChange !== undefined) {
            styleStack.push({ ...style, ...styleChange });
            continue;
        }

        if (token.type.endsWith('_close') && styleStack.length > 1) {
            styleStack.pop();
        }
    }

    return children;
}
