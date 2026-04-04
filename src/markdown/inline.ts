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

function defaultStyle(): InlineStyleState {
    return { bold: false, italic: false, strike: false };
}

export function renderInlineTokens(tokens: Token[], paths: GeneratorPaths): ParagraphChild[] {
    const children: ParagraphChild[] = [];
    const styleStack: InlineStyleState[] = [defaultStyle()];

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
            if (/<br\s*\/?\s*>/i.test(token.content)) {
                children.push(new TextRun({ break: 1 }));
            }
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

        if (token.type === 'link_close') {
            if (styleStack.length > 1) styleStack.pop();
            continue;
        }

        if (token.type === 'strong_open') {
            styleStack.push({ ...style, bold: true });
            continue;
        }

        if (token.type === 'strong_close') {
            if (styleStack.length > 1) styleStack.pop();
            continue;
        }

        if (token.type === 'em_open') {
            styleStack.push({ ...style, italic: true });
            continue;
        }

        if (token.type === 'em_close') {
            if (styleStack.length > 1) styleStack.pop();
            continue;
        }

        if (token.type === 's_open') {
            styleStack.push({ ...style, strike: true });
            continue;
        }

        if (token.type === 's_close') {
            if (styleStack.length > 1) styleStack.pop();
            continue;
        }
    }

    return children;
}
