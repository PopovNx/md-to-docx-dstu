import { type Token } from 'markdown-it';
import { ParagraphChild, TextRun } from 'docx';

import type { RenderOptions } from '../types';
import { createTextRun, type TextStyle } from '../docx';
import { buildImageRun } from './image';

const STYLE_TOGGLES: Partial<Record<string, TextStyle>> = {
    strong_open: { bold: true },
    em_open: { italics: true },
    s_open: { strike: true },
};

const isBr = (html: string) => /<br\s*\/?\s*>/i.test(html);

export function renderInlineTokens(tokens: Token[], opts: RenderOptions): ParagraphChild[] {
    const children: ParagraphChild[] = [];
    const styleStack: TextStyle[] = [{}];

    for (const token of tokens) {
        const style = styleStack.at(-1)!;

        switch (token.type) {
            case 'text':
            case 'code_inline':
                children.push(createTextRun(token.content.replace(/\u00a0/g, ' '), style));
                break;
            case 'softbreak':
            case 'hardbreak':
                children.push(new TextRun({ break: 1 }));
                break;
            case 'html_inline':
                if (isBr(token.content)) children.push(new TextRun({ break: 1 }));
                break;
            case 'image':
                children.push(buildImageRun(token, opts));
                break;
            case 'link_open':
                styleStack.push({ ...style, href: token.attrGet('href') ?? '' });
                break;
            default: {
                const toggle = STYLE_TOGGLES[token.type];
                if (toggle) styleStack.push({ ...style, ...toggle });
                else if (token.type.endsWith('_close') && styleStack.length > 1) styleStack.pop();
            }
        }
    }

    return children;
}

export const inlineText = (tokens: Token[]) =>
    tokens
        .map(t => {
            if (t.type === 'text' || t.type === 'code_inline') return t.content;
            if (t.type === 'softbreak' || t.type === 'hardbreak' || (t.type === 'html_inline' && isBr(t.content))) return ' ';
            return '';
        })
        .join(' ')
        .replace(/\s+/g, ' ')
        .trim();
