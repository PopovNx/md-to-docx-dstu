import type { PageMargins } from '../config';
import type { ListReference } from './paragraphs';
import { AlignmentType, Document, Header, LevelFormat, LevelSuffix, PageNumber, Paragraph, Table, TextRun } from 'docx';

const FONT = 'Times New Roman';
const FONT_SIZE = '14pt';

export const PAGE_WIDTH = 11906;
const PAGE_HEIGHT = 16838;
const PARAGRAPH_INDENT = 567;
const HEADER_DISTANCE = 680;

const sp = () => ({ before: 0, after: 0, line: 360, lineRule: 'auto' as const });

const pageNumberHeader = () => new Header({
    children: [new Paragraph({
        style: 'Normal',
        alignment: AlignmentType.RIGHT,
        spacing: { line: 240, lineRule: 'auto' },
        children: [new TextRun({ children: [PageNumber.CURRENT] })],
    })],
});

export function buildDocument(children: Array<Paragraph | Table>, margins: PageMargins) {
    return new Document({
        sections: [
            {
                properties: {
                    page: {
                        size: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
                        margin: { ...margins, header: HEADER_DISTANCE },
                    },
                },
                headers: { default: pageNumberHeader() },
                children,
            },
        ],
        styles: {
            paragraphStyles: [
                {
                    id: 'Normal',
                    name: 'Normal',
                    run: { font: FONT, size: FONT_SIZE },
                    paragraph: { spacing: { line: 360, lineRule: 'auto' } },
                },
                {
                    id: 'BodyText',
                    name: 'BodyText',
                    basedOn: 'Normal',
                    next: 'BodyText',
                    run: { font: FONT, size: FONT_SIZE },
                    paragraph: {
                        alignment: AlignmentType.JUSTIFIED,
                        indent: { firstLine: PARAGRAPH_INDENT },
                        spacing: sp(),
                    },
                },
                {
                    id: 'ChapterTitle',
                    name: 'ChapterTitle',
                    basedOn: 'Normal',
                    next: 'BodyText',
                    run: { font: FONT, size: FONT_SIZE, bold: true },
                    paragraph: {
                        alignment: AlignmentType.CENTER,
                        keepNext: true,
                        keepLines: true,
                        spacing: sp(),
                        outlineLevel: 0,
                    },
                },
                {
                    id: 'SectionTitle',
                    name: 'SectionTitle',
                    basedOn: 'Normal',
                    next: 'BodyText',
                    run: { font: FONT, size: FONT_SIZE, bold: true },
                    paragraph: {
                        alignment: AlignmentType.LEFT,
                        keepNext: true,
                        keepLines: true,
                        spacing: sp(),
                        outlineLevel: 1,
                        indent: { firstLine: PARAGRAPH_INDENT },
                    },
                },
                {
                    id: 'ImageBlock',
                    name: 'ImageBlock',
                    basedOn: 'Normal',
                    next: 'ImageCaption',
                    run: { font: FONT, size: FONT_SIZE },
                    paragraph: {
                        alignment: AlignmentType.CENTER,
                        indent: { firstLine: 0 },
                        spacing: sp(),
                    },
                },
                {
                    id: 'ImageCaption',
                    name: 'ImageCaption',
                    basedOn: 'Normal',
                    next: 'BodyText',
                    run: { font: FONT, size: FONT_SIZE },
                    paragraph: {
                        alignment: AlignmentType.CENTER,
                        indent: { firstLine: 0 },
                        spacing: sp(),
                    },
                },
                {
                    id: 'CodeBlock',
                    name: 'CodeBlock',
                    basedOn: 'Normal',
                    next: 'BodyText',
                    run: { font: FONT, size: FONT_SIZE },
                    paragraph: {
                        alignment: AlignmentType.LEFT,
                        spacing: sp(),
                        indent: { left: PARAGRAPH_INDENT, right: 0 },
                    },
                },
                {
                    id: 'TableText',
                    name: 'TableText',
                    basedOn: 'Normal',
                    run: { font: FONT, size: '12pt' },
                    paragraph: { spacing: { before: 0, after: 0, line: 240, lineRule: 'auto' } },
                },
            ],
        },
        numbering: {
            config: [
                listNumbering('bullets', LevelFormat.BULLET, () => '–'),
                listNumbering('numbers-dot', LevelFormat.DECIMAL, level => `%${level + 1}.`),
                listNumbering('numbers-paren', LevelFormat.DECIMAL, level => `%${level + 1})`),
            ],
        },
    });
}

function listNumbering(reference: ListReference, format: (typeof LevelFormat)[keyof typeof LevelFormat], text: (level: number) => string) {
    return {
        reference,
        levels: Array.from({ length: 9 }, (_, level) => ({
            level,
            format,
            text: text(level),
            suffix: LevelSuffix.SPACE,
            alignment: AlignmentType.LEFT,
            style: {
                paragraph: { indent: { left: PARAGRAPH_INDENT * level, firstLine: PARAGRAPH_INDENT } },
                run: { font: FONT },
            },
        })),
    };
}
