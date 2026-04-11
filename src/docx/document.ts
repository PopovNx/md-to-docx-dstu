import { AlignmentType, Document, LevelFormat, Paragraph } from 'docx';

const FONT = 'Times New Roman';
const FONT_SIZE = '14pt';

const sp = () => ({ before: 0, after: 0, line: 360, lineRule: 'auto' as const });

const PAGE_MARGINS = {
    top: 1440,
    right: 1440,
    bottom: 1440,
    left: 1440,
};

export function buildDocument(children: Paragraph[]) {
    return new Document({
        sections: [
            {
                properties: {
                    page: {
                        size: {
                            width: 11906,
                            height: 16838,
                        },
                        margin: PAGE_MARGINS,
                    },
                },
                children,
            },
        ],
        styles: {
            default: {
                document: {
                    run: {
                        font: FONT,
                        size: FONT_SIZE,
                    },
                    paragraph: {
                        spacing: {
                            line: 360,
                            lineRule: 'auto',
                        },
                    },
                },
            },
            paragraphStyles: [
                {
                    id: 'BodyText',
                    name: 'BodyText',
                    basedOn: 'Normal',
                    next: 'BodyText',
                    run: { font: FONT, size: FONT_SIZE },
                    paragraph: {
                        alignment: AlignmentType.JUSTIFIED,
                        indent: { firstLine: '1cm' },
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
                        spacing: sp(),
                        outlineLevel: 1,
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
                        indent: { left: 720, right: 0 },
                    },
                },
            ],
        },
        numbering: {
            config: [
                {
                    reference: 'bullets',
                    levels: Array.from({ length: 9 }, (_, level) => ({
                        level,
                        format: LevelFormat.BULLET,
                        text: '\uF02D',
                        alignment: AlignmentType.LEFT,
                        style: {
                            paragraph: {
                                indent: { left: 720 + level * 360, hanging: 360 },
                            },
                            run: { font: 'Symbol' },
                        },
                    })),
                },
                {
                    reference: 'numbers',
                    levels: Array.from({ length: 9 }, (_, level) => ({
                        level,
                        format: LevelFormat.DECIMAL,
                        text: `${Array.from({ length: level + 1 }, (_value, index) => `%${index + 1}`).join('.')}.`,
                        alignment: AlignmentType.LEFT,
                        style: {
                            paragraph: {
                                indent: { left: 720 + level * 360, hanging: 360 },
                            },
                        },
                    })),
                },
            ],
        },
    });
}
