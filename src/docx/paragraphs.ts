import {
    AlignmentType,
    ImageRun,
    PageBreak,
    Paragraph,
    ParagraphChild,
    TextRun,
    UnderlineType,
} from 'docx';

export function createPageBreak(): Paragraph {
    return new Paragraph({ children: [new PageBreak()] });
}

export function createBodyParagraph(children: ParagraphChild[], alignment = AlignmentType.JUSTIFIED) {
    return new Paragraph({ style: 'BodyText', alignment, children });
}

export function createHeadingParagraph(level: 1 | 2, children: ParagraphChild[]) {
    return new Paragraph({
        style: level === 1 ? 'ChapterTitle' : 'SectionTitle',
        children,
    });
}

export function createListParagraph(children: ParagraphChild[], listKind: 'bullet' | 'number', level: number) {
    return new Paragraph({
        style: 'BodyText',
        numbering: {
            reference: listKind === 'bullet' ? 'bullets' : 'numbers',
            level,
        },
        children,
    });
}

export function createCodeParagraph(lines: string[]) {
    const children = lines.flatMap((line, index) => {
        const run = new TextRun({ text: line.length > 0 ? line : ' ' });
        return index < lines.length - 1 ? [run, new TextRun({ break: 1 })] : [run];
    });
    return new Paragraph({ style: 'CodeBlock', children });
}

export function createImageParagraph(image: ImageRun) {
    return new Paragraph({
        style: 'ImageBlock',
        alignment: AlignmentType.CENTER,
        children: [image],
    });
}

export function createCaptionParagraph(text: string) {
    return new Paragraph({
        style: 'ImageCaption',
        children: [new TextRun({ text })],
    });
}

export function createLinkTextRun(text: string, href?: string, bold?: boolean, italics?: boolean, strike?: boolean) {
    return new TextRun({
        text,
        bold,
        italics,
        strike,
        color: href ? '0563C1' : undefined,
        underline: href ? { type: UnderlineType.SINGLE } : undefined,
        font: 'Times New Roman',
        size: 28,
    });
}
