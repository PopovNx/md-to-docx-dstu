import {
    AlignmentType,
    ImageRun,
    PageBreak,
    Paragraph,
    ParagraphChild,
    TableLayoutType,
    WidthType,
    Table,
    TableCell,
    TableRow,
    TextRun,
    UnderlineType,
} from 'docx';

export function createPageBreak(): Paragraph {
    return new Paragraph({ children: [new PageBreak()] });
}

export function createBodyParagraph(children: ParagraphChild[], alignment = AlignmentType.JUSTIFIED) {
    const safeChildren = children.length > 0 ? children : [new TextRun({ text: ' ' })];
    return new Paragraph({ style: 'BodyText', alignment, children: safeChildren });
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

export const TABLE_CONTENT_WIDTH = 9026;

export function createTableCell(children: ParagraphChild[]) {
    return new TableCell({
        children: [new Paragraph({ children })],
    });
}

export function createTableRow(cells: TableCell[]) {
    return new TableRow({ children: cells });
}

export function createTable(rows: TableRow[], columnWidths: number[]) {
    const totalWidth = columnWidths.reduce((sum, value) => sum + value, 0) || TABLE_CONTENT_WIDTH;

    return new Table({
        rows,
        width: { size: totalWidth, type: WidthType.DXA },
        columnWidths,
        layout: TableLayoutType.FIXED,
    });
}

export function chunkPlaceholderStyleId(index: number): string {
    return `InjectChunk_${index}`;
}

export function createChunkPlaceholder(index: number): Paragraph {
    return new Paragraph({ style: chunkPlaceholderStyleId(index) });
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
        bold: bold || undefined,
        italics: italics || undefined,
        strike: strike || undefined,
        color: href ? '0563C1' : undefined,
        underline: href ? { type: UnderlineType.SINGLE } : undefined,
        font: 'Times New Roman',
        size: 28,
    });
}
