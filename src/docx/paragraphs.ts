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
    return new Paragraph({ style: 'Normal', children: [new PageBreak()] });
}

export function createBodyParagraph(children: ParagraphChild[], alignment = AlignmentType.JUSTIFIED) {
    const safeChildren = children.length > 0 ? children : [new TextRun({ text: ' ' })];
    return new Paragraph({ style: 'BodyText', alignment, children: safeChildren });
}

export function createHeadingParagraph(level: number, children: ParagraphChild[]) {
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

export function createTableCell(children: ParagraphChild[]) {
    return new TableCell({
        children: [new Paragraph({ style: 'Normal', children })],
    });
}

export function createTableRow(cells: TableCell[]) {
    return new TableRow({ children: cells });
}

export function createTable(rows: TableRow[], columnWidths: number[]) {
    return new Table({
        rows,
        width: { size: columnWidths.reduce((sum, value) => sum + value, 0), type: WidthType.DXA },
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

export type TextStyle = { bold?: true; italics?: true; strike?: true; href?: string };

export function createTextRun(text: string, { bold, italics, strike, href }: TextStyle = {}) {
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
