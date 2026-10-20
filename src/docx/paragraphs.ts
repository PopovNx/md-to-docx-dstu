import {
    AlignmentType,
    ImageRun,
    Paragraph,
    ParagraphChild,
    Table,
    TableCell,
    TableLayoutType,
    TableRow,
    TextRun,
    WidthType,
} from 'docx';

export type ListRef = { reference: ListReference; instance: number };
export type ListReference = 'bullets' | 'numbers-dot' | 'numbers-paren';

export function createPageBreak(): Paragraph {
    return new Paragraph({ style: 'Normal', pageBreakBefore: true, children: [] });
}

export function createBodyParagraph(children: ParagraphChild[], { keepNext = false } = {}) {
    return new Paragraph({
        style: 'BodyText',
        keepNext: keepNext || undefined,
        children: children.length > 0 ? children : [new TextRun({ text: ' ' })],
    });
}

export function createHeadingParagraph(level: number, children: ParagraphChild[]) {
    return new Paragraph({ style: level === 1 ? 'ChapterTitle' : 'SectionTitle', children });
}

export function createListParagraph(children: ParagraphChild[], { reference, instance }: ListRef, level: number) {
    return new Paragraph({ style: 'BodyText', numbering: { reference, instance, level }, children });
}

export function createCodeParagraph(lines: string[]) {
    const children = lines.flatMap((line, index) => {
        const run = new TextRun({ text: line.length > 0 ? line : ' ' });
        return index < lines.length - 1 ? [run, new TextRun({ break: 1 })] : [run];
    });
    return new Paragraph({ style: 'CodeBlock', children });
}

export function createImageParagraph(image: ImageRun, { keepNext = false } = {}) {
    return new Paragraph({ style: 'ImageBlock', keepNext: keepNext || undefined, children: [image] });
}

export function createCaptionParagraph(text: string) {
    return new Paragraph({ style: 'ImageCaption', children: [new TextRun({ text })] });
}

export function createTableCell(children: ParagraphChild[], header: boolean) {
    return new TableCell({
        children: [new Paragraph({ style: 'TableText', alignment: header ? AlignmentType.CENTER : undefined, children })],
    });
}

export function createTableRow(cells: TableCell[], header: boolean) {
    return new TableRow({ children: cells, tableHeader: header || undefined });
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

export type TextStyle = { bold?: true; italics?: true; strike?: true };

export function createTextRun(text: string, { bold, italics, strike }: TextStyle = {}) {
    return new TextRun({ text, bold, italics, strike });
}
