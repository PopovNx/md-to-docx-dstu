import { Paragraph, Table } from 'docx';

import { createBodyParagraph } from '../docx';

export type DocxBlock = Paragraph | Table;

type Mark = 'empty' | 'heading' | 'pageBreak' | 'chunk';
const marks = new WeakMap<DocxBlock, Mark>();

export const mark = <T extends DocxBlock>(block: T, kind: Mark) => (marks.set(block, kind), block);
export const markOf = (block: DocxBlock | undefined) => (block ? marks.get(block) : undefined);
export const emptyLine = () => mark(createBodyParagraph([]), 'empty');

const startsPage = (block: DocxBlock) => markOf(block) === 'pageBreak' || markOf(block) === 'chunk';

export function collapseEmptyLines(blocks: DocxBlock[]): DocxBlock[] {
    return blocks.filter((block, i) => {
        if (markOf(block) !== 'empty') return true;
        const prev = blocks[i - 1];
        if (markOf(prev) === 'empty') return false;

        let end = i;
        while (markOf(blocks[end + 1]) === 'empty') end++;
        const next = blocks[end + 1];

        if (!prev || !next || startsPage(prev) || startsPage(next)) return false;
        return !(markOf(prev) === 'heading' && markOf(next) === 'heading');
    });
}
