import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { test } from 'node:test';

import JSZip from 'jszip';
import { AlignmentType, Document, Header, ImageRun, LevelFormat, Packer, Paragraph, Table, TableCell, TableRow, TextRun } from 'docx';

import { DEFAULT_MARGINS } from '../config';
import { buildDocument, PAGE_WIDTH } from '../docx';
import { finishPages, loadInputFiles, renderInputFile, type RenderResult } from '../markdown';
import { injectDocxChunks } from '.';
import { PartCopier } from './parts';

const CHUNK_MARGIN = 1440;

type ChunkOptions = { text?: string[]; header?: string; images?: string[]; list?: boolean; endsWithTable?: boolean };

async function chunkDocx({ text = ['Chunk text'], header, images = [], list, endsWithTable }: ChunkOptions = {}): Promise<Buffer> {
    const children: Array<Paragraph | Table> = text.map((t, i) => new Paragraph({ style: i === 0 ? 'Fancy' : undefined, children: [new TextRun(t)] }));
    if (list) children.push(new Paragraph({ numbering: { reference: 'chunk-list', level: 0 }, children: [new TextRun('List item')] }));
    for (const image of images) {
        children.push(new Paragraph({ children: [new ImageRun({ type: 'png', data: Buffer.from(image), transformation: { width: 10, height: 10 } })] }));
    }
    if (endsWithTable) children.push(new Table({ rows: [new TableRow({ children: [new TableCell({ children: [new Paragraph('Cell')] })] })] }));

    return Packer.toBuffer(new Document({
        styles: {
            default: { document: { run: { font: 'Arial', size: 22 }, paragraph: { spacing: { after: 160 } } } },
            paragraphStyles: [
                { id: 'Fancy', name: 'Fancy', basedOn: 'Normal', run: { italics: true } },
                { id: 'Plain', name: 'Plain', run: { bold: true } },
            ],
        },
        numbering: {
            config: [{
                reference: 'chunk-list',
                levels: [{ level: 0, format: LevelFormat.LOWER_LETTER, text: '%1)', alignment: AlignmentType.LEFT }],
            }],
        },
        sections: [{
            properties: { page: { margin: { top: CHUNK_MARGIN, right: CHUNK_MARGIN, bottom: CHUNK_MARGIN, left: CHUNK_MARGIN } } },
            headers: header ? { default: new Header({ children: [new Paragraph(header)] }) } : undefined,
            children,
        }],
    }));
}

async function render(pages: Record<string, string | Buffer>, vars: Record<string, string> = {}) {
    const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'md2docx-'));
    for (const [name, content] of Object.entries(pages)) await fs.writeFile(path.join(dir, name), content);

    const result: RenderResult = { paragraphs: [], chunks: [] };
    const opts = { imageDir: dir, contentWidth: PAGE_WIDTH - DEFAULT_MARGINS.left - DEFAULT_MARGINS.right };
    for (const file of await loadInputFiles(dir)) await renderInputFile(file, opts, vars, result);

    const { paragraphs, chunks } = finishPages(result);
    const packed = await Packer.toBuffer(buildDocument(paragraphs, DEFAULT_MARGINS));
    const zip = await JSZip.loadAsync(await injectDocxChunks(packed, chunks, vars));
    const read = (name: string) => zip.file(name)!.async('string');
    return { zip, read, document: await read('word/document.xml') };
}

const sectPrs = (xml: string) => [...xml.matchAll(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)].map(([s]) => s);
const leftMargin = (sectPr: string) => Number(sectPr.match(/<w:pgMar [^>]*w:left="(\d+)"/)?.[1]);
const texts = (xml: string) => [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(([, t = '']) => t).filter(t => t.trim());

const sectionEnds = (xml: string) => [...xml.matchAll(/<w:p>(?:(?!<w:p>)[\s\S])*?<w:sectPr[\s\S]*?<\/w:sectPr>[\s\S]*?<\/w:p>/g)]
    .map(([p]) => ({ text: texts(p).join(''), left: leftMargin(p) }));

const styleXml = (styles: string, id: string) =>
    styles.match(new RegExp(`<w:style [^>]*w:styleId="${id}"[^>]*>[\\s\\S]*?</w:style>`))?.[0] ?? '';
function styleChain(styles: string, id: string): string[] {
    const chain = [id];
    for (let parent = attrOf(styleXml(styles, id), 'basedOn'); parent && !chain.includes(parent); parent = attrOf(styleXml(styles, parent), 'basedOn')) {
        chain.push(parent);
    }
    return chain;
}
const attrOf = (xml: string, element: string) => xml.match(new RegExp(`<w:${element} w:val="([^"]+)"`))?.[1];

const paragraphStyle = (document: string, text: string) =>
    attrOf(document.match(new RegExp(`<w:p>(?:(?!<w:p>)[\\s\\S])*?>${text}</w:t>`))?.[0] ?? '', 'pStyle') ?? '';

const emptySectionParagraphs = (xml: string) =>
    xml.match(/<w:p><w:pPr>(?:<w:pStyle w:val="[^"]*"\/>)?<w:sectPr[\s\S]*?<\/w:sectPr><\/w:pPr><\/w:p>/g) ?? [];

test('title page at the start keeps its page setup, styles and header', async () => {
    const { document, read } = await render({
        '0.docx': await chunkDocx({ text: ['Title [[NAME]]', 'Second line'], header: 'Header of [[NAME]]' }),
        '1.md': '# Intro\n\nBody text',
    }, { NAME: 'Popov' });

    assert.deepEqual(texts(document), ['Title Popov', 'Second line', 'Intro', 'Body text']);
    assert.deepEqual(sectionEnds(document), [{ text: 'Second line', left: CHUNK_MARGIN }]);
    assert.equal(leftMargin(sectPrs(document).at(-1)!), DEFAULT_MARGINS.left);
    assert.deepEqual(emptySectionParagraphs(document), []);

    const styles = await read('word/styles.xml');
    assert.doesNotMatch(styles.match(/<w:docDefaults>[\s\S]*?<\/w:docDefaults>/)![0], /rFonts|w:sz|spacing/);
    assert.match(styleXml(styles, 'ch0_DocDefaults'), /w:after="160"[\s\S]*?w:ascii="Arial"/);
    assert.equal(styleChain(styles, 'ch0_Fancy').at(-1), 'ch0_DocDefaults');
    assert.deepEqual(styleChain(styles, 'ch0_Plain'), ['ch0_Plain', 'ch0_DocDefaults']);
    assert.equal(paragraphStyle(document, 'Title Popov'), 'ch0_Fancy');
    assert.equal(styleChain(styles, paragraphStyle(document, 'Second line')).at(-1), 'ch0_DocDefaults');

    const [chunkHeaderRef, mainHeaderRef] = sectPrs(document).map(s => s.match(/<w:headerReference [^>]*r:id="([^"]+)"/)?.[1]);
    const rels = await read('word/_rels/document.xml.rels');
    const targetOf = (id: string | undefined) => rels.match(new RegExp(`Id="${id}"[^>]*Target="([^"]+)"`))?.[1];
    assert.match(targetOf(chunkHeaderRef)!, /^ch0_header\d*\.xml$/);
    assert.equal(texts(await read(`word/${targetOf(chunkHeaderRef)}`)).join(''), 'Header of Popov');
    assert.match(await read(`word/${targetOf(mainHeaderRef)}`), /PAGE/);
    assert.match(await read('[Content_Types].xml'), new RegExp(`PartName="/word/${targetOf(chunkHeaderRef)}"`));
});

test('chunk between markdown pages closes sections on real paragraphs', async () => {
    const { document } = await render({
        '1.md': 'Before text',
        '2.docx': await chunkDocx(),
        '3.md': 'After text',
    });

    assert.deepEqual(texts(document), ['Before text', 'Chunk text', 'After text']);
    assert.deepEqual(sectionEnds(document), [
        { text: 'Before text', left: DEFAULT_MARGINS.left },
        { text: 'Chunk text', left: CHUNK_MARGIN },
    ]);
    assert.equal(leftMargin(sectPrs(document).at(-1)!), DEFAULT_MARGINS.left);
    assert.deepEqual(emptySectionParagraphs(document), []);
});

test('chunk at the end becomes the final section instead of leaving an empty one', async () => {
    const { document } = await render({ '1.md': 'Body text', '2.docx': await chunkDocx() });

    assert.deepEqual(sectionEnds(document), [{ text: 'Body text', left: DEFAULT_MARGINS.left }]);
    assert.equal(leftMargin(sectPrs(document).at(-1)!), CHUNK_MARGIN);
    assert.deepEqual(emptySectionParagraphs(document), []);
});

test('consecutive chunks and \\newpage next to a chunk add no blank pages', async () => {
    const { document } = await render({
        '1.md': 'A\n\n\\newpage',
        '2.docx': await chunkDocx({ text: ['First chunk'] }),
        '3.docx': await chunkDocx({ text: ['Second chunk'] }),
        '4.md': '\\newpage\n\nB',
    });

    assert.deepEqual(texts(document), ['A', 'First chunk', 'Second chunk', 'B']);
    assert.doesNotMatch(document, /w:type="page"/);
    assert.deepEqual(sectionEnds(document).map(s => s.text), ['A', 'First chunk', 'Second chunk']);
    assert.equal(sectPrs(document).length, 4);
    assert.deepEqual(emptySectionParagraphs(document), []);
});

test('blank lines collapse across files: none between headings split over two files', async () => {
    const { document } = await render({ '1.md': '# 1 Chapter', '2.md': '## 1.1 Section\n\nText' });
    const between = document.match(/1 Chapter<\/w:t>([\s\S]*?)1\.1 Section/)![1]!;
    assert.doesNotMatch(between, /<w:t[^>]*> <\/w:t>/);
    assert.match(document, /1\.1 Section<\/w:t>[\s\S]*?<w:t[^>]*> <\/w:t>[\s\S]*?Text/);
});

test('chunk ending with a table gets its own section-break paragraph', async () => {
    const { document, read } = await render({ '1.docx': await chunkDocx({ endsWithTable: true }), '2.md': 'After' });

    const sectionStyle = document.match(/<\/w:tbl><w:p><w:pPr><w:pStyle w:val="(ch0_\w+)"\/><w:sectPr/)?.[1];
    assert.ok(sectionStyle, 'section-break paragraph right after the table, in a chunk style');
    assert.equal(styleChain(await read('word/styles.xml'), sectionStyle).at(-1), 'ch0_DocDefaults');
    assert.equal(emptySectionParagraphs(document).length, 1);
});

test('chunk images and lists keep pointing at their own definitions', async () => {
    const images = ['img-a', 'img-b', 'img-c', 'img-d', 'img-e', 'img-f'];
    const { zip, read, document } = await render({
        '1.md': 'Before',
        '2.docx': await chunkDocx({ images, list: true }),
    });

    const rels = await read('word/_rels/document.xml.rels');
    const embedded = await Promise.all([...document.matchAll(/r:embed="([^"]+)"/g)].map(async ([, id]) => {
        const target = rels.match(new RegExp(`Id="${id}"[^>]*Target="([^"]+)"`))![1];
        return zip.file(`word/${target}`)!.async('string');
    }));
    assert.deepEqual(embedded, images);

    const numId = document.match(/<w:numId w:val="(\d+)"\/>[\s\S]*?List item/)![1];
    const numbering = await read('word/numbering.xml');
    const abstractId = numbering.match(new RegExp(`<w:num [^>]*w:numId="${numId}"[^>]*>\\s*<w:abstractNumId w:val="(\\d+)"`))![1];
    assert.match(numbering, new RegExp(`<w:abstractNum [^>]*w:abstractNumId="${abstractId}"[\\s\\S]*?<w:lvlText w:val="%1\\)"/>`));
    assert.ok(numbering.lastIndexOf('<w:abstractNum ') < numbering.indexOf('<w:num '), 'abstractNum elements precede num elements');
});

test('paragraph style missing from the chunk falls back to the chunk default, not to a main style', async () => {
    const zip = await JSZip.loadAsync(await chunkDocx());
    const xml = await zip.file('word/document.xml')!.async('string');
    zip.file('word/document.xml', xml.replace('<w:pStyle w:val="Fancy"/>', '<w:pStyle w:val="BodyText"/>'));

    const { document, read } = await render({ '1.docx': await zip.generateAsync({ type: 'nodebuffer' }), '2.md': 'After' });

    assert.equal(styleChain(await read('word/styles.xml'), paragraphStyle(document, 'Chunk text')).at(-1), 'ch0_DocDefaults');
    assert.doesNotMatch(document, /ch0_BodyText/);
});

test('copied media of a new type registers its content type', async () => {
    const src = new JSZip();
    src.file('[Content_Types].xml', '<Types><Default Extension="emf" ContentType="image/x-emf"/></Types>');
    src.file('word/media/logo.emf', 'emf-bytes');
    const dst = new JSZip();
    dst.file('[Content_Types].xml', '<Types><Default Extension="png" ContentType="image/png"/></Types>');

    const copied = await new PartCopier(src, dst, 'ch0_', xml => xml).copy('word/media/logo.emf');

    assert.equal(copied, 'word/media/ch0_logo.emf');
    assert.equal(await dst.file(copied)!.async('string'), 'emf-bytes');
    assert.match(await dst.file('[Content_Types].xml')!.async('string'), /<Default Extension="emf" ContentType="image\/x-emf"\/>/);
});
