import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { test } from 'node:test';

import JSZip from 'jszip';
import { Packer } from 'docx';

import { DEFAULT_MARGINS } from '../config';
import { buildDocument, PAGE_WIDTH } from '../docx';
import { collapseEmptyLines } from './blocks';
import { renderMarkdown } from './renderer';

const PNG_1PX = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==', 'base64');

async function render(source: string) {
    const imageDir = await fs.mkdtemp(path.join(os.tmpdir(), 'md2docx-md-'));
    await fs.writeFile(path.join(imageDir, 'pic.png'), PNG_1PX);
    const blocks = collapseEmptyLines(renderMarkdown(source, { imageDir, contentWidth: PAGE_WIDTH - DEFAULT_MARGINS.left - DEFAULT_MARGINS.right }));
    const zip = await JSZip.loadAsync(await Packer.toBuffer(buildDocument(blocks, DEFAULT_MARGINS)));
    const body = (await zip.file('word/document.xml')!.async('string')).match(/<w:body>([\s\S]*)<\/w:body>/)![1]!;
    return { body, numbering: await zip.file('word/numbering.xml')!.async('string') };
}

const outline = (body: string) => [...body.matchAll(/<w:p>([\s\S]*?)<\/w:p>/g)].map(([, p = '']) => {
    const style = p.match(/<w:pStyle w:val="([^"]*)"/)?.[1] ?? '';
    return `${style}:${[...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(([, t]) => t).join('').trim()}`;
});

const paragraphWith = (body: string, text: string) => body.match(new RegExp(`<w:p>(?:(?!<w:p>)[\\s\\S])*?>${text}</w:t>[\\s\\S]*?</w:p>`))![0];

test('links render as plain black text', async () => {
    const { body } = await render('See [the docs](https://example.com) here');
    assert.match(body, /the docs/);
    assert.doesNotMatch(body, /<w:color |<w:u /);
});

test('lists: dash bullets and numbers from the paragraph indent, each list restarts', async () => {
    const { body, numbering } = await render('- a\n    - b\n\ntext\n\n1. one\n2. two\n\ntext\n\n1. again\n\ntext\n\n1) paren');

    const numIdOf = (text: string) => paragraphWith(body, text).match(/<w:numId w:val="(\d+)"/)![1];
    const level = (text: string, lvl: number) => {
        const abstractId = numbering.match(new RegExp(`<w:num w:numId="${numIdOf(text)}"[^>]*>\\s*<w:abstractNumId w:val="(\\d+)"`))![1];
        const abstractNum = numbering.match(new RegExp(`<w:abstractNum [^>]*w:abstractNumId="${abstractId}"[\\s\\S]*?</w:abstractNum>`))![0];
        return abstractNum.match(new RegExp(`<w:lvl w:ilvl="${lvl}"[\\s\\S]*?</w:lvl>`))![0];
    };
    assert.match(level('a', 0), /<w:lvlText w:val="–"\/>/);
    assert.match(level('a', 0), /<w:ind w:left="0" w:firstLine="567"\/>/);
    assert.match(level('b', 1), /<w:ind w:left="567" w:firstLine="567"\/>/);
    assert.match(level('a', 0), /w:ascii="Times New Roman"/);

    assert.equal(numIdOf('one'), numIdOf('two'));
    assert.notEqual(numIdOf('one'), numIdOf('again'));
    assert.match(level('paren', 0), /<w:lvlText w:val="%1\)"\/>/);
    assert.match(level('one', 0), /<w:lvlText w:val="%1."\/>/);
});

test('tables: header row repeats and is centered, cells use the table text style, caption keeps with table', async () => {
    const { body } = await render('Таблиця 1 – Caption\n\n| A | B |\n|---|---|\n| 1 | 2 |');

    const [headerRow, bodyRow] = [...body.matchAll(/<w:tr>[\s\S]*?<\/w:tr>/g)].map(([r]) => r);
    assert.match(headerRow!, /<w:tblHeader\/>/);
    assert.match(headerRow!, /<w:jc w:val="center"\/>/);
    assert.doesNotMatch(bodyRow!, /<w:tblHeader\/>|<w:jc /);
    assert.match(bodyRow!, /<w:pStyle w:val="TableText"\/>/);
    assert.match(paragraphWith(body, 'Таблиця 1 – Caption'), /<w:keepNext\/>/);
});

test('table columns fit their longest word and span the text width', async () => {
    const long = 'довгий опис причини з багатьма словами, який займає кілька рядків у комірці таблиці';
    const { body } = await render(`| Причина | Тип |\n|---|---|\n| ${long} | підсилює |`);

    const widths = [...body.matchAll(/<w:gridCol w:w="(\d+)"\/>/g)].map(([, w]) => Number(w));
    assert.equal(widths.reduce((a, b) => a + b, 0), PAGE_WIDTH - DEFAULT_MARGINS.left - DEFAULT_MARGINS.right);
    assert.ok(widths[1]! >= 'підсилює'.length * 130 + 216, `"Тип" column too narrow: ${widths[1]}`);
    assert.ok(widths[0]! > widths[1]!);
});

test('a captioned figure keeps with its caption, an uncaptioned one does not', async () => {
    const { body } = await render('![Рисунок 1 – Pic](pic.png)\n\n![](pic.png)');
    const images = [...body.matchAll(/<w:p>(?:(?!<w:p>)[\s\S])*?<w:drawing>[\s\S]*?<\/w:p>/g)].map(([p]) => p);
    assert.equal(images.length, 2);
    assert.match(images[0]!, /<w:keepNext\/>/);
    assert.doesNotMatch(images[1]!, /<w:keepNext\/>/);
});

test('no blank line between consecutive headings, one between heading and text', async () => {
    const { body } = await render('# 1 Chapter\n\n## 1.1 Section\n\nText');
    assert.deepEqual(outline(body), ['ChapterTitle:1 Chapter', 'SectionTitle:1.1 Section', 'BodyText:', 'BodyText:Text']);
});

test('a heading keeps with its text across the blank line, so it never ends a page alone', async () => {
    const { body } = await render('## 1.1 Section\n\nText\n\nMore');
    const paragraphs = [...body.matchAll(/<w:p>([\s\S]*?)<\/w:p>/g)].map(([p]) => p);
    assert.deepEqual(outline(body), ['SectionTitle:1.1 Section', 'BodyText:', 'BodyText:Text', 'BodyText:More']);
    assert.match(paragraphs[1]!, /<w:keepNext\/>/);
    assert.doesNotMatch(paragraphs[2]!, /<w:keepNext\/>/);
});

test('a paragraph introducing a figure keeps with it, other paragraphs do not', async () => {
    const { body } = await render('Intro\n\n![Рисунок 1 – Pic](pic.png)\n\nAfter\n\nEnd');
    const paragraphs = [...body.matchAll(/<w:p>([\s\S]*?)<\/w:p>/g)].map(([p]) => p);
    assert.deepEqual(outline(body), ['BodyText:Intro', 'BodyText:', 'ImageBlock:', 'ImageCaption:Рисунок 1 – Pic', 'BodyText:', 'BodyText:After', 'BodyText:End']);
    for (const i of [0, 1, 2]) assert.match(paragraphs[i]!, /<w:keepNext\/>/, `paragraph ${i}`);
    for (const i of [4, 5]) assert.doesNotMatch(paragraphs[i]!, /<w:keepNext\/>/, `paragraph ${i}`);
});

test('document properties: author and title from [meta], never the library default "Un-named"', async () => {
    const coreOf = async (meta?: { author: string; title?: string }) => {
        const zip = await JSZip.loadAsync(await Packer.toBuffer(buildDocument([], DEFAULT_MARGINS, meta)));
        return zip.file('docProps/core.xml')!.async('string');
    };
    const core = await coreOf({ author: 'Попов А.А.', title: 'Лабораторна робота №1' });
    assert.match(core, /<dc:creator>Попов А\.А\.<\/dc:creator>/);
    assert.match(core, /<cp:lastModifiedBy>Попов А\.А\.<\/cp:lastModifiedBy>/);
    assert.match(core, /<dc:title>Лабораторна робота №1<\/dc:title>/);
    assert.doesNotMatch(await coreOf(), /Un-named/);
});
