import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';

import { applyVars } from '../config';
import { chunkPlaceholderStyleId } from '../docx';
import { appendRels, nextRIdNumber, parseRels, PartCopier, readOptionalText, readText, relsPathOf, resolveTarget, type Relationship } from './parts';
import { importNumbering, importStyles } from './styles';
import { endSectionAt, lastParagraphStart, mergeNamespaces, splitBody } from './xml';

export type DocxChunk = { index: number; path: string; breakBefore: boolean };

const DOCUMENT = 'word/document.xml';
const DOCUMENT_RELS = relsPathOf(DOCUMENT);
const THEME_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';
const HEADER_FOOTER_TYPE = /wordprocessingml\.(header|footer)\+xml$/;

export async function injectDocxChunks(docxBuffer: Buffer, chunks: DocxChunk[], vars: Record<string, string> = {}): Promise<Buffer> {
    if (chunks.length === 0) return docxBuffer;

    const main = await JSZip.loadAsync(docxBuffer);
    let documentXml = await readText(main, DOCUMENT);
    const mainSectPr = splitBody(documentXml).sectPr ?? '<w:sectPr/>';

    for (const chunk of chunks) {
        const src = await JSZip.loadAsync(await fs.readFile(chunk.path));
        const prefix = `ch${chunk.index}_`;
        const chunkXml = await readText(src, DOCUMENT);

        const remapNumIds = await importNumbering(main, src, prefix);
        const applyStyles = await importStyles(main, src, prefix, remapNumIds);
        const drawingIdShift = (chunk.index + 1) * 100_000;
        const transform = (xml: string) => applyStyles(remapNumIds(applyVars(xml, vars)))
            .replace(/(<wp:docPr [^>]*?\bid=")(\d+)"/g, (_m, head: string, id: string) => `${head}${Number(id) + drawingIdShift}"`);

        const copier = new PartCopier(src, main, prefix, (xml, type) => (HEADER_FOOTER_TYPE.test(type) ? transform(xml) : xml));
        const { content, sectPr } = splitBody(chunkXml);
        const remapRIds = await importDocumentRels(main, src, copier, content + (sectPr ?? ''));
        await importTheme(main, src, copier);

        const body = remapRIds(transform(content));
        const chunkSectPr = sectPr ? remapRIds(sectPr) : mainSectPr;
        documentXml = spliceChunk(mergeNamespaces(documentXml, chunkXml, 'w:document'), chunk, body, chunkSectPr, mainSectPr, transform);
    }

    main.file(DOCUMENT, documentXml);
    return main.generateAsync({ type: 'nodebuffer' });
}

function spliceChunk(
    documentXml: string,
    chunk: DocxChunk,
    body: string,
    chunkSectPr: string,
    mainSectPr: string,
    transform: (xml: string) => string,
): string {
    const marker = documentXml.indexOf(`<w:pStyle w:val="${chunkPlaceholderStyleId(chunk.index)}"/>`);
    if (marker === -1) throw new Error(`Placeholder for ${path.basename(chunk.path)} not found`);
    const start = Math.max(documentXml.lastIndexOf('<w:p>', marker), documentXml.lastIndexOf('<w:p ', marker));
    const end = documentXml.indexOf('</w:p>', marker) + '</w:p>'.length;

    let before = documentXml.slice(0, start);
    let after = documentXml.slice(end);

    if (chunk.breakBefore) {
        before = endSection(before, mainSectPr, '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p>');
    }

    if (after.startsWith('<w:sectPr')) {
        after = after.replace(/^<w:sectPr[\s\S]*?<\/w:sectPr>/, chunkSectPr);
        return before + body + after;
    }
    return before + endSection(body, chunkSectPr, transform('<w:p/>')) + after;
}

function endSection(xml: string, sectPr: string, emptyParagraph: string): string {
    const last = lastParagraphStart(xml);
    return last === -1 ? endSectionAt(xml + emptyParagraph, xml.length, sectPr) : endSectionAt(xml, last, sectPr);
}

async function importDocumentRels(main: JSZip, src: JSZip, copier: PartCopier, referencingXml: string): Promise<(xml: string) => string> {
    const srcRels = parseRels((await readOptionalText(src, DOCUMENT_RELS)) ?? '');
    const mainRels = await readText(main, DOCUMENT_RELS);
    const referenced = new Set([...referencingXml.matchAll(/\br:\w+="([^"]+)"/g)].map(m => m[1]));

    let nextId = nextRIdNumber(mainRels);
    const ids = new Map<string, string>();
    const added: Relationship[] = [];
    for (const rel of srcRels.filter(r => referenced.has(r.id))) {
        const target = rel.external ? rel.target : await copier.copy(resolveTarget(DOCUMENT, rel.target));
        if (target === null) continue;
        const id = `rId${nextId++}`;
        ids.set(rel.id, id);
        added.push({ ...rel, id, target: rel.external ? target : path.posix.relative('word', target) });
    }
    main.file(DOCUMENT_RELS, appendRels(mainRels, added));

    return xml => xml.replace(/\b(r:\w+=")([^"]+)"/g, (m, head: string, id: string) => (ids.has(id) ? `${head}${ids.get(id)}"` : m));
}

async function importTheme(main: JSZip, src: JSZip, copier: PartCopier) {
    const mainRels = await readText(main, DOCUMENT_RELS);
    if (parseRels(mainRels).some(r => r.type === THEME_REL)) return;
    const theme = parseRels((await readOptionalText(src, DOCUMENT_RELS)) ?? '').find(r => r.type === THEME_REL);
    const copied = theme && (await copier.copy(resolveTarget(DOCUMENT, theme.target)));
    if (!theme || !copied) return;
    main.file(DOCUMENT_RELS, appendRels(mainRels, [{ ...theme, id: `rId${nextRIdNumber(mainRels)}`, target: path.posix.relative('word', copied) }]));
}
