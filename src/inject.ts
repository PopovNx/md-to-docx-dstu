import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';

import { chunkPlaceholderStyleId } from './docx';

export type DocxChunk = { index: number; path: string };

export async function injectDocxChunks(docxBuffer: Buffer, chunks: DocxChunk[]): Promise<Buffer> {
    if (chunks.length === 0) return docxBuffer;

    const zip = await JSZip.loadAsync(docxBuffer);

    let documentXml = await zip.file('word/document.xml')!.async('string');
    let relsXml = await zip.file('word/_rels/document.xml.rels')!.async('string');

    const mainSectPr = extractFinalSectPr(documentXml);

    for (const { index, path: chunkPath } of chunks) {
        const chunkZip = await JSZip.loadAsync(await fs.readFile(chunkPath));

        const chunkDocXml = await chunkZip.file('word/document.xml')!.async('string');
        const chunkRelsXml = chunkZip.file('word/_rels/document.xml.rels')
            ? await chunkZip.file('word/_rels/document.xml.rels')!.async('string')
            : null;

        const { bodyContent, sectPr: chunkSectPr } = splitBodyContent(chunkDocXml);

        const cleanBody = stripHeaderFooterRefs(bodyContent);
        const cleanSectPr = chunkSectPr ? stripHeaderFooterRefs(chunkSectPr) : null;

        let mergedContent = cleanBody;

        if (chunkRelsXml) {
            const nextRId = getMaxRId(relsXml) + 1;
            const { remapped, newRels } = await remapMediaRelationships(
                cleanBody, chunkRelsXml, chunkZip, zip, `ch${index}_`, nextRId,
            );
            mergedContent = remapped;
            relsXml = appendRelationships(relsXml, newRels);
        }

        const sectionBefore = `<w:p><w:pPr>${mainSectPr}</w:pPr></w:p>`;
        const sectionAfter = cleanSectPr ? `<w:p><w:pPr>${cleanSectPr}</w:pPr></w:p>` : '';

        documentXml = replacePlaceholderParagraph(documentXml, index, sectionBefore + mergedContent + sectionAfter);
    }

    zip.file('word/document.xml', documentXml);
    zip.file('word/_rels/document.xml.rels', relsXml);

    return zip.generateAsync({ type: 'nodebuffer' });
}

function splitBodyContent(documentXml: string): { bodyContent: string; sectPr: string | null } {
    const bodyMatch = documentXml.match(/<w:body>([\s\S]*)<\/w:body>/);
    if (!bodyMatch) return { bodyContent: '', sectPr: null };

    const body = bodyMatch[1];

    const lastEnd = body.lastIndexOf('</w:sectPr>');
    if (lastEnd === -1) return { bodyContent: body.trim(), sectPr: null };

    const lastStart = body.lastIndexOf('<w:sectPr', lastEnd);
    if (lastStart === -1) return { bodyContent: body.trim(), sectPr: null };

    const sectPr = body.slice(lastStart, lastEnd + '</w:sectPr>'.length);
    const bodyContent = body.slice(0, lastStart).trim();

    return { bodyContent, sectPr };
}

function extractFinalSectPr(documentXml: string): string {
    const lastEnd = documentXml.lastIndexOf('</w:sectPr>');
    if (lastEnd === -1) return '<w:sectPr/>';
    const lastStart = documentXml.lastIndexOf('<w:sectPr', lastEnd);
    if (lastStart === -1) return '<w:sectPr/>';
    return documentXml.slice(lastStart, lastEnd + '</w:sectPr>'.length);
}

function stripHeaderFooterRefs(xml: string): string {
    return xml
        .replace(/<w:headerReference[^>]*\/>/g, '')
        .replace(/<w:footerReference[^>]*\/>/g, '');
}

function replacePlaceholderParagraph(documentXml: string, index: number, replacement: string): string {
    const styleId = chunkPlaceholderStyleId(index);
    return documentXml.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, (match) =>
        match.includes(`w:val="${styleId}"`) ? replacement : match,
    );
}

function getMaxRId(relsXml: string): number {
    const ids = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((m) => parseInt(m[1], 10));
    return ids.length > 0 ? Math.max(...ids) : 0;
}

function appendRelationships(relsXml: string, rels: string[]): string {
    if (rels.length === 0) return relsXml;
    return relsXml.replace('</Relationships>', rels.join('\n') + '\n</Relationships>');
}

type RemapResult = { remapped: string; newRels: string[] };

async function remapMediaRelationships(
    bodyContent: string,
    chunkRelsXml: string,
    chunkZip: JSZip,
    mainZip: JSZip,
    prefix: string,
    startRId: number,
): Promise<RemapResult> {
    const relPattern = /<Relationship\s([^>]+)\/>/g;
    let remapped = bodyContent;
    const newRels: string[] = [];
    let rIdCounter = startRId;

    for (const relMatch of chunkRelsXml.matchAll(relPattern)) {
        const attrs = relMatch[1];
        const oldId = extractAttr(attrs, 'Id');
        const type = extractAttr(attrs, 'Type');
        const target = extractAttr(attrs, 'Target');

        if (!oldId || !type || !target) continue;

        if (!bodyContent.includes(`"${oldId}"`)) continue;

        const newId = `rId${rIdCounter++}`;
        remapped = remapped
            .replaceAll(`r:id="${oldId}"`, `r:id="${newId}"`)
            .replaceAll(`r:embed="${oldId}"`, `r:embed="${newId}"`);

        if (target.startsWith('media/')) {
            const chunkFile = chunkZip.file(`word/${target}`);
            if (chunkFile) {
                const ext = path.extname(target);
                const newTarget = `media/${prefix}${path.basename(target, ext)}${ext}`;
                mainZip.file(`word/${newTarget}`, await chunkFile.async('nodebuffer'));
                newRels.push(`<Relationship Id="${newId}" Type="${type}" Target="${newTarget}"/>`);
            }
        } else if (type.includes('hyperlink')) {
            newRels.push(`<Relationship Id="${newId}" Type="${type}" Target="${target}" TargetMode="External"/>`);
        }
    }

    return { remapped, newRels };
}

function extractAttr(attrs: string, name: string): string | null {
    return attrs.match(new RegExp(`${name}="([^"]+)"`))?.[1] ?? null;
}
