import path from 'node:path';
import JSZip from 'jszip';

import { attr } from './xml';

const CONTENT_TYPES = '[Content_Types].xml';

export type Relationship = { id: string; type: string; target: string; external: boolean };

export async function readText(zip: JSZip, name: string): Promise<string> {
    const file = zip.file(name);
    if (!file) throw new Error(`Invalid .docx: ${name} is missing`);
    return file.async('string');
}

export const readOptionalText = (zip: JSZip, name: string) => zip.file(name)?.async('string') ?? Promise.resolve(null);

export const relsPathOf = (part: string) => path.posix.join(path.posix.dirname(part), '_rels', `${path.posix.basename(part)}.rels`);

export function parseRels(relsXml: string): Relationship[] {
    return [...relsXml.matchAll(/<Relationship\s[^>]*?\/?>/g)].map(([tag]) => ({
        id: attr(tag, 'Id') ?? '',
        type: attr(tag, 'Type') ?? '',
        target: attr(tag, 'Target') ?? '',
        external: attr(tag, 'TargetMode') === 'External',
    }));
}

const relTag = ({ id, type, target, external }: Relationship) =>
    `<Relationship Id="${id}" Type="${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`;

export const appendRels = (relsXml: string, rels: Relationship[]) =>
    relsXml.replace('</Relationships>', rels.map(relTag).join('') + '</Relationships>');

export const nextRIdNumber = (relsXml: string) =>
    Math.max(0, ...[...relsXml.matchAll(/Id="rId(\d+)"/g)].map(m => Number(m[1]))) + 1;

export const resolveTarget = (ownerPart: string, target: string) =>
    target.startsWith('/') ? target.slice(1) : path.posix.join(path.posix.dirname(ownerPart), target);

export class PartCopier {
    private readonly copied = new Map<string, string>();

    constructor(
        private readonly src: JSZip,
        private readonly dst: JSZip,
        private readonly prefix: string,
        private readonly transformXml: (xml: string, contentType: string) => string,
    ) {}

    async copy(srcPart: string): Promise<string | null> {
        const known = this.copied.get(srcPart);
        if (known) return known;
        const file = this.src.file(srcPart);
        if (!file) return null;

        const dstPart = path.posix.join(path.posix.dirname(srcPart), this.prefix + path.posix.basename(srcPart));
        this.copied.set(srcPart, dstPart);

        const contentType = await this.registerContentType(srcPart, dstPart);
        if (/\.xml$/.test(srcPart)) {
            this.dst.file(dstPart, this.transformXml(await file.async('string'), contentType));
            await this.copyRels(srcPart, dstPart);
        } else {
            this.dst.file(dstPart, await file.async('nodebuffer'));
        }
        return dstPart;
    }

    private async copyRels(srcPart: string, dstPart: string) {
        const relsXml = await readOptionalText(this.src, relsPathOf(srcPart));
        if (!relsXml) return;
        const rels: Relationship[] = [];
        for (const rel of parseRels(relsXml)) {
            if (rel.external) {
                rels.push(rel);
                continue;
            }
            const copied = await this.copy(resolveTarget(srcPart, rel.target));
            if (copied) rels.push({ ...rel, target: path.posix.relative(path.posix.dirname(dstPart), copied) });
        }
        this.dst.file(relsPathOf(dstPart), appendRels('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>', rels));
    }

    private async registerContentType(srcPart: string, dstPart: string): Promise<string> {
        const srcTypes = await readText(this.src, CONTENT_TYPES);
        let dstTypes = await readText(this.dst, CONTENT_TYPES);
        const ext = path.posix.extname(srcPart).slice(1).toLowerCase();

        const override = srcTypes.match(new RegExp(`<Override [^>]*PartName="/${escapeRegExp(srcPart)}"[^>]*/>`, 'i'))?.[0];
        const defaultOf = (types: string) => types.match(new RegExp(`<Default [^>]*Extension="${ext}"[^>]*/>`, 'i'))?.[0];
        const contentType = attr(override ?? defaultOf(srcTypes) ?? '', 'ContentType') ?? '';

        if (override) {
            dstTypes = dstTypes.replace('</Types>', `<Override PartName="/${dstPart}" ContentType="${contentType}"/></Types>`);
        } else if (!defaultOf(dstTypes) && contentType) {
            dstTypes = dstTypes.replace('</Types>', `<Default Extension="${ext}" ContentType="${contentType}"/></Types>`);
        }
        this.dst.file(CONTENT_TYPES, dstTypes);
        return contentType;
    }
}

const escapeRegExp = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
