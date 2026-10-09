import JSZip from 'jszip';

import { readOptionalText, readText } from './parts';
import { attr, mergeNamespaces } from './xml';

const STYLES = 'word/styles.xml';
const NUMBERING = 'word/numbering.xml';

export async function importStyles(
    main: JSZip,
    chunk: JSZip,
    prefix: string,
    remapNumIds: (xml: string) => string,
): Promise<(xml: string) => string> {
    const chunkStyles = await readOptionalText(chunk, STYLES);
    const renameRefs = (xml: string) =>
        xml.replace(/(<w:(?:pStyle|rStyle|tblStyle|basedOn|next|link|styleLink|numStyleLink) w:val=")([^"]*)"/g, `$1${prefix}$2"`);

    if (!chunkStyles) return renameRefs;

    const docDefaultsId = `${prefix}DocDefaults`;
    const defaultsOf = (kind: 'rPr' | 'pPr') =>
        chunkStyles.match(new RegExp(`<w:${kind}Default>\\s*<w:${kind}>([\\s\\S]*?)</w:${kind}>`))?.[1] ?? '';
    const docDefaults = `<w:style w:type="paragraph" w:styleId="${docDefaultsId}"><w:name w:val="${docDefaultsId}"/><w:semiHidden/>`
        + `<w:pPr>${defaultsOf('pPr')}</w:pPr><w:rPr>${defaultsOf('rPr')}</w:rPr></w:style>`;

    let defaultParagraphStyle: string | undefined;
    let defaultTableStyle: string | undefined;

    const chunkStyleList = [...chunkStyles.matchAll(/<w:style [\s\S]*?<\/w:style>/g)].map(([style]) => style);
    const paragraphStyleIds = new Set(chunkStyleList
        .filter(style => attr(style, 'w:type') === 'paragraph')
        .map(style => attr(style, 'w:styleId')));

    const styles = chunkStyleList.map(style => {
        const open = style.match(/^<w:style [^>]*>/)![0];
        const type = attr(open, 'w:type');
        const id = attr(open, 'w:styleId') ?? '';
        if (/w:default="(1|true)"/.test(open)) {
            if (type === 'paragraph') defaultParagraphStyle = prefix + id;
            if (type === 'table') defaultTableStyle = prefix + id;
        }

        let imported = remapNumIds(renameRefs(style))
            .replace(open, open.replace(/ w:default="[^"]*"/, '').replace(/w:styleId="/, `w:styleId="${prefix}`))
            .replace(/<w:name w:val="/, `<w:name w:val="${prefix}`);
        const basedOn = attr(style.match(/<w:basedOn [^>]*>/)?.[0] ?? '', 'w:val');
        if (type === 'paragraph' && basedOn !== undefined && !paragraphStyleIds.has(basedOn)) {
            imported = imported.replace(/<w:basedOn [^>]*>/, `<w:basedOn w:val="${docDefaultsId}"/>`);
        } else if (type === 'paragraph' && basedOn === undefined) {
            imported = imported.replace(/(<w:name [^>]*\/>(?:<w:aliases [^>]*\/>)?)/, `$1<w:basedOn w:val="${docDefaultsId}"/>`);
        }
        return imported;
    });

    const mainStyles = mergeNamespaces(await readText(main, STYLES), chunkStyles, 'w:styles');
    main.file(STYLES, mainStyles.replace('</w:styles>', docDefaults + styles.join('') + '</w:styles>'));

    const pStyle = defaultParagraphStyle ?? docDefaultsId;
    return xml => {
        const known = xml.replace(/<w:pStyle w:val="([^"]*)"\/>/g, (m, id: string) => (paragraphStyleIds.has(id) ? m : ''));
        let out = renameRefs(known)
            .replace(/<w:p(?: [^>]*?)?\/>/g, m => `${m.slice(0, -2)}><w:pPr><w:pStyle w:val="${pStyle}"/></w:pPr></w:p>`)
            .replace(/(<w:p(?: [^>]*)?>)(<w:pPr\/>|<w:pPr>(?!<w:pStyle )|(?!<w:pPr[ >/]))/g, (_m, open, pPr) =>
                pPr === '<w:pPr/>' || pPr === ''
                    ? `${open}<w:pPr><w:pStyle w:val="${pStyle}"/></w:pPr>`
                    : `${open}<w:pPr><w:pStyle w:val="${pStyle}"/>`);
        if (defaultTableStyle) {
            out = out.replace(/<w:tblPr>(?!<w:tblStyle )/g, `<w:tblPr><w:tblStyle w:val="${defaultTableStyle}"/>`);
        }
        return out;
    };
}

export async function importNumbering(main: JSZip, chunk: JSZip, prefix: string): Promise<(xml: string) => string> {
    const chunkNumbering = await readOptionalText(chunk, NUMBERING);
    const mainNumbering = await readOptionalText(main, NUMBERING);
    if (!chunkNumbering || !mainNumbering) return xml => xml;

    const maxId = (xml: string, name: string) =>
        Math.max(0, ...[...xml.matchAll(new RegExp(`<w:${name} [^>]*?w:${name}Id="(\\d+)"`, 'g'))].map(m => Number(m[1])));
    const abstractShift = maxId(mainNumbering, 'abstractNum');
    const numShift = maxId(mainNumbering, 'num');
    const shift = (by: number) => (_m: string, head: string, id: string) => `${head}${Number(id) + by}"`;

    const remapNumIds = (xml: string) =>
        xml.replace(/(<w:numId w:val=")(\d+)"/g, (m, head, id) => (id === '0' ? m : shift(numShift)(m, head, id)));

    const abstractNums = [...chunkNumbering.matchAll(/<w:abstractNum [\s\S]*?<\/w:abstractNum>/g)]
        .map(([xml]) => xml
            .replace(/(<w:abstractNum [^>]*?w:abstractNumId=")(\d+)"/, shift(abstractShift))
            .replace(/(<w:(?:pStyle|styleLink|numStyleLink) w:val=")([^"]*)"/g, `$1${prefix}$2"`));
    const nums = [...chunkNumbering.matchAll(/<w:num [\s\S]*?<\/w:num>/g)]
        .map(([xml]) => xml
            .replace(/(<w:num [^>]*?w:numId=")(\d+)"/, shift(numShift))
            .replace(/(<w:abstractNumId w:val=")(\d+)"/, shift(abstractShift)));

    const merged = mergeNamespaces(mainNumbering, chunkNumbering, 'w:numbering');
    const firstNum = merged.search(/<w:num /);
    main.file(NUMBERING, firstNum === -1
        ? merged.replace('</w:numbering>', abstractNums.join('') + nums.join('') + '</w:numbering>')
        : merged.slice(0, firstNum) + abstractNums.join('') + merged.slice(firstNum).replace('</w:numbering>', nums.join('') + '</w:numbering>'));

    return remapNumIds;
}
