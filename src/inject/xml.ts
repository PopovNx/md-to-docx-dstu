export const attr = (xml: string, name: string) => xml.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1];

export function splitBody(documentXml: string): { content: string; sectPr: string | null } {
    const body = documentXml.match(/<w:body>([\s\S]*)<\/w:body>/)?.[1] ?? '';
    const end = body.lastIndexOf('</w:sectPr>');
    const start = end === -1 ? -1 : body.lastIndexOf('<w:sectPr', end);
    if (start === -1 || body.slice(end).includes('<w:p')) return { content: body.trim(), sectPr: null };
    return { content: body.slice(0, start).trim(), sectPr: body.slice(start, end + '</w:sectPr>'.length) };
}

export function lastParagraphStart(xml: string): number {
    if (!xml.endsWith('</w:p>')) return -1;
    let depth = 0;
    let start = -1;
    for (const m of xml.matchAll(/<w:p\/>|<w:p[ >]|<\/w:p>/g)) {
        if (m[0] === '<w:p/>') {
            if (depth === 0) start = m.index;
        } else if (m[0] === '</w:p>') {
            depth--;
        } else if (depth++ === 0) {
            start = m.index;
        }
    }
    return start;
}

export function endSectionAt(xml: string, start: number, sectPr: string): string {
    const head = xml.slice(0, start);
    const paragraph = xml.slice(start);
    const pPrStart = paragraph.match(/^<w:p(?: [^>]*)?>(?=<w:pPr[ >/])/)?.[0].length;

    if (pPrStart === undefined) {
        return head + paragraph.replace(/^<w:p( [^>]*)?(\/?)>/, (_m, attrs = '', selfClosing) =>
            `<w:p${attrs}><w:pPr>${sectPr}</w:pPr>${selfClosing ? '</w:p>' : ''}`);
    }
    if (paragraph.startsWith('<w:pPr/>', pPrStart)) {
        return head + paragraph.slice(0, pPrStart) + `<w:pPr>${sectPr}</w:pPr>` + paragraph.slice(pPrStart + '<w:pPr/>'.length);
    }
    const close = paragraph.slice(pPrStart).search(/<w:pPrChange[ >]|<\/w:pPr>/) + pPrStart;
    return head + paragraph.slice(0, close) + sectPr + paragraph.slice(close);
}

export function mergeNamespaces(target: string, source: string, rootTag: string): string {
    const rootOf = (xml: string) => xml.match(new RegExp(`<${rootTag}\\b[^>]*>`))?.[0] ?? '';
    const ignorableOf = (root: string) => (attr(root, 'mc:Ignorable') ?? '').split(/\s+/).filter(Boolean);
    const targetRoot = rootOf(target);
    const sourceRoot = rootOf(source);

    const missing = [...sourceRoot.matchAll(/\sxmlns:(\w+)="[^"]*"/g)]
        .filter(([, prefix]) => !targetRoot.includes(`xmlns:${prefix}=`))
        .map(([decl]) => decl);
    const ignorable = [...new Set([...ignorableOf(targetRoot), ...ignorableOf(sourceRoot)])].join(' ');
    if (missing.length === 0 && ignorable === ignorableOf(targetRoot).join(' ')) return target;

    const withoutIgnorable = targetRoot.replace(/\smc:Ignorable="[^"]*"/, '').slice(0, -1);
    return target.replace(targetRoot, `${withoutIgnorable}${missing.join('')}${ignorable ? ` mc:Ignorable="${ignorable}"` : ''}>`);
}
