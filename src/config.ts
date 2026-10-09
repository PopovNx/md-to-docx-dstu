import fs from 'node:fs/promises';
import path from 'node:path';
import { parse } from 'smol-toml';

export type Config = {
    dirs: {
        pages: string;
        images: string;
        output: string;
    };
    vars: Record<string, string>;
    margins: PageMargins;
};

export type PageMargins = { top: number; right: number; bottom: number; left: number };

const TWIPS_PER_MM = 56.7;
const MARGIN_SIDES = ['top', 'right', 'bottom', 'left'] as const;

export const DEFAULT_MARGINS: PageMargins = { top: 1134, right: 567, bottom: 1134, left: 1418 };

export function parseMargins(page: Record<string, unknown> | undefined): PageMargins {
    const raw = (page?.['margins'] ?? {}) as Record<string, unknown>;
    const margins = { ...DEFAULT_MARGINS };
    for (const side of MARGIN_SIDES) {
        const mm = raw[side];
        if (mm === undefined) continue;
        if (typeof mm !== 'number' || mm < 0) {
            throw new Error(`${CONFIG_FILE}: [page] margins.${side} must be a non-negative number (mm)`);
        }
        margins[side] = Math.round(mm * TWIPS_PER_MM);
    }
    return margins;
}

export const CONFIG_FILE = 'endeavour.toml';

const INIT_TEMPLATE = `# endeavour.toml — md-to-docx configuration

# Shared settings (author, group, teacher, margins) from another file; values here override them.
# extends = "../common.toml"

[dirs]
pages  = "./pages"    # .md and .docx input files, sorted numerically
images = "./images"   # images referenced in markdown
output = "./output.docx"

[vars]
# Define variables as [[VARNAME]] — replaced in markdown and injected .docx files
# STUDENT_NAME = "Іванов Іван Іванович"
# GROUP        = "ІСТ-41"
# YEAR         = "2025"

[page]
# Page margins in mm for markdown pages (injected .docx, e.g. the title page, keep their own).
# Default is DSTU 3008:2015; override any side:
# margins = { top = 20, right = 10, bottom = 20, left = 25 }
`;

export async function initConfig(configPath: string): Promise<void> {
    await fs.writeFile(configPath, INIT_TEMPLATE, { encoding: 'utf8', flag: 'wx' }).catch((err: NodeJS.ErrnoException) => {
        throw err.code === 'EEXIST' ? new Error(`${CONFIG_FILE} already exists`) : err;
    });
}

export async function loadConfig(configPath: string): Promise<Config> {
    await fs.access(configPath).catch(() => {
        throw new Error(`${CONFIG_FILE} not found — run \`md-to-docx init\` to create one`);
    });
    const config = await readConfigChain(configPath);

    const dirs = config['dirs'] as Partial<Config['dirs']> | undefined;
    if (!dirs?.pages || !dirs.images || !dirs.output) {
        throw new Error(`${CONFIG_FILE} must have [dirs] with pages, images, and output`);
    }

    return {
        dirs: { pages: dirs.pages, images: dirs.images, output: dirs.output },
        vars: (config['vars'] as Record<string, string> | undefined) ?? {},
        margins: parseMargins(config['page'] as Record<string, unknown> | undefined),
    };
}

type Table = Record<string, unknown>;

const isTable = (value: unknown): value is Table => typeof value === 'object' && value !== null && value.constructor === Object;

const deepMerge = (base: Table, override: Table): Table =>
    Object.fromEntries([...new Set([...Object.keys(base), ...Object.keys(override)])].map(key => {
        const [b, o] = [base[key], override[key]];
        return [key, isTable(b) && isTable(o) ? deepMerge(b, o) : (o ?? b)];
    }));

async function readConfigChain(file: string, seen = new Set<string>()): Promise<Table> {
    if (seen.has(file)) throw new Error(`${file}: circular extends`);
    seen.add(file);

    const raw = await fs.readFile(file, 'utf8').catch(() => {
        throw new Error(`${file} not found (referenced by extends)`);
    });
    const { extends: parent, ...own } = parse(raw) as Table;

    const dir = path.dirname(file);
    if (isTable(own['dirs'])) {
        own['dirs'] = Object.fromEntries(Object.entries(own['dirs']).map(([k, v]) => [k, typeof v === 'string' ? path.resolve(dir, v) : v]));
    }

    if (parent === undefined) return own;
    if (typeof parent !== 'string') throw new Error(`${file}: extends must be a path string`);
    return deepMerge(await readConfigChain(path.resolve(dir, parent), seen), own);
}

export function applyVars(text: string, vars: Record<string, string>): string {
    return Object.entries(vars).reduce(
        (t, [key, value]) => t.replaceAll(`[[${key}]]`, value),
        text,
    );
}
