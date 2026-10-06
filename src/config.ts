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
    margins?: PageMargins;
};

export type PageMargins = { top: number; right: number; bottom: number; left: number };

const MM_TO_TWIPS = 1440 / 25.4;
const MARGIN_SIDES = ['top', 'right', 'bottom', 'left'] as const;

function parseMargins(page: Record<string, unknown> | undefined): PageMargins | undefined {
    const raw = page?.['margins'] as Record<string, unknown> | undefined;
    if (!raw) return undefined;
    const margins = {} as PageMargins;
    for (const side of MARGIN_SIDES) {
        const mm = raw[side];
        if (typeof mm !== 'number' || mm < 0) {
            throw new Error(`${CONFIG_FILE}: [page] margins.${side} must be a non-negative number (mm)`);
        }
        margins[side] = Math.round(mm * MM_TO_TWIPS);
    }
    return margins;
}

export const CONFIG_FILE = 'endeavour.toml';

const INIT_TEMPLATE = `# endeavour.toml — md-to-docx configuration

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
# Page margins in mm for markdown pages (injected .docx keep their own). Default: 15/15/15/20.
# DSTU 3008:2015:
# margins = { top = 20, right = 10, bottom = 20, left = 25 }
`;

export async function initConfig(configPath: string): Promise<void> {
    try {
        await fs.access(configPath);
        throw new Error(`${CONFIG_FILE} already exists`);
    } catch (err: unknown) {
        if (err instanceof Error && err.message.includes('already exists')) throw err;
    }
    await fs.writeFile(configPath, INIT_TEMPLATE, 'utf8');
}

export async function loadConfig(configPath: string): Promise<Config> {
    let raw: string;
    try {
        raw = await fs.readFile(configPath, 'utf8');
    } catch {
        throw new Error(`${CONFIG_FILE} not found — run \`md-to-docx init\` to create one`);
    }

    const parsed = parse(raw) as Record<string, unknown>;

    const dirs = parsed['dirs'] as Record<string, string> | undefined;
    if (!dirs?.pages || !dirs.images || !dirs.output) {
        throw new Error(`${CONFIG_FILE} must have [dirs] with pages, images, and output`);
    }

    const base = path.dirname(configPath);
    const vars = (parsed['vars'] as Record<string, string> | undefined) ?? {};

    return {
        dirs: {
            pages: path.resolve(base, dirs.pages),
            images: path.resolve(base, dirs.images),
            output: path.resolve(base, dirs.output),
        },
        vars,
        margins: parseMargins(parsed['page'] as Record<string, unknown> | undefined),
    };
}

export function applyVars(text: string, vars: Record<string, string>): string {
    return Object.entries(vars).reduce(
        (t, [key, value]) => t.replaceAll(`[[${key}]]`, value),
        text,
    );
}
