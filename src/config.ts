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
};

export const CONFIG_FILE = 'endeavour.toml';

const INIT_TEMPLATE = `# endeavour.toml — md-to-docx configuration

[dirs]
pages  = "../pages"    # .md and .docx input files, sorted numerically
images = "../images"   # images referenced in markdown
output = "./dist/output.docx"

[vars]
# Define variables as [[VARNAME]] — replaced in markdown and injected .docx files
# STUDENT_NAME = "Іванов Іван Іванович"
# GROUP        = "ІСТ-41"
# YEAR         = "2025"
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
    };
}

export function applyVars(text: string, vars: Record<string, string>): string {
    return Object.entries(vars).reduce(
        (t, [key, value]) => t.replaceAll(`[[${key}]]`, value),
        text,
    );
}
