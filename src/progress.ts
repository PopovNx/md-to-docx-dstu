import path from 'node:path';
import { Listr } from 'listr2';

import type { Config } from './config';
import type { InputFile, RenderOptions } from './types';
import { renderInputFile, type RenderResult } from './markdown';
import { PAGE_HEIGHT, PAGE_WIDTH } from './docx';
import { log } from './logger';

export function buildTaskList(config: Config, inputFiles: InputFile[]) {
    const opts: RenderOptions = {
        imageDir: config.dirs.images,
        contentWidth: PAGE_WIDTH - config.margins.left - config.margins.right,
        contentHeight: PAGE_HEIGHT - config.margins.top - config.margins.bottom,
    };

    const fileTasks = inputFiles.map(file => {
        const name = path.basename(file.path);
        const title = `${name}  ${file.kind === 'docx' ? '[inject]' : '[markdown]'}`;
        return {
            title: name,
            task: async (ctx: RenderResult, task: { title: string }) => {
                task.title = `${title}  …`;
                await renderInputFile(file, opts, config.vars, ctx);
                log.debug({ file: name, kind: file.kind }, 'rendered');
                task.title = title;
            },
        };
    });

    return new Listr<RenderResult>(
        [{
            title: `Rendering ${inputFiles.length} file(s)  →  ${config.dirs.output}`,
            task: (_ctx, task) => task.newListr(fileTasks, { concurrent: false, rendererOptions: { collapseSubtasks: false } }),
        }],
        {
            ctx: { paragraphs: [], chunks: [] },
            renderer: 'default',
            fallbackRenderer: 'simple',
            rendererOptions: { collapseErrors: false },
        },
    );
}
