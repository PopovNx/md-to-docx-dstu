import path from 'node:path';
import { Listr } from 'listr2';

import type { Config } from './config';
import type { GeneratorPaths, InputFile } from './types';
import type { RenderResult, FileCallback } from './markdown';
import { renderInputFiles } from './markdown';
import { log } from './logger';

export interface RenderCtx {
    result: RenderResult;
}

export function buildTaskList(config: Config, inputFiles: InputFile[]) {
    const paths: GeneratorPaths = { imageDir: config.dirs.images };

    const fileSubtasks = () =>
        inputFiles.map((file) => ({
            title: path.basename(file.path),
            task: async (ctx: RenderCtx, task: { title: string }) => {
                const name = path.basename(file.path);
                const kindLabel = file.kind === 'docx' ? '[inject]' : '[markdown]';

                const onFile: FileCallback = (_name: string, kind: 'markdown' | 'docx') => {
                    task.title = `${name}  ${kind === 'docx' ? '[inject]' : '[markdown]'}  …`;
                    return () => { task.title = name; };
                };

                const partial = await renderInputFiles([file], paths, config.vars, onFile);
                ctx.result.paragraphs.push(...partial.paragraphs);
                ctx.result.chunks.push(...partial.chunks);

                log.debug({ file: name, kind: file.kind }, 'rendered');
                task.title = `${name}  ${kindLabel}`;
            },
        }));

    return new Listr<RenderCtx>(
        [
            {
                title: `Rendering ${inputFiles.length} file(s)  →  ${config.dirs.output}`,
                task: (_ctx, task) =>
                    task.newListr(fileSubtasks(), {
                        concurrent: false,
                        rendererOptions: { collapseSubtasks: false },
                    }),
            },
        ],
        {
            ctx: { result: { paragraphs: [], chunks: [] } },
            renderer: 'default',
            fallbackRenderer: 'simple',
            rendererOptions: { collapseErrors: false },
        },
    );
}
