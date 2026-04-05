import { readFileSync } from 'node:fs';
import path from 'node:path';

import { type Token } from 'markdown-it';
import { imageSize } from 'image-size';
import { ImageRun } from 'docx';

import type { GeneratorPaths } from '../types';

export function buildImageRun(token: Token, paths: GeneratorPaths): ImageRun {
    const source = token.attrGet('src') ?? '';
    const widthAttr = token.attrGet('width');
    const imagePath = path.resolve(paths.imageDir, source);
    const buffer = readFileSync(imagePath);
    const dimensions = imageSize(buffer);
    const type = getImageType(imagePath);

    const naturalWidth = dimensions.width ?? 500;
    const naturalHeight = dimensions.height ?? 300;
    const targetWidth = widthAttr ? Number(widthAttr) : Math.min(naturalWidth, 650);
    const scale = naturalWidth > 0 ? targetWidth / naturalWidth : 1;

    return new ImageRun({
        data: buffer,
        type,
        transformation: {
            width: Math.max(1, Math.round(targetWidth)),
            height: Math.max(1, Math.round(naturalHeight * scale)),
        },
    });
}

function getImageType(imagePath: string): 'png' | 'jpg' | 'gif' | 'bmp' {
    const extension = path.extname(imagePath).slice(1).toLowerCase();

    if (extension === 'jpeg') {
        return 'jpg';
    }

    if (extension === 'jpg' || extension === 'png' || extension === 'gif' || extension === 'bmp') {
        return extension;
    }

    return 'png';
}
