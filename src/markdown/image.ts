import { readFileSync } from 'node:fs';
import path from 'node:path';

import { type Token } from 'markdown-it';
import { imageSize } from 'image-size';
import { ImageRun } from 'docx';

import type { RenderOptions } from '../types';

const FALLBACK_WIDTH = 500;
const FALLBACK_HEIGHT = 300;

const CAPTION_ROOM = 3 * 420 + 300;

const IMAGE_TYPES: Record<string, 'png' | 'jpg' | 'gif' | 'bmp'> = {
    jpeg: 'jpg',
    jpg: 'jpg',
    png: 'png',
    gif: 'gif',
    bmp: 'bmp',
};

export function buildImageRun(token: Token, paths: RenderOptions): ImageRun {
    const source = token.attrGet('src') ?? '';
    const widthAttr = token.attrGet('width');
    const imagePath = path.resolve(paths.imageDir, source);
    const buffer = readFileSync(imagePath);
    const dimensions = imageSize(buffer);
    const ext = path.extname(imagePath).slice(1).toLowerCase();

    const naturalWidth = dimensions.width ?? FALLBACK_WIDTH;
    const naturalHeight = dimensions.height ?? FALLBACK_HEIGHT;
    const maxWidth = Math.floor(paths.contentWidth / 15);
    const maxHeight = Math.floor((paths.contentHeight - CAPTION_ROOM) / 15);
    const requestedWidth = widthAttr ? parseInt(widthAttr, 10) : naturalWidth;
    const scale = Math.min(requestedWidth / naturalWidth, maxWidth / naturalWidth, maxHeight / naturalHeight);

    return new ImageRun({
        data: buffer,
        type: IMAGE_TYPES[ext] ?? 'png',
        transformation: {
            width: Math.max(1, Math.round(naturalWidth * scale)),
            height: Math.max(1, Math.round(naturalHeight * scale)),
        },
    });
}
