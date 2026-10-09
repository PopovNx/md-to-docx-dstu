import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { test } from 'node:test';

import { DEFAULT_MARGINS, loadConfig, parseMargins, parseMeta } from './config';

test('extends: shared config is merged key by key, dirs resolve against their own file', async () => {
    const root = await fs.mkdtemp(path.join(os.tmpdir(), 'md2docx-config-'));
    const lab = path.join(root, 'lab1');
    await fs.mkdir(lab);
    await fs.writeFile(path.join(root, 'common.toml'), [
        '[dirs]', 'images = "./shared-images"',
        '[vars]', 'AUTHOR = "Popov"', 'TEACHER = "Babeshko"',
        '[page]', 'margins = { left = 30, right = 15 }',
    ].join('\n'));
    await fs.writeFile(path.join(lab, 'endeavour.toml'), [
        'extends = "../common.toml"',
        '[dirs]', 'pages = "./pages"', 'output = "./out.docx"',
        '[vars]', 'TEACHER = "Fesenko"', 'LAB = "1"',
        '[page]', 'margins = { right = 10 }',
    ].join('\n'));

    const config = await loadConfig(path.join(lab, 'endeavour.toml'));

    assert.deepEqual(config.dirs, {
        pages: path.join(lab, 'pages'),
        images: path.join(root, 'shared-images'),
        output: path.join(lab, 'out.docx'),
    });
    assert.deepEqual(config.vars, { AUTHOR: 'Popov', TEACHER: 'Fesenko', LAB: '1' });
    assert.deepEqual(config.margins, { ...DEFAULT_MARGINS, left: 1701, right: 567 });

    await fs.writeFile(path.join(root, 'common.toml'), 'extends = "./lab1/endeavour.toml"');
    await assert.rejects(loadConfig(path.join(lab, 'endeavour.toml')), /circular extends/);
});

test('parseMargins: DSTU default, partial override, validation', () => {
    assert.deepEqual(parseMargins(undefined), { top: 1134, right: 567, bottom: 1134, left: 1418 });
    assert.deepEqual(parseMargins({ margins: { left: 30 } }), { ...DEFAULT_MARGINS, left: 1701 });
    assert.throws(() => parseMargins({ margins: { top: -1 } }), /margins\.top/);
    assert.throws(() => parseMargins({ margins: { top: '20' } }), /margins\.top/);
});

test('parseMeta: [[VAR]] substitution, empty author by default, validation', () => {
    const vars = { AUTHOR: 'Попов А.А.', LAB: '1' };
    assert.deepEqual(parseMeta({ author: '[[AUTHOR]]', title: 'Лабораторна робота №[[LAB]]' }, vars),
        { author: 'Попов А.А.', title: 'Лабораторна робота №1' });
    assert.deepEqual(parseMeta(undefined, vars), { author: '' });
    assert.throws(() => parseMeta({ author: 1 }, vars), /\[meta\] author/);
});
