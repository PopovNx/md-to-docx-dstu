# @popovnx/md-to-docx-dstu

Write lab reports and course papers in Markdown, get a DOCX formatted to [DSTU 3008:2015](https://uk.wikipedia.org/wiki/%D0%94%D0%A1%D0%A2%D0%A3_3008:2015) (ДСТУ 3008:2015 «Інформація та документація. Звіти у сфері науки і техніки. Структура та правила оформлювання»).

The tool reads `endeavour.toml`, loads `.md` and `.docx` files from the configured pages directory, applies `[[VARNAME]]` substitutions, and writes a single DOCX. Pages that are easier to make in Word, such as a title page or a task form, can be dropped in as `.docx` files and are merged as they are.

## Quick start

Requires Node.js 18 or newer.

```bash
npx @popovnx/md-to-docx-dstu init      # creates endeavour.toml
mkdir pages
```

```
.
├── endeavour.toml
├── images/
│   └── scheme.png
└── pages/
    ├── 1.docx       # title page made in Word
    ├── 2.md
    └── 3.md
```

```bash
npx @popovnx/md-to-docx-dstu           # writes output.docx
```

## Usage

```bash
npx @popovnx/md-to-docx-dstu init
npx @popovnx/md-to-docx-dstu
```

The default command is `render`, so the second line is equivalent to:

```bash
npx @popovnx/md-to-docx-dstu render
```

If the package has been installed globally or linked locally, the exposed CLI command is:

```bash
md-to-docx init
md-to-docx
```

## Configuration

```toml
extends = "../common.toml"   # optional: shared settings, overridden key by key by this file

[dirs]
pages  = "./pages"           # .md and .docx input files, sorted numerically (2.md before 10.md)
images = "./images"          # images referenced in markdown
output = "./output.docx"

[vars]                       # [[VARNAME]] is replaced in markdown and in injected .docx files
LAB_NUMBER = "1"

[page]                       # optional, mm; default is DSTU 3008:2015 — 20/10/20/25
margins = { left = 30 }      # any side may be overridden

[meta]                       # optional: document properties (File → Info), [[VARNAME]] works here
author = "[[LAB_AUTHOR]]"    # creator / last modified by; empty when not set
title  = "Лабораторна робота №[[LAB_NUMBER]]"
```

Paths are resolved relative to the file that declares them, so a shared `common.toml` one level up can hold the author, group, teacher and margins for every lab:

```toml
# common.toml
[vars]
LAB_AUTHOR  = "Іванов Іван Іванович"
LAB_TEACHER = "доц. Петренко П. П."
```

## Supported content

- Headings (`#` and `##`)
- Paragraphs with inline formatting
- Bullet and numbered lists (`1.` or `1)` as written, every list restarts at 1)
- Tables: full text width, columns sized so words don't break, header row repeated on every page
- Fenced code blocks
- Images (never wider than the text) and centered captions
- Explicit `\newpage` markers
- Horizontal rules (`---`) as an empty line
- `.docx` pages, e.g. a title page or a task form

## Layout

DSTU 3008:2015: A4, Times New Roman 14 pt black for all text (code and links included), 1.5 line spacing,
1 cm paragraph indent, page numbers top right. A title page placed first counts as page 1 but shows no number.
Lists start with a dash or number at the paragraph indent. A figure stays on the page of its caption, a table caption
on the page of its table. No blank line between consecutive headings, also when they come from different files.
Table text is 12 pt single-spaced (the standard leaves it to the author).

## Injected .docx pages

A `.docx` page is merged as its own section and looks exactly as when opened on its own: its page setup, styles,
lists, headers / footers, images and theme come along, and nothing of it leaks into the Markdown pages.
It always starts on a new page, so a `\newpage` right before or after it is dropped instead of producing a blank page.

## Development

```bash
npm install
npm test            # node:test suite (tsx)
npm run typecheck   # tsc --noEmit
npm run build       # bundles dist/index.js
npm run dev         # run the CLI from source
```
