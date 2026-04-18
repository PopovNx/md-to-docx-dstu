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

`endeavour.toml` controls three paths:

- `pages`: input directory for `.md` and `.docx` files
- `images`: image directory used while rendering Markdown
- `output`: final DOCX file path

`npx @popovnx/md-to-docx-dstu init` writes a starter config with comments.

## Supported content

- Headings (`#` and `##`)
- Paragraphs with inline formatting
- Bullet and numbered lists
- Tables
- Fenced code blocks
- Images and centered captions
- Explicit `\newpage` markers

## Notes

- Paths in `endeavour.toml` are resolved relative to that file.
- The layout is conservative: A4 portrait, Arial body text, centered figure captions, and monospaced code blocks.
