# md-to-docx

TypeScript generator that reads the course Markdown fragments in `../pages` and emits a single DOCX with `docx`, without a Pandoc merge step.

## Why this exists

The Markdown sources already encode the thesis structure: chapter headings, section headings, page breaks, figures, and code blocks. Generating DOCX directly keeps the first page and document flow under one renderer instead of splitting layout across multiple tools.

## Run

```bash
cd md-to-docx
yarn build
yarn start -- --input ../pages --images ../images --output ../Result.docx
```

For iterative development:

```bash
yarn dev -- --input ../pages --images ../images --output ../Result.docx
```

## What it handles

- Headings (`#` and `##`)
- Paragraphs with inline formatting
- Bullet and numbered lists
- Fenced code blocks
- Images with Pandoc-style attributes like `.{.center width=500px}`
- Explicit `\newpage` markers
- `.docx` pages inserted between Markdown pages, e.g. a title page

## Notes

- The generator expects images in a separate folder, defaulting to `../images` relative to this project.
- The layout is intentionally conservative: A4 portrait, Arial body text, centered figure captions, and monospaced code blocks.
- If you later want a dedicated title page, this project can be extended with a front-page Markdown file or a separate cover template.