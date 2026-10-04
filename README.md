# Folio

A free, local-first office workspace with document, spreadsheet, and presentation editors. Built with React, TypeScript, Vite, Tiptap, SheetJS / xlsx-js-style, docx, and PptxGenJS.

## Run

```sh
npm install
npm run dev
```

Open http://localhost:5173. `npm run build` produces the production app in `dist/`.

## Deploy to Vercel

The live app is https://folio-office.vercel.app. In GitHub it lives in `BoroGPT/codex` under `folio-office/`; use `folio-office` as the Vercel Root Directory when importing that repository. For a standalone checkout, deploy from the application root. The included `vercel.json` configures the Vite build and `dist/` output. No environment variables or backend services are required.

Browser storage is specific to each site address. Documents saved at localhost or a preview URL will not appear automatically on the production site; export and import them to move them between addresses.

## Features

### Documents

- Rich text, fonts, headings, alignment, line spacing, indentation, subscript/superscript, case conversion, and format painter.
- Tables with row/column controls, merged cells, header rows, and cell shading; images with width and alignment controls.
- Text comments with resolve/reopen, browser spellcheck, find/replace, and document outline.
- Page breaks, A4/Letter/Legal paper, orientation, margins, headers, footers, and page numbers.
- DOCX export with native comments and page settings; HTML, TXT, and print-to-PDF.

### Spreadsheets

- Multiple worksheets with add, rename, duplicate, delete, and cross-sheet references.
- Cell/range text formatting, colors, alignment, number/currency/percentage/date formats, and resizable columns.
- Whole-row sorting, text filtering, fill down/right with relative references, find/replace, frozen top row, gridline/formula views, and zoom.
- Arithmetic, comparisons, logical conditions, text/date functions, conditional sums/counts, and error handling. Formulas run through a restricted parser without JavaScript execution.
- Workbook undo/redo, range selection, keyboard navigation, and tabular paste.
- XLSX export preserves all worksheets, supported formatting, widths, formulas, and frozen rows. CSV exports the active worksheet.

### Presentations

- Title, content, split, and blank layouts; fonts, text sizes, themes, and colors.
- Movable/resizable text boxes, pictures, and rectangle/ellipse/line shapes; object formatting, layering, rotation, duplication, and deletion.
- Speaker notes, hidden slides, fade/push transitions, presentation mode, undo/redo, and keyboard shortcuts.
- PPTX export with editable native text/shapes/images, notes, hidden slides, and transitions; print-to-PDF.

### Workspace

- Import DOCX, XLSX/XLS, CSV, PPTX, TXT, HTML, and basic Markdown.
- Search, favorites, rename, duplicate, trash/restore, list/grid views, and editable templates.
- Save up to 10 named versions per file; restoring a version first saves the current draft.
- Download a complete `.folio.json` workspace backup including versions; restore files as copies.
- IndexedDB storage with automatic migration of earlier files, accurate save indicators, local profile, and responsive layouts.

## Compatibility and storage

Folio is an independent application, not Microsoft Office or a complete replacement for every Office feature. Advanced Office features such as macros, embedded charts, tracked changes, complex animation sequences, and unsupported drawing types are not fully preserved. Formula support is a useful subset of Excel, not its entire calculation engine. The editors and import notices explain these limits. Keep original source files for fidelity-sensitive work.

Files save in the current browser on the current device. There is no server, account, cloud synchronization, or collaboration backend. Download important files or a workspace backup before clearing browser data. Saved versions share the browser storage quota; large pictures and many versions can fill it. Office-format exports are genuine Office Open XML files.

## Verification

`npm run build` validates TypeScript and builds for production. Browser scripts require a local dev server on port 5173 and Chromium at `/usr/bin/chromium` (or update executablePath for your system). The document suite also uses `pdftotext` to inspect actual PDF output.

With the Vite dev server running:

```sh
node tests/verify-workspace.mjs
node tests/workspace-history.mjs
node tests/workspace-storage.mjs
node tests/workspace-async.mjs
node tests/document-tools.mjs
node tests/doc-advanced-exports.mjs
node tests/spreadsheet-editor.mjs
node --test tests/spreadsheet-formulas.test.mjs
node tests/verify-slides.mjs
node tests/office-roundtrip.mjs
```

`BASE_URL=https://folio-office.vercel.app node tests/verify-workspace.mjs` checks the deployed UI using a fresh browser profile. Roundtrip tests inspect native Office ZIP/XML files; they do not replace rendering checks in desktop Microsoft Office.
