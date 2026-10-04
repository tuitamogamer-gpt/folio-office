# Folio

A free, local-first office workspace with document, spreadsheet, and presentation editors. Built with React, TypeScript, Vite, Tiptap, SheetJS, docx, and PptxGenJS.

## Run

```sh
npm install
npm run dev
```

Open http://localhost:5173. `npm run build` produces the production app in `dist/`.

## Deploy to Vercel

Import the Git repository into Vercel and deploy from the repository root. The included `vercel.json` configures the Vite build and `dist/` output. No environment variables or backend services are required.

Browser storage is specific to each site address. Documents saved at localhost or a preview URL will not appear automatically on the production site; export and import them to move them between addresses.

## Features

- Document editor with headings, fonts, text formatting, lists, tables, images, links, find/replace, document layout, and DOCX, HTML, TXT, or print-to-PDF export.
- Spreadsheet editor with formulas, cell ranges, keyboard navigation, paste, undo/redo, XLSX and CSV export. Supports SUM, AVERAGE, MIN, MAX, COUNT, ABS, ROUND and arithmetic.
- Presentation editor with slide layouts, themes, add/duplicate/reorder/delete, presentation mode, PPTX and print-to-PDF export.
- Import DOCX, XLSX/XLS, CSV, PPTX, TXT, HTML, and basic Markdown.
- Search, favorites, rename, duplicate, download, trash/restore, grid/list views, and editable templates.
- Browser localStorage persistence, local profile, responsive mobile workspace.

## Compatibility and storage

Folio is an independent application, not Microsoft Office or a complete replacement for every Office feature. Complex formatting, macros, embedded charts, animations, and arbitrary slide object positions are not preserved. Spreadsheet imports use the first sheet; additional sheets are not loaded. The editors and import notices explain these limits. Keep original source files for fidelity-sensitive work.

Files save in the current browser on the current device. There is no server, account, cloud synchronization, or collaboration backend. Download important files before clearing browser data. Office-format exports are genuine Office Open XML files.

## Verification

`npm run build` validates TypeScript and builds for production. Browser smoke scripts in `tests/` require a local dev server on port 5173 and Chromium at `/usr/bin/chromium` (or update executablePath for your system).
