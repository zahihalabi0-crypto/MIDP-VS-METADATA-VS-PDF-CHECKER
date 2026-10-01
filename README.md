# DrawingCheck

A browser-based BIM / Document Control checker that compares:

- MIDP drawing records
- metadata export records
- digitally generated PDF title blocks

Files are processed locally in the browser and are not sent to a server.

## Run locally

```bash
npm install
npm run dev
```

## Production build

```bash
npm run build
```

## Deploy to Vercel

1. Push this folder to a GitHub repository.
2. Import the repository at Vercel.
3. Vercel will detect Vite automatically.
4. Use `npm run build` as the build command and `dist` as the output directory.

## Current scope

- Supports the supplied MIDP and metadata layouts.
- Supports multiple digitally generated drawing PDFs.
- Normalizes PDF/CAD suffixes, Unicode hyphens, and numeric revisions.
- Compares drawing number, revision, title, issue date, and model reference.
- Exports an Excel discrepancy report.
- Scanned title blocks require a future OCR module.

Before production use, validate field ownership and severity rules with the BIM Coordinator and Document Control team.
