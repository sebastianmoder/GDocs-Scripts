# Google Docs Apps Script Utilities

This repository stores Google Apps Script (`.gs`) utilities for Google Docs, with an emphasis on academic manuscript writing and editing workflows.

Most scripts are intended to be copied into a document-bound Apps Script project or managed with `clasp` when a script grows large enough to need local development.

## Current Scripts

### Menu template: `menu.gs`

Copy `menu.gs` into the same document-bound Apps Script project as the utilities you want to use. It defines an `onOpen()` function that creates the `Document Utilities` menu.

Both `.addItem(...)` lines are commented out by default. Uncomment the line for each script imported into the project:

- `Tidy statistical notation` runs `tidyApaStatNotation` from `tidyApaStatNotation.gs`.
- `Words by heading` runs `showHeadingWordCounts` from `showHeadingWordCounts.gs`.

Save the project and reload the Google Doc to display the enabled menu items. Keep only one `onOpen()` function in the project; if you already have one, merge the menu setup into it. You can also run either utility directly from the Apps Script editor without importing the menu template.

### `tidyApaStatNotation.gs`

The script is designed for APA 7 statistical notation cleanup in manuscript drafts. It:

- adds non-breaking spaces around comparison operators
- converts `>=` and `<=` to APA-friendly comparison symbols
- converts hyphen-minus signs in negative numbers to true minus signs
- formats common statistics such as `R2`, chi-squared, and partial eta squared
- italicizes selected Latin statistical symbols before operators
- normalizes spacing in degrees-of-freedom expressions

Run it on a clean copy of a manuscript. The script deliberately avoids broad text replacement methods and whole-paragraph rewrites, but formatting tools that modify rich Google Docs text should still be tested on a duplicate document first.

### `showHeadingWordCounts.gs`

Opens a read-only word-count report for the active Google Docs tab. Before running it, enable the Google Docs API in the Apps Script editor under `Services (+)` -> `Google Docs API` -> `Add`.

The report includes:

- word counts for H1–H6 sections in the active tab, including nested sections in their parent counts
- totals for the current tab and the entire document, including nested tabs
- text in paragraphs, lists, and tables; generated tables of contents are excluded
- plain text, Markdown, table (TSV), CSV, and nested JSON formats, with copy and download buttons

Counts use a preview with all suggestions accepted. The script never changes the document or accepts or rejects suggestions. Section counts exclude heading text by default; set `INCLUDE_HEADING_TEXT` to `true` in `showHeadingWordCounts()` to include it. Tab and document totals always include headings and text before the first heading.

## Typical Use

1. Open the target Google Doc.
2. Go to `Extensions` -> `Apps Script`.
3. Add or paste the relevant utility `.gs` files and `menu.gs` into the Apps Script editor.
4. Uncomment the matching menu items in `menu.gs`, enable the Google Docs API if using word counts, and save the Apps Script project.
5. Reload the Google Doc.
6. Run the utility from the `Document Utilities` menu.

For scripts that require authorization, Google will prompt for access the first time the function runs.

## Repository Notes

- Keep reusable source files in `.gs` files.
- Keep document-specific drafts, exports, and manuscript files out of this repository unless they are intentionally part of the script documentation.
- Do not commit OAuth tokens, local `clasp` credentials, or environment files.
- Prefer descriptive script filenames, such as `tidyApaStatNotation.gs` or `insertReviewerResponseTable.gs`.

## Optional `clasp` Workflow

This repository can be used without `clasp`. If you later manage scripts with the Apps Script CLI:

- commit source files such as `.gs`, `.js`, `.html`, and `appsscript.json`
- avoid committing local credentials or personal script IDs unless a project is intentionally shared
- use `.claspignore` when a specific Apps Script project should upload only selected files
