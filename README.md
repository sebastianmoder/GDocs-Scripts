# Google Docs Apps Script Utilities

This repository stores Google Apps Script (`.gs`) utilities for Google Docs, with an emphasis on academic manuscript writing and editing workflows.

Most scripts are intended to be copied into a document-bound Apps Script project or managed with `clasp` when a script grows large enough to need local development.

## Current Scripts

### `tidyApaStatNotation.gs`

Adds a Google Docs menu item named `APA Cleanup` with the command `Tidy statistical notation`.

The script is designed for APA 7 statistical notation cleanup in manuscript drafts. It:

- adds non-breaking spaces around comparison operators
- converts `>=` and `<=` to APA-friendly comparison symbols
- converts hyphen-minus signs in negative numbers to true minus signs
- formats common statistics such as `R2`, chi-squared, and partial eta squared
- italicizes selected Latin statistical symbols before operators
- normalizes spacing in degrees-of-freedom expressions

Run it on a clean copy of a manuscript. The script deliberately avoids broad text replacement methods and whole-paragraph rewrites, but formatting tools that modify rich Google Docs text should still be tested on a duplicate document first.

## Typical Use

1. Open the target Google Doc.
2. Go to `Extensions` -> `Apps Script`.
3. Add or paste the relevant `.gs` file into the Apps Script editor.
4. Save the Apps Script project.
5. Reload the Google Doc.
6. Use the custom menu added by the script.

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
