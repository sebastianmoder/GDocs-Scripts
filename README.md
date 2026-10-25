# Google Docs Apps Script Utilities

This repository stores Google Apps Script (`.gs`) utilities for Google Docs, with an emphasis on academic manuscript writing and editing workflows.

Most scripts are intended to be copied into a document-bound Apps Script project or managed with `clasp` when a script grows large enough to need local development.

## Current Scripts

### Menu template: `menu.gs`

Copy `menu.gs` into the same document-bound Apps Script project as the utilities you want to use. It defines an `onOpen()` function that creates the `Document Utilities` menu.

All `.addItem(...)` lines are commented out by default. Uncomment the line for each script imported into the project:

- `Count words by heading` runs `showHeadingWordCounts` from `showHeadingWordCounts.gs`.
- `Tidy APA statistical notation` runs `tidyApaStatNotation` from `tidyApaStatNotation.gs`.
- `Highlight references…` runs `highlightReferences` from `highlightReferences.gs`.

Save the project and reload the Google Doc to display the enabled menu items. Keep only one `onOpen()` function in the project; if you already have one, merge the menu setup into it. You can also run any utility directly from the Apps Script editor without importing the menu template.

### `tidyApaStatNotation.gs`

The script is designed for APA 7 statistical notation cleanup in manuscript drafts. It:

- adds non-breaking spaces around `=`, `>`, `<`, `≥`, `≤`, `≠`, `≈`, and `±`; converts `>=`, `<=`, and `!=` to symbols
- spaces binary `+` and `−` between recognized numeric/statistical operands; converts hyphen subtraction involving a known statistic (such as `n-1`) while preserving numeric hyphen ranges such as `12-24`
- converts hyphen-minus and en-dash signs in negative numbers to true minus signs
- formats squared `R`, `r`, `f`, `s`, `k`, `I`, `T`, `χ`, `η`, `ω`, `ε`, and `σ`, accepting plain `2`, `²`, and `^2` forms; Latin bases are italic, Greek bases are roman, and the exponent is roman superscript
- formats partial/generalized eta squared (`η_p^2`, `η_g²`, also without underscores), correlations (`r_s`, `r_pb`, `r_tb`, also `rₛ`), and Bayes factors (`BF10`, `BF01`, also underscore or Unicode subscript forms); identifying subscripts are roman
- converts every remaining `²` glyph in processed body text to a regular `2` with superscript formatting
- italicizes Latin statistics in expressions, including `Mdn`, `MS`, `MSE`, `SSE`, `SS`, `U`, `W`, `H`, `Q`, `D`, `V`, `g`, `f`, `s`, and `OR`, alongside the original symbols; plural suffixes such as those in `Ms`, `SDs`, and `rs` remain roman
- recognizes phrases such as `Cohen’s d`, `Hedges’ g`, `p values`, `t-test`, and `z scores`, and standalone statistical labels occupying a whole table cell
- makes standalone Greek statistical letters roman
- removes spaces before numeric degrees-of-freedom parentheses (`t (42)` → `t(42)`, `F (2,145)` → `F(2, 145)`) and inserts one non-breaking space after each df comma, including when no space existed

Matching does not span text elements. Standalone Latin letters in ordinary prose are left alone. Whole-cell statistical labels are recognized by their text, without inferring a table's structure or meaning. Reverse italics in italicized titles are not inferred automatically. Numerical values, leading zeros, and precision are unchanged.

Run it on a clean copy of a manuscript. The script deliberately avoids broad text replacement methods and whole-paragraph rewrites, but formatting tools that modify rich Google Docs text should still be tested on a duplicate document first.

#### Notation validation

Run `node tests/tidyApaStatNotation.test.js` from the repository root for local regression checks using a rich-text mock. These verify text, italics, subscripts/superscripts, preservation of surviving characters' bold/color/link attributes, repeated runs, and table traversal. They do not replace testing in Google Docs.

For manual validation in a duplicate Doc, check the following; spaces beside operators and after df commas should be non-breaking. Live Google Docs results have not yet been recorded.

| Input | Expected result |
| --- | --- |
| `t (42)=2.10; F (2,145)=4.2` | `t(42) = 2.10; F(2, 145) = 4.2`, with italic test symbols |
| `Mdn=.5; U=15; OR=2.1; SDs>.1` | Spaced operators, italic statistics, roman plural `s` |
| `r^2=.2; f²=.15; ω2=.1` | Superscript exponents; italic `r`/`f`, roman `ω` |
| `η_g^2=.1; r_pb=.3; BF_10=4` | Roman identifying subscripts, superscript eta exponent, italic Latin bases |
| `p!=.05; b≈-.04; M±SD; n-1; 12-24` | `p ≠ .05; b ≈ −.04; M ± SD; n − 1; 12-24` |
| `Cohen’s d; p values; t-test; z scores` | Only the statistical symbols become italic |
| Standalone `M`/`SD` in table cells and `M` in ordinary prose | Table labels become italic; standalone prose letter stays unchanged |
| Mixed bold/color/link formatting across `η_p^2` and `t (42)` | Surviving characters retain those attributes; only notation italics/alignment change |
| Repeat a run; include text in another tab/header/footer | Same result on rerun; excluded areas remain unchanged |

### `showHeadingWordCounts.gs`

Opens a read-only word-count report for the active Google Docs tab. Before running it, enable the Google Docs API in the Apps Script editor under `Services (+)` -> `Google Docs API` -> `Add`.

The report includes:

- word counts for H1–H6 sections in the active tab, including nested sections in their parent counts
- totals for the current tab and the entire document, including nested tabs
- text in paragraphs, lists, and tables; generated tables of contents are excluded
- plain text, Markdown, table (TSV), CSV, and nested JSON formats, with copy and download buttons

Counts use a preview with all suggestions accepted. The script never changes the document or accepts or rejects suggestions. Section counts exclude heading text by default; set `INCLUDE_HEADING_TEXT` to `true` in `showHeadingWordCounts()` to include it. Tab and document totals always include headings and text before the first heading.

### `highlightReferences.gs`

Run `highlightReferences` to open a dialog with three independently selectable categories. Citations are selected initially. The highlighter, dialog HTML, and all helpers are contained in **one `.gs` file**: you can paste only `highlightReferences.gs` into a document-bound project and run it from the editor. The menu template is optional. No advanced Google Docs API service is required.

| Category | Color | Examples |
| --- | --- | --- |
| Citations | Pale yellow `#FFF3B0` | `(Smith, 2020)`, `(Jones, n. d.)` |
| Materials | Pale blue `#B8E8FF` | `Table 1a`, `Figure S3a`, `Appendix 2`, `Supplementary Material` |
| Page numbers | Pale pink `#FFD1E8` | `(p. 12)`, `[p. 12]`, `p.[12]`, `pp. iv–vi` |

These three distinct colors were checked against `avoidcolors.csv`; none appears in that file. Colors are configured in `REFERENCE_HIGHLIGHT_COLORS` at the top of the script. The CSV is a development reference and does not need to be uploaded to Apps Script.

**Citations:** A parenthetical passage matches if it contains a year from 1900 to 2099, optionally followed by a letter, or `n.d.` / `n. d.`. The entire passage, including parentheses, is highlighted and counted once, even with multiple sources. This retains the original date-based heuristic: `(the project began in 2020)` also matches. Square-bracket citations and passages spanning paragraph breaks are skipped. With nested parentheses, only qualifying inner passages match. The original `highlightParentheticalCitations` function remains available for citation-only runs and existing menus.

**Materials:** Case-insensitive matching covers Table(s), Figure(s), Appendix/Appendices, `Tab./Tabs.`, `Fig./Figs.`, `App.`, and Supplemental/Supplementary Material(s). Parentheses are not required. Identifiers can be Arabic numbers, letters, S-prefixed numbers, or identifiers such as `A1`, with panel suffixes such as `1a`, `S3a`, or `2(a)`. Lists and ranges include `Tables 1 and 2`, `Figures 1a, 2b, and S3c`, `Figures 1a–c`, `Figures 1 and Figure 2`, and `Appendices A–C` or `Appendices 1–3`. Supplemental/Supplementary Material(s) also matches without an identifier; other labels require one. The label and reference sequence are highlighted, excluding surrounding parentheses and sentence punctuation.

**Pages:** A `p.` or `pp.` label is required, case-insensitively. Optional spaces, including nonbreaking spaces, are accepted: `p.12`, `p. 12`, `[p. 12]`, `p. [12]`, `pp. [12–15]`, and `(p. 12)` all match. Arabic and valid Roman numerals, comma/conjunction lists, repeated labels, and hyphen/en-dash/em-dash ranges are supported. Balanced brackets or parentheses are included when they enclose only the locator; `(see p. 12)` highlights just `p. 12`. Bare numbers and years do not match. Symbolic placeholders such as `p. x` are interpreted as Roman numerals when valid; `[x]` without a page label is skipped.

The dialog reports matches found for each selected category. Each complete list/range counts as one match; counts are occurrences found, not newly changed highlights. If categories overlap, citation color takes priority: `(Smith, 2020, p. 12)` remains entirely yellow when both Citations and Page numbers are selected. Materials take priority over pages in any remaining overlap. Counts still include matches covered by a higher-priority category. With only Page numbers selected, the same passage highlights just `p. 12`.

Processing is restricted to the active tab's body, including ordinary table text. Headers, footers, footnotes, other tabs, and non-text objects are not processed. Matches do not cross paragraph boundaries. Matching uses textual patterns rather than validating references against captions or a bibliography, so a matching caption or generic phrase such as `Appendix A` can also be highlighted.

Run it on a duplicate manuscript first. Only backgrounds within matches are overwritten; text, italics, bold, links, superscripts, and other formatting are preserved. Reruns do not clear existing highlights, including highlights from a previously selected category. Use Google Docs Undo to revert a run.

To remove highlights, select the categories to clear and click the red **Clear highlights** button in the same dialog. It scans background colors in the current tab's body, including tables, independently of reference matching, so edited text is cleared too. All instances of the selected colors are removed, including manually applied identical colors. Other background colors and text formatting remain unchanged. Clearing removes the background; it cannot restore an earlier background overwritten by highlighting. Its counts represent contiguous spans of each color cleared, rather than reference matches. Highlight and Clear highlights are both disabled while either action is running.

#### Highlighter validation

Local checks exercise exact matching offsets, exclusions, all seven category combinations, overlap priority, color exclusions, mocked background-only writes, repeat runs, legacy entry-point behavior, and dialog submission/success/failure handling. These checks do not replace validation in Google Docs.

For manual validation, paste a mixed-format sample into a duplicate Google Doc and record expected and observed results:

| Scenario | Expected result |
| --- | --- |
| `(Smith, 2020, p. 12; see Figure S3a)` with all categories selected | One match per category; whole passage yellow |
| Same passage with Materials and Page numbers only | `Figure S3a` blue and `p. 12` pink |
| `Tables 1a–c; Appendices 1–3; Supplementary Material` | Three blue material matches |
| `(p. 12); [p. 12]; p.[12]; pp. iv–vi` | Four pink page matches, including their wrappers |
| Bold/italic/link/superscript text within a reference | Background changes; original formatting and text remain |
| References in a table and in two document tabs | Active tab's body/table matches change; other tab remains unchanged |
| No selected categories, then a run with no matching text | Empty selection rejected; no-match run reports zero |
| Repeat a run, then select a different category | Counts repeat; previous highlights remain |
| Select Materials and click Clear highlights after editing highlighted text | All blue spans in the active tab's body are cleared; other colors remain |
| Clear a span split by bold/italic/link formatting | Entire matching color span is cleared; original formatting remains |
| Clear again, or clear an empty tab | Zero spans cleared; other tabs remain unchanged |
| Authorization or server failure | Error appears in the dialog; both action buttons become available again |

Live Google Docs results have not been recorded in this checkout; there is no configured document-bound deployment target.

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
