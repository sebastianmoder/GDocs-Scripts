/**
 * Applies APA 7 statistical notation cleanup to the active Google Docs tab.
 *
 * Processes body paragraphs, lists, and tables, within individual text elements.
 * Excludes headers, footers, footnotes, and other tabs. Changes are local:
 * surviving characters retain their rich-text attributes except for the
 * notation-specific italics and alignment described below.
 * Visible URLs and linked text are excluded, including their formatting.
 *
 * - Spaces comparisons (=, >, <, ≥, ≤, ≠, ≈), ±, and recognized binary +/−
 *   expressions with non-breaking spaces; converts >=, <=, and !=.
 * - Converts unary hyphen/en-dash negatives and hyphen subtraction involving
 *   a statistic to true minus signs; keeps numeric hyphen ranges unchanged.
 *   Removes spaces before numeric t/F/χ² df parentheses
 *   and inserts a non-breaking space after df commas, even if none exists.
 * - Italicizes common Latin statistics in expressions, recognized phrases
 *   (Cohen's d, p values, t-test, etc.), and standalone table-cell labels.
 *   Plural s stays roman. Ordinary prose letters are not globally italicized.
 * - Normalizes squared R/r/f/s/k/I/T/χ/η/ω/ε/σ notation (2, ², or ^2),
 *   indexed η_p²/η_g², r_s/r_pb/r_tb, and BF10/BF01 (also BF_10/BF_01).
 *   Greek bases and identifying subscripts/exponents are roman.
 * - Converts every remaining ² glyph to a roman superscript 2.
 * - Makes standalone Greek statistical letters roman. Reverse italics in
 *   already italicized titles are not inferred automatically.
 *
 * Does not change numerical precision, leading zeros, or reported values.
 */

const APA_LATIN_STATS_ = [
  "Mdn", "MSE", "SSE", "SD", "SE", "MS", "SS", "OR", "df",
  "p", "r", "R", "t", "F", "z", "M", "b", "B", "d", "g", "f",
  "s", "k", "n", "N", "U", "W", "H", "Q", "D", "V"
];
const APA_GREEK_STATS_ = "αβγδΔεηθκλμνξπρσΣτφΦχψωΩ";
const APA_HORIZONTAL_SPACE_ = "[ \\t\\u00A0\\u202F]";

function tidyApaStatNotation(confirmed) {
  if (confirmed !== true) {
    showApaStatConfirmation_();
    return;
  }
  const document = DocumentApp.getActiveDocument();
  const tab = document && document.getActiveTab();
  if (!tab) throw new Error('Open an editable Google Doc and run APA cleanup from Extensions.');
  processElement_(tab.asDocumentTab().getBody());
}

function showApaStatConfirmation_() {
  const html = HtmlService.createHtmlOutput(`
    <!doctype html>
    <html><head><base target="_top"><style>
      body { font: 14px Arial, sans-serif; color: #202124; padding: 12px; }
      p { line-height: 1.5; margin: 0 0 20px; }
      .actions { display: flex; justify-content: flex-end; gap: 8px; }
      button { padding: 9px 16px; border: 0; border-radius: 4px;
        background: #1a73e8; color: #fff; font: inherit; cursor: pointer; }
      #cancel { background: #b3261e; }
      #cancel:hover:not(:disabled) { background: #8c1d18; }
      button:focus-visible { outline: 2px solid #202124; outline-offset: 3px; }
      button:disabled { opacity: 0.55; cursor: default; }
    </style></head><body>
      <p id="status" role="status" aria-live="polite">This script will change text
        and statistical formatting in the active tab of your document.</p><p>It might be wise to test-run this script on a duplicate copy first.</p>
      <div class="actions">
        <button id="cancel" type="button" autofocus>Cancel</button>
        <button id="run" type="button">Confirm &amp; run</button>
        <button id="close" type="button" hidden>Close</button>
      </div>
      <script>
        const status = document.getElementById("status");
        const cancel = document.getElementById("cancel");
        const run = document.getElementById("run");
        const close = document.getElementById("close");
        cancel.addEventListener("click", function() { google.script.host.close(); });
        close.addEventListener("click", function() { google.script.host.close(); });
        function finish(message) {
          status.textContent = message;
          run.hidden = cancel.hidden = true;
          close.hidden = false;
          close.focus();
        }
        run.addEventListener("click", function() {
          if (run.disabled) return;
          run.disabled = cancel.disabled = true;
          status.textContent = "Cleaning up APA statistical notation…";
          google.script.run.withSuccessHandler(function() {
            finish("APA statistical notation cleanup complete.");
          }).withFailureHandler(function(error) {
            status.textContent = "Cleanup failed: " + error.message +
              " Some edits may have completed. Check the document or use Undo before trying again. " +
              "If access was denied, reopen the tool from Extensions and allow the requested permissions.";
            run.disabled = cancel.disabled = false;
            run.focus();
          }).tidyApaStatNotation(true);
        });
      </script>
    </body></html>
  `).setWidth(420).setHeight(240);
  DocumentApp.getUi().showModalDialog(html, "Confirm APA cleanup");
}

function processElement_(element, isTableLabel) {
  if (element.getType() === DocumentApp.ElementType.TABLE_CELL) {
    // A whole cell must be a label; a lone text fragment in prose is not enough.
    isTableLabel = isStatLabel_(element.getText().trim());
  }
  if (element.getType() === DocumentApp.ElementType.TEXT) {
    processTextElement_(element.asText(), isTableLabel);
    return;
  }
  if (!element.getNumChildren) return;
  for (let i = 0; i < element.getNumChildren(); i++) {
    processElement_(element.getChild(i), isTableLabel);
  }
}

function processTextElement_(textElement, isTableLabel) {
  const ranges = protectedUrlRanges_(textElement);
  if (!ranges.length) {
    processStatText_(textElement, isTableLabel);
    return;
  }
  // Work backwards so edits never shift the original offsets still to process.
  let end = textElement.getText().length;
  for (let i = ranges.length - 1; i >= 0; i--) {
    if (ranges[i].end < end) {
      processStatText_(statTextRange_(textElement, ranges[i].end, end), isTableLabel);
    }
    end = ranges[i].start;
  }
  if (end > 0) processStatText_(statTextRange_(textElement, 0, end), isTableLabel);
}

/** Returns merged, end-exclusive spans for visible URLs and hyperlink text. */
function protectedUrlRanges_(textElement) {
  const text = textElement.getText();
  // Consume URL tails conservatively, including query operators and punctuation.
  // Also recognize addresses without a scheme (e.g., doi.org/10.1234/r^2).
  const tail = "[^\\s\"“”]";
  const scheme = "(?:[a-z][a-z\\d+.-]*://|(?:mailto|tel|urn|doi|data):|www\\.)";
  const host = "(?:(?:[\\p{L}\\p{N}-]+\\.)+[\\p{L}]{2,}|(?:\\d{1,3}\\.){3}\\d{1,3}|localhost)";
  // Include an opening angle bracket so a wrapped URL is not spaced as '<'.
  const pattern = new RegExp("<?(?:" + scheme + tail + "+|" + host +
    "(?::\\d+)?(?:[/?#]" + tail + "*)?)", "giu");
  const ranges = [...text.matchAll(pattern)].map(match => ({
    start: match.index, end: match.index + match[0].length
  }));
  const indices = textElement.getTextAttributeIndices();
  for (let i = 0; i < indices.length; i++) {
    const start = indices[i];
    if (start < text.length && textElement.getLinkUrl(start)) {
      ranges.push({start, end: i + 1 < indices.length ? indices[i + 1] : text.length});
    }
  }
  ranges.sort((a, b) => a.start - b.start);
  const merged = [];
  for (const range of ranges) {
    const last = merged[merged.length - 1];
    if (last && range.start <= last.end) {
      last.end = Math.max(last.end, range.end);
    } else {
      merged.push({start: range.start, end: range.end});
    }
  }
  return merged;
}

/** A local-offset Text view keeps matching and edits outside protected spans. */
function statTextRange_(textElement, start, end) {
  return {
    getText: () => textElement.getText().slice(start, end),
    getAttributes: index => textElement.getAttributes(start + index),
    setAttributes: (first, last, attrs) => textElement.setAttributes(start + first, start + last, attrs),
    setItalic: (first, last, value) => textElement.setItalic(start + first, start + last, value),
    setTextAlignment: (first, last, value) => textElement.setTextAlignment(start + first, start + last, value),
    deleteText: (first, last) => {
      textElement.deleteText(start + first, start + last);
      end -= last - first + 1;
    },
    insertText: (index, value) => {
      textElement.insertText(start + index, value);
      end += value.length;
    }
  };
}

function processStatText_(textElement, isTableLabel) {
  normalizeMinusSigns_(textElement);
  normalizeComparisonOperators_(textElement);
  normalizeArithmeticOperators_(textElement);
  formatIndexedStatistics_(textElement);
  formatSquaredStatistics_(textElement);
  normalizeSquaredGlyphs_(textElement);
  normalizeDfSpacing_(textElement);
  italicizeStatSymbols_(textElement, isTableLabel);
  formatGreekStatistics_(textElement);
}

/** Unicode-aware boundaries avoid matching statistics inside names/Greek words. */
function statMatches_(text, source) {
  const pattern = new RegExp(
    "(^|[^\\p{L}\\p{N}_])(" + source + ")(?![\\p{L}\\p{N}_])", "gu"
  );
  return [...text.matchAll(pattern)].map(match => ({
    index: match.index + match[1].length,
    value: match[2]
  }));
}

function isStatLabel_(text) {
  const source = APA_LATIN_STATS_.join("|") + "|[" + APA_GREEK_STATS_ + "]";
  return new RegExp("^(?:" + source + ")(?:s)?$").test(text);
}

/** Converts unary negatives, including after the newly supported operators. */
function normalizeMinusSigns_(textElement) {
  const text = textElement.getText();
  const pattern = /(^|[\s([{,;:=<>≤≥≠≈±+−×÷*/])[-–](?=(?:\.\d|\d))/g;
  const matches = [...text.matchAll(pattern)];
  for (let i = matches.length - 1; i >= 0; i--) {
    const index = matches[i].index + matches[i][1].length;
    replaceRangePreservingStyle_(textElement, index, index, "−");
  }
}

/** Comparisons retain the original broad body-text scope; ± is also binary. */
function normalizeComparisonOperators_(textElement) {
  const matches = [...textElement.getText().matchAll(/>=|<=|!=|[=<>≥≤≠≈±]/g)];
  const replacements = {">=": "≥", "<=": "≤", "!=": "≠"};
  for (let i = matches.length - 1; i >= 0; i--) {
    const match = matches[i];
    spaceOperator_(textElement, match.index, match[0], replacements[match[0]] || match[0]);
  }
}

/**
 * Distinguishes binary +/− from unary signs using complete numeric/statistical
 * operands. A hyphen between two numbers remains a potential range; a hyphen
 * involving a known statistic is subtraction. Prose hyphens are untouched.
 */
function normalizeArithmeticOperators_(textElement) {
  const text = textElement.getText();
  const matches = [...text.matchAll(/[+−-]/g)];
  for (let i = matches.length - 1; i >= 0; i--) {
    const index = matches[i].index;
    const left = text.slice(0, index).match(/([\p{L}\p{N}_.²^]+)[ \t\u00A0\u202F]*$/u);
    const right = text.slice(index + 1).match(/^[ \t\u00A0\u202F]*[−+]?([\p{L}\p{N}_.²^]+)/u);
    if (!left || !right || !isStatOperand_(left[1]) || !isStatOperand_(right[1])) continue;
    const operator = matches[i][0];
    if (operator === "-" && /^[\d.]+$/.test(left[1]) && /^[\d.]+$/.test(right[1])) continue;
    spaceOperator_(textElement, index, operator, operator === "-" ? "−" : operator);
  }
}

function isStatOperand_(token) {
  return /^(?:\d+(?:\.\d+)?|\.\d+)$/.test(token) ||
    isStatLabel_(token) ||
    /^(?:[RrfsIkTχηωεσ](?:2|²|\^2)|η_?[pg](?:2|²|\^2)|r_?(?:s|pb|tb)|BF_?(?:10|01))$/.test(token);
}

/** Edits the operator and each adjacent space separately, keeping operand styles. */
function spaceOperator_(textElement, index, original, replacement) {
  const text = textElement.getText();
  const before = text.slice(0, index).match(/[ \t\u00A0\u202F]*$/)[0];
  const after = text.slice(index + original.length).match(/^[ \t\u00A0\u202F]*/)[0];
  replaceRangePreservingStyle_(textElement, index + original.length,
    index + original.length + after.length - 1, "\u00A0");
  replaceRangePreservingStyle_(textElement, index, index + original.length - 1, replacement);
  setNotationStyle_(textElement, index, index + replacement.length - 1, false,
    DocumentApp.TextAlignment.NORMAL);
  replaceRangePreservingStyle_(textElement, index - before.length, index - 1, "\u00A0");
  // Spaces should not inherit a neighboring operand's italics or exponent.
  const operatorStart = index - before.length + 1;
  setNotationStyle_(textElement, operatorStart - 1, operatorStart - 1, false,
    DocumentApp.TextAlignment.NORMAL);
  setNotationStyle_(textElement, operatorStart + replacement.length,
    operatorStart + replacement.length, false, DocumentApp.TextAlignment.NORMAL);
}

/**
 * Deletes notation markers individually, preserving each surviving character's
 * bold/color/link attributes. Also accepts already normalized text on reruns.
 */
function normalizeNotationToken_(textElement, start, token) {
  for (let i = token.length - 1; i >= 0; i--) {
    const character = token[i];
    if (character === "_" || character === "^") {
      textElement.deleteText(start + i, start + i);
    } else if ("²₀₁ₛ".includes(character)) {
      const replacement = {"²": "2", "₀": "0", "₁": "1", "ₛ": "s"}[character];
      replaceRangePreservingStyle_(textElement, start + i, start + i, replacement);
    }
  }
  return token.replace(/[_^]/g, "").replace(/[²₀₁ₛ]/g,
    character => ({"²": "2", "₀": "0", "₁": "1", "ₛ": "s"}[character]));
}

function formatIndexedStatistics_(textElement) {
  const source = "η_?[pg](?:2|²|\\^2)|r(?:_(?:s|pb|tb)|pb|tb|ₛ)|BF_?(?:10|01|₁₀|₀₁)";
  const matches = statMatches_(textElement.getText(), source);
  for (let i = matches.length - 1; i >= 0; i--) {
    const match = matches[i];
    const token = normalizeNotationToken_(textElement, match.index, match.value);
    const isEta = token[0] === "η";
    const baseLength = token.startsWith("BF") ? 2 : 1;
    setNotationStyle_(textElement, match.index, match.index + baseLength - 1,
      !isEta, DocumentApp.TextAlignment.NORMAL);
    const subEnd = match.index + token.length - (isEta ? 2 : 1);
    setNotationStyle_(textElement, match.index + baseLength, subEnd,
      false, DocumentApp.TextAlignment.SUBSCRIPT);
    if (isEta) {
      setNotationStyle_(textElement, subEnd + 1, subEnd + 1,
        false, DocumentApp.TextAlignment.SUPERSCRIPT);
    }
    resetFollowingOperatorRunToNormal_(textElement, match.index + token.length);
  }
}

function formatSquaredStatistics_(textElement) {
  const matches = statMatches_(textElement.getText(), "[RrfsIkTχηωεσ](?:2|²|\\^2)");
  for (let i = matches.length - 1; i >= 0; i--) {
    const match = matches[i];
    normalizeNotationToken_(textElement, match.index, match.value);
    setNotationStyle_(textElement, match.index, match.index,
      !APA_GREEK_STATS_.includes(match.value[0]), DocumentApp.TextAlignment.NORMAL);
    setNotationStyle_(textElement, match.index + 1, match.index + 1,
      false, DocumentApp.TextAlignment.SUPERSCRIPT);
    resetFollowingOperatorRunToNormal_(textElement, match.index + 2);
  }
}

function normalizeSquaredGlyphs_(textElement) {
  const matches = [...textElement.getText().matchAll(/²/g)];
  for (let i = matches.length - 1; i >= 0; i--) {
    const index = matches[i].index;
    replaceRangePreservingStyle_(textElement, index, index, "2");
    setNotationStyle_(textElement, index, index, false, DocumentApp.TextAlignment.SUPERSCRIPT);
  }
}

function setNotationStyle_(textElement, start, end, italic, alignment) {
  textElement.setItalic(start, end, italic);
  textElement.setTextAlignment(start, end, alignment);
}

function resetFollowingOperatorRunToNormal_(textElement, start) {
  const match = textElement.getText().slice(start).match(/^[ \t\u00A0\u202F=<>≤≥≠≈±+−]+/);
  if (match) {
    setNotationStyle_(textElement, start, start + match[0].length - 1,
      false, DocumentApp.TextAlignment.NORMAL);
  }
}

/** t (42) -> t(42); F (2,145) -> F(2, NBSP 145), including fractional df. */
function normalizeDfSpacing_(textElement) {
  const space = APA_HORIZONTAL_SPACE_;
  const number = "(?:\\d+(?:\\.\\d+)?)";
  const source = "(?:[tF]|χ2)" + space + "*\\(" + space + "*" + number +
    "(?:" + space + "*," + space + "*" + number + ")*" + space + "*\\)";
  const matches = statMatches_(textElement.getText(), source);
  for (let i = matches.length - 1; i >= 0; i--) {
    const match = matches[i];
    const commas = [...match.value.matchAll(/,[ \t\u00A0\u202F]*(?=\d)/g)];
    for (let j = commas.length - 1; j >= 0; j--) {
      const comma = commas[j];
      replaceRangePreservingStyle_(textElement, match.index + comma.index + 1,
        match.index + comma.index + comma[0].length - 1, "\u00A0");
    }
    const gap = match.value.match(/^(?:[tF]|χ2)([ \t\u00A0\u202F]+)\(/);
    if (gap) {
      const start = match.index + (match.value.startsWith("χ") ? 2 : 1);
      textElement.deleteText(start, start + gap[1].length - 1);
    }
  }
}

/**
 * Uses explicit expression/phrase contexts rather than italicizing all letters.
 * Plural s is roman; an existing subscript (e.g., normalized r_s) is preserved.
 */
function italicizeStatSymbols_(textElement, isTableLabel) {
  const text = textElement.getText();
  const source = "(?:" + APA_LATIN_STATS_.join("|") + ")(?:s)?";
  const matches = statMatches_(text, source);
  for (const match of matches) {
    const before = text.slice(0, match.index);
    const after = text.slice(match.index + match.value.length);
    const expression = /^[ \t\u00A0\u202F]*[=<>≤≥≠≈±+−]/.test(after) ||
      /[±+−][ \t\u00A0\u202F]*$/.test(before);
    const phrase = /^(?:[ \t\u00A0\u202F]+|[-–])(?:values?|tests?|scores?|statistics?|ratios?|coefficients?)\b/.test(after);
    const namedEffect = /(?:Cohen['’]s|Hedges['’]s?|Glass['’]s?)[ \t\u00A0\u202F]+$/.test(before) &&
      /^(?:d|f|g)$/.test(match.value);
    const df = /^[tF]$/.test(match.value) && /^\([ \t\u00A0\u202F]*\d/.test(after);
    if (!expression && !phrase && !namedEffect && !df && !isTableLabel) continue;
    const base = APA_LATIN_STATS_.includes(match.value) ? match.value : match.value.slice(0, -1);
    textElement.setItalic(match.index, match.index + base.length - 1, true);
    if (base.length !== match.value.length) {
      textElement.setItalic(match.index + base.length, match.index + base.length, false);
    }
  }
}

function formatGreekStatistics_(textElement) {
  const matches = statMatches_(textElement.getText(), "[" + APA_GREEK_STATS_ + "]");
  for (const match of matches) {
    textElement.setItalic(match.index, match.index, false);
  }
}

/**
 * Preserves unchanged prefixes/suffixes and replaces only the differing middle.
 * Supports insertion (end = start - 1) without delete/reinsert of adjacent text.
 */
function replaceRangePreservingStyle_(textElement, start, end, replacement) {
  const text = textElement.getText();
  const original = text.slice(start, end + 1);
  let prefix = 0;
  while (prefix < original.length && prefix < replacement.length &&
    original[prefix] === replacement[prefix]) prefix++;
  let suffix = 0;
  while (suffix < original.length - prefix && suffix < replacement.length - prefix &&
    original[original.length - suffix - 1] === replacement[replacement.length - suffix - 1]) suffix++;
  const removeLength = original.length - prefix - suffix;
  const inserted = replacement.slice(prefix, replacement.length - suffix);
  if (!removeLength && !inserted) return;
  const position = start + prefix;
  const styleIndex = Math.min(position, text.length - 1);
  const attrs = styleIndex >= 0 ? textElement.getAttributes(styleIndex) : {};
  if (removeLength) textElement.deleteText(position, position + removeLength - 1);
  if (inserted) {
    textElement.insertText(position, inserted);
    textElement.setAttributes(position, position + inserted.length - 1, attrs);
  }
}
