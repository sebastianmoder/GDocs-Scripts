/**
 * Applies APA 7 statistical notation cleanup to the active Google Docs tab.
 *
 * SCOPE:
 * Run from a document-bound project. Processes body text, including paragraphs,
 * lists, and tables; excludes headers, footers, footnotes, and other tabs.
 * Matches are processed within individual text elements.
 *
 * SPACING AND SIGNS:
 * - Places one non-breaking space on each side of =, >, <, ≥, and ≤ throughout
 *   the processed text, and converts >= and <= to ≥ and ≤.
 * - Converts hyphen-minus before negative numbers to a true minus sign
 *   (e.g., -.04 becomes −.04), leaving numeric ranges such as 12-24 unchanged.
 * - Replaces whitespace after commas before digits in t(...) and F(...)
 *   degrees-of-freedom expressions with a single non-breaking space.
 *
 * SYMBOL FORMATTING:
 * - Italicizes p, r, R, t, F, z, M, SD, SE, b, B, d, df, k, n, and N before
 *   comparison operators, with or without surrounding parentheses. For ps,
 *   italicizes only p and makes s roman. Also italicizes t and F before "(".
 * - Formats R2 / R² / R^2 as italic R with a roman superscript 2, and resets
 *   immediately following spaces and comparison operators to roman baseline.
 * - Formats χ2 / χ² / χ^2 as roman χ with a roman superscript 2.
 * - Formats ηp2 / ηp² / ηp^2 (also with η_p) as roman η, roman subscript p,
 *   and roman superscript 2.
 * - Converts every remaining ² glyph to a roman superscript 2, regardless of
 *   the preceding symbol.
 *
 * Makes local text replacements and formatting changes, preserving surrounding
 * rich text. Replacement text inherits the first replaced character's attributes
 * before the notation-specific formatting is applied. Alerts when complete.
 */


function tidyApaStatNotation() {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    
    processElement_(body);
    
    DocumentApp.getUi().alert("APA statistical notation cleanup complete.");
}


function processElement_(element) {
    if (element.getType() === DocumentApp.ElementType.TEXT) {
        processTextElement_(element.asText());
        return;
    }
    
    if (!element.getNumChildren) return;
    
    for (let i = 0; i < element.getNumChildren(); i++) {
        processElement_(element.getChild(i));
    }
}


function processTextElement_(textElement) {
    normalizeMinusSigns_(textElement);
    normalizeComparisonOperators_(textElement);
    
    formatPartialEtaSquared_(textElement);
    formatRSquared_(textElement);
    formatChiSquared_(textElement);
    normalizeSquaredGlyphs_(textElement);
    
    normalizeDfCommaSpaces_(textElement);
    
    italicizeStatSymbolsBeforeOperators_(textElement);
    italicizeTestStatisticsWithDf_(textElement);
}


/**
 * Converts every remaining ² glyph to a regular 2 with superscript formatting.
 *
 * Runs after the statistic-specific formatters so their symbol/subscript rules
 * still apply. Covers other squared statistics, such as η², ω², ε², f², and σ²,
 * without a symbol allowlist. Applies anywhere in the processed body text.
 * Only replaces the glyph itself, preserving its other attributes and leaving
 * surrounding characters untouched.
 */
function normalizeSquaredGlyphs_(textElement) {
  const text = textElement.getText();
  const matches = [...text.matchAll(/²/g)];

  for (let i = matches.length - 1; i >= 0; i--) {
    const index = matches[i].index;
    replaceRangePreservingStyle_(textElement, index, index, "2");
    textElement.setItalic(index, index, false);
    textElement.setTextAlignment(
      index,
      index,
      DocumentApp.TextAlignment.SUPERSCRIPT
    );
  }
}


/**
* Converts hyphen-minus to true minus sign for negative numeric values.
*
* Examples:
* -.04   -> −.04
* -24.3  -> −24.3
*
* Avoids ranges such as 12-24.
*/
function normalizeMinusSigns_(textElement) {
    const text = textElement.getText();
    if (!text || !/-(?:\.\d|\d)/.test(text)) return;
    
    const pattern = /(^|[\s\u00A0\(\[\{,;:=<>≤≥])-(?=(?:\.\d|\d))/g;
    const matches = [...text.matchAll(pattern)];
    
    for (let i = matches.length - 1; i >= 0; i--) {
        const match = matches[i];
        const hyphenIndex = match.index + match[1].length;
        
        replaceRangePreservingStyle_(textElement, hyphenIndex, hyphenIndex, "−");
    }
}


/**
* Adds non-breaking spaces around comparison operators.
* Converts >= and <= to ≥ and ≤.
*
* Examples:
* p=.032   -> p = .032
* p <.001  -> p < .001
* b=-.04   -> b = −.04
*/
function normalizeComparisonOperators_(textElement) {
    const text = textElement.getText();
    if (!text || !/[=<>≤≥]/.test(text)) return;
    
    const nbsp = "\u00A0";
    const pattern = /\s*(>=|<=|=|>|<|≥|≤)\s*/g;
    const matches = [...text.matchAll(pattern)];
    
    for (let i = matches.length - 1; i >= 0; i--) {
        const match = matches[i];
        
        let operator = match[1];
        if (operator === ">=") operator = "≥";
        if (operator === "<=") operator = "≤";
        
        replaceRangePreservingStyle_(
            textElement,
            match.index,
            match.index + match[0].length - 1,
            nbsp + operator + nbsp
        );
    }
}


/**
* R2, R², R^2 -> italic R + roman superscript 2.
*
* Explicitly resets nearby operator spacing to NORMAL to prevent
* superscript bleed such as R[2 = ].04.
*/
function formatRSquared_(textElement) {
    const text = textElement.getText();
    if (!text || !/R(?:2|²|\^2)/.test(text)) return;
    
    // ² is not a regex word character, so a trailing \b misses R² before spaces.
    const pattern = /\bR(?:2|²|\^2)(?!\w)/g;
    const matches = [...text.matchAll(pattern)];
    
    for (let i = matches.length - 1; i >= 0; i--) {
        const match = matches[i];
        const start = match.index;
        const end = start + match[0].length - 1;
        
        replaceRangePreservingStyle_(textElement, start, end, "R2");
        
        // R: italic, normal baseline
        textElement.setItalic(start, start, true);
        textElement.setTextAlignment(
            start,
            start,
            DocumentApp.TextAlignment.NORMAL
        );
        
        // 2: roman, superscript
        textElement.setItalic(start + 1, start + 1, false);
        textElement.setTextAlignment(
            start + 1,
            start + 1,
            DocumentApp.TextAlignment.SUPERSCRIPT
        );
        
        // Defensive cleanup: make sure spaces/operators after R2 are not superscript.
        resetFollowingOperatorRunToNormal_(textElement, start + 2);
    }
}


/**
* Resets following operator-related characters to normal baseline.
*
* This prevents superscript bleed in cases such as:
* R2 = .04
*
* Only touches a narrow run of spaces/non-breaking spaces and comparison
* operators immediately after the statistic.
*/
function resetFollowingOperatorRunToNormal_(textElement, start) {
    const text = textElement.getText();
    const charsToReset = /^[\s\u00A0=<>≤≥]+/;
    const rest = text.slice(start);
    const match = rest.match(charsToReset);
    
    if (!match) return;
    
    const resetStart = start;
    const resetEnd = start + match[0].length - 1;
    
    textElement.setTextAlignment(
        resetStart,
        resetEnd,
        DocumentApp.TextAlignment.NORMAL
    );
    
    textElement.setItalic(resetStart, resetEnd, false);
}


/**
* χ2, χ², χ^2 -> roman χ + superscript 2.
*/
function formatChiSquared_(textElement) {
    const text = textElement.getText();
    if (!text || !/χ(?:2|²|\^2)/.test(text)) return;
    
    const pattern = /χ(?:2|²|\^2)/g;
    const matches = [...text.matchAll(pattern)];
    
    for (let i = matches.length - 1; i >= 0; i--) {
        const match = matches[i];
        const start = match.index;
        const end = start + match[0].length - 1;
        
        replaceRangePreservingStyle_(textElement, start, end, "χ2");
        
        textElement.setItalic(start, start, false);
        textElement.setItalic(start + 1, start + 1, false);
        textElement.setTextAlignment(
            start + 1,
            start + 1,
            DocumentApp.TextAlignment.SUPERSCRIPT
        );
    }
}


/**
* ηp2, ηp², ηp^2, η_p2, η_p², η_p^2
* -> roman η + subscript p + superscript 2.
*/
function formatPartialEtaSquared_(textElement) {
    const text = textElement.getText();
    if (!text || !/η_?p(?:2|²|\^2)/.test(text)) return;
    
    const pattern = /η_?p(?:2|²|\^2)/g;
    const matches = [...text.matchAll(pattern)];
    
    for (let i = matches.length - 1; i >= 0; i--) {
        const match = matches[i];
        const start = match.index;
        const end = start + match[0].length - 1;
        
        replaceRangePreservingStyle_(textElement, start, end, "ηp2");
        
        // η roman
        textElement.setItalic(start, start, false);
        textElement.setTextAlignment(
            start,
            start,
            DocumentApp.TextAlignment.NORMAL
        );
        
        // p subscript, roman
        textElement.setItalic(start + 1, start + 1, false);
        textElement.setTextAlignment(
            start + 1,
            start + 1,
            DocumentApp.TextAlignment.SUBSCRIPT
        );
        
        // 2 superscript, roman
        textElement.setItalic(start + 2, start + 2, false);
        textElement.setTextAlignment(
            start + 2,
            start + 2,
            DocumentApp.TextAlignment.SUPERSCRIPT
        );
    }
}


/**
* Converts regular spaces after commas inside t(...) and F(...)
* degrees-of-freedom expressions to non-breaking spaces.
*
* Examples:
* F(3, 437) -> F(3, 437)
* F(2, 145) -> F(2, 145)
*/
function normalizeDfCommaSpaces_(textElement) {
    const text = textElement.getText();
    if (!text || !/\b[tF]\([^)]*,\s+\d/.test(text)) return;
    
    const pattern = /\b[tF]\([^)]*\)/g;
    const matches = [...text.matchAll(pattern)];
    const nbsp = "\u00A0";
    
    for (let i = matches.length - 1; i >= 0; i--) {
        const match = matches[i];
        const segmentStart = match.index;
        const segment = match[0];
        
        const innerPattern = /,\s+(?=\d)/g;
        const innerMatches = [...segment.matchAll(innerPattern)];
        
        for (let j = innerMatches.length - 1; j >= 0; j--) {
            const innerMatch = innerMatches[j];
            
            // The comma is at innerMatch.index.
            // The replaceable spaces begin one character after the comma.
            const spacesStart = segmentStart + innerMatch.index + 1;
            const spacesEnd = segmentStart + innerMatch.index + innerMatch[0].length - 1;
            
            replaceRangePreservingStyle_(
                textElement,
                spacesStart,
                spacesEnd,
                nbsp
            );
        }
    }
}


/**
* Italicizes selected statistical symbols before comparison operators.
*
* Examples:
* p = .032       -> p italic
* ps ≥ .148      -> only p italic, s roman
* r = .26        -> r italic
* M = 4.21       -> M italic
* SD = 1.03      -> SD italic
* (k = 2)        -> k italic (parentheses are optional)
* adjusted R2 = .03 -> R and 2 italic, 2 superscript
*/
function italicizeStatSymbolsBeforeOperators_(textElement) {
    const text = textElement.getText();
    if (!text || !/[=<>≤≥]/.test(text)) return;
    
    const space = "[\\s\\u00A0]";
    const op = "(?:=|>|<|≥|≤)";
    
    // Longest patterns first.
    const patterns = [
        { regex: new RegExp("\\bSD(?=" + space + op + space + ")", "g"), italicStart: 0, italicLength: 2 },
        { regex: new RegExp("\\bSE(?=" + space + op + space + ")", "g"), italicStart: 0, italicLength: 2 },
        { regex: new RegExp("\\bdf(?=" + space + op + space + ")", "g"), italicStart: 0, italicLength: 2 },
        
        // Plural p-values: ps ≥ .148. Italicize only p.
        { regex: new RegExp("\\bps(?=" + space + op + space + ")", "g"), italicStart: 0, italicLength: 1, romanOffsets: [1] },
        
        // Single-character symbols.
        { regex: new RegExp("\\bp(?=" + space + op + space + ")", "g"), italicStart: 0, italicLength: 1 },
        { regex: new RegExp("\\br(?=" + space + op + space + ")", "g"), italicStart: 0, italicLength: 1 },
        { regex: new RegExp("\\bR(?=" + space + op + space + ")", "g"), italicStart: 0, italicLength: 1 },
        { regex: new RegExp("\\bt(?=" + space + op + space + ")", "g"), italicStart: 0, italicLength: 1 },
        { regex: new RegExp("\\bF(?=" + space + op + space + ")", "g"), italicStart: 0, italicLength: 1 },
        { regex: new RegExp("\\bz(?=" + space + op + space + ")", "g"), italicStart: 0, italicLength: 1 },
        { regex: new RegExp("\\bM(?=" + space + op + space + ")", "g"), italicStart: 0, italicLength: 1 },
        { regex: new RegExp("\\bb(?=" + space + op + space + ")", "g"), italicStart: 0, italicLength: 1 },
        { regex: new RegExp("\\bB(?=" + space + op + space + ")", "g"), italicStart: 0, italicLength: 1 },
        { regex: new RegExp("\\bd(?=" + space + op + space + ")", "g"), italicStart: 0, italicLength: 1 },
        { regex: new RegExp("\\bk(?=" + space + op + space + ")", "g"), italicStart: 0, italicLength: 1 },
        { regex: new RegExp("\\bn(?=" + space + op + space + ")", "g"), italicStart: 0, italicLength: 1 },
        { regex: new RegExp("\\bN(?=" + space + op + space + ")", "g"), italicStart: 0, italicLength: 1 }
    ];
    
    patterns.forEach(spec => {
        const currentText = textElement.getText();
        const matches = [...currentText.matchAll(spec.regex)];
        
        for (let i = matches.length - 1; i >= 0; i--) {
            const start = matches[i].index + spec.italicStart;
            const end = start + spec.italicLength - 1;
            
            textElement.setItalic(start, end, true);
            
            if (spec.romanOffsets) {
                spec.romanOffsets.forEach(offset => {
                    textElement.setItalic(matches[i].index + offset, matches[i].index + offset, false);
                });
            }
            
            if (typeof spec.superscriptOffset === "number") {
                const pos = matches[i].index + spec.superscriptOffset;
                textElement.setTextAlignment(
                    pos,
                    pos,
                    DocumentApp.TextAlignment.SUPERSCRIPT
                );
            }
        }
    });
}


/**
* Italicizes t and F before degrees of freedom.
*
* Examples:
* t(148) = 2.34
* F(2, 145) = 4.19
*
* χ remains roman.
*/
function italicizeTestStatisticsWithDf_(textElement) {
    const text = textElement.getText();
    if (!text || !/[tF]\(/.test(text)) return;
    
    const pattern = /\b(t|F)(?=\()/g;
    const matches = [...text.matchAll(pattern)];
    
    for (let i = matches.length - 1; i >= 0; i--) {
        const start = matches[i].index;
        textElement.setItalic(start, start, true);
    }
}


/**
* Replaces a text range while preserving the attributes of the first character.
*
* This is used only for small, local replacements.
*/
function replaceRangePreservingStyle_(textElement, start, end, replacement) {
    const originalText = textElement.getText();
    const textLength = originalText.length;
    
    let attrs = {};
    if (textLength > 0 && start < textLength) {
        attrs = textElement.getAttributes(start);
    }
    
    textElement.deleteText(start, end);
    textElement.insertText(start, replacement);
    
    const newEnd = start + replacement.length - 1;
    
    textElement.setAttributes(start, newEnd, attrs);
}


/**
* Utility for debugging selected text elements, if needed.
*/
function logText_(textElement) {
    Logger.log(textElement.getText());
}
