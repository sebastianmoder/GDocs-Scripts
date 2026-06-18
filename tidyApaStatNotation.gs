/**
 * APA 7 statistical notation cleanup for Google Docs.
 *
 * Fresh script.
 * Does NOT repair damage from prior faulty versions.
 * Does NOT use Text.replaceText().
 * Does NOT use whole-paragraph setText().
 *
 * Intended functions:
 * 1. Add non-breaking spaces around comparison operators:
 *    =, >, <, ≥, ≤
 *
 * 2. Convert >= and <= to ≥ and ≤.
 *
 * 3. Italicize selected Latin statistical symbols before operators:
 *    p, ps, r, R, R2, t, F, z, M, SD, SE, b, B, d, df, n, N
 *
 * 4. Format:
 *    R2 / R² / R^2       -> italic R + superscript 2
 *    χ2 / χ² / χ^2       -> roman χ + superscript 2
 *    ηp2 / ηp² / ηp^2
 *    η_p2 / η_p² / η_p^2 -> roman η + subscript p + superscript 2
 *
 * 5. Convert hyphen-minus in negative numeric values to a true minus sign:
 *    -.04  -> −.04
 *    -24.3 -> −24.3
 *
 * Run on a clean copy of the manuscript.
 */


function onOpen() {
  DocumentApp.getUi()
    .createMenu("APA Cleanup")
    .addItem("Tidy statistical notation", "tidyApaStatNotation")
    .addToUi();
}


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

  normalizeDfCommaSpaces_(textElement);

  italicizeStatSymbolsBeforeOperators_(textElement);
  italicizeTestStatisticsWithDf_(textElement);
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

  const pattern = /\bR(?:2|²|\^2)\b/g;
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