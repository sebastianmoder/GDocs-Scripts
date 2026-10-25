/**
* Highlights parenthetical passages in the currently selected Google Docs tab
* in pale yellow, then displays an alert with the number highlighted.
*
* MATCHES:
* Parentheses containing either:
* - A year from 1900 to 2099, optionally followed by a letter.
* - "n.d." or "n. d." (allows whitespace between the two parts).
*
* Matching is case-insensitive. The whole passage, including its parentheses,
* is highlighted.
*
* Examples that match:
*   (Smith, 2020)
*   (Smith et al., 2020a)
*   (see Author et al., 2020, for a review)
*   (Smith et al., n.d.)
*   (Smith et al., n. d.)
*   (Smith, 2020; Jones, n.d.)
*
* DOES NOT MATCH:
* - Parentheses without a qualifying year or no-date marker: (see Appendix A).
* - Other years, unless another qualifying date is present: (Smith, 1899).
* - No-date markers without both periods: (Smith, nd) or (Smith, n.d).
* - Citations outside parentheses: Smith (only the author name) or [Smith, 2020].
* - An entire outer passage containing nested parentheses; qualifying inner
*   parentheses can still match.
*
* SCOPE AND LIMITATIONS:
* - Processes only the selected tab's body, including ordinary table text.
* - Uses a date-based heuristic, not author or citation-style validation:
*   (the project began in 2020) and (2020) will also match.
* - Counts each matched parenthetical passage once, even with multiple sources.
* - Overwrites background colors within matches; preserves other formatting.
* - Does not remove existing highlights, including those from previous runs.
*/

function highlightParentheticalCitations() {
    const HIGHLIGHT_COLOR = '#FFF3B0';
    
    const doc = DocumentApp.getActiveDocument();
    const tab = doc && doc.getActiveTab();
    
    if (!tab) {
        throw new Error(
            'Run this script from your Google Doc via Extensions → Apps Script.'
        );
    }
    
    const textElement = tab.asDocumentTab().getBody().editAsText();
    const text = textElement.getText();
    
    // Find parentheses without nesting or paragraph breaks.
    const parenthesesRegex = /\([^()\r\n]*\)/g;
    
    // Recognize 1900–2099, optional year suffixes, and n.d. / n. d.
    const dateRegex = /\b(?:19|20)\d{2}[a-z]?\b|\bn\.\s*d\./i;
    
    let match;
    let count = 0;
    
    while ((match = parenthesesRegex.exec(text)) !== null) {
        if (!dateRegex.test(match[0])) {
            continue;
        }
        
        textElement.setBackgroundColor(
            match.index,
            match.index + match[0].length - 1,
            HIGHLIGHT_COLOR
        );
        
        count++;
    }
    
    DocumentApp.getUi().alert(
        'Highlighted ' + count + ' parenthetical passage' +
        (count === 1 ? '.' : 's.')
    );
}