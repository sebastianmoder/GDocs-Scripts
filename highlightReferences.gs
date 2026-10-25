/**
* Select citations, material references, and/or page locators to highlight.
* Everything, including the dialog, lives in this one document-bound script.
* Only the active tab's body (including table text) is processed. Existing
* backgrounds within matches are overwritten; text and other attributes remain.
* Citation matches take color priority. Reruns do not clear old highlights.
*/
const REFERENCE_HIGHLIGHT_COLORS = {
    citations: '#FFF3B0',
    materials: '#B8E8FF',
    pages: '#FFD1E8'
};

function highlightReferences() {
    const colors = REFERENCE_HIGHLIGHT_COLORS;
    const html = HtmlService.createHtmlOutput(`
    <!doctype html>
    <html><head><base target="_top"><style>
        body { font: 14px Arial, sans-serif; color: #202124; padding: 12px; }
        label { display: block; margin: 16px 0; }
        small { display: block; margin: 5px 0 0 25px; color: #5f6368; }
        .swatch { display: inline-block; width: 14px; height: 14px;
        border: 1px solid #777; vertical-align: middle; margin: 0 5px; }
        p { line-height: 1.5; }
        button { padding: 8px 16px; margin-right: 8px; cursor: pointer; }
        #status { white-space: pre-line; }
    </style></head><body>
        <p>Highlight references in the active tab, including tables.</p>
        <form id="options">
            <label><input type="checkbox" id="citations" checked>
                <span class="swatch" style="background:${colors.citations}"></span>Citations
                <small>(Smith, 2020); (Jones, n.d.)</small></label>
            <label><input type="checkbox" id="materials">
                <span class="swatch" style="background:${colors.materials}"></span>Materials
                <small>Table 1a; Figure S3a; Appendices A–C; Supplementary Material</small></label>
            <label><input type="checkbox" id="pages">
                <span class="swatch" style="background:${colors.pages}"></span>Page numbers
                <small>p.12; [p. 12]; p. [12]; pp. iv–vi</small></label>
            <p>Only backgrounds inside matches change. Existing highlights elsewhere
                remain. Citation color wins when selected categories overlap.</p>
            <button id="run" type="submit">Highlight</button>
            <button type="button" onclick="google.script.host.close()">Close</button>
        </form>
        <p id="status" role="status" aria-live="polite"></p>
        <script>
            document.getElementById('options').addEventListener('submit', function(event) {
                event.preventDefault();
                var options = {};
                ['citations', 'materials', 'pages'].forEach(function(key) {
            options[key] = document.getElementById(key).checked;
            });
            var status = document.getElementById('status');
            var button = document.getElementById('run');
            if (!options.citations && !options.materials && !options.pages) {
                status.textContent = 'Select at least one category.';
                return;
            }
            button.disabled = true;
            status.textContent = 'Highlighting…';
            google.script.run.withSuccessHandler(function(result) {
                button.disabled = false;
                var labels = {citations: 'Citations', materials: 'Materials', pages: 'Page numbers'};
                var lines = ['Matches found:'];
                Object.keys(labels).forEach(function(key) {
                    if (options[key]) lines.push(labels[key] + ': ' + result.counts[key]);
                });
                lines.push('Overlapping matches can share highlighted text.');
                status.textContent = lines.join('\\n');
                }).withFailureHandler(function(error) {
                    button.disabled = false;
                    status.textContent = 'Highlighting failed: ' + error.message;
                }).highlightSelectedReferences(options);
            });
        </script>
    </body></html>
    `).setWidth(480).setHeight(510);
        DocumentApp.getUi().showModalDialog(html, 'Highlight references');
    }
    
    // Backwards-compatible entry point for existing menus and editor runs.
    function highlightParentheticalCitations() {
        const result = highlightSelectedReferences({ citations: true });
        const count = result.counts.citations;
        DocumentApp.getUi().alert(
            'Highlighted ' + count + ' parenthetical passage' + (count === 1 ? '.' : 's.')
        );
    }
    
    function highlightSelectedReferences(options) {
        const categories = Object.keys(REFERENCE_HIGHLIGHT_COLORS).filter(function(key) {
            return options && options[key] === true;
        });
        if (!categories.length) {
            throw new Error('Select at least one category.');
        }
        const doc = DocumentApp.getActiveDocument();
        const tab = doc && doc.getActiveTab();
        if (!tab) {
            throw new Error('Run this script from your Google Doc via Extensions → Apps Script.');
        }
        const textElement = tab.asDocumentTab().getBody().editAsText();
        const result = findReferenceHighlights_(textElement.getText(), categories);
        result.ranges.forEach(function(range) {
            // Match offsets are half-open; DocumentApp's end offset is inclusive.
            textElement.setBackgroundColor(range.start, range.end - 1,
                REFERENCE_HIGHLIGHT_COLORS[range.category]);
            });
            return { counts: result.counts };
        }
        
        function findReferenceHighlights_(text, categories) {
            const finders = {
                citations: findCitationReferences_,
                materials: findMaterialReferences_,
                pages: findPageReferences_
            };
            const counts = { citations: 0, materials: 0, pages: 0 };
            let matches = [];
            categories.forEach(function(category) {
                const found = finders[category](text);
                counts[category] = found.length;
                matches = matches.concat(found.map(function(range) {
                    return { start: range.start, end: range.end, category: category };
                }));
            });
            return { counts: counts, ranges: resolveReferenceOverlaps_(matches) };
        }
        
        function findCitationReferences_(text) {
            const parentheses = /\([^()\r\n]*\)/g;
            const date = /\b(?:19|20)\d{2}[a-z]?\b|\bn\.\s*d\./i;
            const ranges = [];
            let match;
            while ((match = parentheses.exec(text)) !== null) {
                if (date.test(match[0])) {
                    ranges.push({ start: match.index, end: parentheses.lastIndex });
                }
            }
            return ranges;
        }
        
        function findMaterialReferences_(text) {
            const label = /(?:supplement(?:al|ary)[ \t\u00a0\u202f]+materials?|appendices|appendix|tables?|figures?|tabs?\.|figs?\.|app\.)/gi;
            const ranges = [];
            let match;
            while ((match = label.exec(text)) !== null) {
                if (isReferenceWordChar_(text[match.index - 1]) ||
                isReferenceWordChar_(text[label.lastIndex])) {
                    continue;
                }
                const idStart = skipReferenceSpaces_(text, label.lastIndex);
                const idEnd = idStart > label.lastIndex ? readMaterialId_(text, idStart) : null;
                const isSupplement = /^supplement/i.test(match[0]);
                if (idEnd === null && !isSupplement) continue;
                const end = idEnd === null ? label.lastIndex :
                extendReferenceSequence_(text, idEnd, readMaterialId_, label.source);
                ranges.push({ start: match.index, end: end });
                label.lastIndex = end;
            }
            return ranges;
        }
        
        function findPageReferences_(text) {
            const label = /pp?\./gi;
            const ranges = [];
            let match;
            while ((match = label.exec(text)) !== null) {
                if (isReferenceWordChar_(text[match.index - 1])) continue;
                let start = skipReferenceSpaces_(text, label.lastIndex);
                const opener = text[start];
                const closer = opener === '[' ? ']' : opener === '(' ? ')' : null;
                if (closer) start = skipReferenceSpaces_(text, start + 1);
                const idEnd = readPageId_(text, start);
                if (idEnd === null) continue;
                let end = extendReferenceSequence_(text, idEnd, readPageId_, 'pp?\\.');
                if (closer) {
                    const closingIndex = skipReferenceSpaces_(text, end);
                    if (text[closingIndex] !== closer) continue;
                    end = closingIndex + 1;
                }
                // Include balanced wrappers only when their contents are just the locator.
                const range = expandPageWrapper_(text, match.index, end);
                ranges.push(range);
                label.lastIndex = range.end;
            }
            return ranges;
        }
        
        function readMaterialId_(text, start) {
            // Arabic/S-prefixed numbers, letter suffixes, A1 identifiers, and panels.
            const match = /^(?:s?\d+[a-z]?|[a-z](?:\d+[a-z]?)?)/i.exec(text.slice(start));
            if (!match) return null;
            let end = start + match[0].length;
            if (!isReferenceIdEnd_(text, end)) return null;
            const panelStart = skipReferenceSpaces_(text, end);
            const panel = /^\([a-z](?:[ \t\u00a0\u202f]*[-–—][ \t\u00a0\u202f]*[a-z])?\)/i.exec(text.slice(panelStart));
            if (panel && !isReferenceWordChar_(text[panelStart + panel[0].length])) {
                end = panelStart + panel[0].length;
            }
            return end;
        }
        
        function readPageId_(text, start) {
            const match = /^(?:\d+|[ivxlcdm]+)/i.exec(text.slice(start));
            if (!match) return null;
            const end = start + match[0].length;
            if (!isReferenceIdEnd_(text, end)) return null;
            // Validate Roman numerals rather than accepting any sequence of Roman letters.
            if (!/^\d+$/.test(match[0]) &&
            !/^m{0,3}(?:cm|cd|d?c{0,3})(?:xc|xl|l?x{0,3})(?:ix|iv|v?i{0,3})$/i.test(match[0])) {
                return null;
            }
            return end;
        }
        
        function extendReferenceSequence_(text, end, readId, labelSource) {
            // Require a valid next identifier before including a separator or label.
            const connector = /^[ \t\u00a0\u202f]*(?:[-–—]|,(?:[ \t\u00a0\u202f]*(?:and|or)\b)?|\b(?:and|or)\b|&)[ \t\u00a0\u202f]*/i;
            const repeatedLabel = new RegExp('^(?:' + labelSource + ')', 'i');
            while (true) {
                const join = connector.exec(text.slice(end));
                if (!join) return end;
                let start = end + join[0].length;
                // A following page label is not a single-letter material identifier.
                if (readId === readMaterialId_ && /^pp?\./i.test(text.slice(start))) return end;
                const label = repeatedLabel.exec(text.slice(start));
                if (label && !isReferenceWordChar_(text[start + label[0].length])) {
                    start = skipReferenceSpaces_(text, start + label[0].length);
                }
                const nextEnd = readId(text, start);
                if (nextEnd === null) return end;
                end = nextEnd;
            }
        }
        
        function expandPageWrapper_(text, start, end) {
            while (true) {
                let left = start;
                while (left > 0 && /[ \t\u00a0\u202f]/.test(text[left - 1])) left--;
                const right = skipReferenceSpaces_(text, end);
                if ((text[left - 1] === '(' && text[right] === ')') ||
                (text[left - 1] === '[' && text[right] === ']')) {
                    start = left - 1;
                    end = right + 1;
                } else {
                    return { start: start, end: end };
                }
            }
        }
        
        function skipReferenceSpaces_(text, start) {
            while (start < text.length && /[ \t\u00a0\u202f]/.test(text[start])) start++;
            return start;
        }
        
        function isReferenceWordChar_(character) {
            return !!character && /[\p{L}\p{N}_]/u.test(character);
        }
        
        function isReferenceIdEnd_(text, end) {
            // Do not highlight numeric prefixes of words or decimals.
            return !isReferenceWordChar_(text[end]) &&
            !(text[end] === '.' && /\d/.test(text[end + 1] || ''));
        }
        
        function resolveReferenceOverlaps_(matches) {
            // Sweep boundaries to emit disjoint ranges; citations win deterministically.
            const priority = { pages: 1, materials: 2, citations: 3 };
            const events = [];
            matches.forEach(function(match) {
                events.push({ offset: match.start, category: match.category, delta: 1 });
                events.push({ offset: match.end, category: match.category, delta: -1 });
            });
            events.sort(function(a, b) { return a.offset - b.offset; });
            const active = { citations: 0, materials: 0, pages: 0 };
            const ranges = [];
            let index = 0;
            while (index < events.length) {
                const start = events[index].offset;
                while (index < events.length && events[index].offset === start) {
                    active[events[index].category] += events[index].delta;
                    index++;
                }
                if (index === events.length) break;
                const end = events[index].offset;
                const category = Object.keys(active).filter(function(key) {
                    return active[key] > 0;
                }).sort(function(a, b) { return priority[b] - priority[a]; })[0];
                if (!category) continue;
                const previous = ranges[ranges.length - 1];
                if (previous && previous.end === start && previous.category === category) {
                    previous.end = end;
                } else {
                    ranges.push({ start: start, end: end, category: category });
                }
            }
            return ranges;
        }
