/**
 * Opens a read-only word-count report for headings in the active Google Docs tab.
 *
 * SETUP:
 * Enable the advanced Google Docs API service in the Apps Script editor:
 * Services (+) > Google Docs API > Add. Run from a document-bound project.
 *
 * COUNTING:
 * - Lists H1–H6 sections in document order; each parent includes its subsections.
 * - Excludes heading text from section counts by default. Set
 *   INCLUDE_HEADING_TEXT to true to include both own and nested heading text.
 * - Reports separate totals for the active tab and all tabs, including child
 *   tabs. These totals always include headings and text before the first heading.
 * - Reads body paragraphs, lists, and tables; skips generated tables of contents.
 *   Headers, footers, and footnotes are outside the counting scope.
 * - Counts Unicode letter/number sequences, keeping internal apostrophes and
 *   hyphens within words. Counts may differ from Google Docs' native word count.
 *
 * SUGGESTIONS AND OUTPUT:
 * Reads a preview with all suggested changes accepted, without modifying the
 * document or accepting/rejecting suggestions. Displays plain text, Markdown,
 * table (TSV), CSV, and nested JSON formats with copy and download controls.
 *
 * Alerts and returns if the advanced service is unavailable.
 *
 * @throws {Error} If there is no active tab or it is missing from the preview.
 */
function showHeadingWordCounts() {
    // Set to true to include each heading's own text and nested heading text.
    const INCLUDE_HEADING_TEXT = false;
    
    const document = DocumentApp.getActiveDocument();
    const tab = document.getActiveTab();
    if (!tab) throw new Error('Run this script from its Google Docs document.');
    
    if (typeof Docs === 'undefined') {
        DocumentApp.getUi().alert(
            'Enable Google Docs API in the Apps Script editor: Services (+) > ' +
            'Google Docs API > Add. Then run the word count again.'
        );
        return;
    }
    
    // GET only: suggested deletions disappear and suggested insertions remain
    // in the returned preview. No suggestions are actually accepted or rejected.
    const preview = Docs.Documents.get(document.getId(), {
        suggestionsViewMode: 'PREVIEW_SUGGESTIONS_ACCEPTED',
        includeTabsContent: true
    });
    const previewTab = findPreviewTab_(preview.tabs || [], tab.getId());
    if (!previewTab || !previewTab.documentTab || !previewTab.documentTab.body) {
        throw new Error('The active tab was not found in the preview. Try again.');
    }
    
    const headings = [];
    const active = [];
    
    function paragraphText(paragraph) {
        // Join runs without spaces: formatting and suggestions can split a word.
        return (paragraph.elements || []).map(function (element) {
            return element.textRun ? element.textRun.content || '' : '';
        }).join('');
    }
    
    function processParagraph(paragraph) {
        const text = paragraphText(paragraph);
        const style = (paragraph.paragraphStyle || {}).namedStyleType || '';
        const match = /^HEADING_([1-6])$/.exec(style);
        const level = match ? Number(match[1]) : 0;
        
        if (level > 0) {
            // A same-level or higher-level heading closes the preceding section.
            while (active.length && active[active.length - 1].level >= level) {
                active.pop();
            }
            const heading = {
                level: level,
                title: text.replace(/\s+/g, ' ').trim() || '(untitled heading)',
                words: 0
            };
            headings.push(heading);
            active.push(heading);
            if (!INCLUDE_HEADING_TEXT) return;
        }
        
        // Credit this text to its heading and every still-open parent heading.
        const count = countWords_(text);
        active.forEach(function (heading) { heading.words += count; });
    }
    
    function walk(content, visitParagraph) {
        (content || []).forEach(function (element) {
            if (element.paragraph) {
                // List items also appear as paragraphs in the Docs API.
                visitParagraph(element.paragraph);
            } else if (element.table) {
                (element.table.tableRows || []).forEach(function (row) {
                    (row.tableCells || []).forEach(function (cell) {
                        walk(cell.content, visitParagraph);
                    });
                });
            }
            // Skip generated tables of contents and other non-body-text elements.
        });
    }
    
    // Count each paragraph once across all tabs, including nested tabs. This
    // total always includes headings and text before the first heading.
    let documentWords = 0;
    function countTabs(tabs) {
        (tabs || []).forEach(function (previewDocumentTab) {
            const body = (previewDocumentTab.documentTab || {}).body;
            if (body) {
                walk(body.content, function (paragraph) {
                    documentWords += countWords_(paragraphText(paragraph));
                });
            }
            countTabs(previewDocumentTab.childTabs);
        });
    }
    countTabs(preview.tabs);
    
    // Like the document total, include headings and introductory text in this
    // total, independently of the per-heading INCLUDE_HEADING_TEXT setting.
    let currentTabWords = 0;
    walk(previewTab.documentTab.body.content, function (paragraph) {
        currentTabWords += countWords_(paragraphText(paragraph));
        processParagraph(paragraph);
    });
    
    const lines = headings.map(function (heading) {
        return '  '.repeat(heading.level - 1) + 'H' + heading.level +
        '\t' + heading.title + '\t' + heading.words;
    });
    const report = 'Tab: ' + tab.getTitle() + '\n' +
    'Suggestions: preview with all accepted (document unchanged)\n' +
    'Heading text included: ' + (INCLUDE_HEADING_TEXT ? 'yes\n\n' : 'no\n\n') +
    'Current tab (suggestions accepted): ' + currentTabWords + ' words\n' +
    'Entire document (all tabs, suggestions accepted): ' +
    documentWords + ' words' +
    '\n\nLevel\tHeading\tWords\n' +
    (lines.length ? lines.join('\n') : 'No H1–H6 headings found.');
    
    const markdown = [
        'Tab: ' + escapeWordCountMarkdown_(tab.getTitle()),
        'Suggestions: preview with all accepted (document unchanged)',
        'Heading text included: ' + (INCLUDE_HEADING_TEXT ? 'yes' : 'no'),
        '',
        'Current tab (suggestions accepted): ' + currentTabWords + ' words',
        'Entire document (all tabs, suggestions accepted): ' + documentWords + ' words',
        '',
        '| Level | Heading | Words |',
        '| --- | --- | ---: |'
    ].concat(headings.map(function (heading) {
        return '| H' + heading.level + ' | ' +
        escapeWordCountMarkdown_(heading.title) + ' | ' + heading.words + ' |';
    }));
    if (!headings.length) markdown.push('', 'No H1–H6 headings found.');
    
    const nestedHeadings = [];
    const parentHeadings = [];
    headings.forEach(function (heading) {
        const node = {
            level: heading.level, title: heading.title, words: heading.words, children: []
        };
        while (parentHeadings.length &&
            parentHeadings[parentHeadings.length - 1].level >= node.level) {
                parentHeadings.pop();
            }
            const siblings = parentHeadings.length ?
            parentHeadings[parentHeadings.length - 1].children : nestedHeadings;
            siblings.push(node);
            parentHeadings.push(node);
        });
        
        const json = JSON.stringify({
            tab: { id: tab.getId(), title: tab.getTitle() },
            suggestionsViewMode: 'PREVIEW_SUGGESTIONS_ACCEPTED',
            headingTextIncluded: INCLUDE_HEADING_TEXT,
            currentTabWordCount: currentTabWords,
            currentTabWordCountScope: 'Current tab only; includes heading text and introductory text',
            documentWordCount: documentWords,
            documentWordCountScope: 'All tabs; includes heading text and introductory text',
            headings: nestedHeadings
        }, null, 2);
        
        const summary = report.slice(0, report.indexOf('\n\nLevel\tHeading\tWords'));
        const tableContent = summary + '\n\n' + ['Level\tHeading\tWords'].concat(
            headings.map(function (heading) {
                return ['H' + heading.level, heading.title, heading.words].join('\t');
            })
        ).join('\n');
        
        // Keep CSV rectangular and distinguish tab/document totals from headings.
        const csvRows = [
            ['Scope', 'Tab', 'Level', 'Heading', 'Words'],
            ['Tab', tab.getTitle(), '', 'Current tab', currentTabWords],
            ['Document', '', '', 'Entire document (all tabs)', documentWords]
        ].concat(headings.map(function (heading) {
            return ['Heading', tab.getTitle(), 'H' + heading.level, heading.title, heading.words];
        }));
        const csv = csvRows.map(function (row) {
            return row.map(escapeWordCountCsv_).join(',');
        }).join('\r\n');
        
        // Serialize as data, escaping characters that could terminate an HTML script.
        const formats = JSON.stringify({
            text: { content: report, extension: 'txt', mime: 'text/plain' },
            markdown: { content: markdown.join('\n'), extension: 'md', mime: 'text/markdown' },
            table: { content: tableContent, extension: 'tsv', mime: 'text/tab-separated-values',
                summary: summary, rows: headings },
                csv: { content: csv, extension: 'csv', mime: 'text/csv' },
                json: { content: json, extension: 'json', mime: 'application/json' }
            }).replace(/</g, '\\u003c')
            .replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
            
            const html = HtmlService.createHtmlOutput(`
    <!doctype html>
    <html><head><meta charset="utf-8">
    <style>
      body { box-sizing: border-box; margin: 0; padding: 16px;
        height: 100vh; display: flex; flex-direction: column; gap: 12px;
        font: 14px Arial, sans-serif; }
      p { margin: 0; line-height: 1.5; }
      textarea { box-sizing: border-box; width: 100%; flex: 1;
        min-height: 200px; resize: none;
        font: 13px monospace; white-space: pre; }
      .buttons { display: flex; gap: 8px; flex-wrap: wrap; }
      button { padding: 7px 12px; cursor: pointer; }
      button[aria-pressed="true"] { background: #e8f0fe; color: #174ea6;
        border: 2px solid #174ea6; }
      button:disabled { cursor: default; }
      #tableResults { flex: 1; min-height: 200px; overflow: auto;
        border: 1px solid #888; padding: 12px; }
      #tableSummary { white-space: pre-wrap; margin-bottom: 16px; }
      table { width: 100%; border-collapse: collapse; }
      th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
      th { background: #f5f5f5; }
      td:last-child, th:last-child { text-align: right; }
      [hidden] { display: none !important; }
      #toast { position: fixed; bottom: 76px; left: 50%;
        transform: translateX(-50%); background: #263238; color: white;
        border-radius: 6px; padding: 10px 16px; pointer-events: none;
        box-shadow: 0 2px 8px #0003; }
    </style></head><body>
      <p>Counts the preview with all suggestions accepted. Document unchanged.<br>
      Choose a format, then copy or download it. Table copies as tab-separated text.</p>
      <textarea id="results" aria-label="Heading word counts" readonly></textarea>
      <div id="tableResults" hidden>
        <p id="tableSummary"></p>
        <table aria-label="Heading word counts">
          <thead><tr><th scope="col">Level</th><th scope="col">Heading</th>
          <th scope="col">Words</th></tr></thead>
          <tbody id="tableRows"></tbody>
        </table>
        <p id="emptyTable" hidden>No H1–H6 headings found.</p>
      </div>
      <div class="buttons" role="group" aria-label="Report format">
        <button data-format="text" onclick="showFormat('text')">Plain text</button>
        <button data-format="markdown" onclick="showFormat('markdown')">Markdown</button>
        <button data-format="table" onclick="showFormat('table')">Table</button>
        <button data-format="csv" onclick="showFormat('csv')">CSV</button>
        <button data-format="json" onclick="showFormat('json')">JSON</button>
      </div>
      <div class="buttons">
        <button id="copy" onclick="copyReport()">Copy to clipboard</button>
        <button id="download" onclick="downloadReport()">Download .txt</button>
      </div>
      <div id="toast" role="status" aria-live="polite" hidden></div>
      <script>
        const formats = ${formats};
        let currentFormat = 'text';
        let toastTimer;
        let copySequence = 0;
                
        function showFormat(name) {
          currentFormat = name;
          const results = document.getElementById('results');
          results.value = formats[name].content;
          results.hidden = name === 'table';
          document.getElementById('tableResults').hidden = name !== 'table';
          document.querySelectorAll('[data-format]').forEach(function (button) {
            const selected = button.dataset.format === name;
            button.disabled = selected;
            button.setAttribute('aria-pressed', String(selected));
          });
          document.getElementById('download').textContent =
            'Download .' + formats[name].extension;
        }
                
        function showToast(message, duration) {
          const toast = document.getElementById('toast');
          clearTimeout(toastTimer);
          toast.textContent = message;
          toast.hidden = false;
          toastTimer = setTimeout(function () { toast.hidden = true; }, duration);
        }
                
        async function copyReport() {
          const sequence = ++copySequence;
          const text = formats[currentFormat].content;
          const previousFocus = document.activeElement;
          const helper = document.createElement('textarea');
          helper.value = text;
          helper.readOnly = true;
          helper.style.cssText = 'position:fixed;left:-9999px;top:0;opacity:0';
          document.body.appendChild(helper);
          helper.focus();
          helper.select();
          let copied = false;
          try {
            // Synchronous copying works in dialogs where the newer clipboard
            // API is blocked by the surrounding iframe's permissions policy.
            copied = document.execCommand('copy');
          } catch (error) {
            copied = false;
          } finally {
            helper.remove();
            if (previousFocus) previousFocus.focus({ preventScroll: true });
          }
          if (!copied && navigator.clipboard && navigator.clipboard.writeText) {
            try {
              await navigator.clipboard.writeText(text);
              copied = true;
            } catch (error) {
              copied = false;
            }
          }
          if (sequence !== copySequence) return;
          if (copied) {
            showToast('Copied to clipboard', 1000);
          } else {
            // Provide selectable text even if the browser denies clipboard access.
            document.getElementById('tableResults').hidden = true;
            const results = document.getElementById('results');
            results.hidden = false;
            results.focus();
            results.select();
            showToast('Copy blocked. Press Ctrl+C or ⌘C to copy the selected text.', 4000);
          }
        }
                
        function downloadReport() {
          const format = formats[currentFormat];
          const blob = new Blob([format.content], { type: format.mime + ';charset=utf-8' });
          const url = URL.createObjectURL(blob);
          const link = document.createElement('a');
          link.href = url;
          link.download = 'heading-word-counts.' + format.extension;
          document.body.appendChild(link);
          link.click();
          link.remove();
          setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
        }
                
        document.getElementById('tableSummary').textContent = formats.table.summary;
        const tableRows = document.getElementById('tableRows');
        formats.table.rows.forEach(function (heading) {
          const row = document.createElement('tr');
          ['H' + heading.level, heading.title, heading.words].forEach(function (value, index) {
            const cell = document.createElement('td');
            cell.textContent = value;
            if (index === 1) cell.style.paddingLeft = (8 + 12 * (heading.level - 1)) + 'px';
            row.appendChild(cell);
          });
          tableRows.appendChild(row);
        });
        document.getElementById('emptyTable').hidden = formats.table.rows.length > 0;
        showFormat('text');
      </script>
    </body></html>
  `).setWidth(720).setHeight(560);
                
                DocumentApp.getUi().showModalDialog(html, 'Words by heading');
            }
            
            function escapeWordCountMarkdown_(value) {
                return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;')
                .replace(/>/g, '&gt;').replace(/[\\`*_{}\[\]()#+.!|>-]/g, '\\$&')
                .replace(/[\r\n]+/g, ' ');
            }
            
            function escapeWordCountCsv_(value) {
                return '"' + String(value).replace(/"/g, '""') + '"';
            }
            
            function findPreviewTab_(tabs, tabId) {
                for (const tab of tabs) {
                    if (tab.tabProperties && tab.tabProperties.tabId === tabId) return tab;
                    const found = findPreviewTab_(tab.childTabs || [], tabId);
                    if (found) return found;
                }
                return null;
            }
            
            function countWords_(text) {
                // Letters/numbers form words; internal apostrophes and hyphens stay together.
                const words = text.match(/[\p{L}\p{N}][\p{L}\p{M}\p{N}]*(?:['’‐‑-][\p{L}\p{N}][\p{L}\p{M}\p{N}]*)*/gu);
                return words ? words.length : 0;
            }
