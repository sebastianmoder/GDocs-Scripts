// Setup: Apps Script editor > Services (+) > Google Docs API > Add.
// Reads a preview with suggestions accepted; never modifies the document.

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

  walk(previewTab.documentTab.body.content, processParagraph);

  const lines = headings.map(function (heading) {
    return '  '.repeat(heading.level - 1) + 'H' + heading.level +
      '\t' + heading.title + '\t' + heading.words;
  });
  const report = 'Tab: ' + tab.getTitle() + '\n' +
    'Suggestions: preview with all accepted (document unchanged)\n' +
    'Heading text included: ' + (INCLUDE_HEADING_TEXT ? 'yes\n\n' : 'no\n\n') +
    'Entire document (all tabs, suggestions accepted): ' +
    documentWords + ' words' +
    '\n\nLevel\tHeading\tWords\n' +
    (lines.length ? lines.join('\n') : 'No H1–H6 headings found.');

  const markdown = [
    'Tab: ' + escapeWordCountMarkdown_(tab.getTitle()),
    'Suggestions: preview with all accepted (document unchanged)',
    'Heading text included: ' + (INCLUDE_HEADING_TEXT ? 'yes' : 'no'),
    '',
    'Entire document (all tabs, suggestions accepted): ' + documentWords + ' words',
    '',
    '| Level | Heading | Words |',
    '| --- | --- | ---: |'
  ].concat(headings.map(function (heading) {
    return '| H' + heading.level + ' | ' +
      escapeWordCountMarkdown_(heading.title) + ' | ' + heading.words + ' |';
  }));
  if (!headings.length) markdown.push('', 'No H1–H6 headings found.');

  const json = JSON.stringify({
    tab: { id: tab.getId(), title: tab.getTitle() },
    suggestionsViewMode: 'PREVIEW_SUGGESTIONS_ACCEPTED',
    headingTextIncluded: INCLUDE_HEADING_TEXT,
    documentWordCount: documentWords,
    documentWordCountScope: 'All tabs; includes heading text and introductory text',
    headings: headings
  }, null, 2);

  // Serialize as data, escaping characters that could terminate an HTML script.
  const formats = JSON.stringify({
    text: { content: report, extension: 'txt', mime: 'text/plain' },
    markdown: { content: markdown.join('\n'), extension: 'md', mime: 'text/markdown' },
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
      button:disabled { cursor: default; opacity: 0.55; }
    </style></head><body>
      <p>Counts the preview with all suggestions accepted. Document unchanged.<br>
      Select the results, then press Ctrl+C or ⌘C.</p>
      <textarea id="results" aria-label="Heading word counts" readonly></textarea>
      <div class="buttons" role="group" aria-label="Report format">
        <button data-format="text" onclick="showFormat('text')" disabled>Plain text</button>
        <button data-format="markdown" onclick="showFormat('markdown')">Markdown table</button>
        <button data-format="json" onclick="showFormat('json')">JSON</button>
      </div>
      <div class="buttons">
        <button onclick="document.getElementById('results').select()">Select all</button>
        <button id="download" onclick="downloadReport()">Download .txt</button>
        <button onclick="google.script.host.close()">Close</button>
      </div>
      <script>
        const formats = ${formats};
        let currentFormat = 'text';

        function showFormat(name) {
          currentFormat = name;
          document.getElementById('results').value = formats[name].content;
          document.querySelectorAll('[data-format]').forEach(function (button) {
            const selected = button.dataset.format === name;
            button.disabled = selected;
            button.setAttribute('aria-pressed', String(selected));
          });
          document.getElementById('download').textContent =
            'Download .' + formats[name].extension;
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
