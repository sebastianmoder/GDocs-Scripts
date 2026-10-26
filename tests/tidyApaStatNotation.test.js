// Run from the repository root: node tests/tidyApaStatNotation.test.js
// Exercises rich-text edits locally; Google Docs integration still needs a duplicate Doc.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const context = {
  DocumentApp: {
    TextAlignment: {NORMAL: "normal", SUPERSCRIPT: "super", SUBSCRIPT: "sub"},
    ElementType: {TEXT: "text", TABLE_CELL: "cell", PARAGRAPH: "paragraph", BODY_SECTION: "body"}
  }
};
vm.createContext(context);
vm.runInContext(fs.readFileSync(path.join(__dirname, "..", "tidyApaStatNotation.gs"), "utf8"), context);

class MockText {
  constructor(value) {
    this.value = value;
    // Document offsets, like JavaScript regex indices, count UTF-16 code units.
    this.attrs = Array.from({length: value.length}, (_, index) => ({
      italic: false, alignment: "normal", bold: index % 2 === 0,
      color: `color-${index}`, link: null, marker: `character-${index}`
    }));
  }
  getType() { return "text"; }
  asText() { return this; }
  getText() { return this.value; }
  getAttributes(index) { this.range(index, index); return {...this.attrs[index]}; }
  getLinkUrl(index) { this.range(index, index); return this.attrs[index].link; }
  getTextAttributeIndices() {
    return this.attrs.reduce((indices, attrs, index) => {
      if (!index || JSON.stringify(attrs) !== JSON.stringify(this.attrs[index - 1])) indices.push(index);
      return indices;
    }, []);
  }
  range(start, end) {
    assert(start >= 0 && end >= start && end < this.value.length,
      `Invalid range ${start}:${end} for ${JSON.stringify(this.value)}`);
  }
  deleteText(start, end) {
    this.range(start, end);
    this.value = this.value.slice(0, start) + this.value.slice(end + 1);
    this.attrs.splice(start, end - start + 1);
    return this;
  }
  insertText(index, value) {
    assert(index >= 0 && index <= this.value.length);
    this.value = this.value.slice(0, index) + value + this.value.slice(index);
    this.attrs.splice(index, 0, ...Array.from({length: value.length}, () => ({})));
    return this;
  }
  setAttributes(start, end, attrs) {
    this.range(start, end);
    for (let index = start; index <= end; index++) this.attrs[index] = {...attrs};
    return this;
  }
  setItalic(start, end, value) { return this.setProperty(start, end, "italic", value); }
  setTextAlignment(start, end, value) { return this.setProperty(start, end, "alignment", value); }
  setProperty(start, end, key, value) {
    this.range(start, end);
    for (let index = start; index <= end; index++) this.attrs[index][key] = value;
    return this;
  }
}

let cases = 0;
function check(input, expected, verify = () => {}, isTableLabel = false, prepare = () => {}) {
  const text = new MockText(input);
  prepare(text);
  context.processTextElement_(text, isTableLabel);
  assert.equal(text.value, expected, input);
  verify(text);
  const attrs = JSON.stringify(text.attrs);
  context.processTextElement_(text, isTableLabel);
  assert.equal(text.value, expected, `Rerun text: ${input}`);
  assert.equal(JSON.stringify(text.attrs), attrs, `Rerun attributes: ${input}`);
  cases++;
  return text;
}
const nb = "\u00A0";
function italicToken(text, token, value = true) {
  const start = text.value.indexOf(token);
  assert(start >= 0, token);
  for (let i = start; i < start + token.length; i++) assert.equal(text.attrs[i].italic, value, token);
}
function aligned(text, token, offset, value) {
  assert.equal(text.attrs[text.value.indexOf(token) + offset].alignment, value, token);
}

for (const symbol of ["IQR", "Mdn", "s", "MS", "MSE", "SSE", "SS", "U", "W", "H", "Q", "g", "f", "OR", "p", "df", "SD", "SE", "b", "N"]) {
  check(`${symbol}=.50`, `${symbol}${nb}=${nb}.50`, text => italicToken(text, symbol));
}
for (const symbol of ["IQRs", "Ms", "SDs", "SEs", "rs", "ts", "Fs", "ps", "ns", "Ns"]) {
  check(`${symbol}≥.10`, `${symbol}${nb}≥${nb}.10`, text => {
    italicToken(text, symbol.slice(0, -1));
    assert.equal(text.attrs[symbol.length - 1].italic, false);
  });
}
for (const phrase of ["Cohen’s d", "Cohen's f", "Hedges’ g", "Hedges's g", "p values", "t-test", "z scores", "F tests"]) {
  check(phrase, phrase, text => {
    const match = phrase.match(/(?:^| )([dfgptzF])(?:[- ]|$)/);
    assert.equal(text.attrs[match.index + match[0].indexOf(match[1])].italic, true, phrase);
  });
}
for (const symbol of ["R", "r", "f", "s", "k", "I", "T", "χ", "η", "ω", "ε", "σ"]) {
  for (const exponent of ["2", "²", "^2"]) {
    check(`${symbol}${exponent}=.12`, `${symbol}2${nb}=${nb}.12`, text => {
      aligned(text, `${symbol}2`, 0, "normal");
      aligned(text, `${symbol}2`, 1, "super");
      assert.equal(text.attrs[0].italic, !"χηωεσ".includes(symbol));
      assert.equal(text.attrs[1].italic, false);
      assert.equal(text.attrs[2].alignment, "normal");
    });
  }
}
for (const token of ["ηp2", "η_p²", "η_p^2", "ηg2", "η_g²", "η_g^2"]) {
  const output = token.includes("g") ? "ηg2" : "ηp2";
  check(`${token}=.12`, `${output}${nb}=${nb}.12`, text => {
    assert.deepEqual(text.attrs.slice(0, 3).map(a => a.alignment), ["normal", "sub", "super"]);
    text.attrs.slice(0, 3).forEach(a => assert.equal(a.italic, false));
  });
}
for (const token of ["r_s", "rₛ", "r_pb", "r_tb", "BF10", "BF01", "BF_10", "BF_01", "BF₁₀", "BF₀₁"]) {
  const output = token.replace(/_/g, "").replace(/ₛ/g, "s").replace(/₁/g, "1").replace(/₀/g, "0");
  check(`${token}=.3`, `${output}${nb}=${nb}.3`, text => {
    const base = token.startsWith("BF") ? "BF" : "r";
    italicToken(text, base);
    for (let i = base.length; i < output.length; i++) {
      assert.equal(text.attrs[i].alignment, "sub");
      assert.equal(text.attrs[i].italic, false);
    }
  });
}
check("t (42)=2.10; F (2,145)=4.2; F(1.5,  42.7)=3.1; χ^2 (1,100)=5.1",
  `t(42)${nb}=${nb}2.10; F(2,${nb}145)${nb}=${nb}4.2; F(1.5,${nb}42.7)${nb}=${nb}3.1; χ2(1,${nb}100)${nb}=${nb}5.1`,
  text => { italicToken(text, "t"); italicToken(text, "F"); });
check("p!=.05; b≈-.04; M±SD", `p${nb}≠${nb}.05; b${nb}≈${nb}−.04; M${nb}±${nb}SD`, text => {
  for (const token of ["p", "b", "M", "SD"]) italicToken(text, token);
});
check("M+SD; n−1; 2+3; b=−.04; 12-24; p>=.05; p<=.10",
  `M${nb}+${nb}SD; n${nb}−${nb}1; 2${nb}+${nb}3; b${nb}=${nb}−.04; 12-24; p${nb}≥${nb}.05; p${nb}≤${nb}.10`);
check("M-SD; n-1; 12-24; t-test; 1−-2; p≈−.4; 1+2+3",
  `M${nb}−${nb}SD; n${nb}−${nb}1; 12-24; t-test; 1${nb}−${nb}−2; p${nb}≈${nb}−.4; 1${nb}+${nb}2${nb}+${nb}3`);
check("[-.04, –.02]; p<-.01; x=+2", `[−.04, −.02]; p${nb}<${nb}−.01; x${nb}=${nb}+2`);
check("t (test); F (condition); after2; rifle2; χ20; ηp20; BF100; résumé; 12-24",
  "t (test); F (condition); after2; rifle2; χ20; ηp20; BF100; résumé; 12-24", text => {
    assert(text.attrs.every(a => !a.italic && a.alignment === "normal"));
  });
check("a d in prose; As expected; gas; SDscale; état; alphabet; αλφα",
  "a d in prose; As expected; gas; SDscale; état; alphabet; αλφα", text => {
    assert(text.attrs.every(a => !a.italic));
  });
check("α=.8; β=.3; κ=.7; ρ=.4", `α${nb}=${nb}.8; β${nb}=${nb}.3; κ${nb}=${nb}.7; ρ${nb}=${nb}.4`,
  text => { for (const token of ["α", "β", "κ", "ρ"]) italicToken(text, token, false); }, false,
  text => text.setItalic(0, text.value.length - 1, true));
check("cm²", "cm2", text => aligned(text, "cm2", 2, "super"));
check("M", "M", text => italicToken(text, "M", false));
check(" SD ", " SD ", text => italicToken(text, "SD"), true);
check("interquartile range (IQR); IQR describes spread; preIQR; IQRscale",
  "interquartile range (IQR); IQR describes spread; preIQR; IQRscale", text => {
    assert(text.attrs.every(a => !a.italic));
  });
check("IQR-1", `IQR${nb}−${nb}1`, text => italicToken(text, "IQR"));
check("\np=.05\np=.10\n", `\np${nb}=${nb}.05\np${nb}=${nb}.10\n`);
check("📊 t (42)=2.1; résumé; η_g^2=.1", `📊 t(42)${nb}=${nb}2.1; résumé; ηg2${nb}=${nb}.1`,
  text => { italicToken(text, "t"); aligned(text, "ηg2", 1, "sub"); });
check("", "");

// Unique attributes follow surviving characters across multiple local edits.
const input = "prefix η_p^2 = .12; r_s; BF_10; t (42); F(2,145) suffix";
const styled = check(input, `prefix ηp2${nb}=${nb}.12; rs; BF10; t(42); F(2,${nb}145) suffix`);
let cursor = 0;
for (let i = 0; i < input.length; i++) {
  if (input[i] === "_" || input[i] === "^" || input[i] === " ") continue;
  const at = styled.attrs.findIndex((a, j) => j >= cursor && a.marker === `character-${i}`);
  assert(at >= cursor, `Original style lost for ${input[i]} at ${i}`);
  assert.equal(styled.attrs[at].bold, i % 2 === 0);
  assert.equal(styled.attrs[at].color, `color-${i}`);
  cursor = at + 1;
}

// URLs must retain every character and attribute through all cleanup passes.
function checkProtected(input, expected, spans, prepare = () => {}, isTableLabel = false) {
  const snapshots = [];
  return check(input, expected, text => {
    for (const snapshot of snapshots) {
      const start = text.value.indexOf(snapshot.value);
      assert(start >= 0, `Protected text lost: ${snapshot.value}`);
      assert.deepEqual(text.attrs.slice(start, start + snapshot.value.length), snapshot.attrs,
        `Protected formatting changed: ${snapshot.value}`);
    }
  }, isTableLabel, text => {
    prepare(text);
    for (const value of spans) {
      const start = input.indexOf(value);
      assert(start >= 0, value);
      // Deliberately use non-default italics/alignment to catch formatting changes.
      text.setItalic(start, start + value.length - 1, true);
      text.setTextAlignment(start, start + value.length - 1, "sub");
      snapshots.push({value, attrs: text.attrs.slice(start, start + value.length).map(a => ({...a}))});
    }
  });
}

for (const url of [
  "https://example.org/search?p=.05&n=-1&r^2=.2#η_p²",
  "HTTP://example.org/r_s/BF_10/χ²/α?q=2+3&x!=.1",
  "https://example.org/p>=.05&x<=.1&b≈-.04&n-1&cm²",
  "www.example.org/r_pb?F(2,145)=3&x=±2",
  "doi.org/10.1234/r^2?p=.01",
  "example.org:8080/η_g^2?BF₁₀=2",
  "127.0.0.1:8080/r^2?p=.05",
  "localhost:8080/r^2?p=.05",
  "https://例子.测试/σ²?p=.05",
  "ftp://example.org/-2?p=.05",
  "mailto:stats@example.org?subject=p=.05",
  "doi:10.1234/r^2",
  "data:text/plain,p=.05"
]) {
  checkProtected(url, url, [url]);
  checkProtected(`📊 η_p^2=.1; ${url} t (42)=2.1; cm²`,
    `📊 ηp2${nb}=${nb}.1; ${url} t(42)${nb}=${nb}2.1; cm2`, [url]);
}

const firstUrl = "https://example.org/?p=.05";
const secondUrl = "www.example.com/r^2?x=-2";
checkProtected(`p=.1; ${firstUrl} F (2,145)=3; ${secondUrl} n-1`,
  `p${nb}=${nb}.1; ${firstUrl} F(2,${nb}145)${nb}=${nb}3; ${secondUrl} n${nb}−${nb}1`,
  [firstUrl, secondUrl]);
checkProtected(`<${firstUrl}>; p=.1`, `<${firstUrl}>; p${nb}=${nb}.1`, [firstUrl]);
checkProtected(`"${firstUrl}"; p=.1`, `"${firstUrl}"; p${nb}=${nb}.1`, [firstUrl]);

function linkSpan(text, value) {
  const start = text.value.indexOf(value);
  for (let i = start; i < start + value.length; i++) text.attrs[i].link = firstUrl;
}
const linkLabel = "η_p^2=.1; t (42); cm²; α; p values";
checkProtected(`r^2=.2; ${linkLabel} F (2,145)=3`,
  `r2${nb}=${nb}.2; ${linkLabel} F(2,${nb}145)${nb}=${nb}3`, [linkLabel],
  text => linkSpan(text, linkLabel));
// Adjacent protected runs merge even when they have different link/style attributes.
checkProtected("αr^2p=.1", `αr^2p${nb}=${nb}.1`, ["αr^2"], text => {
  linkSpan(text, "α");
  linkSpan(text, "r^2");
  for (let i = 1; i < 4; i++) text.attrs[i].link = secondUrl;
});
checkProtected("SD", "SD", ["SD"], text => linkSpan(text, "SD"), true);
checkProtected(`p=.1 ${firstUrl} F (1,2)=3`,
  `p${nb}=${nb}.1 ${firstUrl} F(1,${nb}2)${nb}=${nb}3`, [firstUrl], text => {
    // A single link run overlaps only part of a visible URL.
    linkSpan(text, "example.org");
  });

// Traversal only treats complete table-cell labels as standalone statistics.
function container(type, children) {
  return {
    getType: () => type, getNumChildren: () => children.length,
    getChild: index => children[index], getText: () => children.map(c => c.getText()).join("")
  };
}
const heading = new MockText("SD");
const iqrHeading = new MockText("IQR");
const prose = new MockText("SD is an abbreviation");
const bodyLetter = new MockText("M");
const tree = container("body", [container("cell", [heading]), container("cell", [iqrHeading]),
  container("cell", [prose]), bodyLetter]);
context.processElement_(tree);
italicToken(heading, "SD");
italicToken(iqrHeading, "IQR");
italicToken(prose, "SD", false);
italicToken(bodyLetter, "M", false);
console.log(`Passed ${cases} notation cases, rerun checks, rich-text preservation, and table traversal.`);
