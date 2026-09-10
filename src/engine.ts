import JSZip from 'jszip';
import { diffArrays, diffChars } from 'diff';
import { normalizeBlock, formatIssue, effectiveFormat } from './formatting';

export type Side = 'left' | 'right';
export type Choice = Side | 'both' | 'omit';
export type FormatChoice = 'default' | 'left' | 'right' | 'smart';
export type Kind = 'equal' | 'format' | 'modified' | 'added' | 'removed' | 'complex';
export interface Block {
  index: number;
  element: Element;
  text: string;
  type: string;
  safe: boolean;
  fingerprint: string;
}
export interface WordFile {
  name: string;
  bytes: Uint8Array;
  zip: JSZip;
  xml: Document;
  blocks: Block[];
  warnings: string[];
  locked: boolean;
  styles?: Document;
  theme?: Document;
}
export interface Row {
  id: string;
  left?: Block;
  right?: Block;
  kind: Kind;
}
export interface Plan {
  base: Side;
  choices: Record<string, Choice>;
  formatMode?: 'preserve' | 'smart';
  formats?: Record<string, FormatChoice>;
}
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML = 'http://www.w3.org/XML/1998/namespace';
const MAIN = 'word/document.xml';
const MAX_FILE = 25 * 1024 * 1024;
const MAX_EXPANDED = 100 * 1024 * 1024;
const serial = (el: Node) => new XMLSerializer().serializeToString(el);
const children = (el: Element, name: string) =>
  Array.from(el.children).filter((e) => e.namespaceURI === W && e.localName === name);
const descendant = (el: Element, name: string) => Array.from(el.getElementsByTagNameNS(W, name));
const first = (el: Element, name: string) => children(el, name)[0];

function parseXML(value: string): Document {
  if (/<!DOCTYPE|<!ENTITY/i.test(value)) throw new Error('文档包含不支持的 XML 实体声明。');
  const xml = new DOMParser().parseFromString(value, 'application/xml');
  if (xml.getElementsByTagName('parsererror').length) throw new Error('Word XML 损坏，无法读取。');
  return xml;
}

export function textOf(el: Element): string {
  let result = '';
  for (const node of Array.from(el.children)) {
    if (node.namespaceURI === W && node.localName === 't') result += node.textContent ?? '';
    else if (node.namespaceURI === W && node.localName === 'tab') result += '\t';
    else if (node.namespaceURI === W && ['br', 'cr'].includes(node.localName)) result += '\n';
    else {
      result += textOf(node);
      if (node.namespaceURI === W && node.localName === 'p' && el.localName !== 'p') result += '\n';
      if (node.namespaceURI === W && node.localName === 'tc') result += '\t';
    }
  }
  return result;
}

// Whitelist ordinary WordprocessingML; unknown content is retained in the base,
// but never flattened or copied across packages without relationship handling.
function simpleParagraph(el: Element): boolean {
  if (el.namespaceURI !== W || el.localName !== 'p') return false;
  for (const node of Array.from(el.getElementsByTagName('*'))) {
    if (node.localName === 'sectPr' || node.localName.endsWith('Change')) return false;
    if (
      Array.from(node.attributes).some(
        (a) =>
          a.namespaceURI === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      )
    )
      return false;
  }
  return Array.from(el.children).every((node) => {
    if (node.namespaceURI !== W) return false;
    if (['pPr', 'proofErr'].includes(node.localName)) return true;
    if (node.localName !== 'r') return false;
    return Array.from(node.children).every(
      (n) =>
        n.namespaceURI === W &&
        ['rPr', 't', 'tab', 'cr', 'lastRenderedPageBreak', 'br'].includes(n.localName) &&
        (n.localName !== 'br' ||
          !n.getAttributeNS(W, 'type') ||
          n.getAttributeNS(W, 'type') === 'textWrapping'),
    );
  });
}

function simpleBlock(el: Element): boolean {
  if (el.localName === 'p') return simpleParagraph(el);
  if (el.namespaceURI !== W || el.localName !== 'tbl') return false;
  if (
    !Array.from(el.children).every(
      (n) => n.namespaceURI === W && ['tblPr', 'tblGrid', 'tr'].includes(n.localName),
    )
  )
    return false;
  return (
    children(el, 'tr').every(
      (tr) =>
        Array.from(tr.children).every(
          (n) => n.namespaceURI === W && ['trPr', 'tc'].includes(n.localName),
        ) &&
        children(tr, 'tc').every((tc) =>
          Array.from(tc.children).every(
            (n) => n.namespaceURI === W && (n.localName === 'tcPr' || simpleParagraph(n)),
          ),
        ),
    ) && !Array.from(el.getElementsByTagName('*')).some((n) => n.localName.endsWith('Change'))
  );
}

function fingerprint(el: Element): string {
  const copy = el.cloneNode(true) as Element;
  for (const node of [copy, ...Array.from(copy.getElementsByTagName('*'))]) {
    for (const attr of Array.from(node.attributes)) {
      if (/^(rsid|paraId|textId)/.test(attr.localName)) node.removeAttributeNode(attr);
    }
    if (['proofErr', 'lastRenderedPageBreak'].includes(node.localName)) node.remove();
  }
  return serial(copy);
}

export async function readWord(name: string, data: ArrayBuffer | Uint8Array): Promise<WordFile> {
  if (!/\.docx$/i.test(name))
    throw new Error('请使用 .docx 文件；旧版 .doc 请先在 Word 中另存为 .docx。');
  const bytes = new Uint8Array(data);
  if (bytes.byteLength > MAX_FILE) throw new Error('单个文件不能超过 25 MB。');
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    throw new Error('无法打开 DOCX：文件可能已损坏、加密，或不是真正的 .docx。');
  }
  const entries = Object.values(zip.files);
  if (entries.length > 5000) throw new Error('文档附件数量超过支持范围。');
  // JSZip exposes the central-directory uncompressed sizes before inflation.
  let expanded = 0;
  for (const entry of entries) {
    const size =
      (entry as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? 0;
    expanded += size;
    if (size > MAX_EXPANDED || expanded > MAX_EXPANDED)
      throw new Error('文档解压后超过 100 MB，请拆分后再试。');
  }
  if (entries.some((e) => /vbaProject|_xmlsignatures/i.test(e.name)))
    throw new Error('不支持带宏或数字签名的文档，请使用普通 .docx 副本。');
  const main = zip.file(MAIN);
  if (!main || !zip.file('[Content_Types].xml')) throw new Error('缺少必要的 Word 文件结构。');
  const xml = parseXML(await main.async('string'));
  const body = xml.getElementsByTagNameNS(W, 'body')[0];
  if (!body) throw new Error('目前仅支持常见的 Transitional DOCX，请在 Word 中另存为普通 .docx。');
  const elements = Array.from(body.children).filter(
    (el) => !(el.namespaceURI === W && el.localName === 'sectPr'),
  );
  if (elements.length > 2500) throw new Error('第一版最多支持 2500 个正文块，请拆分文档。');
  const blocks = elements.map(
    (element, index): Block => ({
      index,
      element,
      text: textOf(element),
      type: element.localName,
      safe: simpleBlock(element),
      fingerprint: fingerprint(element),
    }),
  );
  if (
    blocks.some((b) => b.text.length > 20000) ||
    blocks.reduce((n, b) => n + b.text.length, 0) > 500000
  )
    throw new Error('正文超出本版本的对比规模限制。');
  const warnings: string[] = [];
  const locked = [
    'fldChar',
    'ins',
    'del',
    'moveFrom',
    'moveTo',
    'commentRangeStart',
    'bookmarkStart',
    'permStart',
  ].some((tag) => xml.getElementsByTagNameNS(W, tag).length > 0);
  if (locked)
    warnings.push(
      '含修订、跨段域、书签或批注范围：本版仅允许原样导出此文档，请先在 Word 的副本中处理这些结构。',
    );
  const complex = blocks.filter((b) => !b.safe).length;
  if (complex)
    warnings.push(
      `${complex} 个正文块含图片、公式、链接、分节或其他复杂结构；可以保留底稿中的原块，暂不支持跨侧合入这些块。`,
    );
  warnings.push(
    '页眉页脚、页面设置、样式定义和其他非正文部分始终使用格式基准；差异列表只比较正文，不比较这些部分。',
  );
  const stylesText = await zip.file('word/styles.xml')?.async('string');
  const themeText = await zip.file('word/theme/theme1.xml')?.async('string');
  return {
    name,
    bytes,
    zip,
    xml,
    blocks,
    warnings,
    locked,
    styles: stylesText ? parseXML(stylesText) : undefined,
    theme: themeText ? parseXML(themeText) : undefined,
  };
}

export function compareWords(left: WordFile, right: WordFile): Row[] {
  const key = (b: Block) => `${b.type}:${b.text}`;
  const diffs = diffArrays(left.blocks, right.blocks, {
    comparator: (a, b) => key(a) === key(b),
    timeout: 2000,
  });
  if (!diffs) throw new Error('对比超时，请拆分文档后重试。');
  const rows: Row[] = [];
  const push = (l?: Block, r?: Block) => {
    let kind: Kind = !l
      ? 'added'
      : !r
        ? 'removed'
        : l.text !== r.text || l.type !== r.type
          ? 'modified'
          : l.fingerprint !== r.fingerprint
            ? 'format'
            : 'equal';
    if ((l && !l.safe) || (r && !r.safe)) kind = 'complex';
    rows.push({ id: `row-${rows.length}`, left: l, right: r, kind });
  };
  let li = 0,
    ri = 0;
  for (let i = 0; i < diffs.length; i++) {
    const part = diffs[i];
    if (!part.added && !part.removed) {
      for (let j = 0; j < part.value.length; j++) push(left.blocks[li++], right.blocks[ri++]);
    } else {
      const removed: Block[] = [],
        added: Block[] = [];
      while (i < diffs.length && (diffs[i].added || diffs[i].removed)) {
        const change = diffs[i];
        if (change.removed)
          for (let j = 0; j < change.value.length; j++) removed.push(left.blocks[li++]);
        if (change.added)
          for (let j = 0; j < change.value.length; j++) added.push(right.blocks[ri++]);
        i++;
      }
      i--;
      // Pair replacement runs in order; all unmatched blocks stay visible.
      for (let j = 0; j < Math.max(removed.length, added.length); j++) push(removed[j], added[j]);
    }
  }
  return rows;
}

function topology(el: Element): string {
  return JSON.stringify(
    children(el, 'tr').map((tr) =>
      children(tr, 'tc').map((tc) => ({
        p: children(tc, 'p').length,
        span: first(first(tc, 'tcPr') ?? tc, 'gridSpan')?.getAttributeNS(W, 'val') ?? '1',
        merge: serial(
          first(first(tc, 'tcPr') ?? tc, 'vMerge') ?? el.ownerDocument.createElement('none'),
        ),
      })),
    ),
  );
}

export function choiceIssue(
  row: Row,
  choice: Choice,
  base: Side,
  files: Record<Side, WordFile>,
): string | undefined {
  if (choice === base) return;
  if (files[base].locked || files[base === 'left' ? 'right' : 'left'].locked)
    return '存在跨段域、修订或批注范围，请先处理源文档。';
  const target = row[base];
  if (target && !target.safe) return '底稿此块含复杂结构，本版仅支持原样保留。';
  if (choice === 'omit') return;
  const sources = choice === 'both' ? [row.left, row.right] : [row[choice]];
  for (const source of sources) {
    if (!source || source === target) continue;
    if (!source.safe) return '此块含图片、公式、链接或其他复杂结构，暂不支持跨侧合入。';
    if (
      source.type === 'tbl' &&
      (!target || target.type !== 'tbl' || topology(source.element) !== topology(target.element))
    )
      return '表格需要两侧行列、合并单元格及段落结构一致；结构变化请在 Word 中处理。';
    if (target && source.type !== target.type) return '段落与表格互换暂不支持。';
  }
}

function writeRun(doc: Document, text: string, properties?: Element): Element {
  const run = doc.createElementNS(W, 'w:r');
  if (properties) run.appendChild(doc.importNode(properties, true));
  for (const part of text.split(/([\t\n])/)) {
    if (!part) continue;
    const node = doc.createElementNS(W, part === '\t' ? 'w:tab' : part === '\n' ? 'w:br' : 'w:t');
    if (node.localName === 't') {
      node.textContent = part;
      node.setAttributeNS(XML, 'xml:space', 'preserve');
    }
    run.appendChild(node);
  }
  return run;
}

// Preserve base formatting at character granularity for unchanged text. New
// characters inherit the run at the insertion/replacement point in the base.
function patchParagraph(base: Element, value: string): Element {
  const result = base.cloneNode(true) as Element;
  if (textOf(base) === value) return result;
  const runs = children(base, 'r');
  const spans: { start: number; end: number; properties?: Element }[] = [];
  let offset = 0;
  for (const run of runs) {
    const text = textOf(run);
    spans.push({ start: offset, end: offset + text.length, properties: first(run, 'rPr') });
    offset += text.length;
  }
  const at = (position: number) => {
    const span = spans.find((s) => position >= s.start && position < s.end);
    // An unformatted run is meaningful: do not borrow the final run's formatting.
    return span ? span.properties : spans[spans.length - 1]?.properties;
  };
  for (const child of Array.from(result.children)) if (child.localName !== 'pPr') child.remove();
  const changes = diffChars(textOf(base), value, { timeout: 1000 });
  if (!changes) throw new Error('段落字符对比超时，请缩短该段落后重试。');
  offset = 0;
  let replacementStart: number | undefined;
  for (const change of changes) {
    if (change.removed) {
      replacementStart = offset;
      offset += change.value.length;
      continue;
    }
    if (change.added) {
      result.appendChild(
        writeRun(
          base.ownerDocument,
          change.value,
          at(Math.min(replacementStart ?? offset, Math.max(0, textOf(base).length - 1))),
        ),
      );
      replacementStart = undefined;
      continue;
    }
    replacementStart = undefined;
    let consumed = 0;
    while (consumed < change.value.length) {
      const span = spans.find((s) => offset >= s.start && offset < s.end);
      const count = Math.min(
        change.value.length - consumed,
        span ? span.end - offset : change.value.length,
      );
      result.appendChild(
        writeRun(
          base.ownerDocument,
          change.value.slice(consumed, consumed + count),
          span?.properties,
        ),
      );
      offset += count;
      consumed += count;
    }
  }
  return result;
}

function transplant(
  source: Block,
  target: Block | undefined,
  baseFile: WordFile,
  row: Row,
  rows: Row[],
  base: Side,
): Element {
  if (target && source.type === 'tbl') {
    const table = target.element.cloneNode(true) as Element;
    const paragraphs = descendant(table, 'p');
    descendant(source.element, 'p').forEach((p, i) =>
      paragraphs[i].replaceWith(patchParagraph(paragraphs[i], textOf(p))),
    );
    return table;
  }
  if (target) return patchParagraph(target.element, source.text);
  // Inserted paragraphs use the closest ordinary paragraph in the base.
  const position = rows.indexOf(row);
  let template: Block | undefined;
  for (let distance = 1; distance < rows.length; distance++) {
    template = [rows[position - distance]?.[base], rows[position + distance]?.[base]].find(
      (b) => b?.safe && b.type === 'p' && b.text.trim(),
    );
    if (template) break;
  }
  const empty = baseFile.xml.createElementNS(W, 'w:p');
  if (template) {
    const props = first(template.element, 'pPr');
    if (props) {
      const copy = props.cloneNode(true) as Element;
      // New blocks do not inherit a list item or a section boundary accidentally.
      for (const n of children(copy, 'numPr')) n.remove();
      empty.appendChild(copy);
    }
    const rp = children(template.element, 'r')
      .map((r) => first(r, 'rPr'))
      .find(Boolean);
    empty.appendChild(writeRun(baseFile.xml, source.text, rp));
  } else empty.appendChild(writeRun(baseFile.xml, source.text));
  return empty;
}

export async function mergeWords(
  files: Record<Side, WordFile>,
  rows: Row[],
  plan: Plan,
): Promise<Uint8Array> {
  const baseFile = files[plan.base];
  if (
    rows.every(
      (row) =>
        (!plan.choices[row.id] || plan.choices[row.id] === plan.base) &&
        effectiveFormat(row, plan) === 'default',
    )
  )
    return baseFile.bytes.slice();
  const xml = baseFile.xml.cloneNode(true) as Document;
  const body = xml.getElementsByTagNameNS(W, 'body')[0];
  const section = children(body, 'sectPr')[0]?.cloneNode(true);
  for (const child of Array.from(body.childNodes)) body.removeChild(child);
  for (const row of rows) {
    const choice = plan.choices[row.id] ?? plan.base;
    const issue = choiceIssue(row, choice, plan.base, files);
    if (issue) throw new Error(`第 ${rows.indexOf(row) + 1} 块：${issue}`);
    const format = effectiveFormat(row, plan);
    const styleIssue = formatIssue(row, format, files, plan.base);
    if (choice !== 'omit' && styleIssue)
      throw new Error(`第 ${rows.indexOf(row) + 1} 块格式：${styleIssue}`);
    const selected =
      choice === 'omit' ? [] : choice === 'both' ? [row.left, row.right] : [row[choice]];
    for (const block of selected) {
      if (!block) continue;
      const element =
        format !== 'default'
          ? normalizeBlock(block, row, files, plan.base, format)
          : block === row[plan.base]
            ? block.element
            : transplant(block, row[plan.base], baseFile, row, rows, plan.base);
      body.appendChild(xml.importNode(element, true));
    }
  }
  // Word expects a body paragraph even if the user discards every block.
  if (!body.children.length) body.appendChild(xml.createElementNS(W, 'w:p'));
  if (section) body.appendChild(section);
  const output = await JSZip.loadAsync(baseFile.bytes);
  output.file(MAIN, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + serial(xml));
  return output.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

export function inlineDiff(left: string, right: string, side: Side) {
  return (
    diffChars(left, right, { timeout: 100 }) ?? [
      { value: side === 'left' ? left : right, added: false, removed: false },
    ]
  )
    .filter((part) => (side === 'left' ? !part.added : !part.removed))
    .map((part) => ({ text: part.value, changed: !!(part.added || part.removed) }));
}
