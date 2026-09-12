import { t } from './i18n';
import { textOf, type WordFile } from './engine';
import { paragraphProps, dominantRun } from './formatting';

export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const all = (node: Element | Document, name: string) =>
  Array.from(node.getElementsByTagNameNS(W, name));
export const child = (node: Element | undefined, name: string) =>
  Array.from(node?.children ?? []).find((e) => e.namespaceURI === W && e.localName === name);
export const value = (node?: Element, name = 'val') => node?.getAttributeNS(W, name) ?? '';
export const make = (doc: Document, name: string, attrs: Record<string, string> = {}) => {
  const node = doc.createElementNS(W, `w:${name}`);
  for (const [key, val] of Object.entries(attrs)) node.setAttributeNS(W, `w:${key}`, val);
  return node;
};
export const typeLabel = (level: number) =>
  level === -1
    ? t('保持原样')
    : level === 0
      ? t('正文')
      : level === 10
        ? t('图表题注')
        : level === 11
          ? t('表格文字')
          : t('{0} 级标题', level);
export interface OutlineItem {
  index: number;
  text: string;
  level: number;
  reason: string;
  group: string;
  inTable: boolean;
}
export interface Chapter {
  id: string;
  file: WordFile;
  outline: OutlineItem[];
  overrides: Record<number, number>;
}
export interface TypeFormat {
  enabled: boolean;
  font: string;
  size: number;
  bold: boolean;
  align: 'left' | 'center' | 'both';
  line: number;
  after: number;
  indent: number;
}
export type TypeFormats = Record<number, TypeFormat>;
export interface ChapterOptions {
  mode: 'original' | 'unified';
  pageBreak: boolean;
  formats: TypeFormats;
  /** Explicit assembly order. References use positions in the original file.blocks array. */
  assembly?: ({ chapterId: string; blockIndex: number } | { title: string; level: number })[];
}
export function defaultFormats(preset: 'report' | 'simple' = 'report'): TypeFormats {
  return Object.fromEntries(
    Array.from({ length: 12 }, (_, level) => [
      level,
      {
        enabled: true,
        font: preset === 'simple' ? '微软雅黑' : level >= 1 && level <= 9 ? '黑体' : '宋体',
        size:
          level === 1
            ? 18
            : level === 2
              ? 16
              : level === 3
                ? 14
                : level === 10 || level === 11
                  ? 10.5
                  : 12,
        bold: level >= 1 && level <= 9,
        align: level === 10 ? 'center' : 'left',
        line: level === 11 ? 1.15 : 1.5,
        after: level >= 1 && level <= 9 ? 8 : 0,
        indent: level === 0 && preset === 'report' ? 2 : 0,
      },
    ]),
  );
}
export function paragraphs(file: WordFile): Element[] {
  return all(file.xml, 'p').filter((p) => {
    let ancestor = p.parentElement;
    while (ancestor) {
      if (ancestor.namespaceURI === W && ancestor.localName === 'txbxContent') return false;
      ancestor = ancestor.parentElement;
    }
    return true;
  });
}
export function detectOutline(file: WordFile): OutlineItem[] {
  const ps = paragraphs(file);
  const sizes = new Map<number, number>();
  for (const p of ps) {
    const size = Number(value(child(dominantRun(file, p), 'sz')));
    if (size) sizes.set(size, (sizes.get(size) ?? 0) + textOf(p).length);
  }
  const bodySize = [...sizes].sort((a, b) => b[1] - a[1])[0]?.[0] ?? 24;
  return ps.map((p, index) => {
    const text = textOf(p).trim();
    const props = paragraphProps(file, p),
      run = dominantRun(file, p);
    const id = value(child(props, 'pStyle'));
    const style = file.styles && all(file.styles, 'style').find((s) => value(s, 'styleId') === id);
    const name = value(child(style, 'name'));
    const outline = value(child(props, 'outlineLvl'));
    const size = Number(value(child(run, 'sz'))) || bodySize;
    const bold = child(run, 'b') && !['0', 'false', 'off'].includes(value(child(run, 'b')));
    let a = p.parentElement,
      inTable = false;
    while (a) {
      if (a.localName === 'tc' && a.namespaceURI === W) inTable = true;
      a = a.parentElement;
    }
    let level = 0,
      reason = '普通正文',
      group = 'body';
    // Table cells and captions must not become navigation headings by typography alone.
    if (inTable) {
      level = 11;
      reason = '表格单元格';
      group = 'table';
    } else if (/^[图表]\s*[\d一二三四五六七八九十]+(?:[-－.．]\d+)*[\s：:、]/.test(text)) {
      level = 10;
      reason = '图表题注';
      group = 'caption';
    } else if (outline !== '' && Number(outline) < 9 && text) {
      level = Number(outline) + 1;
      reason = '已有大纲级别';
      group = `style:${id}:outline:${outline}`;
    } else if (
      /^(?:heading\s*|标题\s*)([1-9])$/i.test(name) ||
      /^(?:heading|标题)([1-9])$/i.test(id)
    ) {
      level = Number((name + ' ' + id).match(/(?:heading\s*|标题\s*)([1-9])/i)?.[1] ?? 1);
      reason = '已有标题样式';
      group = `style:${id}`;
    } else if (text && text.length <= 80 && !/[。；;！？!?]$/.test(text)) {
      const decimal = text.match(/^(\d+(?:[.．]\d+){1,8})(?:[.．、]?\s*[^\d.．\s])/);
      if (
        /^第[零〇一二三四五六七八九十百\d]+[章节篇部]/.test(text) ||
        /^[一二三四五六七八九十]+[、．.]/.test(text)
      ) {
        level = 1;
        reason = '中文章节编号';
        group = 'chinese-chapter';
      } else if (/^[（(][一二三四五六七八九十]+[）)]/.test(text)) {
        level = 2;
        reason = '中文括号编号（请核对）';
        group = 'chinese-paren';
      } else if (decimal) {
        level = Math.min(9, decimal[1].split(/[.．]/).length);
        reason = '多级数字编号';
        group = `decimal:${level}`;
      } else if (/^\d+[、．.]\s*\S/.test(text)) {
        level = 3;
        reason = '单层数字编号（请核对）';
        group = 'number-item';
      } else if (size > bodySize * 1.18) {
        level = size >= bodySize * 1.5 ? 1 : 2;
        reason = '字号推断（请核对）';
        group = `visual:${size}:${!!bold}`;
      } else if (bold && text.length <= 35 && !child(props, 'numPr') && !/[：:]$/.test(text)) {
        level = 3;
        reason = '短句加粗（请核对）';
        group = `visual:${size}:bold`;
      }
    }
    return { index, text, level, reason, group, inTable };
  });
}
export function createChapter(file: WordFile): Chapter {
  return { id: crypto.randomUUID(), file, outline: detectOutline(file), overrides: {} };
}
export const levelOf = (chapter: Chapter, item: OutlineItem) =>
  chapter.overrides[item.index] ?? item.level;

// Keep properties in the schema order used by Word, including when overriding direct formatting.
const PP =
  'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange'.split(
    ' ',
  );
const RP =
  'rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange'.split(
    ' ',
  );
export function orderProps(props: Element) {
  const names = props.localName === 'pPr' ? PP : RP;
  const current = Array.from(props.children);
  const sorted = [...current].sort(
    (a, b) =>
      (names.indexOf(a.localName) < 0 ? 999 : names.indexOf(a.localName)) -
      (names.indexOf(b.localName) < 0 ? 999 : names.indexOf(b.localName)),
  );
  if (sorted.some((e, i) => e !== current[i])) sorted.forEach((e) => props.appendChild(e));
}
export function put(parent: Element, name: string, attrs: Record<string, string>) {
  const element = make(parent.ownerDocument, name, attrs);
  const old = child(parent, name);
  if (old) old.replaceWith(element);
  else parent.appendChild(element);
  return element;
}
export function formatProperties(doc: Document, level: number, f: TypeFormat) {
  const pp = make(doc, 'pPr'),
    rp = make(doc, 'rPr');
  if (level >= 1 && level <= 9) {
    put(pp, 'keepNext', { val: '1' });
    put(pp, 'keepLines', { val: '1' });
  }
  put(pp, 'snapToGrid', { val: '0' });
  put(pp, 'spacing', {
    before: level >= 1 && level <= 9 ? '160' : '0',
    after: String(f.after * 20),
    line: String(Math.round(f.line * 240)),
    lineRule: 'auto',
  });
  put(pp, 'ind', {
    left: '0',
    right: '0',
    firstLine: String(Math.round(f.indent * f.size * 20)),
    firstLineChars: String(f.indent * 100),
  });
  put(pp, 'jc', { val: f.align });
  put(pp, 'outlineLvl', { val: String(level >= 1 && level <= 9 ? level - 1 : 9) });
  put(rp, 'rFonts', { ascii: f.font, hAnsi: f.font, eastAsia: f.font, cs: f.font });
  put(rp, 'b', { val: f.bold ? '1' : '0' });
  put(rp, 'bCs', { val: f.bold ? '1' : '0' });
  put(rp, 'color', { val: '000000' });
  put(rp, 'sz', { val: String(f.size * 2) });
  put(rp, 'szCs', { val: String(f.size * 2) });
  return { pp, rp };
}
export function unifyParagraph(p: Element, level: number, f?: TypeFormat) {
  if (level < 0) return;
  const pp = child(p, 'pPr') ?? make(p.ownerDocument, 'pPr');
  if (!pp.parentNode) p.insertBefore(pp, p.firstChild);
  put(pp, 'outlineLvl', { val: String(level >= 1 && level <= 9 ? level - 1 : 9) });
  if (!f?.enabled) {
    orderProps(pp);
    return;
  }
  const props = formatProperties(p.ownerDocument, level, f);
  put(pp, 'pStyle', { val: `WMUnified${level}` });
  // Keep explicit list numbering and its indentation; do not renumber chapter text.
  for (const prop of Array.from(props.pp.children)) {
    if (prop.localName === 'ind' && child(pp, 'numPr')) continue;
    child(pp, prop.localName)?.remove();
    pp.appendChild(prop.cloneNode(true));
  }
  const applyRun = (rp: Element) => {
    // Preserve italic, emphasis, super/subscripts and hyperlinks, but enforce chosen typography.
    for (const prop of Array.from(props.rp.children)) {
      child(rp, prop.localName)?.remove();
      rp.appendChild(prop.cloneNode(true));
    }
    orderProps(rp);
  };
  const marker = child(pp, 'rPr') ?? make(p.ownerDocument, 'rPr');
  if (!marker.parentNode) pp.appendChild(marker);
  applyRun(marker);
  for (const r of all(p, 'r')) {
    let parent = r.parentElement;
    while (parent && parent.localName !== 'p') parent = parent.parentElement;
    if (parent !== p) continue;
    const rp = child(r, 'rPr') ?? make(p.ownerDocument, 'rPr');
    if (!rp.parentNode) r.insertBefore(rp, r.firstChild);
    applyRun(rp);
  }
  orderProps(pp);
}
