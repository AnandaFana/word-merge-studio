import JSZip from 'jszip';
import { themeFonts, paragraphProps, runProps } from './formatting';
import {
  W,
  all,
  child,
  value,
  make,
  put,
  orderProps,
  paragraphs,
  levelOf,
  unifyParagraph,
  formatProperties,
  typeLabel,
  type Chapter,
  type ChapterOptions,
} from './chapterOutline';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const MAIN = 'word/document.xml';
const xml = (text: string) => {
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw new Error('附件包含不支持的 XML 实体声明。');
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('文档附件 XML 损坏。');
  return doc;
};
const serial = (node: Node) => new XMLSerializer().serializeToString(node);
function copyNamespaces(source: Element, target: Element) {
  for (const a of Array.from(source.attributes))
    if (a.namespaceURI === 'http://www.w3.org/2000/xmlns/') {
      // Default namespaces are scoped by XMLSerializer on imported elements.
      if (a.name === 'xmlns') continue;
      const existing = target.getAttribute(a.name);
      if (existing && existing !== a.value)
        throw new Error('章节间存在冲突的 XML 命名空间，请在 Word 中统一另存格式。');
      target.setAttributeNS(a.namespaceURI, a.name, a.value);
    }
  const ignorable = [
    ...(target.getAttributeNS(MC, 'Ignorable') ?? '').split(/\s+/),
    ...(source.getAttributeNS(MC, 'Ignorable') ?? '').split(/\s+/),
  ].filter(Boolean);
  if (ignorable.length)
    target.setAttributeNS(MC, 'mc:Ignorable', [...new Set(ignorable)].join(' '));
}
// Snapshot with a TreeWalker: indexing a large live HTMLCollection in some DOMs
// becomes quadratic after many edits (notably a formatted document's run properties).
const elements = (root: Element): Element[] => {
  const result: Element[] = [root];
  const walker = root.ownerDocument.createTreeWalker(root, 1);
  while (walker.nextNode()) result.push(walker.currentNode as Element);
  return result;
};
const relPath = (part: string) =>
  `${part.slice(0, part.lastIndexOf('/') + 1)}_rels/${part.slice(part.lastIndexOf('/') + 1)}.rels`;
const resolve = (part: string, target: string) => {
  let decoded: string;
  try {
    decoded = decodeURI(target);
  } catch {
    throw new Error('附件路径编码损坏。');
  }
  if (/[\\?#]/.test(decoded) || /^[a-z]+:/i.test(decoded))
    throw new Error('不支持的内部附件路径。');
  const path = decoded.startsWith('/')
    ? decoded.slice(1)
    : part.slice(0, part.lastIndexOf('/') + 1) + decoded;
  const stack: string[] = [];
  for (const segment of path.split('/')) {
    if (segment === '..') {
      if (!stack.length) throw new Error('附件路径超出文档范围。');
      stack.pop();
    } else if (segment && segment !== '.') stack.push(segment);
  }
  return stack.join('/');
};
export function chapterIssues(chapter: Chapter): string[] {
  const root = chapter.file.xml;
  const unsupported: [string, string][] = [
    ['footnoteReference', '脚注'],
    ['endnoteReference', '尾注'],
    ['commentReference', '批注'],
    ['commentRangeStart', '批注'],
    ['ins', '修订'],
    ['del', '修订'],
    ['moveFrom', '移动修订'],
    ['moveTo', '移动修订'],
    ['fldChar', '域（例如自动目录或交叉引用）'],
    ['fldSimple', '域'],
    ['altChunk', '嵌入文档'],
    ['object', '嵌入对象'],
    ['pict', '旧式 VML 图形'],
    ['sdt', '内容控件'],
    ['subDoc', '子文档'],
    ['customXml', '自定义 XML 正文'],
    ['permStart', '受保护范围'],
  ];
  return [
    ...new Set(unsupported.filter(([tag]) => all(root, tag).length).map(([, label]) => label)),
  ];
}
export function chapterWarnings(chapter: Chapter): string[] {
  const warnings: string[] = [];
  if (all(chapter.file.xml, 'sectPr').length > 1)
    warnings.push('文档内有分节；合并后按首章页面设置排版。');
  if (all(chapter.file.xml, 'drawing').length)
    warnings.push('图片会复制；浮动位置、图表主题及分页请在 Word 中核对。');
  return warnings;
}

/** Compose OPC packages locally. Each chapter has isolated style, numbering and media identities. */
export async function mergeChapters(
  chapters: Chapter[],
  options: ChapterOptions,
): Promise<Uint8Array> {
  if (!chapters.length) throw new Error('请先添加章节文档。');
  if (
    chapters.length > 20 ||
    chapters.reduce((n, c) => n + c.file.bytes.length, 0) > 100 * 1024 * 1024
  )
    throw new Error('最多合并 20 份文件，总大小不超过 100 MB。');
  let expanded = 0;
  for (const chapter of chapters) {
    const issues = chapterIssues(chapter);
    if (issues.length)
      throw new Error(
        `${chapter.file.name} 含${issues.join('、')}；请在 Word 副本中处理这些内容后再添加。`,
      );
    for (const part of Object.values(chapter.file.zip.files))
      expanded +=
        (part as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? 0;
  }
  if (expanded > 250 * 1024 * 1024) throw new Error('所有文档解压后超过 250 MB，请分批合并。');
  if (chapters.reduce((n, c) => n + c.outline.length, 0) > 10000)
    throw new Error('总段落数超过 10000，请分批合并。');
  for (const [level, f] of Object.entries(options.formats)) {
    if (options.mode !== 'unified' || !f.enabled) continue;
    if (
      !f.font.trim() ||
      f.font.length > 80 ||
      !Number.isFinite(f.size) ||
      f.size < 5 ||
      f.size > 72 ||
      f.size * 2 !== Math.round(f.size * 2) ||
      !Number.isFinite(f.line) ||
      f.line < 1 ||
      f.line > 3 ||
      !Number.isFinite(f.after) ||
      f.after < 0 ||
      f.after > 72 ||
      !Number.isFinite(f.indent) ||
      f.indent < 0 ||
      f.indent > 4
    )
      throw new Error(
        `${typeLabel(Number(level))}的字体或排版参数无效。字号 5–72 pt（半点步进），行距 1–3，段后 0–72 pt，缩进 0–4 字。`,
      );
  }
  const zip = new JSZip();
  const main = xml(`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body/></w:document>`);
  const body = all(main, 'body')[0];
  const styles = xml(`<w:styles xmlns:w="${W}"/>`);
  const numbering = xml(`<w:numbering xmlns:w="${W}" xmlns:r="${R}"/>`);
  const rels = xml(`<Relationships xmlns="${REL}"/>`);
  const types = xml(
    `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`,
  );
  const addType = (path: string, type: string) => {
    const node = types.createElementNS(CT, 'Override');
    node.setAttribute('PartName', '/' + path);
    node.setAttribute('ContentType', type);
    types.documentElement.appendChild(node);
  };
  let nextRel = 1,
    nextNum = 1,
    nextAbstract = 0,
    nextPic = 0,
    nextDrawing = 1,
    nextBookmark = 0;
  const addRel = (type: string, target: string, external = false) => {
    const node = rels.createElementNS(REL, 'Relationship'),
      id = `wmRel${nextRel++}`;
    node.setAttribute('Id', id);
    node.setAttribute('Type', type);
    node.setAttribute('Target', target);
    if (external) node.setAttribute('TargetMode', 'External');
    rels.documentElement.appendChild(node);
    return id;
  };
  addRel(R + '/styles', 'styles.xml');
  addRel(R + '/numbering', 'numbering.xml');
  let finalSection: Element | undefined;
  for (let ci = 0; ci < chapters.length; ci++) {
    const chapter = chapters[ci],
      file = chapter.file,
      prefix = `WM${ci}_`;
    const sourceTypes = xml(await file.zip.file('[Content_Types].xml')!.async('string'));
    const sourceRels = file.zip.file(relPath(MAIN));
    const relations = sourceRels
      ? Array.from(xml(await sourceRels.async('string')).documentElement.children)
      : [];
    const relationById = new Map(relations.map((r) => [r.getAttribute('Id')!, r]));
    const typeFor = (path: string) => {
      const entries = Array.from(sourceTypes.documentElement.children);
      return (
        entries
          .find((e) => e.localName === 'Override' && e.getAttribute('PartName') === '/' + path)
          ?.getAttribute('ContentType') ??
        entries
          .find(
            (e) =>
              e.localName === 'Default' && e.getAttribute('Extension') === path.split('.').pop(),
          )
          ?.getAttribute('ContentType') ??
        'application/octet-stream'
      );
    };
    const stylesSource = file.styles;
    if (stylesSource) copyNamespaces(stylesSource.documentElement, styles.documentElement);
    const styleMap = new Map<string, string>();
    for (const s of stylesSource ? all(stylesSource, 'style') : [])
      styleMap.set(value(s, 'styleId'), prefix + value(s, 'styleId'));
    const numMap = new Map<string, string>([['0', '0']]),
      abstractMap = new Map<string, string>(),
      picMap = new Map<string, string>();
    const numRel = relations.find((r) => r.getAttribute('Type') === R + '/numbering');
    const numPath = numRel ? resolve(MAIN, numRel.getAttribute('Target')!) : 'word/numbering.xml';
    const numPart = file.zip.file(numPath);
    const sourceNumbering = numPart ? xml(await numPart.async('string')) : undefined;
    if (sourceNumbering) {
      copyNamespaces(sourceNumbering.documentElement, numbering.documentElement);
      all(sourceNumbering, 'num').forEach((n) => numMap.set(value(n, 'numId'), String(nextNum++)));
      all(sourceNumbering, 'abstractNum').forEach((n) =>
        abstractMap.set(value(n, 'abstractNumId'), String(nextAbstract++)),
      );
      all(sourceNumbering, 'numPicBullet').forEach((n) =>
        picMap.set(value(n, 'numPicBulletId'), String(nextPic++)),
      );
    }
    const defaults = (type: string) =>
      (stylesSource ? all(stylesSource, 'style') : []).find(
        (s) => value(s, 'type') === type && ['1', 'true', 'on'].includes(value(s, 'default')),
      );
    const defaultId = (type: string) =>
      styleMap.get(value(defaults(type), 'styleId')) ?? prefix + `Default${type}`;
    const bookmarks = new Map<string, string>(),
      bookmarkNames = new Map<string, string>();
    all(file.xml, 'bookmarkStart').forEach((b) => {
      bookmarks.set(value(b, 'id'), String(nextBookmark++));
      bookmarkNames.set(value(b, 'name'), prefix + value(b, 'name'));
    });
    const rewrite = (root: Element, assignDefaults = false) => {
      const nodes = elements(root);
      for (const n of nodes) {
        if (n.namespaceURI === W) {
          const tag = n.localName;
          if (
            [
              'pStyle',
              'rStyle',
              'tblStyle',
              'basedOn',
              'next',
              'link',
              'styleLink',
              'numStyleLink',
            ].includes(tag)
          ) {
            const id = value(n);
            if (styleMap.has(id)) n.setAttributeNS(W, 'w:val', styleMap.get(id)!);
            else if (['pStyle', 'rStyle', 'tblStyle'].includes(tag)) n.remove();
          }
          if (tag === 'numId') {
            const id = numMap.get(value(n));
            if (!id) throw new Error(`${file.name} 的列表编号定义缺失。`);
            n.setAttributeNS(W, 'w:val', id);
          }
          if (tag === 'abstractNumId')
            n.setAttributeNS(W, 'w:val', abstractMap.get(value(n)) ?? value(n));
          if (tag === 'lvlPicBulletId')
            n.setAttributeNS(W, 'w:val', picMap.get(value(n)) ?? value(n));
          if (tag === 'bookmarkStart' || tag === 'bookmarkEnd')
            n.setAttributeNS(W, 'w:id', bookmarks.get(value(n, 'id')) ?? value(n, 'id'));
          if (tag === 'bookmarkStart')
            n.setAttributeNS(W, 'w:name', bookmarkNames.get(value(n, 'name'))!);
          if (tag === 'hyperlink' && n.hasAttributeNS(W, 'anchor'))
            n.setAttributeNS(
              W,
              'w:anchor',
              bookmarkNames.get(value(n, 'anchor')) ?? value(n, 'anchor'),
            );
          if (tag === 'rPr') {
            themeFonts(file, n);
            orderProps(n);
          }
          if (assignDefaults && tag === 'p') {
            const pp = child(n, 'pPr') ?? make(n.ownerDocument, 'pPr');
            if (!pp.parentNode) n.insertBefore(pp, n.firstChild);
            if (!child(pp, 'pStyle')) put(pp, 'pStyle', { val: defaultId('paragraph') });
            orderProps(pp);
          }
          if (assignDefaults && tag === 'tbl') {
            const tp = child(n, 'tblPr') ?? make(n.ownerDocument, 'tblPr');
            if (!tp.parentNode) n.insertBefore(tp, n.firstChild);
            if (!child(tp, 'tblStyle'))
              tp.insertBefore(
                make(n.ownerDocument, 'tblStyle', { val: defaultId('table') }),
                tp.firstChild,
              );
          }
        }
        if (['docPr', 'cNvPr'].includes(n.localName) && n.hasAttribute('id'))
          n.setAttribute('id', String(nextDrawing++));
        for (const a of Array.from(n.attributes))
          if (['paraId', 'textId'].includes(a.localName)) n.removeAttributeNode(a);
      }
    };
    const copied = new Map<string, string>();
    const importPart = async (path: string): Promise<string> => {
      if (copied.has(path)) return copied.get(path)!;
      const part = file.zip.file(path);
      if (!part) throw new Error(`${file.name} 的附件缺失：${path}`);
      const target = `word/chapters/c${ci}/${path}`;
      copied.set(path, target);
      addType(target, typeFor(path));
      const bytes = await part.async('uint8array');
      if (/\.xml$/i.test(path)) {
        const doc = xml(new TextDecoder().decode(bytes));
        rewrite(doc.documentElement, true);
        zip.file(target, serial(doc));
      } else zip.file(target, bytes);
      const rp = file.zip.file(relPath(path));
      if (rp) {
        const rs = xml(await rp.async('string'));
        for (const r of Array.from(rs.documentElement.children)) {
          if (r.getAttribute('TargetMode') === 'External') {
            if (r.getAttribute('Type') !== R + '/hyperlink')
              throw new Error('附件包含外链资源，请在 Word 中将其嵌入后再合并。');
          } else
            r.setAttribute(
              'Target',
              '/' + (await importPart(resolve(path, r.getAttribute('Target')!))),
            );
        }
        zip.file(relPath(target), serial(rs));
      }
      return target;
    };
    const relMap = new Map<string, string>();
    const importMainRel = async (id: string) => {
      if (relMap.has(id)) return relMap.get(id)!;
      const r = relationById.get(id);
      if (!r) throw new Error(`${file.name} 包含无效的附件关系。`);
      const type = r.getAttribute('Type')!,
        external = r.getAttribute('TargetMode') === 'External';
      if (external && type !== R + '/hyperlink')
        throw new Error('存在外链图片，请在 Word 中将图片嵌入后再合并。');
      if (
        ![
          'image',
          'hyperlink',
          'header',
          'footer',
          'chart',
          'diagramData',
          'diagramLayout',
          'diagramQuickStyle',
          'diagramColors',
        ].includes(type.split('/').pop()!)
      )
        throw new Error(`暂不支持这种正文附件：${type.split('/').pop()}。`);
      const target = external
        ? r.getAttribute('Target')!
        : '/' + (await importPart(resolve(MAIN, r.getAttribute('Target')!)));
      const mapped = addRel(type, target, external);
      relMap.set(id, mapped);
      return mapped;
    };
    const remapBodyRels = async (root: Element) => {
      for (const e of elements(root))
        for (const a of Array.from(e.attributes))
          if (a.namespaceURI === R) e.setAttributeNS(R, a.name, await importMainRel(a.value));
    };
    // No document-wide default from one source is allowed to leak into another chapter.
    for (const type of ['paragraph', 'character', 'table']) {
      const root = make(styles, 'style', { type, styleId: prefix + `Default${type}` });
      root.appendChild(make(styles, 'name', { val: `章节 ${ci + 1} 默认 ${type}` }));
      const pp = stylesSource ? child(all(stylesSource, 'pPrDefault')[0], 'pPr') : undefined;
      const rp = stylesSource ? child(all(stylesSource, 'rPrDefault')[0], 'rPr') : undefined;
      if (type === 'paragraph' && pp) root.appendChild(pp.cloneNode(true));
      if (type !== 'table') {
        const rprops = rp ? (rp.cloneNode(true) as Element) : make(styles, 'rPr');
        themeFonts(file, rprops);
        orderProps(rprops);
        root.appendChild(rprops);
      }
      styles.documentElement.appendChild(root);
    }
    for (const source of stylesSource ? all(stylesSource, 'style') : []) {
      const s = source.cloneNode(true) as Element;
      s.setAttributeNS(W, 'w:styleId', styleMap.get(value(source, 'styleId'))!);
      s.removeAttributeNS(W, 'default');
      const name = child(s, 'name');
      if (name) name.setAttributeNS(W, 'w:val', `章节 ${ci + 1} · ${value(name)}`);
      rewrite(s);
      if (!child(s, 'basedOn') && ['paragraph', 'character', 'table'].includes(value(s, 'type'))) {
        const based = make(styles, 'basedOn', { val: prefix + `Default${value(s, 'type')}` });
        s.insertBefore(
          based,
          Array.from(s.children).find((e) => !['name', 'aliases'].includes(e.localName)) ?? null,
        );
      }
      styles.documentElement.appendChild(s);
    }
    if (sourceNumbering) {
      const nr = file.zip.file(relPath(numPath));
      const nrs = nr ? xml(await nr.async('string')) : undefined;
      const outputRels = xml(`<Relationships xmlns="${REL}"/>`);
      // Numbering images use relationship IDs scoped to numbering.xml, not document.xml.
      const oldOutput = zip.file('word/_rels/numbering.xml.rels');
      if (oldOutput)
        for (const e of Array.from(xml(await oldOutput.async('string')).documentElement.children))
          outputRels.documentElement.appendChild(outputRels.importNode(e, true));
      const numberRelMap = new Map<string, string>();
      for (const r of nrs ? Array.from(nrs.documentElement.children) : []) {
        if (r.getAttribute('TargetMode') === 'External')
          throw new Error('不支持外链图片项目符号。');
        const copy = r.cloneNode(true) as Element,
          id = prefix + r.getAttribute('Id');
        numberRelMap.set(r.getAttribute('Id')!, id);
        copy.setAttribute('Id', id);
        copy.setAttribute(
          'Target',
          '/' + (await importPart(resolve(numPath, r.getAttribute('Target')!))),
        );
        outputRels.documentElement.appendChild(copy);
      }
      if (outputRels.documentElement.children.length)
        zip.file('word/_rels/numbering.xml.rels', serial(outputRels));
      for (const source of Array.from(sourceNumbering.documentElement.children)) {
        if (!['num', 'abstractNum', 'numPicBullet'].includes(source.localName)) continue;
        const n = source.cloneNode(true) as Element;
        rewrite(n);
        const key =
          n.localName === 'num'
            ? 'numId'
            : n.localName === 'abstractNum'
              ? 'abstractNumId'
              : 'numPicBulletId';
        const map = key === 'numId' ? numMap : key === 'abstractNumId' ? abstractMap : picMap;
        n.setAttributeNS(W, 'w:' + key, map.get(value(source, key))!);
        if (n.localName === 'abstractNum') {
          const nsid = child(n, 'nsid') ?? make(numbering, 'nsid');
          nsid.setAttributeNS(
            W,
            'w:val',
            (0xa0000000 + Number(value(n, 'abstractNumId'))).toString(16).toUpperCase(),
          );
          if (!nsid.parentNode) n.insertBefore(nsid, n.firstChild);
        }
        for (const e of elements(n))
          for (const a of Array.from(e.attributes))
            if (a.namespaceURI === R && numberRelMap.has(a.value))
              e.setAttributeNS(R, a.name, numberRelMap.get(a.value)!);
        numbering.documentElement.appendChild(n);
      }
    }
    // Keep namespace bindings used by markup-compatibility attributes on imported content.
    copyNamespaces(file.xml.documentElement, main.documentElement);

    const sourcePs = paragraphs(file);
    const copies = file.blocks.map((b) => b.element.cloneNode(true) as Element);
    const container = make(main, 'body');
    copies.forEach((b) => container.appendChild(b));
    // Paragraphs in text boxes are excluded from classification and stay untouched.
    const copyPs = all(container, 'p').filter((p) => {
      let a = p.parentElement;
      while (a) {
        if (a.namespaceURI === W && a.localName === 'txbxContent') return false;
        a = a.parentElement;
      }
      return true;
    });
    // Lists inherited from source styles remain lists when assigning a unified paragraph style.
    copyPs.forEach((p, i) => {
      if (options.mode !== 'unified') return;
      const level = levelOf(chapter, chapter.outline[i]);
      if (level < 0 || !options.formats[level]?.enabled) return;
      const effective = paragraphProps(file, sourcePs[i]);
      if (child(effective, 'numPr')) {
        const pp = child(p, 'pPr') ?? make(main, 'pPr');
        if (!pp.parentNode) p.insertBefore(pp, p.firstChild);
        if (!child(pp, 'numPr')) pp.appendChild(child(effective, 'numPr')!.cloneNode(true));
        if (!child(pp, 'ind') && child(effective, 'ind'))
          pp.appendChild(child(effective, 'ind')!.cloneNode(true));
      }
      const sourceRuns = all(sourcePs[i], 'r');
      all(p, 'r').forEach((r, ri) => {
        const resolved = runProps(file, sourcePs[i], sourceRuns[ri]);
        for (const tag of ['i', 'iCs', 'vertAlign']) {
          const prop = child(resolved, tag);
          if (!prop) continue;
          const rp = child(r, 'rPr') ?? make(main, 'rPr');
          if (!rp.parentNode) r.insertBefore(rp, r.firstChild);
          if (!child(rp, tag)) rp.appendChild(prop.cloneNode(true));
          orderProps(rp);
        }
      });
    });

    for (const block of Array.from(container.children)) {
      const placeholder = main.createComment('block');
      block.replaceWith(placeholder);
      rewrite(block, true);
      placeholder.replaceWith(block);
    }

    all(container, 'sectPr').forEach((s) => s.remove());
    if (options.mode === 'unified')
      copyPs.forEach((p, i) => {
        // Edit detached paragraphs so XML DOM implementations do not repeatedly invalidate
        // live collections for the entire chapter on every run-property mutation.
        const placeholder = main.createComment('paragraph');
        p.replaceWith(placeholder);
        unifyParagraph(
          p,
          levelOf(chapter, chapter.outline[i]),
          options.formats[levelOf(chapter, chapter.outline[i])],
        );
        placeholder.replaceWith(p);
      });
    await remapBodyRels(container);
    if (ci > 0 && options.pageBreak) {
      const first = container.firstElementChild;
      if (first?.namespaceURI === W && first.localName === 'p') {
        const pp = child(first, 'pPr') ?? make(main, 'pPr');
        if (!pp.parentNode) first.insertBefore(pp, first.firstChild);
        put(pp, 'pageBreakBefore', { val: '1' });
        orderProps(pp);
      } else if (first) {
        const breakP = make(main, 'p'),
          props = make(main, 'pPr');
        put(props, 'spacing', { before: '0', after: '0', line: '20', lineRule: 'exact' });
        breakP.appendChild(props);
        const r = make(main, 'r');
        r.appendChild(make(main, 'br', { type: 'page' }));
        breakP.appendChild(r);
        body.appendChild(breakP);
      }
    }
    Array.from(container.children).forEach((e) => body.appendChild(e));
    if (ci === 0) {
      const section = child(all(file.xml, 'body')[0], 'sectPr');
      if (section) {
        finalSection = section.cloneNode(true) as Element;
        await remapBodyRels(finalSection);
      }
      // Global settings and theme follow the first chapter, as stated in the UI.
      for (const kind of ['theme', 'fontTable']) {
        const r = relations.find((r) => r.getAttribute('Type') === R + '/' + kind);
        if (r)
          addRel(
            R + '/' + kind,
            '/' + (await importPart(resolve(MAIN, r.getAttribute('Target')!))),
          );
      }
      const settingRel = relations.find((r) => r.getAttribute('Type') === R + '/settings');
      const settingPart =
        settingRel && file.zip.file(resolve(MAIN, settingRel.getAttribute('Target')!));
      const settings = xml(`<w:settings xmlns:w="${W}"/>`);
      if (settingPart) {
        const original = xml(await settingPart.async('string'));
        for (const tag of [
          'evenAndOddHeaders',
          'defaultTabStop',
          'characterSpacingControl',
          'compat',
        ]) {
          const e = child(original.documentElement, tag);
          if (e) settings.documentElement.appendChild(e.cloneNode(true));
        }
      }
      zip.file('word/settings.xml', serial(settings));
      addRel(R + '/settings', 'settings.xml');
      addType(
        'word/settings.xml',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
      );
    }
  }
  if (options.mode === 'unified')
    for (const [level, f] of Object.entries(options.formats))
      if (f.enabled) {
        const s = make(styles, 'style', {
          type: 'paragraph',
          styleId: `WMUnified${level}`,
          customStyle: '1',
        });
        s.appendChild(make(styles, 'name', { val: `统一 · ${typeLabel(Number(level))}` }));
        s.appendChild(make(styles, 'qFormat'));
        const props = formatProperties(styles, Number(level), f);
        s.append(props.pp, props.rp);
        styles.documentElement.appendChild(s);
      }
  if (finalSection) body.appendChild(finalSection);
  const numOrder = ['numPicBullet', 'abstractNum', 'num'];
  Array.from(numbering.documentElement.children)
    .sort((a, b) => numOrder.indexOf(a.localName) - numOrder.indexOf(b.localName))
    .forEach((e) => numbering.documentElement.appendChild(e));
  zip.file(MAIN, serial(main));
  zip.file('word/styles.xml', serial(styles));
  zip.file('word/numbering.xml', serial(numbering));
  zip.file(relPath(MAIN), serial(rels));
  zip.file(
    '_rels/.rels',
    `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  addType(MAIN, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml');
  addType(
    'word/styles.xml',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
  );
  addType(
    'word/numbering.xml',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
  );
  zip.file('[Content_Types].xml', serial(types));
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
