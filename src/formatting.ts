import type { Block, FormatChoice, Plan, Row, Side, WordFile } from './engine';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const all = (el: Element | Document, name: string) =>
  Array.from(el.getElementsByTagNameNS(W, name));
const child = (el: Element | undefined, name: string) =>
  Array.from(el?.children ?? []).find((e) => e.namespaceURI === W && e.localName === name);
const val = (el?: Element) => el?.getAttributeNS(W, 'val') ?? '';
const text = (el: Element) =>
  all(el, 't')
    .map((t) => t.textContent ?? '')
    .join('');
const make = (doc: Document, name: string) => doc.createElementNS(W, `w:${name}`);
const clone = (el: Element) => el.cloneNode(true) as Element;

function cascade(doc: Document, name: string, layers: (Element | undefined)[]): Element {
  const result = make(doc, name);
  for (const layer of layers)
    for (const prop of Array.from(layer?.children ?? [])) {
      const old = child(result, prop.localName);
      // Composite properties can override individual attributes, e.g. only eastAsia.
      if (old && ['rFonts', 'lang', 'spacing', 'ind'].includes(prop.localName)) {
        for (const attr of Array.from(prop.attributes))
          old.setAttributeNS(attr.namespaceURI, attr.name, attr.value);
        if (prop.localName === 'rFonts')
          for (const script of ['ascii', 'hAnsi', 'eastAsia', 'cs']) {
            if (prop.hasAttributeNS(W, script)) old.removeAttributeNS(W, `${script}Theme`);
          }
      } else if (old) old.replaceWith(doc.importNode(prop, true));
      else result.appendChild(doc.importNode(prop, true));
    }
  return result;
}

const styleChains = new WeakMap<WordFile, Map<string, Element[]>>();
function styleChain(file: WordFile, id: string, type = 'paragraph'): Element[] {
  let cached = styleChains.get(file);
  if (!cached) {
    cached = new Map();
    styleChains.set(file, cached);
  }
  const key = `${type}:${id}`;
  if (cached.has(key)) return cached.get(key)!;
  const styles = file.styles ? all(file.styles, 'style') : [];
  let style = id
    ? styles.find((s) => s.getAttributeNS(W, 'styleId') === id)
    : styles.find(
        (s) =>
          s.getAttributeNS(W, 'type') === type &&
          ['1', 'true'].includes(s.getAttributeNS(W, 'default') ?? ''),
      );
  const chain: Element[] = [],
    seen = new Set<Element>();
  while (style && !seen.has(style)) {
    seen.add(style);
    chain.unshift(style);
    const parent = val(child(style, 'basedOn'));
    style = parent ? styles.find((s) => s.getAttributeNS(W, 'styleId') === parent) : undefined;
  }
  cached.set(key, chain);
  return chain;
}

function paragraphProps(file: WordFile, p: Element): Element {
  const defaults = file.styles ? all(file.styles, 'pPrDefault')[0] : undefined;
  return cascade(file.xml, 'pPr', [
    child(defaults, 'pPr'),
    ...styleChain(file, val(child(child(p, 'pPr'), 'pStyle'))).map((s) => child(s, 'pPr')),
    child(p, 'pPr'),
  ]);
}

function themeFonts(file: WordFile, props: Element): void {
  const fonts = child(props, 'rFonts');
  for (const script of fonts ? ['ascii', 'hAnsi', 'eastAsia', 'cs'] : []) {
    const theme = fonts!.getAttributeNS(W, `${script}Theme`);
    if (!theme) continue;
    const family = file.theme?.getElementsByTagNameNS(
      A,
      theme.startsWith('major') ? 'majorFont' : 'minorFont',
    )[0];
    const tag = script === 'eastAsia' ? 'ea' : script === 'cs' ? 'cs' : 'latin';
    let face = family?.getElementsByTagNameNS(A, tag)[0]?.getAttribute('typeface');
    if (!face && script === 'eastAsia') {
      const lang = child(props, 'lang')?.getAttributeNS(W, 'eastAsia') ?? 'zh-CN';
      const languageScript = lang.startsWith('ja')
        ? 'Jpan'
        : lang.startsWith('ko')
          ? 'Hang'
          : /TW|HK|Hant/i.test(lang)
            ? 'Hant'
            : 'Hans';
      face = Array.from(family?.getElementsByTagNameNS(A, 'font') ?? [])
        .find((f) => f.getAttribute('script') === languageScript)
        ?.getAttribute('typeface');
    }
    if (face) {
      fonts!.setAttributeNS(W, `w:${script}`, face);
      fonts!.removeAttributeNS(W, `${script}Theme`);
    }
  }
  const color = child(props, 'color');
  const name = color?.getAttributeNS(W, 'themeColor');
  if (color && name) {
    const aliases: Record<string, string> = {
      text1: 'dk1',
      text2: 'dk2',
      background1: 'lt1',
      background2: 'lt2',
      dark1: 'dk1',
      light1: 'lt1',
      dark2: 'dk2',
      light2: 'lt2',
      hyperlink: 'hlink',
      followedHyperlink: 'folHlink',
    };
    const source = file.theme?.getElementsByTagNameNS(A, aliases[name] ?? name)[0]
      ?.firstElementChild;
    let rgb = source?.getAttribute('lastClr') ?? source?.getAttribute('val');
    if (rgb && /^[\da-f]{6}$/i.test(rgb)) {
      for (const modifier of ['themeShade', 'themeTint']) {
        const fraction = color.getAttributeNS(W, modifier);
        if (fraction) {
          const factor = parseInt(fraction, 16) / 255;
          rgb = rgb
            .match(/../g)!
            .map((c) => {
              const n = parseInt(c, 16);
              return Math.round(
                modifier === 'themeShade' ? n * factor : n * factor + 255 * (1 - factor),
              )
                .toString(16)
                .padStart(2, '0');
            })
            .join('');
        }
      }
      color.setAttributeNS(W, 'w:val', rgb);
      for (const attr of ['themeColor', 'themeShade', 'themeTint'])
        color.removeAttributeNS(W, attr);
    }
  }
}

function runProps(file: WordFile, p: Element, run?: Element): Element {
  const defaults = file.styles ? all(file.styles, 'rPrDefault')[0] : undefined;
  const pStyles = styleChain(file, val(child(child(p, 'pPr'), 'pStyle')));
  const rStyles = styleChain(file, val(child(child(run, 'rPr'), 'rStyle')), 'character');
  const result = cascade(file.xml, 'rPr', [
    child(defaults, 'rPr'),
    ...pStyles.map((s) => child(s, 'rPr')),
    ...rStyles.map((s) => child(s, 'rPr')),
    child(run, 'rPr'),
  ]);
  child(result, 'rStyle')?.remove();
  themeFonts(file, result);
  return result;
}

const dominantRuns = new WeakMap<WordFile, WeakMap<Element, Element>>();
function dominantRun(file: WordFile, p: Element): Element {
  let cache = dominantRuns.get(file);
  if (!cache) {
    cache = new WeakMap();
    dominantRuns.set(file, cache);
  }
  const previous = cache.get(p);
  if (previous) return clone(previous);
  const groups = new Map<string, { props: Element; weight: number }>();
  for (const run of Array.from(p.children).filter((e) => e.localName === 'r')) {
    const weight = text(run).trim().length;
    if (!weight) continue;
    const props = runProps(file, p, run);
    // Ignore language/proofing metadata when counting visual styles.
    const key = Array.from(props.children)
      .filter((e) => !['lang', 'noProof'].includes(e.localName))
      .map(
        (e) =>
          `${e.localName}:${Array.from(e.attributes)
            .map((a) => `${a.localName}=${a.value}`)
            .sort()
            .join(',')}`,
      )
      .sort()
      .join('|');
    const old = groups.get(key);
    if (old) old.weight += weight;
    else groups.set(key, { props, weight });
  }
  const result =
    [...groups.values()].sort((a, b) => b.weight - a.weight)[0]?.props ?? runProps(file, p);
  cache.set(p, clone(result));
  return result;
}

const bodySizes = new WeakMap<WordFile, number>();
function usualBodySize(file: WordFile): number {
  if (bodySizes.has(file)) return bodySizes.get(file)!;
  const sizes = new Map<number, number>();
  for (const block of file.blocks) {
    if (!block.safe || block.type !== 'p' || !block.text.trim()) continue;
    const size = Number(val(child(dominantRun(file, block.element), 'sz')));
    if (size) sizes.set(size, (sizes.get(size) ?? 0) + block.text.trim().length);
  }
  const result = [...sizes.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? 0;
  bodySizes.set(file, result);
  return result;
}

export function paragraphRole(file: WordFile, p: Element): string {
  const props = paragraphProps(file, p);
  const id = val(child(child(p, 'pPr'), 'pStyle'));
  const name = styleChain(file, id)
    .map((s) => val(child(s, 'name')))
    .join(' ');
  const outline = val(child(props, 'outlineLvl'));
  if ((outline && Number(outline) < 9) || /heading|title|标题/i.test(name + ' ' + id))
    return 'heading';
  const value = text(p).trim();
  if (
    value.length < 70 &&
    /^(?:第.{1,10}[章节篇]|[一二三四五六七八九十]+[、.]|\d+(?:\.\d+)+\s*\S)/.test(value)
  )
    return 'heading';
  if (child(props, 'numPr') || /^[·•▪●]\s*/.test(value)) return 'list';
  const size = Number(val(child(dominantRun(file, p), 'sz')));
  const usual = usualBodySize(file);
  if (
    value.length > 0 &&
    value.length < 70 &&
    !/[。！？]$/.test(value) &&
    usual &&
    size > usual * 1.18
  )
    return 'heading';
  return 'body';
}

export function effectiveFormat(row: Row, plan: Plan): FormatChoice {
  const explicit = plan.formats?.[row.id] ?? 'default';
  if (explicit !== 'default') return explicit;
  // Smart is opt-in and never rewrites protected/unknown blocks automatically.
  return plan.formatMode === 'smart' && [row.left, row.right].every((b) => !b || b.safe)
    ? 'smart'
    : 'default';
}

export function formatIssue(
  row: Row,
  format: FormatChoice,
  files: Record<Side, WordFile>,
  base: Side,
): string | undefined {
  if (format === 'default') return;
  if (files.left.locked || files.right.locked) return '跨段结构锁定，暂不能重排格式。';
  if ([row.left, row.right].some((b) => b && !b.safe)) return '复杂内容仅支持原样保留格式。';
  if (format === 'left' || format === 'right') {
    const template = row[format];
    if (!template) return '该侧没有对应段落，请使用智能匹配。';
    if (row.left && row.right && row.left.type !== row.right.type)
      return '段落与表格之间不能直接套用格式。';
    if (row.left?.type === 'tbl' && row.right?.type === 'tbl') {
      const shape = (b: Block) =>
        all(b.element, 'tr')
          .map((tr) =>
            Array.from(tr.children)
              .filter((e) => e.localName === 'tc')
              .map((tc) =>
                [
                  all(tc, 'p').length,
                  val(all(tc, 'gridSpan')[0]),
                  all(tc, 'vMerge')
                    .map((e) => val(e) || 'continue')
                    .join(','),
                ].join(':'),
              )
              .join('|'),
          )
          .join('/');
      if (shape(row.left) !== shape(row.right)) return '表格结构不一致，不能直接套用格式。';
    }
    if (format !== base) {
      const ps = template.type === 'p' ? [template.element] : all(template.element, 'p');
      if (ps.some((p) => child(paragraphProps(files[format], p), 'numPr')))
        return '该侧含自动编号，尚不支持跨文档编号格式迁移。';
      if (
        ps.some((p) =>
          Array.from(paragraphProps(files[format], p).getElementsByTagName('*')).some((e) =>
            ['framePr', 'cnfStyle'].includes(e.localName),
          ),
        )
      )
        return '该侧含浮动框架或条件样式，暂不能跨侧迁移格式。';
      if (
        ps.some((p) =>
          Array.from(paragraphProps(files[format], p).getElementsByTagName('*')).some((e) =>
            Array.from(e.attributes).some((a) => /Theme|^theme/.test(a.localName)),
          ),
        )
      )
        return '段落边框或底纹包含主题格式，暂不能跨侧迁移。';
    }
  }
}

function smartTemplate(
  source: Element,
  sourceFile: WordFile,
  target: Element | undefined,
  baseFile: WordFile,
  position: number,
): Element {
  const role = paragraphRole(sourceFile, source);
  if (target && text(target).trim() && paragraphRole(baseFile, target) === role) return target;
  const candidates = baseFile.blocks.filter(
    (b) => b.safe && b.type === 'p' && b.text.trim() && paragraphRole(baseFile, b.element) === role,
  );
  if (!candidates.length) {
    // No compatible template: use the base default paragraph style, never a nearby heading.
    return make(baseFile.xml, 'p');
  }
  const groups = new Map<string, Block[]>();
  for (const b of candidates) {
    const props = paragraphProps(baseFile, b.element);
    const fonts = dominantRun(baseFile, b.element);
    const key = [
      val(child(props, 'pStyle')),
      val(child(fonts, 'sz')),
      child(fonts, 'rFonts')?.getAttributeNS(W, 'eastAsia'),
      val(child(props, 'jc')),
    ].join('|');
    groups.set(key, [...(groups.get(key) ?? []), b]);
  }
  const common = [...groups.values()].sort((a, b) => b.length - a.length)[0];
  return [...common].sort((a, b) => Math.abs(a.index - position) - Math.abs(b.index - position))[0]
    .element;
}

export function formatSummary(file: WordFile, block?: Block): string {
  if (!block) return '无对应段落';
  if (!block.safe) return '复杂结构';
  if (block.type === 'tbl') return '表格文字格式';
  const props = dominantRun(file, block.element);
  const name = { heading: '标题', body: '正文', list: '列表' }[paragraphRole(file, block.element)];
  const font =
    child(props, 'rFonts')?.getAttributeNS(W, 'eastAsia') ||
    child(props, 'rFonts')?.getAttributeNS(W, 'ascii') ||
    '默认字体';
  const size = val(child(props, 'sz'));
  return `${name} · ${font}${size ? ` · ${Number(size) / 2}pt` : ''}`;
}

function uniformParagraph(
  value: Element,
  template: Element,
  templateFile: WordFile,
  baseFile: WordFile,
  inserted: boolean,
): Element {
  const result = make(baseFile.xml, 'p');
  const pp = paragraphProps(templateFile, template);
  const rp = dominantRun(templateFile, template);
  // A paragraph marker's rPr must not reintroduce a different font on later edits.
  child(pp, 'rPr')?.remove();
  if (inserted) child(pp, 'numPr')?.remove();
  if (templateFile !== baseFile) {
    child(pp, 'pStyle')?.remove();
    // Explicit resets prevent the destination default style from leaking through.
    const resets: Record<string, string> = {
      b: '0',
      bCs: '0',
      i: '0',
      iCs: '0',
      caps: '0',
      smallCaps: '0',
      strike: '0',
      dstrike: '0',
      outline: '0',
      shadow: '0',
      emboss: '0',
      imprint: '0',
      vanish: '0',
      webHidden: '0',
      u: 'none',
      highlight: 'none',
      vertAlign: 'baseline',
      color: 'auto',
      sz: '22',
      szCs: '22',
      spacing: '0',
      position: '0',
      w: '100',
      kern: '0',
    };
    for (const [name, value] of Object.entries(resets))
      if (!child(rp, name)) {
        const prop = make(baseFile.xml, name);
        prop.setAttributeNS(W, 'w:val', value);
        rp.appendChild(prop);
      }
    // Resolve common theme font references before crossing into the other package.
    if (
      Array.from(rp.getElementsByTagName('*')).some((e) =>
        Array.from(e.attributes).some((a) => /Theme|^theme/.test(a.localName)),
      )
    )
      throw new Error('所选格式包含无法解析的主题引用，请改用底稿格式。');
    const fonts = child(rp, 'rFonts') ?? make(baseFile.xml, 'rFonts');
    for (const script of ['ascii', 'hAnsi', 'eastAsia', 'cs'])
      if (!fonts.hasAttributeNS(W, script))
        fonts.setAttributeNS(W, `w:${script}`, script === 'eastAsia' ? '宋体' : 'Calibri');
    if (!fonts.parentNode) rp.insertBefore(fonts, rp.firstChild);
    if (!child(pp, 'jc')) {
      const jc = make(baseFile.xml, 'jc');
      jc.setAttributeNS(W, 'w:val', 'left');
      pp.appendChild(jc);
    }
    const spacing = child(pp, 'spacing') ?? make(baseFile.xml, 'spacing');
    for (const [name, v] of Object.entries({
      before: '0',
      after: '0',
      line: '240',
      lineRule: 'auto',
    }))
      if (!spacing.hasAttributeNS(W, name)) spacing.setAttributeNS(W, `w:${name}`, v);
    if (!spacing.parentNode) pp.appendChild(spacing);
    const ind = child(pp, 'ind') ?? make(baseFile.xml, 'ind');
    for (const name of ['left', 'right', 'firstLine'])
      if (
        !ind.hasAttributeNS(W, name) &&
        !(name === 'firstLine' && ind.hasAttributeNS(W, 'hanging'))
      )
        ind.setAttributeNS(W, `w:${name}`, '0');
    if (!ind.parentNode) pp.appendChild(ind);
  }
  // Property cascades can change insertion order; emit Word's canonical order.
  const order = (properties: Element, names: string[]) => {
    for (const el of Array.from(properties.children).sort(
      (a, b) =>
        (names.indexOf(a.localName) < 0 ? 999 : names.indexOf(a.localName)) -
        (names.indexOf(b.localName) < 0 ? 999 : names.indexOf(b.localName)),
    ))
      properties.appendChild(el);
  };
  order(pp, [
    'pStyle',
    'keepNext',
    'keepLines',
    'pageBreakBefore',
    'framePr',
    'widowControl',
    'numPr',
    'suppressLineNumbers',
    'pBdr',
    'shd',
    'tabs',
    'suppressAutoHyphens',
    'kinsoku',
    'wordWrap',
    'overflowPunct',
    'topLinePunct',
    'autoSpaceDE',
    'autoSpaceDN',
    'bidi',
    'adjustRightInd',
    'snapToGrid',
    'spacing',
    'ind',
    'contextualSpacing',
    'mirrorIndents',
    'suppressOverlap',
    'jc',
    'textDirection',
    'textAlignment',
    'textboxTightWrap',
    'outlineLvl',
    'divId',
    'cnfStyle',
    'rPr',
    'sectPr',
    'pPrChange',
  ]);
  order(rp, [
    'rStyle',
    'rFonts',
    'b',
    'bCs',
    'i',
    'iCs',
    'caps',
    'smallCaps',
    'strike',
    'dstrike',
    'outline',
    'shadow',
    'emboss',
    'imprint',
    'noProof',
    'snapToGrid',
    'vanish',
    'webHidden',
    'color',
    'spacing',
    'w',
    'kern',
    'position',
    'sz',
    'szCs',
    'highlight',
    'u',
    'effect',
    'bdr',
    'shd',
    'fitText',
    'vertAlign',
    'rtl',
    'cs',
    'em',
    'lang',
    'eastAsianLayout',
    'specVanish',
    'oMath',
    'rPrChange',
  ]);
  result.appendChild(baseFile.xml.importNode(pp, true));
  const run = make(baseFile.xml, 'r');
  run.appendChild(baseFile.xml.importNode(rp, true));
  // Clone only text/tab/newline content after eligibility has been checked.
  for (const originalRun of Array.from(value.children).filter((e) => e.localName === 'r'))
    for (const node of Array.from(originalRun.children)) {
      if (['t', 'tab', 'br', 'cr'].includes(node.localName))
        run.appendChild(baseFile.xml.importNode(node, true));
    }
  result.appendChild(run);
  return result;
}

export function normalizeBlock(
  source: Block,
  row: Row,
  files: Record<Side, WordFile>,
  base: Side,
  format: Exclude<FormatChoice, 'default'>,
): Element {
  const sourceFile = source === row.left ? files.left : files.right;
  const baseBlock = row[base];
  const explicit = format === 'smart' ? undefined : row[format];
  const templateFile = format === 'smart' ? files[base] : files[format];
  if (source.type === 'tbl') {
    const result = clone(baseBlock!.element);
    const sourcePs = all(source.element, 'p'),
      targetPs = all(result, 'p');
    const templatePs = explicit ? all(explicit.element, 'p') : targetPs;
    if (templatePs.length !== sourcePs.length)
      throw new Error('两侧表格段落结构不一致，无法选取格式。');
    sourcePs.forEach((p, i) =>
      targetPs[i].replaceWith(uniformParagraph(p, templatePs[i], templateFile, files[base], false)),
    );
    return result;
  }
  const template =
    explicit?.element ??
    smartTemplate(
      source.element,
      sourceFile,
      baseBlock?.element,
      files[base],
      baseBlock?.index ?? source.index,
    );
  return uniformParagraph(source.element, template, templateFile, files[base], !baseBlock);
}
