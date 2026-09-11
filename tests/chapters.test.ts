import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { readFileSync, mkdirSync, writeFileSync } from 'node:fs';
import { readWord, textOf } from '../src/engine';
import { createDemo } from '../src/demo';
import { mergeChapters, chapterIssues } from '../src/chapters';
import {
  W,
  all,
  child,
  value,
  createChapter,
  defaultFormats,
  type ChapterOptions,
} from '../src/chapterOutline';
import { chapterSamples } from './localSamples';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const parse = (s: string) => new DOMParser().parseFromString(s, 'application/xml');
const p = (text: string, props = '', run = '') =>
  `<w:p><w:pPr>${props}</w:pPr><w:r><w:rPr>${run}</w:rPr><w:t>${text}</w:t></w:r></w:p>`;
const options = (mode: 'original' | 'unified' = 'original'): ChapterOptions => ({
  mode,
  pageBreak: false,
  formats: defaultFormats(),
});
async function fixture(body: string, font = '宋体', image = 1) {
  const zip = await JSZip.loadAsync(await createDemo('left'));
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`,
  );
  zip.file(
    'word/styles.xml',
    `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:eastAsia="${font}"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="H2"><w:name w:val="标题 2"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style></w:styles>`,
  );
  zip.file(
    'word/numbering.xml',
    `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
  );
  zip.file(
    'word/_rels/document.xml.rels',
    `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/image" Target="media/image.png"/><Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/><Relationship Id="numbering" Type="${R}/numbering" Target="numbering.xml"/><Relationship Id="link" Type="${R}/hyperlink" Target="https://example.com/${image}" TargetMode="External"/></Relationships>`,
  );
  zip.file('word/media/image.png', new Uint8Array([image, 2, 3, 4]));
  return createChapter(
    await readWord(`chapter-${image}.docx`, await zip.generateAsync({ type: 'uint8array' })),
  );
}
async function output(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes);
  return {
    zip,
    doc: parse(await zip.file('word/document.xml')!.async('string')),
    styles: parse(await zip.file('word/styles.xml')!.async('string')),
    rels: parse(await zip.file('word/_rels/document.xml.rels')!.async('string')),
  };
}
async function assertPackageLinks(zip: JSZip) {
  for (const [name, part] of Object.entries(zip.files)) {
    if (part.dir || !name.endsWith('.rels')) continue;
    const source = name === '_rels/.rels' ? '' : name.replace(/_rels\/([^/]+)\.rels$/, '$1');
    const dir = source.slice(0, source.lastIndexOf('/') + 1);
    for (const r of Array.from(parse(await part.async('string')).documentElement.children)) {
      if (r.getAttribute('TargetMode') === 'External') continue;
      const target = r.getAttribute('Target')!;
      const path = target.startsWith('/') ? target.slice(1) : dir + target;
      expect(zip.file(path), `${name} -> ${path}`).not.toBeNull();
    }
  }
}
describe('chapter outline and package composition', () => {
  it('recognizes existing styles, Chinese and decimal headings, excluding captions and table cells', async () => {
    const c = await fixture(
      p('一、项目说明') +
        p('1.1 背景') +
        p('1.1.1 范围') +
        p('已有标题', '<w:pStyle w:val="H2"/>') +
        p('图 1-1 总体架构') +
        p('2026 年的计划是完成工作。') +
        `<w:tbl><w:tr><w:tc>${p('1.1 单元格')}</w:tc></w:tr></w:tbl>`,
    );
    expect(c.outline.map((p) => p.level)).toEqual([1, 2, 3, 2, 10, 0, 11]);
  });
  it('isolates conflicting paragraph styles and default fonts across chapters without changing sources', async () => {
    const a = await fixture(p('第一份'), '宋体'),
      b = await fixture(p('第二份'), '微软雅黑', 2);
    const before = a.file.bytes.slice(),
      original = new XMLSerializer().serializeToString(a.file.xml);
    const out = await output(await mergeChapters([a, b], options()));
    expect(all(out.doc, 'p').map((p) => value(child(child(p, 'pPr'), 'pStyle')))).toEqual([
      'WM0_Normal',
      'WM1_Normal',
    ]);
    const style = all(out.styles, 'style');
    expect(
      value(
        child(
          child(
            style.find((s) => value(s, 'styleId') === 'WM1_Defaultparagraph'),
            'rPr',
          ),
          'rFonts',
        ),
        'eastAsia',
      ),
    ).toBe('微软雅黑');
    expect(all(out.styles, 'style').filter((s) => value(s, 'default'))).toHaveLength(0);
    expect(a.file.bytes).toEqual(before);
    expect(new XMLSerializer().serializeToString(a.file.xml)).toBe(original);
    await assertPackageLinks(out.zip);
  });
  it('copies colliding image filenames with byte parity and scopes hyperlink relationships', async () => {
    const content = (name: string) =>
      p(name) +
      `<w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="image"/><a:blip r:embed="rId1"/></wp:inline></w:drawing></w:r><w:hyperlink r:id="link"><w:r><w:t>链接</w:t></w:r></w:hyperlink></w:p>`;
    const a = await fixture(content('A'), '宋体', 1),
      b = await fixture(content('B'), '微软雅黑', 2);
    const out = await output(await mergeChapters([a, b], options()));
    expect(
      await out.zip.file('word/chapters/c0/word/media/image.png')!.async('uint8array'),
    ).toEqual(new Uint8Array([1, 2, 3, 4]));
    expect(
      await out.zip.file('word/chapters/c1/word/media/image.png')!.async('uint8array'),
    ).toEqual(new Uint8Array([2, 2, 3, 4]));
    const drawings = Array.from(out.doc.getElementsByTagNameNS('*', 'docPr')).map((n) =>
      n.getAttribute('id'),
    );
    expect(new Set(drawings).size).toBe(2);
    const links = all(out.doc, 'hyperlink').map((n) => n.getAttributeNS(R, 'id'));
    expect(new Set(links).size).toBe(2);
    await assertPackageLinks(out.zip);
  });
  it('keeps numbered lists independent even when source numIds are identical', async () => {
    const a = await fixture(
      p('条目 A', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'),
    );
    const b = await fixture(
      p('条目 B', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'),
      '微软雅黑',
      2,
    );
    const out = await output(await mergeChapters([a, b], options('unified')));
    expect(all(out.doc, 'numId').map((n) => value(n))).toEqual(['1', '2']);
    const nums = parse(await out.zip.file('word/numbering.xml')!.async('string'));
    expect(all(nums, 'num').map((n) => value(n, 'numId'))).toEqual(['1', '2']);
    expect(all(nums, 'abstractNum').map((n) => value(n, 'abstractNumId'))).toEqual(['0', '1']);
  });
  it('writes real outline levels, honors manual overrides and normalizes chosen types only', async () => {
    const c = await fixture(
      p('一、标题', '', '<w:sz w:val="48"/>') +
        p('普通正文', '', '<w:sz w:val="40"/><w:i/>') +
        p('保留这一段', '', '<w:sz w:val="32"/>'),
    );
    c.overrides = { 0: 2, 1: 0, 2: -1 };
    const plan = options('unified');
    plan.formats[2].enabled = false;
    plan.formats[0].size = 11;
    const out = await output(await mergeChapters([c], plan));
    const ps = all(out.doc, 'p');
    expect(all(ps[0], 'outlineLvl').map((n) => value(n))).toEqual(['1']);
    expect(all(ps[0], 'sz').map((n) => value(n))).toEqual(['48']);
    expect(all(ps[1], 'outlineLvl').map((n) => value(n))).toEqual(['9']);
    expect(all(ps[1], 'sz').every((n) => value(n) === '22')).toBe(true);
    expect(all(ps[1], 'i')).toHaveLength(1);
    expect(all(ps[2], 'outlineLvl')).toHaveLength(0);
    expect(all(ps[2], 'sz').map((n) => value(n))).toEqual(['32']);
    expect(all(out.styles, 'style').some((s) => value(s, 'styleId') === 'WMUnified0')).toBe(true);
  });
  it('supports reorder, optional page breaks and source tables without dropping text', async () => {
    const a = await fixture(p('A')),
      b = await fixture(
        `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tr><w:tc>${p('B')}</w:tc></w:tr></w:tbl>`,
        '宋体',
        2,
      );
    const out = await output(await mergeChapters([b, a], { ...options(), pageBreak: true }));
    expect(all(out.doc, 't').map((n) => n.textContent)).toEqual(['B', 'A']);
    expect(all(out.doc, 'pageBreakBefore').filter((n) => value(n) === '1')).toHaveLength(1);
    expect(value(all(out.doc, 'tblW')[0], 'w')).toBe('9000');
    expect(all(out.doc, 'sectPr')).toHaveLength(1);
  });
  it('remaps bookmark anchors instead of colliding duplicate source names', async () => {
    const b = `<w:p><w:bookmarkStart w:id="0" w:name="mark"/><w:r><w:t>目标</w:t></w:r><w:bookmarkEnd w:id="0"/><w:hyperlink w:anchor="mark"><w:r><w:t>跳转</w:t></w:r></w:hyperlink></w:p>`;
    const out = await output(
      await mergeChapters([await fixture(b), await fixture(b, '宋体', 2)], options()),
    );
    expect(all(out.doc, 'bookmarkStart').map((n) => value(n, 'name'))).toEqual([
      'WM0_mark',
      'WM1_mark',
    ]);
    expect(all(out.doc, 'hyperlink').map((n) => value(n, 'anchor'))).toEqual([
      'WM0_mark',
      'WM1_mark',
    ]);
  });
  it('uses the first chapter page settings and header with nested image relationships', async () => {
    const a = await fixture(p('A'));
    const z = a.file.zip;
    const document = parse(await z.file('word/document.xml')!.async('string'));
    const ref = document.createElementNS(W, 'w:headerReference');
    ref.setAttributeNS(W, 'w:type', 'default');
    ref.setAttributeNS(R, 'r:id', 'header');
    all(document, 'sectPr')[0].prepend(ref);
    z.file('word/document.xml', new XMLSerializer().serializeToString(document));
    const relationships = parse(await z.file('word/_rels/document.xml.rels')!.async('string'));
    const rel = relationships.createElementNS(REL, 'Relationship');
    rel.setAttribute('Id', 'header');
    rel.setAttribute('Type', R + '/header');
    rel.setAttribute('Target', 'header1.xml');
    relationships.documentElement.appendChild(rel);
    z.file('word/_rels/document.xml.rels', new XMLSerializer().serializeToString(relationships));
    z.file(
      'word/header1.xml',
      `<w:hdr xmlns:w="${W}" xmlns:r="${R}"><w:p><w:r><w:t>首章页眉</w:t><w:drawing r:id="pic"/></w:r></w:p></w:hdr>`,
    );
    z.file(
      'word/_rels/header1.xml.rels',
      `<Relationships xmlns="${REL}"><Relationship Id="pic" Type="${R}/image" Target="media/image.png"/></Relationships>`,
    );
    const first = createChapter(
      await readWord('first.docx', await z.generateAsync({ type: 'uint8array' })),
    );
    const out = await output(await mergeChapters([first, await fixture(p('B'))], options()));
    expect(all(out.doc, 'headerReference')).toHaveLength(1);
    expect(await out.zip.file('word/chapters/c0/word/header1.xml')!.async('string')).toContain(
      '首章页眉',
    );
    await assertPackageLinks(out.zip);
  });
  it('preserves inherited italic when replacing the paragraph style in unified mode', async () => {
    const c = await fixture(p('正文', '<w:pStyle w:val="Normal"/>'));
    const normal = all(c.file.styles!, 'style').find((s) => value(s, 'styleId') === 'Normal')!;
    const props = c.file.styles!.createElementNS(W, 'w:rPr');
    props.appendChild(c.file.styles!.createElementNS(W, 'w:i'));
    normal.appendChild(props);
    // Use a new WordFile so its source-style cache represents the edited fixture.
    c.file.zip.file('word/styles.xml', new XMLSerializer().serializeToString(c.file.styles!));
    const updated = createChapter(
      await readWord('italic.docx', await c.file.zip.generateAsync({ type: 'uint8array' })),
    );
    const out = await output(await mergeChapters([updated], options('unified')));
    expect(all(all(out.doc, 'r')[0], 'i')).toHaveLength(1);
  });
  it('rejects unsupported content, missing relationships and invalid format parameters explicitly', async () => {
    const c = await fixture('<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>');
    expect(chapterIssues(c)).toEqual(['脚注']);
    await expect(mergeChapters([c], options())).rejects.toThrow('脚注');
    await expect(mergeChapters([], options())).rejects.toThrow('添加');
    const broken = await fixture(
      '<w:p><w:hyperlink r:id="missing"><w:r><w:t>X</w:t></w:r></w:hyperlink></w:p>',
    );
    await expect(mergeChapters([broken], options())).rejects.toThrow('附件关系');
    const plan = options('unified');
    plan.formats[0].size = 0;
    await expect(mergeChapters([await fixture(p('正文'))], plan)).rejects.toThrow('参数无效');
  });
  it.skipIf(chapterSamples.length !== 3)(
    'composes three private example chapters with text, table, image and relationship parity in both modes',
    async () => {
      const chapters = [];
      for (const path of chapterSamples)
        chapters.push(createChapter(await readWord(path, new Uint8Array(readFileSync(path)))));

      const expected = chapters.flatMap((c) => all(c.file.xml, 't').map((t) => t.textContent));
      mkdirSync('.qa/chapters', { recursive: true });
      for (const mode of ['original', 'unified'] as const) {
        const bytes = await mergeChapters(chapters, { ...options(mode), pageBreak: true });

        const out = await output(bytes);
        expect(all(out.doc, 't').map((t) => t.textContent)).toEqual(expected);
        expect(all(out.doc, 'tbl').length).toBe(
          chapters.reduce((n, c) => n + all(c.file.xml, 'tbl').length, 0),
        );
        expect(all(out.doc, 'drawing').length).toBe(
          chapters.reduce((n, c) => n + all(c.file.xml, 'drawing').length, 0),
        );
        expect(all(out.doc, 'body')[0].textContent).toBeTruthy();
        await assertPackageLinks(out.zip);
        for (let i = 0; i < chapters.length; i++)
          for (const [path, part] of Object.entries(chapters[i].file.zip.files))
            if (path.startsWith('word/media/') && !part.dir)
              expect(
                await out.zip.file(`word/chapters/c${i}/${path}`)!.async('uint8array'),
              ).toEqual(await part.async('uint8array'));
        if (mode === 'unified')
          expect(
            all(out.doc, 'outlineLvl').filter((n) => value(n) === '0').length,
          ).toBeGreaterThanOrEqual(3);
        writeFileSync(`.qa/chapters/${mode}.docx`, bytes);
      }
      writeFileSync(
        '.qa/chapters/outline.json',
        JSON.stringify(
          chapters.map((c) => ({
            name: c.file.name,
            headings: c.outline
              .filter((p) => p.level >= 1 && p.level <= 9)
              .map((p) => ({ text: p.text, level: p.level, reason: p.reason })),
          })),
          null,
          2,
        ),
      );
    },
    30000,
  );
});
