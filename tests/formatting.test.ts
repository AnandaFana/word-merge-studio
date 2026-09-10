import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { readWord, compareWords, mergeWords, type Plan, type WordFile } from '../src/engine';
import { formatIssue, formatSummary } from '../src/formatting';
import { readFileSync, readdirSync } from 'node:fs';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (t: string, pr = '') => `<w:r>${pr ? `<w:rPr>${pr}</w:rPr>` : ''}<w:t>${t}</w:t></w:r>`;
const p = (t: string, pr = '', pp = '') => `<w:p><w:pPr>${pp}</w:pPr>${run(t, pr)}</w:p>`;
const attrs = (f: WordFile, name: string, attr = 'val') =>
  Array.from(f.xml.getElementsByTagNameNS(W, name)).map((e) => e.getAttributeNS(W, attr));
async function word(body: string, styles = '', theme = '') {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
  );
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgMar w:top="900"/></w:sectPr></w:body></w:document>`,
  );
  if (styles) zip.file('word/styles.xml', styles);
  if (theme) zip.file('word/theme/theme1.xml', theme);
  return readWord('sample.docx', await zip.generateAsync({ type: 'uint8array' }));
}
async function execute(left: WordFile, right: WordFile, plan: Plan) {
  return readWord(
    'result.docx',
    await mergeWords({ left, right }, compareWords(left, right), plan),
  );
}
describe('independent paragraph formatting', () => {
  it('keeps right content but uses only the left dominant font, eliminating accidental run mixtures', async () => {
    const left = await word(
      `<w:p>${run('标题', '<w:b/><w:sz w:val="40"/>')}${run('这是一段足够长的正常正文', '<w:rFonts w:eastAsia="宋体"/><w:sz w:val="24"/>')}</w:p>`,
    );
    const right = await word(
      p('右侧最新正文内容', '<w:i/><w:rFonts w:eastAsia="黑体"/><w:sz w:val="32"/>'),
    );
    const out = await execute(left, right, {
      base: 'left',
      choices: { 'row-0': 'right' },
      formats: { 'row-0': 'left' },
    });
    expect(out.blocks[0].text).toBe(right.blocks[0].text);
    expect(attrs(out, 'sz')).toEqual(['24']);
    expect(attrs(out, 'rFonts', 'eastAsia')).toEqual(['宋体']);
    expect(attrs(out, 'b')).toEqual([]);
    expect(attrs(out, 'i')).toEqual([]);
  });
  it('applies format-only decisions even when content is unchanged and defaults to the base', async () => {
    const left = await word(p('相同文字', '<w:sz w:val="24"/>'));
    const right = await word(p('相同文字', '<w:sz w:val="36"/>'));
    const out = await execute(left, right, {
      base: 'left',
      choices: {},
      formats: { 'row-0': 'right' },
    });
    expect(out.blocks[0].text).toBe('相同文字');
    expect(attrs(out, 'sz')).toEqual(['36']);
  });
  it('smart mode also normalizes base text without any content decisions', async () => {
    const left = await word(
      `<w:p>${run('主文本主文本主文本', '<w:sz w:val="24"/>')}${run('异', '<w:sz w:val="40"/>')}</w:p>`,
    );
    const out = await execute(left, left, { base: 'left', choices: {}, formatMode: 'smart' });
    expect(out.blocks[0].text).toBe(left.blocks[0].text);
    expect(attrs(out, 'sz')).toEqual(['24']);
  });
  it('does not use a nearby heading for an inserted body paragraph', async () => {
    const left = await word(
      p('一、标题', '<w:b/><w:sz w:val="40"/>') + p('已有正文', '<w:sz w:val="24"/>'),
    );
    const right = await word(p('一、标题') + p('新加入的正文') + p('已有正文'));
    const rows = compareWords(left, right);
    const insertion = rows.find((r) => !r.left)!;
    const out = await execute(left, right, {
      base: 'left',
      choices: { [insertion.id]: 'right' },
      formats: { [insertion.id]: 'smart' },
    });
    expect(out.blocks[1].text).toBe('新加入的正文');
    expect(out.blocks[1].element.getElementsByTagNameNS(W, 'sz')[0].getAttributeNS(W, 'val')).toBe(
      '24',
    );
    expect(out.blocks[1].element.getElementsByTagNameNS(W, 'b')).toHaveLength(0);
  });
  it('resolves conflicting style IDs, basedOn inheritance and source defaults before applying right formatting', async () => {
    const styles = (face: string, size: string) =>
      `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:eastAsia="${face}"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Parent"><w:rPr><w:sz w:val="${size}"/></w:rPr><w:pPr><w:spacing w:after="200"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Body"><w:basedOn w:val="Parent"/></w:style></w:styles>`;
    const left = await word(p('左文', '', '<w:pStyle w:val="Body"/>'), styles('宋体', '24'));
    const right = await word(p('右文', '', '<w:pStyle w:val="Body"/>'), styles('黑体', '32'));
    const out = await execute(left, right, {
      base: 'left',
      choices: { 'row-0': 'left' },
      formats: { 'row-0': 'right' },
    });
    expect(out.blocks[0].text).toBe('左文');
    expect(attrs(out, 'rFonts', 'eastAsia')).toEqual(['黑体']);
    expect(attrs(out, 'sz')).toEqual(['32']);
    expect(attrs(out, 'pStyle')).toEqual([]);
    expect(attrs(out, 'spacing', 'after')).toContain('200');
    expect(await out.zip.file('word/styles.xml')!.async('string')).toBe(
      await left.zip.file('word/styles.xml')!.async('string'),
    );
  });
  it('materializes theme colors even without a themed font', async () => {
    const theme =
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="123456"/></a:accent1></a:clrScheme></a:themeElements></a:theme>';
    const left = await word(p('左文'));
    const right = await word(p('右文', '<w:color w:themeColor="accent1"/>'), '', theme);
    const out = await execute(left, right, {
      base: 'left',
      choices: {},
      formats: { 'row-0': 'right' },
    });
    expect(attrs(out, 'color')).toEqual(['123456']);
    expect(attrs(out, 'color', 'themeColor')).toEqual([null]);
  });
  it('resolves source theme east-Asian fonts before crossing packages', async () => {
    const theme =
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:font script="Hans" typeface="黑体"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>';
    const left = await word(p('左文'));
    const right = await word(
      p('右文', '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"/>'),
      '',
      theme,
    );
    const out = await execute(left, right, {
      base: 'left',
      choices: {},
      formats: { 'row-0': 'right' },
    });
    expect(attrs(out, 'rFonts', 'eastAsia')).toEqual(['黑体']);
    expect(attrs(out, 'rFonts', 'ascii')).toEqual(['Arial']);
  });
  it('does not pull formatting from the final bold run into a plain run when inserting text', async () => {
    const left = await word(`<w:p>${run('正常正文')}${run('粗体', '<w:b/>')}</w:p>`);
    const right = await word(p('正常新增正文粗体'));
    const out = await execute(left, right, { base: 'left', choices: { 'row-0': 'right' } });
    const inserted = Array.from(out.xml.getElementsByTagNameNS(W, 'r')).find(
      (e) => e.textContent === '新增',
    )!;
    expect(inserted.getElementsByTagNameNS(W, 'b')).toHaveLength(0);
  });
  it('explicit left/right choices override global smart mode and work for both retained contents', async () => {
    const left = await word(p('左侧', '<w:sz w:val="24"/>'));
    const right = await word(p('右侧', '<w:sz w:val="32"/>'));
    const out = await execute(left, right, {
      base: 'left',
      choices: { 'row-0': 'both' },
      formatMode: 'smart',
      formats: { 'row-0': 'right' },
    });
    expect(out.blocks.map((b) => b.text)).toEqual(['左侧', '右侧']);
    expect(attrs(out, 'sz')).toEqual(['32', '32']);
  });
  it('rejects cross-package numbering and protected format edits', async () => {
    const left = await word(p('左侧'));
    const right = await word(p('右侧', '', '<w:numPr><w:numId w:val="3"/></w:numPr>'));
    const row = compareWords(left, right)[0];
    expect(formatIssue(row, 'right', { left, right }, 'left')).toContain('编号');
    await expect(
      execute(left, right, { base: 'left', choices: {}, formats: { 'row-0': 'right' } }),
    ).rejects.toThrow('编号');
    const complex = await word('<w:p><w:r><w:drawing/></w:r></w:p>');
    await expect(
      execute(complex, right, { base: 'left', choices: {}, formats: { 'row-0': 'left' } }),
    ).rejects.toThrow('复杂');
  });
  it('describes a local paragraph format without sending the document anywhere', async () => {
    const file = await word(p('正文', '<w:rFonts w:eastAsia="宋体"/><w:sz w:val="24"/>'));
    expect(formatSummary(file, file.blocks[0])).toBe('正文 · 宋体 · 12pt');
  });
  it('recognizes an unnumbered large title relative to ordinary body text', async () => {
    const file = await word(
      p('项目方案', '<w:b/><w:sz w:val="40"/>') +
        p('这是比较长的普通正文内容，用来识别主体字号。', '<w:sz w:val="24"/>'),
    );
    expect(formatSummary(file, file.blocks[0])).toContain('标题');
    expect(formatSummary(file, file.blocks[1])).toContain('正文');
  });
  it('normalizes table text with foreign formatting while retaining base table geometry and shading', async () => {
    const table = (value: string, size: string, fill: string) =>
      `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tr><w:tc><w:tcPr><w:shd w:fill="${fill}"/></w:tcPr>${p(value, `<w:sz w:val="${size}"/>`)}</w:tc></w:tr></w:tbl>`;
    const left = await word(table('左侧数据', '24', 'AAAAAA'));
    const right = await word(table('右侧数据', '32', 'FFFFFF'));
    const out = await execute(left, right, {
      base: 'left',
      choices: { 'row-0': 'left' },
      formats: { 'row-0': 'right' },
    });
    expect(out.blocks[0].text).toContain('左侧数据');
    expect(attrs(out, 'sz')).toEqual(['32']);
    expect(attrs(out, 'shd', 'fill')).toEqual(['AAAAAA']);
    expect(attrs(out, 'tblW', 'w')).toEqual(['9000']);
  });
  const samples = readdirSync('.').filter((n) => n.endsWith('.docx'));
  it.skipIf(samples.length !== 2)(
    'normalizes actual user documents in both directions without changing selected text or non-body parts',
    async () => {
      const files = {
        left: await readWord(samples[0], new Uint8Array(readFileSync(samples[0]))),
        right: await readWord(samples[1], new Uint8Array(readFileSync(samples[1]))),
      };
      const rows = compareWords(files.left, files.right);
      for (const base of ['left', 'right'] as const)
        for (const format of ['left', 'right', 'smart'] as const) {
          const other = base === 'left' ? 'right' : 'left';
          const out = await execute(files.left, files.right, {
            base,
            choices: Object.fromEntries(rows.map((r) => [r.id, other])),
            formats: Object.fromEntries(
              rows.map((r) => [r.id, format === 'smart' || !r[format] ? 'smart' : format]),
            ),
          });
          expect(out.blocks.map((b) => b.text)).toEqual(files[other].blocks.map((b) => b.text));
          for (const b of out.blocks)
            expect(b.element.getElementsByTagNameNS(W, 'r').length).toBeLessThanOrEqual(1);
          for (const [name, part] of Object.entries(files[base].zip.files))
            if (!part.dir && name !== 'word/document.xml')
              expect(await out.zip.file(name)?.async('uint8array')).toEqual(
                await part.async('uint8array'),
              );
        }
    },
    20000, // Six real-document exports and byte-by-byte resource checks.
  );
});
