import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { readFileSync, readdirSync, existsSync, mkdirSync, writeFileSync } from 'node:fs';
import {
  readWord,
  compareWords,
  mergeWords,
  choiceIssue,
  type Choice,
  type Side,
} from '../src/engine';
import { createDemo } from '../src/demo';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const p = (text: string, props = '<w:b/><w:color w:val="AA0033"/>') =>
  `<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:rPr>${props}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
async function file(body: string, name = 'test.docx') {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
  );
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}<w:sectPr><w:pgMar w:top="1000"/></w:sectPr></w:body></w:document>`,
  );
  zip.file('word/styles.xml', '<styles>baseline-styles</styles>');
  zip.file('word/header1.xml', '<header>base-header</header>');
  zip.file('word/media/image1.png', new Uint8Array([1, 2, 3]));
  return readWord(name, await zip.generateAsync({ type: 'uint8array' }));
}
async function merge(a: string, b: string, base: Side = 'left', choice: Choice = 'right') {
  const files = { left: await file(a), right: await file(b) };
  const rows = compareWords(files.left, files.right);
  const bytes = await mergeWords(files, rows, {
    base,
    choices: Object.fromEntries(rows.map((r) => [r.id, choice])),
  });
  return { files, rows, bytes, output: await readWord('output.docx', bytes) };
}
const table = (text: string) =>
  `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:shd w:fill="AAAAAA"/></w:tcPr>${p(text)}</w:tc></w:tr></w:tbl>`;

describe('OOXML merge contract', () => {
  it('returns byte-identical base when there are no decisions, for either base', async () => {
    const files = { left: await file(p('左边')), right: await file(p('右边')) };
    const rows = compareWords(files.left, files.right);
    for (const base of ['left', 'right'] as Side[])
      expect(await mergeWords(files, rows, { base, choices: {} })).toEqual(files[base].bytes);
  });
  it('changes Chinese text while retaining base run and paragraph formatting', async () => {
    const { output } = await merge(p('这是旧版本'), p('这是新版本', '<w:i/>'));
    expect(output.blocks[0].text).toBe('这是新版本');
    expect(output.blocks[0].element.getElementsByTagNameNS(W, 'b').length).toBeGreaterThan(0);
    expect(output.blocks[0].element.getElementsByTagNameNS(W, 'i')).toHaveLength(0);
    expect(
      output.blocks[0].element.getElementsByTagNameNS(W, 'spacing')[0].getAttributeNS(W, 'after'),
    ).toBe('120');
  });
  it('keeps all other package parts byte-identical, including styles and media', async () => {
    const { files, output } = await merge(p('before'), p('after'));
    for (const name of Object.keys(files.left.zip.files).filter((n) => n !== 'word/document.xml')) {
      expect(await output.zip.file(name)?.async('uint8array')).toEqual(
        await files.left.zip.file(name)?.async('uint8array'),
      );
    }
    expect(output.xml.getElementsByTagNameNS(W, 'pgMar')[0].getAttributeNS(W, 'top')).toBe('1000');
  });
  it('supports right-format base while accepting left content', async () => {
    const { output } = await merge(p('左文', '<w:b/>'), p('右文', '<w:i/>'), 'right', 'left');
    expect(output.blocks[0].text).toBe('左文');
    expect(output.blocks[0].element.getElementsByTagNameNS(W, 'i').length).toBeGreaterThan(0);
    expect(output.blocks[0].element.getElementsByTagNameNS(W, 'b')).toHaveLength(0);
  });
  it('retains different base run styles on unchanged characters', async () => {
    const a =
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>标题</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>正文</w:t></w:r></w:p>';
    const { output } = await merge(a, p('标题正文增加'));
    const runs = output.blocks[0].element.getElementsByTagNameNS(W, 'r');
    expect(runs[0].textContent).toBe('标题');
    expect(runs[0].getElementsByTagNameNS(W, 'b')).toHaveLength(1);
    expect(runs[1].textContent).toBe('正文');
    expect(runs[1].getElementsByTagNameNS(W, 'i')).toHaveLength(1);
  });
  it('replacement at a run boundary inherits the replaced run, not the following run', async () => {
    const a =
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>旧标题</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>正文</w:t></w:r></w:p>';
    const { output } = await merge(a, p('新名称正文'));
    const runs = output.blocks[0].element.getElementsByTagNameNS(W, 'r');
    expect(runs[0].textContent).toBe('新名称');
    expect(runs[0].getElementsByTagNameNS(W, 'b')).toHaveLength(1);
    expect(runs[runs.length - 1].getElementsByTagNameNS(W, 'i')).toHaveLength(1);
  });
  it('accepts inserted/deleted paragraphs without losing any other-side text', async () => {
    const { output } = await merge(
      p('开头') + p('删除') + p('结尾'),
      p('新增') + p('开头') + p('结尾') + p('末尾'),
    );
    expect(output.blocks.map((b) => b.text)).toEqual(['新增', '开头', '结尾', '末尾']);
  });
  it('inserts plain paragraphs with the base font rather than foreign font', async () => {
    const { output } = await merge(p('稳定段落'), p('稳定段落') + p('新增段落', '<w:i/>'));
    expect(output.blocks[1].element.getElementsByTagNameNS(W, 'b')).toHaveLength(1);
    expect(output.blocks[1].element.getElementsByTagNameNS(W, 'i')).toHaveLength(0);
  });
  it('supports keeping both in left-then-right order', async () => {
    const { output } = await merge(p('左边'), p('右边'), 'right', 'both');
    expect(output.blocks.map((b) => b.text)).toEqual(['左边', '右边']);
  });
  it('creates a valid empty paragraph when omitting everything', async () => {
    const { output } = await merge(p('左边'), p('右边'), 'left', 'omit');
    expect(output.blocks.map((b) => b.text)).toEqual(['']);
    expect(output.xml.getElementsByTagNameNS(W, 'sectPr')).toHaveLength(1);
  });
  it('merges table text while preserving base cell shading and geometry', async () => {
    const { output } = await merge(table('旧数据'), table('新数据').replace('AAAAAA', 'FFFFFF'));
    expect(output.blocks[0].text).toContain('新数据');
    expect(output.xml.getElementsByTagNameNS(W, 'shd')[0].getAttributeNS(W, 'fill')).toBe('AAAAAA');
    expect(output.xml.getElementsByTagNameNS(W, 'tblW')[0].getAttributeNS(W, 'w')).toBe('9000');
  });
  it('blocks table topology changes', async () => {
    const files = {
      left: await file(table('a')),
      right: await file(table('b').replace('</w:tc>', `${p('extra')}</w:tc>`)),
    };
    const rows = compareWords(files.left, files.right);
    expect(choiceIssue(rows[0], 'right', 'left', files)).toContain('表格');
    await expect(
      mergeWords(files, rows, { base: 'left', choices: { [rows[0].id]: 'right' } }),
    ).rejects.toThrow('表格');
  });
  it('preserves complex base blocks but refuses to flatten drawings', async () => {
    const files = {
      left: await file('<w:p><w:r><w:drawing/></w:r></w:p>'),
      right: await file(p('text')),
    };
    const rows = compareWords(files.left, files.right);
    expect(rows[0].kind).toBe('complex');
    expect(choiceIssue(rows[0], 'right', 'left', files)).toBeTruthy();
    expect(await mergeWords(files, rows, { base: 'left', choices: {} })).toEqual(files.left.bytes);
  });
  it('blocks importing hyperlinks and relationship-bearing runs', async () => {
    const files = {
      left: await file(p('a')),
      right: await file(
        '<w:p><w:hyperlink r:id="rId1"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>',
      ),
    };
    expect(
      choiceIssue(compareWords(files.left, files.right)[0], 'right', 'left', files),
    ).toBeTruthy();
  });
  it('locks cross-side merging for spanning fields and tracked changes', async () => {
    for (const special of [
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
      '<w:ins><w:r><w:t>new</w:t></w:r></w:ins>',
      '<w:bookmarkStart w:id="1" w:name="x"/>',
    ]) {
      const files = {
        left: await file(`<w:p>${special}</w:p>${p('正文')}`),
        right: await file(p('another')),
      };
      expect(files.left.locked).toBe(true);
      expect(
        choiceIssue(compareWords(files.left, files.right).at(-1)!, 'right', 'left', files),
      ).toBeTruthy();
    }
  });
  it('handles literal XML characters, whitespace, tabs, linebreaks and emoji', async () => {
    const { output } = await merge(
      p('旧文字'),
      p(' A &amp; B &lt; C 😀 ').replace(
        '</w:t>',
        '</w:t><w:tab/><w:t>缩进</w:t><w:br/><w:t>下一行</w:t>',
      ),
    );
    expect(output.blocks[0].text).toBe(' A & B < C 😀 \t缩进\n下一行');
  });
  it('distinguishes text-identical formatting changes', async () => {
    const a = await file(p('一样', '<w:b/>')),
      b = await file(p('一样', '<w:i/>'));
    expect(compareWords(a, b)[0].kind).toBe('format');
  });
  it('does not mutate inputs across repeated merges', async () => {
    const files = { left: await file(p('before')), right: await file(p('after')) };
    const rows = compareWords(files.left, files.right);
    const plan = { base: 'left' as const, choices: { [rows[0].id]: 'right' as const } };
    await mergeWords(files, rows, plan);
    await mergeWords(files, rows, plan);
    expect(files.left.blocks[0].text).toBe('before');
    expect(files.left.xml.getElementsByTagNameNS(W, 't')[0].textContent).toBe('before');
  });
  it('rejects non-docx, invalid zip and macro documents', async () => {
    await expect(readWord('a.doc', new Uint8Array())).rejects.toThrow('.doc');
    await expect(readWord('a.docx', new Uint8Array([1, 2]))).rejects.toThrow('无法打开');
    const zip = new JSZip();
    zip.file('word/vbaProject.bin', 'macro');
    await expect(
      readWord('a.docx', await zip.generateAsync({ type: 'uint8array' })),
    ).rejects.toThrow('宏');
  });
  it('runs a full merge on the built-in demonstration', async () => {
    const files = {
      left: await readWord('a.docx', await createDemo('left')),
      right: await readWord('b.docx', await createDemo('right')),
    };
    const rows = compareWords(files.left, files.right);
    const output = await readWord(
      'merged.docx',
      await mergeWords(files, rows, {
        base: 'left',
        choices: Object.fromEntries(rows.map((r) => [r.id, 'right'])),
      }),
    );
    expect(output.blocks.map((b) => b.text)).toEqual(files.right.blocks.map((b) => b.text));
  });
});

describe('local user documents (optional, never bundled into the app)', () => {
  const names = readdirSync('.').filter((n) => n.endsWith('.docx'));
  it.skipIf(names.length !== 2)(
    'merges both directions with complete text and package preservation',
    async () => {
      const files = {
        left: await readWord(names[0], new Uint8Array(readFileSync(names[0]))),
        right: await readWord(names[1], new Uint8Array(readFileSync(names[1]))),
      };
      const rows = compareWords(files.left, files.right);
      const metrics: object[] = [];
      for (const base of ['left', 'right'] as Side[]) {
        const other = base === 'left' ? 'right' : 'left';
        expect(rows.every((r) => !choiceIssue(r, other, base, files))).toBe(true);
        const bytes = await mergeWords(files, rows, {
          base,
          choices: Object.fromEntries(rows.map((r) => [r.id, other])),
        });
        const output = await readWord('result.docx', bytes);
        expect(output.blocks.map((b) => b.text)).toEqual(files[other].blocks.map((b) => b.text));
        for (const name of Object.keys(files[base].zip.files).filter(
          (n) => n !== 'word/document.xml' && !files[base].zip.files[n].dir,
        ))
          expect(await output.zip.file(name)?.async('uint8array')).toEqual(
            await files[base].zip.file(name)?.async('uint8array'),
          );
        metrics.push({
          base,
          inputBlocks: files[base].blocks.length,
          otherBlocks: files[other].blocks.length,
          outputBlocks: output.blocks.length,
          rows: rows.length,
          changes: rows.filter((r) => r.kind !== 'equal').length,
        });
        if (process.env.WRITE_QA === '1') {
          if (!existsSync('.qa')) mkdirSync('.qa');
          writeFileSync(`.qa/merged-${base}.docx`, bytes);
        }
      }
      if (process.env.WRITE_QA === '1')
        writeFileSync('.qa/metrics.json', JSON.stringify(metrics, null, 2));
    },
  );
});
