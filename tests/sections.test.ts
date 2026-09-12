import { describe, expect, it } from 'vitest';
import { dominantRun } from '../src/formatting';
import { createSectionDemo } from '../src/sectionDemo';
import { readWord, textOf } from '../src/engine';
import { all, child, value, createChapter, defaultFormats } from '../src/chapterOutline';
import {
  groupSections,
  mergeSections,
  sectionKey,
  splitSections,
  sectionPlan,
} from '../src/sections';
const makeSources = () =>
  Promise.all(
    ['garden', 'library', 'cafe'].map(async (k) =>
      createChapter(await readWord(k + '.docx', await createSectionDemo(k))),
    ),
  );
describe('section assembly', () => {
  it('matches numbered titles and retains every source block in exactly one card', async () => {
    const sources = await makeSources(),
      groups = groupSections(sources);
    expect(groups).toHaveLength(3);
    expect(groups.map((g) => g.cards.length)).toEqual([3, 3, 3]);
    for (const source of sources) {
      const indices = splitSections(source).flatMap((c) =>
        Array.from({ length: c.end - c.start }, (_, i) => c.start + i),
      );
      expect(indices).toEqual(source.file.blocks.map((_, i) => i));
    }
    expect(sectionKey('一、项目目标')).toBe(sectionKey('第二章 项目目标'));
    expect(sectionKey('Chapter 2: Goals')).toBe(sectionKey('1. Goals'));
  });
  for (const mode of ['original', 'unified'] as const)
    it(`exports reordered text, tables and images with three real levels (${mode})`, async () => {
      const sources = await makeSources(),
        groups = groupSections(sources).reverse();
      groups[0].cards.reverse();
      const output = await readWord(
        'out.docx',
        await mergeSections(sources, groups, { mode, formats: defaultFormats(), pageBreak: true }),
      );
      const text = textOf(output.xml.documentElement);
      expect(text.indexOf('3. Schedule')).toBeLessThan(text.indexOf('1. Overview'));
      expect(text.indexOf('cafe pilot')).toBeLessThan(text.indexOf('garden pilot'));
      expect(all(output.xml, 'tbl')).toHaveLength(3);
      expect(all(output.xml, 'drawing')).toHaveLength(3);
      const levels = all(output.xml, 'p')
        .map((p) => value(child(child(p, 'pPr'), 'outlineLvl')))
        .filter(Boolean);
      expect([...new Set(levels)].sort()).toEqual(['0', '1', '2', '9']);
      expect(levels.filter((l) => l === '0')).toHaveLength(3);
      expect(levels.filter((l) => l === '1')).toHaveLength(9);
      expect(levels.filter((l) => l === '2')).toHaveLength(9);
      const media = Object.values(output.zip.files).filter((f) => f.name.endsWith('.png'));
      expect(media).toHaveLength(3);
      const original = await sources[0].file.zip.file('word/media/stages.png')!.async('uint8array');
      for (const part of media) expect(await part.async('uint8array')).toEqual(original);
      const sourceTexts = sources.flatMap((c) =>
        c.file.blocks
          .filter(
            (b) =>
              b.element.localName !== 'p' ||
              !all(b.element, 'pStyle').some((s) => value(s) === 'Heading1'),
          )
          .map((b) => b.text),
      );
      for (const s of sourceTexts) if (s.trim()) expect(text).toContain(s);
      // All package relationship targets resolve, including relocated image parts.
      for (const part of Object.values(output.zip.files).filter((f) => f.name.endsWith('.rels'))) {
        const rels = new DOMParser().parseFromString(await part.async('string'), 'application/xml');
        for (const rel of Array.from(rels.documentElement.children)) {
          if (rel.getAttribute('TargetMode') === 'External') continue;
          const target = rel.getAttribute('Target')!;
          const base =
            part.name === '_rels/.rels' ? '' : part.name.replace(/_rels\/[^/]+\.rels$/, '');
          const resolved = new URL(target, 'https://local/' + base).pathname.slice(1);
          expect(output.zip.file(resolved), resolved).toBeTruthy();
        }
      }
    });
  it('excludes a card and its image without losing neighbouring cards', async () => {
    const sources = await makeSources(),
      groups = groupSections(sources);
    groups[0].cards[1].included = false;
    const result = await readWord(
      'out.docx',
      await mergeSections(sources, groups, {
        mode: 'original',
        formats: defaultFormats(),
        pageBreak: false,
      }),
    );
    expect(textOf(result.xml.documentElement)).not.toContain('Create a quiet reading space');
    expect(all(result.xml, 'drawing')).toHaveLength(2);
    expect(Object.keys(result.zip.files).filter((n) => n.endsWith('.png'))).toHaveLength(2);
    expect(textOf(result.xml.documentElement)).toContain('library pilot');
  });
  it('preserves preambles and documents with no matching split headings', async () => {
    const sources = await makeSources();
    sources[0].overrides[0] = 0;
    const cards = splitSections(sources[0]);
    expect(cards[0].headingBlock).toBeUndefined();
    expect(cards[0].start).toBe(0);
    const unstructured = splitSections(sources[0], 9);
    expect(unstructured).toHaveLength(1);
    expect(unstructured[0].end).toBe(sources[0].file.blocks.length);
    const result = await mergeSections([sources[0]], groupSections([sources[0]], 9), {
      mode: 'original',
      formats: defaultFormats(),
      pageBreak: false,
    });
    expect(all((await readWord('all.docx', result)).xml, 'drawing')).toHaveLength(1);
  });
  it('rejects duplicates and empty plans and keeps source state immutable', async () => {
    const sources = await makeSources(),
      groups = groupSections(sources),
      before = sources.map((c) => c.overrides);
    sectionPlan(sources, groups);
    expect(sources.map((c) => c.overrides)).toEqual(before);
    groups[0].cards.push(groups[0].cards[0]);
    expect(() => sectionPlan(sources, groups)).toThrow(/重复/);
    expect(() => sectionPlan(sources, [])).toThrow(/至少/);
  });
  it('keeps source typography in preserve mode and applies checked unified types', async () => {
    const sources = await makeSources();
    const options = { pageBreak: false, formats: defaultFormats() };
    const original = await readWord(
      'original.docx',
      await mergeSections(sources, groupSections(sources), { ...options, mode: 'original' }),
    );
    const unified = await readWord(
      'unified.docx',
      await mergeSections(sources, groupSections(sources), { ...options, mode: 'unified' }),
    );
    const target = (file: typeof original) =>
      all(file.xml, 'p').find((p) => textOf(p).startsWith('Create a quiet'))!;
    expect(value(child(child(target(original), 'pPr'), 'pStyle'))).toBe('WM1_Normal');
    expect(value(child(child(target(unified), 'pPr'), 'pStyle'))).toBe('WMUnified0');
    expect(
      child(dominantRun(original, target(original)), 'rFonts')?.getAttributeNS(
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'ascii',
      ),
    ).toBe('Arial');
  });
});
