import { t } from './i18n';
import { levelOf, paragraphs, type Chapter, type ChapterOptions } from './chapterOutline';
import { mergeChapters } from './chapters';

export interface SectionCard {
  id: string;
  chapterId: string;
  title: string;
  source: string;
  ordinal: number;
  start: number;
  end: number;
  headingBlock?: number;
  included: boolean;
  excerpt: string;
}
export interface SectionGroup {
  id: string;
  title: string;
  cards: SectionCard[];
}

/** Split only top-level body blocks. Tables and their internal paragraphs remain atomic. */
export function splitSections(chapter: Chapter, splitLevel = 1): SectionCard[] {
  const ps = paragraphs(chapter.file);
  const index = new Map(ps.map((p, i) => [p, i]));
  const roots = chapter.file.blocks.flatMap((block, bi) => {
    const pi = index.get(block.element);
    const item = pi === undefined ? undefined : chapter.outline[pi];
    return item && !item.inTable && item.text && levelOf(chapter, item) === splitLevel
      ? [{ bi, title: item.text }]
      : [];
  });
  const boundaries: { bi: number; title: string; heading?: number }[] = roots.map((r) => ({
    ...r,
    heading: r.bi,
  }));
  if (!roots.length || roots[0].bi > 0)
    boundaries.unshift({ bi: 0, title: 'Front matter / 前置内容' });
  return boundaries
    .map((b, i) => ({
      id: `${chapter.id}:${b.bi}`,
      chapterId: chapter.id,
      title: b.title,
      source: chapter.file.name.replace(/\.docx$/i, ''),
      ordinal: b.heading === undefined ? -1 : roots.findIndex((r) => r.bi === b.bi),
      start: b.bi,
      end: boundaries[i + 1]?.bi ?? chapter.file.blocks.length,
      headingBlock: b.heading,
      included: true,
      excerpt: chapter.file.blocks
        .slice(b.bi + (b.heading === undefined ? 0 : 1), boundaries[i + 1]?.bi)
        .map((block) => block.text)
        .filter(Boolean)
        .join('\n')
        .slice(0, 600),
    }))
    .filter((card) => card.end > card.start);
}

export function sectionKey(title: string): string {
  return title
    .normalize('NFKC')
    .replace(/^第[零〇一二三四五六七八九十百\d]+[章节篇部]\s*/, '')
    .replace(/^(?:chapter|section)\s+\d+[\s.:、-]*/i, '')
    .replace(/^(?:\d+(?:\.\d+)*|[一二三四五六七八九十百]+)[\s.、:：)）-]+/, '')
    .replace(/[\s\p{P}]/gu, '')
    .toLocaleLowerCase();
}

/** Exact normalized titles by default; positional matching is an explicit user choice. */
export function groupSections(
  chapters: Chapter[],
  splitLevel = 1,
  by: 'title' | 'position' = 'title',
): SectionGroup[] {
  const groups = new Map<string, SectionGroup>();
  for (const chapter of chapters)
    for (const card of splitSections(chapter, splitLevel)) {
      // Unstructured documents/front matter are never silently matched with chapter one.
      const key =
        card.ordinal < 0
          ? `front:${card.id}`
          : by === 'title'
            ? sectionKey(card.title)
            : `position:${card.ordinal}`;
      if (!groups.has(key))
        groups.set(key, { id: crypto.randomUUID(), title: card.title, cards: [] });
      groups.get(key)!.cards.push(card);
    }
  return [...groups.values()];
}

export function sectionPlan(chapters: Chapter[], groups: SectionGroup[]) {
  const assembly: NonNullable<ChapterOptions['assembly']> = [];
  const active = new Set(
    groups.flatMap((g) => g.cards.filter((c) => c.included).map((c) => c.chapterId)),
  );
  // Keep source order, so page/header baseline is stable when users reorder groups.
  const sources = chapters
    .filter((c) => active.has(c.id))
    .map((c) => ({ ...c, overrides: { ...c.overrides } }));
  const seen = new Set<string>();
  for (const group of groups) {
    const cards = group.cards.filter((c) => c.included);
    if (!cards.length) continue;
    assembly.push({ title: group.title, level: 1 });
    for (const card of cards) {
      const chapter = sources.find((c) => c.id === card.chapterId);
      if (
        !chapter ||
        seen.has(card.id) ||
        card.start < 0 ||
        card.end > chapter.file.blocks.length ||
        card.start >= card.end
      )
        throw new Error(t('章节卡片无效或重复。'));
      seen.add(card.id);
      assembly.push({ title: card.source, level: 2 });
      const ps = paragraphs(chapter.file);
      const byParagraph = new Map(ps.map((p, i) => [p, i]));
      const includedPs = new Set<Element>();
      for (let bi = card.start; bi < card.end; bi++) {
        if (bi === card.headingBlock) continue; // Replaced by the group/source navigation headings.
        assembly.push({ chapterId: chapter.id, blockIndex: bi });
        const block = chapter.file.blocks[bi].element;
        if (block.localName === 'p') includedPs.add(block);
        for (const p of Array.from(block.getElementsByTagNameNS(block.namespaceURI, 'p')))
          includedPs.add(p);
      }
      const levels = [...includedPs]
        .map((p) => byParagraph.get(p))
        .filter((i): i is number => i !== undefined)
        .map((i) => levelOf(chapter, chapter.outline[i]))
        .filter((l) => l >= 1 && l <= 9);
      const top = Math.min(...levels);
      for (const p of includedPs) {
        const i = byParagraph.get(p);
        if (i === undefined) continue;
        const l = levelOf(chapter, chapter.outline[i]);
        if (l >= 1 && l <= 9) chapter.overrides[i] = l === top ? 3 : 0;
      }
    }
  }
  if (!assembly.length) throw new Error(t('请至少保留一个章节卡片。'));
  return { sources, assembly };
}

export async function mergeSections(
  chapters: Chapter[],
  groups: SectionGroup[],
  options: ChapterOptions,
) {
  const { sources, assembly } = sectionPlan(chapters, groups);
  return mergeChapters(sources, { ...options, assembly });
}

export function sectionReport(
  chapters: Chapter[],
  groups: SectionGroup[],
  options: ChapterOptions,
) {
  return {
    version: 1,
    mode: options.mode,
    pageBreak: options.pageBreak,
    formats: options.formats,
    sources: chapters.map((c) => ({ id: c.id, name: c.file.name, blocks: c.file.blocks.length })),
    groups: groups.map((g) => ({
      title: g.title,
      cards: g.cards.map(({ excerpt: _, ...card }) => card),
    })),
    note: 'Layout audit only, not a reloadable project. No body text; source names and headings are included.',
  };
}
