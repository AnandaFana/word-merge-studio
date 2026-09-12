import { t } from './i18n';
import { useEffect, useRef, useState } from 'react';
import { ArrowDown, ArrowUp, Download, Eye, Plus, RotateCcw, X } from 'lucide-react';
import { readWord } from './engine';
import { createChapter, defaultFormats, type Chapter, type TypeFormats } from './chapterOutline';
import { chapterIssues } from './chapters';
import { groupSections, mergeSections, sectionReport, type SectionGroup } from './sections';
import { createSectionDemo } from './sectionDemo';
import Preview from './Preview';
import './sections.css';

function save(bytes: Uint8Array, name: string, type: string, retain = false) {
  const url = URL.createObjectURL(new Blob([bytes.slice().buffer], { type }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  if (!retain) setTimeout(() => URL.revokeObjectURL(url), 30000);
  return url;
}

export default function SectionWorkspace() {
  const [chapters, setChapters] = useState<Chapter[]>([]);
  const [groups, setGroups] = useState<SectionGroup[]>([]);
  const [history, setHistory] = useState<SectionGroup[][]>([]);
  const [level, setLevel] = useState(1);
  const [match, setMatch] = useState<'title' | 'position'>('title');
  const [mode, setMode] = useState<'original' | 'unified'>('unified');
  const [formats, setFormats] = useState<TypeFormats>(() => defaultFormats());
  const [pageBreak, setPageBreak] = useState(false);
  const [name, setName] = useState('section-merge');
  const [busy, setBusy] = useState(false),
    [error, setError] = useState('');
  const [preview, setPreview] = useState<Uint8Array>();
  const [saved, setSaved] = useState<{ url: string; name: string }>();
  const input = useRef<HTMLInputElement>(null);
  const cards = groups.flatMap((g) => g.cards);
  const included = cards.filter((c) => c.included).length;
  const options = { mode, formats, pageBreak };
  function commit(next: SectionGroup[]) {
    setHistory((h) => [...h.slice(-29), groups]);
    setGroups(next);
    setPreview(undefined);
  }
  useEffect(() => {
    setPreview(undefined);
    setSaved(undefined);
  }, [groups, mode, formats, pageBreak, name]);
  useEffect(
    () => () => {
      if (saved) URL.revokeObjectURL(saved.url);
    },
    [saved],
  );
  useEffect(() => {
    if (!preview) return;
    const previous = document.activeElement as HTMLElement | null;
    const overflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    const listener = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setPreview(undefined);
    };
    window.addEventListener('keydown', listener);
    return () => {
      document.body.style.overflow = overflow;
      window.removeEventListener('keydown', listener);
      previous?.focus();
    };
  }, [preview]);
  async function add(files: File[]) {
    if (busy) return;
    setBusy(true);
    setError('');
    const next = [...chapters],
      errors: string[] = [];
    try {
      for (const file of files) {
        try {
          if (
            next.length >= 20 ||
            next.reduce((n, c) => n + c.file.bytes.length, 0) + file.size > 100 * 1024 * 1024
          )
            throw new Error(t('最多 20 份文档，总大小 100 MB。'));
          if (file.size > 25 * 1024 * 1024) throw new Error(t('单个文件不能超过 25 MB。'));
          const chapter = createChapter(await readWord(file.name, await file.arrayBuffer()));
          const issues = chapterIssues(chapter);
          if (issues.length) throw new Error(t('请先在 Word 副本中处理：{0}', issues.join('、')));
          next.push(chapter);
        } catch (e) {
          errors.push(`${file.name}: ${(e as Error).message}`);
        }
      }
      const fresh = groupSections(next.slice(chapters.length), level, match);
      // Adding files never discards manual edits or excluded cards.
      setChapters(next);
      commit([...groups, ...fresh]);
      setError(errors.join('\n'));
    } finally {
      setBusy(false);
    }
  }
  async function demo() {
    setBusy(true);
    setError('');
    try {
      const demoChapters = await Promise.all(
        ['garden', 'library', 'cafe'].map(async (key) =>
          createChapter(await readWord(`${key}.docx`, await createSectionDemo(key))),
        ),
      );
      setChapters(demoChapters);
      setGroups(groupSections(demoChapters));
      setHistory([]);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  }
  function moveGroup(i: number, delta: number) {
    const next = [...groups];
    [next[i], next[i + delta]] = [next[i + delta], next[i]];
    commit(next);
  }
  function moveCard(groupId: string, cardId: string, delta: number) {
    commit(
      groups.map((g) => {
        if (g.id !== groupId) return g;
        const next = [...g.cards],
          i = next.findIndex((c) => c.id === cardId);
        [next[i], next[i + delta]] = [next[i + delta], next[i]];
        return { ...g, cards: next };
      }),
    );
  }
  function reassign(from: string, cardId: string, to: string) {
    const card = groups.find((g) => g.id === from)!.cards.find((c) => c.id === cardId)!;
    commit(
      groups.map((g) =>
        g.id === from
          ? { ...g, cards: g.cards.filter((c) => c.id !== cardId) }
          : g.id === to
            ? { ...g, cards: [...g.cards, card] }
            : g,
      ),
    );
  }
  async function exportDoc(show: boolean) {
    setBusy(true);
    setError('');
    try {
      const bytes = await mergeSections(chapters, groups, options);
      if (show) setPreview(bytes);
      else {
        const filename = `${name.trim() || 'section-merge'}.docx`;
        setSaved({
          url: save(
            bytes,
            filename,
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            true,
          ),
          name: filename,
        });
      }
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  }
  return (
    <main className="section-main">
      <div className="page-heading">
        <div>
          <div className="eyebrow">SECTION STUDIO</div>
          <h1>
            {t('对应章节，汇成一篇')}
            <span>{t('。')}</span>
          </h1>
          <p>{t('把相似结构的报告按章节归组，用卡片安排内容，再输出可导航的 Word。')}</p>
        </div>
        <button className="button" disabled={busy} onClick={demo}>
          {t('体验三个示例')}
        </button>
      </div>
      <div
        className="section-upload"
        onDragOver={(e) => e.preventDefault()}
        onDrop={(e) => {
          e.preventDefault();
          void add([...e.dataTransfer.files]);
        }}
      >
        <button className="button primary" disabled={busy} onClick={() => input.current?.click()}>
          <Plus size={16} />
          {t('添加 Word 文档')}
        </button>
        <span>{t('支持拖入多份 .docx · 仅在浏览器处理')}</span>
        <input
          ref={input}
          type="file"
          accept=".docx"
          multiple
          hidden
          onChange={(e) => {
            void add([...(e.target.files ?? [])]);
            e.target.value = '';
          }}
        />
      </div>
      <fieldset disabled={busy} className="section-controls">
        <label>
          {t('拆分层级')}
          <select value={level} onChange={(e) => setLevel(Number(e.target.value))}>
            {[1, 2, 3].map((l) => (
              <option key={l} value={l}>
                {l}
                {t('级标题')}
              </option>
            ))}
          </select>
        </label>
        <label>
          {t('配对依据')}
          <select value={match} onChange={(e) => setMatch(e.target.value as typeof match)}>
            <option value="title">{t('标题相同（忽略编号）')}</option>
            <option value="position">{t('章节顺序（需人工核对）')}</option>
          </select>
        </label>
        <button
          className="button"
          disabled={!chapters.length}
          onClick={() => commit(groupSections(chapters, level, match))}
        >
          {t('重新识别与配对')}
        </button>
        <button
          className="button"
          onClick={() =>
            commit([...groups, { id: crypto.randomUUID(), title: 'New section', cards: [] }])
          }
        >
          <Plus size={15} />
          {t('新建分组')}
        </button>
        <button
          className="button"
          disabled={!history.length}
          onClick={() => {
            setGroups(history.at(-1)!);
            setHistory((h) => h.slice(0, -1));
          }}
        >
          <RotateCcw size={15} />
          {t('撤销布局')}
        </button>
        <button
          className="button"
          disabled={!chapters.length}
          onClick={() => {
            setChapters([]);
            setGroups([]);
            setHistory([]);
            setError('');
          }}
        >
          {t('清空工作区')}
        </button>
      </fieldset>
      <p className="section-hint">
        {t(
          '重新识别会重建卡片布局，可撤销。新增文档先单独分组；点击重新配对可与已有文档对齐。无标题内容也会保留，请检查识别结果。',
        )}
      </p>
      {error && (
        <div className="error" role="alert">
          {error}
        </div>
      )}
      {busy && (
        <div className="notice" role="status">
          {t('正在处理文档…')}
        </div>
      )}
      {saved && (
        <div className="download-ready" role="status">
          <span>{t('合并文件已就绪')}</span>
          <a href={saved.url} download={saved.name}>
            {t('保存')} {saved.name}
          </a>
        </div>
      )}
      <div className="section-layout">
        <div className="section-board">
          <div className="section-summary">
            {chapters.length}
            {t('份来源 ·')}
            {groups.length}
            {t('个分组 ·')}
            {included}/{cards.length}
            {t('张卡片保留')}
          </div>
          {!groups.length && (
            <div className="section-empty">
              <h2>{t('从整篇排列，到按章节组织')}</h2>
              <p>{t('例如：报告 A 的“目标” + 报告 B 的“目标” → 同一个“目标”章节。')}</p>
              <p>{t('先加载示例，体验移动、归组和格式选择。')}</p>
            </div>
          )}
          {groups.map((group, gi) => (
            <section className="section-group" key={group.id}>
              <div className="section-group-head">
                <span className="section-number">{String(gi + 1).padStart(2, '0')}</span>
                <input
                  aria-label={t('分组标题')}
                  value={group.title}
                  disabled={busy}
                  onChange={(e) =>
                    commit(
                      groups.map((g) => (g.id === group.id ? { ...g, title: e.target.value } : g)),
                    )
                  }
                />
                <button
                  className="icon-button"
                  aria-label={t('上移分组')}
                  disabled={busy || gi === 0}
                  onClick={() => moveGroup(gi, -1)}
                >
                  <ArrowUp size={17} />
                </button>
                <button
                  className="icon-button"
                  aria-label={t('下移分组')}
                  disabled={busy || gi === groups.length - 1}
                  onClick={() => moveGroup(gi, 1)}
                >
                  <ArrowDown size={17} />
                </button>
                {!group.cards.length && (
                  <button
                    className="icon-button"
                    aria-label={t('删除空分组')}
                    onClick={() => commit(groups.filter((g) => g.id !== group.id))}
                  >
                    <X size={16} />
                  </button>
                )}
              </div>
              <div className="section-card-list">
                {group.cards.map((card, ci) => (
                  <article
                    className={`section-card ${card.included ? '' : 'excluded'}`}
                    key={card.id}
                  >
                    <div className="section-card-top">
                      <label>
                        <input
                          type="checkbox"
                          aria-label={t('保留此卡片')}
                          checked={card.included}
                          disabled={busy}
                          onChange={(e) =>
                            commit(
                              groups.map((g) => ({
                                ...g,
                                cards: g.cards.map((c) =>
                                  c.id === card.id ? { ...c, included: e.target.checked } : c,
                                ),
                              })),
                            )
                          }
                        />
                        {t('保留')}
                      </label>
                      <small>
                        {card.end - card.start}
                        {t('个内容块')}
                      </small>
                      <div className="spacer" />
                      <button
                        className="icon-button"
                        aria-label={t('上移卡片')}
                        disabled={busy || ci === 0}
                        onClick={() => moveCard(group.id, card.id, -1)}
                      >
                        <ArrowUp size={15} />
                      </button>
                      <button
                        className="icon-button"
                        aria-label={t('下移卡片')}
                        disabled={busy || ci === group.cards.length - 1}
                        onClick={() => moveCard(group.id, card.id, 1)}
                      >
                        <ArrowDown size={15} />
                      </button>
                    </div>
                    <input
                      className="section-source"
                      aria-label={t('来源小标题')}
                      value={card.source}
                      disabled={busy}
                      onChange={(e) =>
                        commit(
                          groups.map((g) => ({
                            ...g,
                            cards: g.cards.map((c) =>
                              c.id === card.id ? { ...c, source: e.target.value } : c,
                            ),
                          })),
                        )
                      }
                    />
                    <div className="section-original">{card.title}</div>
                    <details>
                      <summary>{t('查看文字摘录')}</summary>
                      <p className="section-excerpt">{card.excerpt || t('（图片或空白内容）')}</p>
                    </details>
                    <label className="section-destination">
                      {t('移动到')}
                      <select
                        aria-label={t('目标分组')}
                        disabled={busy}
                        value={group.id}
                        onChange={(e) => reassign(group.id, card.id, e.target.value)}
                      >
                        {groups.map((g) => (
                          <option key={g.id} value={g.id}>
                            {g.title}
                          </option>
                        ))}
                      </select>
                    </label>
                  </article>
                ))}
              </div>
            </section>
          ))}
        </div>
        <aside className="section-export">
          <h2>{t('输出设置')}</h2>
          <fieldset disabled={busy}>
            <label>
              {t('格式模式')}
              <select value={mode} onChange={(e) => setMode(e.target.value as typeof mode)}>
                <option value="unified">{t('按类型统一格式')}</option>
                <option value="original">{t('保留来源正文格式')}</option>
              </select>
            </label>
            <p>
              {t(
                '大章节 → 来源小标题 → 原有子标题，生成三级导航。更深标题转为正文；原章节标题由分组标题替代。',
              )}
            </p>
            {mode === 'unified' && (
              <>
                <label>
                  {t('排版预设')}
                  <select
                    defaultValue="report"
                    onChange={(e) =>
                      setFormats(defaultFormats(e.target.value as 'report' | 'simple'))
                    }
                  >
                    <option value="report">{t('中文报告')}</option>
                    <option value="simple">{t('简洁方案')}</option>
                  </select>
                </label>
                {[0, 1, 2, 3, 10, 11].map((l) => (
                  <div className="section-format" key={l}>
                    <label>
                      <input
                        type="checkbox"
                        checked={formats[l].enabled}
                        onChange={(e) =>
                          setFormats({
                            ...formats,
                            [l]: { ...formats[l], enabled: e.target.checked },
                          })
                        }
                      />
                      {l === 0
                        ? t('正文')
                        : l === 10
                          ? t('题注')
                          : l === 11
                            ? t('表格文字')
                            : t('{0} 级标题', l)}
                    </label>
                    <input
                      aria-label={t('{0} 字体', l)}
                      value={formats[l].font}
                      onChange={(e) =>
                        setFormats({ ...formats, [l]: { ...formats[l], font: e.target.value } })
                      }
                    />
                    <input
                      aria-label={t('{0} 字号', l)}
                      type="number"
                      min={5}
                      max={72}
                      step={0.5}
                      value={formats[l].size}
                      onChange={(e) =>
                        setFormats({
                          ...formats,
                          [l]: { ...formats[l], size: Number(e.target.value) },
                        })
                      }
                    />
                  </div>
                ))}
              </>
            )}
            <label className="check-label">
              <input
                type="checkbox"
                checked={pageBreak}
                onChange={(e) => setPageBreak(e.target.checked)}
              />
              {t('每个大章节另起一页')}
            </label>
            <label>
              {t('输出文件名')}
              <input value={name} onChange={(e) => setName(e.target.value)} />
            </label>
            <button className="button" disabled={!included} onClick={() => exportDoc(true)}>
              <Eye size={16} />
              {t('预览结果')}
            </button>
            <button
              className="button primary"
              disabled={!included}
              onClick={() => exportDoc(false)}
            >
              <Download size={16} />
              {t('导出 Word')}
            </button>
            <button
              className="text-button"
              disabled={!chapters.length}
              onClick={() =>
                save(
                  new TextEncoder().encode(
                    JSON.stringify(sectionReport(chapters, groups, options), null, 2),
                  ),
                  'section-layout.json',
                  'application/json',
                )
              }
            >
              {t('下载布局记录')}
            </button>
          </fieldset>
          <p className="section-hint">
            {t(
              '页面、页眉页脚采用第一个保留的来源文档。图片与表格随内容迁移；题注编号和交叉引用需人工核对。浏览器预览分页可能与 Word 不同。',
            )}
          </p>
          <p className="section-hint">
            {t(
              '本地规则识别，不自动润色或删重。刷新页面会清空工作区；布局记录用于追踪，不包含正文，也不能恢复项目。',
            )}
          </p>
        </aside>
      </div>
      {preview && (
        <div className="modal-backdrop">
          <div
            className="modal preview-modal"
            role="dialog"
            aria-modal="true"
            aria-label={t('章节融合预览')}
          >
            <div className="modal-header">
              <h2>{t('章节融合预览')}</h2>
              <button
                autoFocus
                className="icon-button"
                aria-label={t('关闭预览')}
                onClick={() => setPreview(undefined)}
              >
                <X />
              </button>
            </div>
            <Preview bytes={preview} title={t('章节融合预览')} />
          </div>
        </div>
      )}
    </main>
  );
}
