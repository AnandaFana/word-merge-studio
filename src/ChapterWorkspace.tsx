import { t } from './i18n';
import { useEffect, useMemo, useRef, useState } from 'react';
import {
  ArrowDown,
  ArrowUp,
  Download,
  Eye,
  Files,
  FileText,
  Plus,
  Trash2,
  WandSparkles,
  X,
  Search,
  ListTree,
  RotateCcw,
} from 'lucide-react';
import { readWord } from './engine';
import { createDemo } from './demo';
import {
  createChapter,
  defaultFormats,
  levelOf,
  typeLabel,
  type Chapter,
  type TypeFormat,
  type TypeFormats,
} from './chapterOutline';
import { chapterIssues, chapterWarnings, mergeChapters } from './chapters';
import Preview from './Preview';

const types = Array.from({ length: 12 }, (_, i) => i);
export default function ChapterWorkspace() {
  const [chapters, setChapters] = useState<Chapter[]>([]);
  const [mode, setMode] = useState<'original' | 'unified'>('original');
  const [formats, setFormats] = useState<TypeFormats>(() => defaultFormats());
  const [pageBreak, setPageBreak] = useState(true);
  const [busy, setBusy] = useState(''),
    [error, setError] = useState(''),
    [notice, setNotice] = useState('');
  const [query, setQuery] = useState(''),
    [filter, setFilter] = useState('headings'),
    [page, setPage] = useState(0);
  const [active, setActive] = useState('');
  const [name, setName] = useState(t('章节合并'));
  const [preview, setPreview] = useState<Uint8Array>(),
    [saved, setSaved] = useState<{ url: string; name: string }>();
  const input = useRef<HTMLInputElement>(null);
  useEffect(() => {
    if (!preview) return;
    const previous = document.activeElement as HTMLElement | null;
    const overflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    const escape = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setPreview(undefined);
    };
    window.addEventListener('keydown', escape);
    return () => {
      document.body.style.overflow = overflow;
      window.removeEventListener('keydown', escape);
      previous?.focus();
    };
  }, [preview]);
  useEffect(
    () => () => {
      if (saved) URL.revokeObjectURL(saved.url);
    },
    [saved],
  );
  useEffect(() => {
    setPage(0);
  }, [query, filter, chapters]);
  useEffect(() => {
    setSaved(undefined);
    setPreview(undefined);
  }, [chapters, mode, formats, pageBreak]);
  const stats = useMemo(() => {
    const counts: Record<number, number> = {};
    let total = 0;
    for (const c of chapters)
      for (const p of c.outline) {
        const l = levelOf(c, p);
        if (p.text) {
          counts[l] = (counts[l] ?? 0) + 1;
          total++;
        }
      }
    return {
      counts,
      total,
      headings: types.filter((l) => l >= 1 && l <= 9).reduce((n, l) => n + (counts[l] ?? 0), 0),
    };
  }, [chapters]);
  const rows = useMemo(
    () =>
      chapters
        .flatMap((c) =>
          c.outline.filter((p) => p.text).map((p) => ({ c, p, level: levelOf(c, p) })),
        )
        .filter(
          ({ c, p, level }) =>
            (filter === 'all' ||
              (filter === 'headings' ? level >= 1 && level <= 9 : String(level) === filter)) &&
            (!query || `${c.file.name} ${p.text}`.toLowerCase().includes(query.toLowerCase())),
        ),
    [chapters, query, filter],
  );
  async function addFiles(files: File[]) {
    if (busy) return;
    setBusy(t('正在读取章节…'));
    setError('');
    setNotice('');
    const added: Chapter[] = [],
      errors: string[] = [];
    try {
      for (const f of files.sort((a, b) =>
        a.name.localeCompare(b.name, 'zh-CN', { numeric: true }),
      )) {
        if (chapters.length + added.length >= 20) {
          errors.push(t('最多添加 20 份文件。'));
          break;
        }
        if (f.size > 25 * 1024 * 1024) {
          errors.push(t('{0}：单个文件不能超过 25 MB。', f.name));
          continue;
        }
        if (
          chapters.concat(added).reduce((n, c) => n + c.file.bytes.length, 0) + f.size >
          100 * 1024 * 1024
        ) {
          errors.push(t('总文件大小不能超过 100 MB。'));
          continue;
        }
        try {
          const data = new Uint8Array(await f.arrayBuffer());
          if (
            chapters
              .concat(added)
              .some(
                (c) =>
                  c.file.name === f.name &&
                  c.file.bytes.length === data.length &&
                  c.file.bytes.every((byte, i) => byte === data[i]),
              )
          ) {
            errors.push(t('{0} 的相同副本已在列表中，已跳过。', f.name));
            continue;
          }
          const c = createChapter(await readWord(f.name, data));
          const issues = chapterIssues(c);
          if (issues.length)
            throw new Error(t('含{0}，请在 Word 副本中处理后再添加。', issues.join('、')));
          added.push(c);
        } catch (e) {
          errors.push(`${f.name}：${e instanceof Error ? e.message : String(e)}`);
        }
      }
      if (added.length) {
        setChapters((old) => [...old, ...added]);
        setNotice(t('已添加 {0} 章，按文件名排序。可用箭头调整最终顺序。', added.length));
      }
      setError(errors.join('\n'));
    } finally {
      setBusy('');
    }
  }
  async function demo() {
    setBusy(t('正在载入示例…'));
    setError('');
    try {
      const result = [];
      for (const side of ['left', 'right'] as const)
        result.push(
          createChapter(
            await readWord(
              side === 'left' ? t('01_方案初稿.docx') : t('02_补充材料.docx'),
              await createDemo(side),
            ),
          ),
        );
      setChapters(result);
      setMode('unified');
      setNotice(t('已载入两份虚构示例。先检查大纲，再选择需要统一的类型。'));
    } catch (e) {
      setError(String(e));
    } finally {
      setBusy('');
    }
  }
  function move(id: string, direction: number) {
    setChapters((old) => {
      const next = [...old],
        i = next.findIndex((c) => c.id === id),
        j = i + direction;
      if (j < 0 || j >= next.length) return old;
      [next[i], next[j]] = [next[j], next[i]];
      return next;
    });
  }
  function changeLevel(id: string, index: number, level: number) {
    setChapters((old) =>
      old.map((c) => (c.id === id ? { ...c, overrides: { ...c.overrides, [index]: level } } : c)),
    );
    setActive(`${id}:${index}`);
  }
  function applyGroup(chapter: Chapter, index: number) {
    const item = chapter.outline[index],
      level = levelOf(chapter, item);
    setChapters((old) =>
      old.map((c) => {
        const overrides = { ...c.overrides };
        for (const p of c.outline)
          if (
            p.group === item.group &&
            (item.group.startsWith('style:') ? c.id === chapter.id : true)
          )
            overrides[p.index] = level;
        return { ...c, overrides };
      }),
    );
    setNotice(
      t(
        '已将同类识别结果设为“{0}”。样式名称分组仅作用于当前章；编号模式分组作用于所有章。',
        typeLabel(level),
      ),
    );
  }
  function updateFormat(level: number, patch: Partial<TypeFormat>) {
    setFormats((old) => ({ ...old, [level]: { ...old[level], ...patch } }));
  }
  async function exportFile(showPreview: boolean) {
    setBusy(showPreview ? t('正在生成预览…') : t('正在合并并导出…'));
    setError('');
    try {
      const bytes = await mergeChapters(chapters, { mode, pageBreak, formats });
      if (showPreview) setPreview(bytes);
      else {
        const filename =
          (name
            .trim()
            .replace(/[\\/:*?"<>|]/g, '_')
            .replace(/\.docx$/i, '') || t('章节合并')) + '.docx';
        const url = URL.createObjectURL(
          new Blob([bytes.slice().buffer], {
            type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          }),
        );
        setSaved({ url, name: filename });
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setNotice(t('合并文件已生成。若浏览器没有自动保存，请点击下方“保存合并文件”。'));
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy('');
    }
  }
  return (
    <main className="chapter-main">
      <div className="page-heading">
        <div>
          <div className="eyebrow">CHAPTERS TO DOCUMENT</div>
          <h1>
            {t('散落的章节，完整的文档')}
            <span>{t('。')}</span>
          </h1>
          <p>{t('排好顺序，整理大纲，让每一章都各就其位。')}</p>
        </div>
        <button className="button subtle" disabled={!!busy} onClick={demo}>
          {t('试用示例')}
        </button>
      </div>
      <div className="chapter-layout">
        <aside className="chapter-sidebar">
          <section className="chapter-panel">
            <div className="panel-title">
              <h2>
                <Files size={18} />
                {t('章节清单')}
                <small>{chapters.length} / 20</small>
              </h2>
              <button
                className="icon-button"
                onClick={() => input.current?.click()}
                disabled={!!busy}
                aria-label={t('添加章节')}
              >
                <Plus size={18} />
              </button>
            </div>
            <input
              ref={input}
              type="file"
              accept=".docx"
              multiple
              hidden
              onChange={(e) => {
                void addFiles(Array.from(e.target.files ?? []));
                e.target.value = '';
              }}
            />
            <div
              className="chapter-drop"
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => {
                e.preventDefault();
                if (!busy) void addFiles(Array.from(e.dataTransfer.files));
              }}
            >
              <button className="button" onClick={() => input.current?.click()} disabled={!!busy}>
                <Plus size={16} />
                {t('添加 Word 文档')}
              </button>
              <p>{t('支持多选或拖入 · .docx · 每份 ≤ 25 MB')}</p>
            </div>
            <ol className="chapter-files">
              {chapters.map((c, i) => (
                <li key={c.id}>
                  <span className="chapter-number">{String(i + 1).padStart(2, '0')}</span>
                  <div>
                    <strong title={c.file.name}>{c.file.name}</strong>
                    <small>
                      {c.outline.length}
                      {t('段')}
                      {i === 0 && t(' · 页面与页眉页脚基准')}
                    </small>
                  </div>
                  <div className="chapter-file-actions">
                    <button
                      className="icon-button"
                      disabled={!!busy || i === 0}
                      onClick={() => move(c.id, -1)}
                      aria-label={t('上移 {0}', c.file.name)}
                    >
                      <ArrowUp size={15} />
                    </button>
                    <button
                      className="icon-button"
                      disabled={!!busy || i === chapters.length - 1}
                      onClick={() => move(c.id, 1)}
                      aria-label={t('下移 {0}', c.file.name)}
                    >
                      <ArrowDown size={15} />
                    </button>
                    <button
                      className="icon-button"
                      disabled={!!busy}
                      onClick={() => setChapters((old) => old.filter((x) => x.id !== c.id))}
                      aria-label={t('移除 {0}', c.file.name)}
                    >
                      <Trash2 size={15} />
                    </button>
                  </div>
                </li>
              ))}
            </ol>
            {!chapters.length && (
              <div className="chapter-empty-small">
                {t('把第一章、第二章、第三章…')}
                <br />
                {t('按最终阅读顺序放在这里。')}
              </div>
            )}
          </section>
          <section className="chapter-panel export-panel">
            <h2>{t('合并与导出')}</h2>
            <fieldset disabled={!!busy}>
              <label className="check-label">
                <input
                  type="checkbox"
                  checked={pageBreak}
                  onChange={(e) => setPageBreak(e.target.checked)}
                />
                {t('每一章另起一页')}
              </label>
              <label className="field-label">
                {t('文件名')}
                <input
                  value={name}
                  onChange={(e) => setName(e.target.value)}
                  maxLength={100}
                  aria-label={t('合并文件名')}
                />
              </label>
              <p className="chapter-note">
                {t('页面尺寸、页边距、主题和页眉页脚采用清单第一章。原有正文分节不保留。')}
              </p>
              <div className="export-buttons">
                <button
                  className="button"
                  disabled={!chapters.length}
                  onClick={() => exportFile(true)}
                >
                  <Eye size={16} />
                  {t('预览')}
                </button>
                <button
                  className="button primary"
                  disabled={!chapters.length}
                  onClick={() => exportFile(false)}
                >
                  <Download size={16} />
                  {t('导出 Word')}
                </button>
              </div>
            </fieldset>
            {saved && (
              <a className="chapter-save" href={saved.url} download={saved.name}>
                {t('保存合并文件 ·')}
                {saved.name}
              </a>
            )}
          </section>
        </aside>
        <div className="chapter-content">
          <section className="chapter-mode-grid" aria-label={t('合并模式')}>
            <button
              disabled={!!busy}
              className={`chapter-mode ${mode === 'original' ? 'chosen' : ''}`}
              aria-pressed={mode === 'original'}
              onClick={() => setMode('original')}
            >
              <FileText size={23} />
              <div>
                <strong>{t('保留各章格式')}</strong>
                <span>{t('按顺序拼接，保留文字、表格和图片。')}</span>
              </div>
              <span className="radio-dot" />
            </button>
            <button
              disabled={!!busy}
              className={`chapter-mode ${mode === 'unified' ? 'chosen' : ''}`}
              aria-pressed={mode === 'unified'}
              onClick={() => setMode('unified')}
            >
              <WandSparkles size={23} />
              <div>
                <strong>{t('智能大纲与统一格式')}</strong>
                <span>{t('识别标题层级，按类型统一文字格式。')}</span>
              </div>
              <span className="radio-dot" />
            </button>
          </section>
          {(busy || error || notice) && (
            <div
              className={`chapter-message ${error ? 'has-error' : ''}`}
              role={error ? 'alert' : 'status'}
            >
              {busy || error || notice}
            </div>
          )}
          {chapters.flatMap((c) =>
            chapterWarnings(c).map((w) => (
              <div className="chapter-note warning-note" key={c.id + w}>
                {c.file.name}：{w}
              </div>
            )),
          )}
          {mode === 'unified' ? (
            <>
              <section className="chapter-panel">
                <div className="panel-title">
                  <h2>
                    <WandSparkles size={18} />
                    {t('统一哪些类型')}
                  </h2>
                  <label className="preset-picker">
                    {t('格式预设')}
                    <select
                      aria-label={t('格式预设')}
                      defaultValue=""
                      disabled={!!busy}
                      onChange={(e) => {
                        if (e.target.value)
                          setFormats(defaultFormats(e.target.value as 'report' | 'simple'));
                        e.target.value = '';
                      }}
                    >
                      <option value="">{t('选择预设…')}</option>
                      <option value="report">{t('中文报告（宋体 / 黑体）')}</option>
                      <option value="simple">{t('简洁方案（微软雅黑）')}</option>
                    </select>
                  </label>
                </div>
                <p className="chapter-note">
                  {t(
                    '勾选的类型统一字体、字号、粗细、颜色和段落排版。未勾选的类型保留外观，仍写入已确认的大纲级别；斜体、上下标和链接保留。',
                  )}
                </p>
                <fieldset disabled={!!busy} className="format-types">
                  {types
                    .filter((l) => l <= 3 || l >= 10 || stats.counts[l])
                    .map((level) => (
                      <div
                        className={`type-card ${formats[level].enabled ? 'enabled' : ''}`}
                        key={level}
                      >
                        <label className="type-toggle">
                          <input
                            type="checkbox"
                            checked={formats[level].enabled}
                            onChange={(e) => updateFormat(level, { enabled: e.target.checked })}
                          />
                          <strong>{typeLabel(level)}</strong>
                          <small>
                            {stats.counts[level] ?? 0}
                            {t('段')}
                          </small>
                        </label>
                        <div className="type-basic">
                          <label>
                            {t('字体')}
                            <input
                              list="chapter-fonts"
                              value={formats[level].font}
                              disabled={!formats[level].enabled}
                              onChange={(e) => updateFormat(level, { font: e.target.value })}
                              aria-label={t('{0}字体', typeLabel(level))}
                            />
                          </label>
                          <label>
                            {t('字号 pt')}
                            <input
                              type="number"
                              min="5"
                              max="72"
                              step="0.5"
                              value={formats[level].size}
                              disabled={!formats[level].enabled}
                              onChange={(e) =>
                                updateFormat(level, { size: Number(e.target.value) })
                              }
                              aria-label={t('{0}字号', typeLabel(level))}
                            />
                          </label>
                        </div>
                        <details>
                          <summary>{t('更多排版设置')}</summary>
                          <div className="type-details">
                            <label className="check-label">
                              <input
                                type="checkbox"
                                checked={formats[level].bold}
                                onChange={(e) => updateFormat(level, { bold: e.target.checked })}
                              />
                              {t('加粗')}
                            </label>
                            <label>
                              {t('对齐')}
                              <select
                                value={formats[level].align}
                                onChange={(e) =>
                                  updateFormat(level, {
                                    align: e.target.value as TypeFormat['align'],
                                  })
                                }
                              >
                                <option value="left">{t('左对齐')}</option>
                                <option value="center">{t('居中')}</option>
                                <option value="both">{t('两端对齐')}</option>
                              </select>
                            </label>
                            <label>
                              {t('行距倍数')}
                              <input
                                type="number"
                                min="1"
                                max="3"
                                step="0.05"
                                value={formats[level].line}
                                onChange={(e) =>
                                  updateFormat(level, { line: Number(e.target.value) })
                                }
                              />
                            </label>
                            <label>
                              {t('段后 pt')}
                              <input
                                type="number"
                                min="0"
                                max="72"
                                value={formats[level].after}
                                onChange={(e) =>
                                  updateFormat(level, { after: Number(e.target.value) })
                                }
                              />
                            </label>
                            <label>
                              {t('首行缩进（字）')}
                              <input
                                type="number"
                                min="0"
                                max="4"
                                step="0.5"
                                value={formats[level].indent}
                                onChange={(e) =>
                                  updateFormat(level, { indent: Number(e.target.value) })
                                }
                              />
                            </label>
                          </div>
                        </details>
                      </div>
                    ))}
                </fieldset>
                <datalist id="chapter-fonts">
                  {[
                    t('宋体'),
                    t('黑体'),
                    t('微软雅黑'),
                    t('仿宋'),
                    t('楷体'),
                    'Arial',
                    'Times New Roman',
                    'Calibri',
                  ].map((f) => (
                    <option key={f} value={f} />
                  ))}
                </datalist>
              </section>
              <section className="chapter-panel outline-panel">
                <div className="panel-title">
                  <h2>
                    <ListTree size={18} />
                    {t('检查文档大纲')}
                    <small>
                      {stats.headings}
                      {t('个标题')}
                    </small>
                  </h2>
                  <button
                    className="button subtle"
                    disabled={!!busy || !chapters.some((c) => Object.keys(c.overrides).length)}
                    onClick={() => setChapters((old) => old.map((c) => ({ ...c, overrides: {} })))}
                  >
                    <RotateCcw size={14} />
                    {t('重置级别')}
                  </button>
                </div>
                <p className="chapter-note">
                  {t(
                    '识别在本机完成，不调用 AI 服务。可按编号模式批量改级别；切换“全部段落”也能把漏识别的正文改成标题。标题会进入 Word 的导航窗格。',
                  )}
                </p>
                <div className="outline-tools">
                  <label>
                    <Search size={15} />
                    <input
                      placeholder={t('搜索段落或章节…')}
                      value={query}
                      onChange={(e) => setQuery(e.target.value)}
                      aria-label={t('搜索章节段落')}
                    />
                  </label>
                  <select
                    aria-label={t('大纲筛选')}
                    value={filter}
                    onChange={(e) => setFilter(e.target.value)}
                  >
                    <option value="headings">{t('仅标题')}</option>
                    <option value="all">{t('全部段落')}</option>
                    {[-1, ...types].map((l) => (
                      <option key={l} value={l}>
                        {typeLabel(l)}
                      </option>
                    ))}
                  </select>
                  <span>
                    {rows.length}
                    {t('段')}
                  </span>
                </div>
                <div className="outline-list">
                  {rows.slice(page * 50, (page + 1) * 50).map(({ c, p, level }) => (
                    <div
                      key={`${c.id}:${p.index}`}
                      className={`outline-row ${active === `${c.id}:${p.index}` ? 'active' : ''}`}
                      onClick={() => setActive(`${c.id}:${p.index}`)}
                    >
                      <span className="outline-level">
                        {level >= 1 && level <= 9
                          ? `H${level}`
                          : level === 11
                            ? t('表')
                            : level === 10
                              ? t('注')
                              : '¶'}
                      </span>
                      <div
                        className="outline-text"
                        style={{
                          paddingLeft: level >= 1 && level <= 9 ? Math.min(level - 1, 4) * 12 : 0,
                        }}
                      >
                        <strong>{p.text}</strong>
                        <small>
                          {c.file.name}
                          {t('· 第')}
                          {p.index + 1}
                          {t('段 ·')}
                          {t(p.reason)}
                          {c.overrides[p.index] !== undefined ? t(' · 已手动设定') : ''}
                        </small>
                      </div>
                      <div className="outline-actions">
                        <select
                          aria-label={t(
                            '第 {0} 章第 {1} 段级别',
                            chapters.indexOf(c) + 1,
                            p.index + 1,
                          )}
                          disabled={!!busy}
                          value={level}
                          onChange={(e) => changeLevel(c.id, p.index, Number(e.target.value))}
                        >
                          {[-1, ...types]
                            .filter((l) => !p.inTable || [-1, 11, 0].includes(l))
                            .map((l) => (
                              <option value={l} key={l}>
                                {typeLabel(l)}
                              </option>
                            ))}
                        </select>
                        <button
                          className="button subtle"
                          disabled={!!busy}
                          onClick={() => applyGroup(c, p.index)}
                          title={t('将相同识别依据的段落设为此级别')}
                        >
                          {t('同类应用')}
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
                {!rows.length && (
                  <div className="chapter-empty-small">
                    {chapters.length
                      ? t('没有符合条件的段落。可切换“全部段落”检查。')
                      : t('添加文档后，在这里检查和调整大纲。')}
                  </div>
                )}
                {rows.length > 50 && (
                  <div className="outline-pagination">
                    <button
                      className="button"
                      disabled={page === 0}
                      onClick={() => setPage((p) => p - 1)}
                    >
                      {t('上一页')}
                    </button>
                    <span>
                      {page + 1} / {Math.ceil(rows.length / 50)}
                    </span>
                    <button
                      className="button"
                      disabled={(page + 1) * 50 >= rows.length}
                      onClick={() => setPage((p) => p + 1)}
                    >
                      {t('下一页')}
                    </button>
                  </div>
                )}
              </section>
            </>
          ) : (
            <section className="chapter-panel original-guide">
              <div className="original-icon">
                <Files size={34} />
              </div>
              <h2>{t('像拼接章节一样简单')}</h2>
              <p>{t('添加文档 → 调整顺序 → 导出完整 Word')}</p>
              <div className="original-facts">
                <div>
                  <strong>{chapters.length || '01'}</strong>
                  <span>{chapters.length ? t('份章节待合并') : t('添加多个章节')}</span>
                </div>
                <div>
                  <strong>{stats.total || '02'}</strong>
                  <span>{stats.total ? t('段内容已读取') : t('排好阅读顺序')}</span>
                </div>
                <div>
                  <strong>DOCX</strong>
                  <span>{t('继续在 Word 中编辑')}</span>
                </div>
              </div>
              <p className="chapter-note">
                {t(
                  '各章样式、列表编号与图片分别处理，减少同名样式相互覆盖。需要统一正文、标题和表格文字时，选择上方“智能大纲与统一格式”。',
                )}
              </p>
            </section>
          )}
          <details className="chapter-limits">
            <summary>{t('使用范围与小提示')}</summary>
            <p>
              {t(
                '适用于普通 DOCX 段落、表格、嵌入图片、列表和链接。不会自动续编或重写正文中的章节编号。表格内文字不会被自动识别成大纲标题。',
              )}
            </p>
            <p>
              {t(
                '含修订、批注、脚注尾注、正文域、内容控件、嵌入对象或旧式图形的文件会提示处理后再添加。字体需在本机安装；浏览器预览与 Word 的分页可能不同。建议导出后在 Word 中检查图片位置和分页。',
              )}
            </p>
            <p>
              {t(
                '文件仅保存在当前浏览器页面内；切换工作台会保留选择，刷新或关闭页面会清空，请及时导出。',
              )}
            </p>
          </details>
        </div>
      </div>
      {preview && (
        <div className="modal-backdrop">
          <section
            className="modal preview-modal"
            role="dialog"
            aria-modal="true"
            aria-label={t('章节合并预览')}
          >
            <header className="modal-header">
              <h2>{t('章节合并预览')}</h2>
              <button
                className="icon-button"
                onClick={() => setPreview(undefined)}
                aria-label={t('关闭章节预览')}
                autoFocus
              >
                <X size={20} />
              </button>
            </header>
            <p className="chapter-note preview-note">
              {t('此为浏览器近似预览，最终分页请在 Word 中核对。')}
            </p>
            <Preview bytes={preview} title={t('章节合并结果')} />
          </section>
        </div>
      )}
    </main>
  );
}
