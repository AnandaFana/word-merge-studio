import { useEffect, useMemo, useRef, useState } from 'react';
import {
  ArrowDown,
  ArrowUp,
  Check,
  CheckCheck,
  ChevronDown,
  Download,
  Eye,
  FileDiff,
  FileText,
  FolderOpen,
  GitMerge,
  HelpCircle,
  RotateCcw,
  Search,
  ShieldCheck,
  X,
} from 'lucide-react';
import {
  choiceIssue,
  compareWords,
  inlineDiff,
  mergeWords,
  readWord,
  type Choice,
  type Plan,
  type Row,
  type Side,
  type WordFile,
  type FormatChoice,
} from './engine';
import { effectiveFormat, formatIssue, formatSummary } from './formatting';
import Preview from './Preview';
import { createDemo } from './demo';

const labels = {
  equal: '相同',
  format: '格式差异',
  modified: '内容修改',
  added: '右侧新增',
  removed: '左侧独有',
  complex: '复杂内容',
};
const sideLabel = (side: Side) => (side === 'left' ? '左侧' : '右侧');
function download(bytes: Uint8Array, name: string, type: string, retain = false) {
  const url = URL.createObjectURL(new Blob([bytes.slice().buffer], { type }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  if (!retain) setTimeout(() => URL.revokeObjectURL(url), 10000);
  return url;
}

export default function App({ embedded = false }: { embedded?: boolean }) {
  const [files, setFiles] = useState<Partial<Record<Side, WordFile>>>({});
  const [plan, setPlan] = useState<Plan>({ base: 'left', choices: {} });
  const [history, setHistory] = useState<Plan[]>([]);
  const [mode, setMode] = useState<'diff' | 'original'>('diff');
  const [onlyDiff, setOnlyDiff] = useState(false);
  const [search, setSearch] = useState('');
  const [active, setActive] = useState('');
  const [busy, setBusy] = useState('');
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [help, setHelp] = useState(false);
  const [result, setResult] = useState<Uint8Array>();
  const [downloadReady, setDownloadReady] = useState<{ url: string; name: string }>();
  useEffect(
    () => () => {
      if (downloadReady) URL.revokeObjectURL(downloadReady.url);
    },
    [downloadReady],
  );
  const [showWarnings, setShowWarnings] = useState(false);
  const inputs = { left: useRef<HTMLInputElement>(null), right: useRef<HTMLInputElement>(null) };
  const loadVersion = useRef(0);
  const comparison = useMemo(() => {
    if (!files.left || !files.right) return { rows: [] as Row[], error: '' };
    try {
      return { rows: compareWords(files.left, files.right), error: '' };
    } catch (e) {
      return { rows: [] as Row[], error: (e as Error).message };
    }
  }, [files]);
  const { rows } = comparison;
  const loaded = !!files.left && !!files.right && !comparison.error;
  const changes = rows.filter((r) => r.kind !== 'equal');
  const reviewed = changes.filter(
    (r) => !!plan.choices[r.id] || (plan.formats?.[r.id] ?? 'default') !== 'default',
  ).length;
  const visible = rows.filter(
    (row) =>
      (!onlyDiff || row.kind !== 'equal') &&
      (!search ||
        `${row.left?.text ?? ''}\n${row.right?.text ?? ''}`
          .toLowerCase()
          .includes(search.toLowerCase())),
  );
  const commit = (next: Plan) => {
    setError('');
    setHistory((h) => [...h.slice(-99), plan]);
    setPlan(next);
    setResult(undefined);
    setDownloadReady(undefined);
    setNotice('');
  };
  const resetReview = () => {
    setPlan({ base: 'left', choices: {} });
    setHistory([]);
    setResult(undefined);
    setDownloadReady(undefined);
    setActive('');
    setNotice('');
  };

  async function load(side: Side, file?: File) {
    if (!file) return;
    const version = ++loadVersion.current;
    setBusy(`正在读取${sideLabel(side)}文档…`);
    setError('');
    try {
      const parsed = await readWord(file.name, await file.arrayBuffer());
      if (loadVersion.current !== version) return;
      setFiles((prev) => ({ ...prev, [side]: parsed }));
      resetReview();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      if (loadVersion.current === version) setBusy('');
    }
  }
  async function demo() {
    setBusy('正在载入示例…');
    setError('');
    try {
      const [left, right] = await Promise.all([createDemo('left'), createDemo('right')]);
      setFiles({
        left: await readWord('项目方案 · 初稿.docx', left),
        right: await readWord('项目方案 · 修订稿.docx', right),
      });
      resetReview();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy('');
    }
  }
  function choose(row: Row, choice: Choice) {
    const issue = choiceIssue(row, choice, plan.base, files as Record<Side, WordFile>);
    if (issue) {
      setError(issue);
      return;
    }
    commit({ ...plan, choices: { ...plan.choices, [row.id]: choice } });
    setActive(row.id);
  }
  function chooseAll(side: Side) {
    const choices = { ...plan.choices };
    let skipped = 0;
    for (const row of changes) {
      if (choiceIssue(row, side, plan.base, files as Record<Side, WordFile>)) skipped++;
      else choices[row.id] = side;
    }
    commit({ ...plan, choices });
    setNotice(
      `已采纳${sideLabel(side)}${changes.length - skipped} 项${skipped ? `；${skipped} 项复杂内容保持原选择` : ''}。`,
    );
  }
  function chooseFormat(row: Row, format: FormatChoice) {
    const issue = formatIssue(row, format, files as Record<Side, WordFile>, plan.base);
    if (issue) {
      setError(issue);
      return;
    }
    commit({ ...plan, formats: { ...plan.formats, [row.id]: format } });
    setActive(row.id);
  }
  async function exportWord(preview = false) {
    setBusy(preview ? '正在生成合并预览…' : '正在打包 Word…');
    setError('');
    try {
      const bytes = await mergeWords(files as Record<Side, WordFile>, rows, plan);
      if (preview) setResult(bytes);
      else {
        const name = `${files[plan.base]!.name.replace(/\.docx$/i, '')}_合并.docx`;
        const url = download(
          bytes,
          name,
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          true,
        );
        setDownloadReady({ url, name });
        setResult(undefined);
        setNotice(
          `Word 已生成。${changes.length - reviewed} 项未确认差异使用${sideLabel(plan.base)}底稿；若未开始下载，请点击下面的保存链接。`,
        );
      }
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy('');
    }
  }
  function navigate(direction: number) {
    const candidates = visible.filter((r) => r.kind !== 'equal');
    if (!candidates.length) return;
    const index = candidates.findIndex((r) => r.id === active);
    const next = candidates[(index + direction + candidates.length) % candidates.length];
    setActive(next.id);
    document.getElementById(next.id)?.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }
  function saveReport() {
    const content = {
      version: 2,
      exportedAt: new Date().toISOString(),
      left: files.left?.name,
      right: files.right?.name,
      base: plan.base,
      formatMode: plan.formatMode ?? 'preserve',
      note: '正文块选择清单；未确认项使用底稿；不包含正文。',
      rows: rows.map((row) => ({
        id: row.id,
        leftIndex: row.left?.index,
        rightIndex: row.right?.index,
        kind: row.kind,
        choice: plan.choices[row.id] ?? plan.base,
        reviewed: !!plan.choices[row.id] || (plan.formats?.[row.id] ?? 'default') !== 'default',
        format: effectiveFormat(row, plan),
      })),
    };
    download(
      new TextEncoder().encode(JSON.stringify(content, null, 2)),
      '合并选择记录.json',
      'application/json',
    );
  }

  return (
    <div className="app-shell">
      {!embedded && (
        <header className="topbar">
          <div className="brand">
            <span className="brand-mark">
              <GitMerge size={22} />
            </span>
            <strong>Word Merge</strong>
            <span className="edition">本地工作台</span>
          </div>
          <div className="topbar-right">
            <span className="privacy">
              <span /> 文件只在本机处理
            </span>
            <button className="icon-button" onClick={() => setHelp(true)} aria-label="使用说明">
              <HelpCircle size={19} />
            </button>
          </div>
        </header>
      )}
      <main>
        <div className="page-heading">
          <div>
            <div className="eyebrow">DOCUMENT COMPARISON & MERGE</div>
            <h1>
              两份文档，一份定稿<span>。</span>
            </h1>
            <p>并排查看差异，逐项选择内容，让格式跟随同一份底稿。</p>
          </div>
          {embedded && (
            <button className="button subtle" onClick={() => setHelp(true)}>
              <HelpCircle size={17} />
              使用说明
            </button>
          )}
          <button className="button subtle" onClick={demo} disabled={!!busy}>
            <FileText size={16} /> 体验示例
          </button>
        </div>
        <section className="file-grid">
          {(['left', 'right'] as Side[]).map((side) => (
            <div
              className={`file-card ${side}`}
              key={side}
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => {
                e.preventDefault();
                if (!busy) void load(side, e.dataTransfer.files[0]);
              }}
            >
              <span className="file-icon">
                <FileText size={25} />
              </span>
              <div className="file-detail">
                <div className="file-label">
                  {sideLabel(side)}文档 <span>{side === 'left' ? 'A' : 'B'}</span>
                  {plan.base === side && loaded && <em>格式基准</em>}
                </div>
                <strong title={files[side]?.name}>
                  {files[side]?.name ?? '选择或拖入 Word 文档'}
                </strong>
                <small>
                  {files[side]
                    ? `${files[side]!.blocks.length} 个正文块 · ${(files[side]!.bytes.length / 1024).toFixed(1)} KB`
                    : '.docx 格式，最大 25 MB'}
                </small>
              </div>
              <button
                className="button file-open"
                onClick={() => inputs[side].current?.click()}
                disabled={!!busy}
                aria-label={`选择${sideLabel(side)}文档`}
              >
                <FolderOpen size={17} />
                {files[side] ? '更换' : '打开'}
              </button>
              <input
                ref={inputs[side]}
                type="file"
                accept=".docx"
                hidden
                onChange={(e) => {
                  void load(side, e.target.files?.[0]);
                  e.target.value = '';
                }}
              />
            </div>
          ))}
        </section>
        <section className="merge-toolbar">
          <div className="base-control">
            <span className="toolbar-label">文档底稿</span>
            <div className="segmented">
              {(['left', 'right'] as Side[]).map((side) => (
                <button
                  disabled={!!busy}
                  key={side}
                  className={plan.base === side ? 'selected' : ''}
                  onClick={() => {
                    if (side !== plan.base) {
                      commit({ base: side, choices: {}, formatMode: plan.formatMode });
                      setNotice('已切换格式基准并重置选择，可撤销恢复。');
                    }
                  }}
                >
                  {plan.base === side && <Check size={14} />} {sideLabel(side)}文档
                </button>
              ))}
            </div>
            <span className="base-hint">样式 · 页眉页脚 · 页面设置</span>
          </div>
          <div className="action-group">
            <button
              className="icon-button"
              title="撤销最近一次选择"
              disabled={!history.length || !!busy}
              onClick={() => {
                setPlan(history[history.length - 1]);
                setHistory((h) => h.slice(0, -1));
                setResult(undefined);
                setDownloadReady(undefined);
                setNotice('已撤销');
              }}
            >
              <RotateCcw size={17} />
            </button>
            <button
              className="button"
              disabled={!loaded || !!busy}
              onClick={() => exportWord(true)}
            >
              <Eye size={16} /> 合并预览
            </button>
            <button
              className="button primary"
              disabled={!loaded || !!busy}
              onClick={() => exportWord()}
            >
              <Download size={16} /> 导出 Word
            </button>
          </div>
        </section>
        <section className="format-toolbar">
          <label>
            全局文字格式
            <select
              aria-label="全局文字格式"
              value={plan.formatMode ?? 'preserve'}
              disabled={!!busy || !!files.left?.locked || !!files.right?.locked}
              onChange={(e) =>
                commit({ ...plan, formatMode: e.target.value as 'preserve' | 'smart' })
              }
            >
              <option value="preserve">精确保留底稿（保留局部强调）</option>
              <option value="smart">智能匹配底稿（整段统一）</option>
            </select>
          </label>
          <span>
            内容和格式独立选择。整段统一会清除该段局部粗体、斜体等差异；页眉页脚仍跟随底稿。
          </span>
        </section>
        {(error || comparison.error) && (
          <div className="error" role="alert">
            {error || comparison.error}
            <button className="icon-button" onClick={() => setError('')} aria-label="关闭错误">
              <X size={16} />
            </button>
          </div>
        )}
        {(busy || notice) && (
          <div className="notice" role="status">
            {busy || notice}
          </div>
        )}
        {downloadReady && (
          <div className="download-ready">
            <Check size={16} />
            <span>合并文件已就绪</span>
            <a href={downloadReady.url} download={downloadReady.name}>
              保存 {downloadReady.name}
            </a>
          </div>
        )}
        {loaded ? (
          <>
            <div className="review-bar">
              <div className="view-tabs">
                <button className={mode === 'diff' ? 'active' : ''} onClick={() => setMode('diff')}>
                  <FileDiff size={16} />
                  差异审阅
                </button>
                <button
                  className={mode === 'original' ? 'active' : ''}
                  onClick={() => setMode('original')}
                >
                  <FileText size={16} />
                  原版预览
                </button>
              </div>
              <span className="review-count">
                <strong>{changes.length}</strong> 处待审差异 <span>·</span> 已确认 {reviewed}/
                {changes.length}
              </span>
              <div className="progress-track">
                <div
                  style={{ width: `${changes.length ? (reviewed / changes.length) * 100 : 100}%` }}
                />
              </div>
            </div>
            <div className="document-workspace">
              <div className="workspace-controls">
                <label className="search">
                  <Search size={15} />
                  <input
                    aria-label="搜索正文"
                    placeholder="搜索正文…"
                    value={search}
                    onChange={(e) => setSearch(e.target.value)}
                  />
                </label>
                <label className="check-label">
                  <input
                    type="checkbox"
                    checked={onlyDiff}
                    onChange={(e) => setOnlyDiff(e.target.checked)}
                  />
                  只看差异
                </label>
                <div className="spacer" />
                <button className="text-button" onClick={() => navigate(-1)} title="上一处差异">
                  <ArrowUp size={15} />
                </button>
                <button className="text-button" onClick={() => navigate(1)} title="下一处差异">
                  <ArrowDown size={15} />
                </button>
                <span className="vertical-rule" />
                <button className="text-button" disabled={!!busy} onClick={() => chooseAll('left')}>
                  全部选左
                </button>
                <button
                  className="text-button"
                  disabled={!!busy}
                  onClick={() => chooseAll('right')}
                >
                  全部选右
                </button>
              </div>
              <div className="column-head">
                <div>
                  <span className="dot left-dot" />
                  左侧 / A <small>{files.left!.name}</small>
                </div>
                <div>
                  <span className="dot right-dot" />
                  右侧 / B <small>{files.right!.name}</small>
                </div>
              </div>
              {mode === 'original' ? (
                <>
                  <div className="preview-note">
                    原版预览用于核对版式，操作请切回差异审阅。浏览器分页可能与 Word 不同。
                  </div>
                  <div className="original-grid">
                    <Preview bytes={files.left!.bytes} title="左侧 Word 原版预览" />
                    <Preview bytes={files.right!.bytes} title="右侧 Word 原版预览" />
                  </div>
                </>
              ) : (
                <div className="diff-scroll">
                  {!visible.length && <div className="empty-filter">没有符合条件的正文块。</div>}
                  {visible.map((row) => (
                    <div
                      id={row.id}
                      key={row.id}
                      className={`diff-row ${row.kind} ${active === row.id ? 'focused' : ''}`}
                      onClick={() => setActive(row.id)}
                    >
                      <div className="row-content">
                        {(['left', 'right'] as Side[]).map((side) => (
                          <div
                            key={side}
                            className={`diff-cell ${plan.choices[row.id] === side ? 'chosen' : ''}`}
                          >
                            <span className="line-number">
                              {row[side] ? String(row[side]!.index + 1).padStart(2, '0') : '—'}
                            </span>
                            <div className="block-text">
                              {row[side] ? (
                                <>
                                  {row[side]!.type === 'tbl' && (
                                    <span className="block-type">表格</span>
                                  )}
                                  {row.kind === 'modified' && row.left && row.right
                                    ? inlineDiff(row.left.text, row.right.text, side).map(
                                        (part, i) =>
                                          part.changed ? (
                                            <mark key={i} className={side}>
                                              {part.text}
                                            </mark>
                                          ) : (
                                            <span key={i}>{part.text}</span>
                                          ),
                                      )
                                    : row[side]!.text || (
                                        <span className="muted">
                                          {row[side]!.safe
                                            ? '空段落'
                                            : '非文本内容，请查看原版预览'}
                                        </span>
                                      )}
                                </>
                              ) : (
                                <span className="absent">此处没有内容</span>
                              )}
                            </div>
                          </div>
                        ))}
                      </div>
                      {row.kind !== 'equal' && (
                        <div className="row-actions">
                          <span className={`status-tag ${row.kind}`}>{labels[row.kind]}</span>
                          {row.kind === 'format' && (
                            <span className="format-note">文字相同，可单独选择格式</span>
                          )}
                          <span className="spacer" />
                          <span className="decision">
                            {plan.choices[row.id] ||
                            (plan.formats?.[row.id] ?? 'default') !== 'default' ? (
                              <>
                                <CheckCheck size={13} />
                                已确认
                              </>
                            ) : (
                              `默认${sideLabel(plan.base)}`
                            )}
                          </span>
                          {(['left', 'right', 'both', 'omit'] as Choice[]).map((choice) => {
                            const issue = choiceIssue(
                              row,
                              choice,
                              plan.base,
                              files as Record<Side, WordFile>,
                            );
                            return (
                              <button
                                key={choice}
                                title={
                                  issue ??
                                  (choice === 'both'
                                    ? '按左、右顺序保留两块内容，格式由下方选项决定'
                                    : '')
                                }
                                disabled={!!issue || !!busy}
                                className={`choice-button ${plan.choices[row.id] === choice ? 'selected' : ''}`}
                                onClick={(e) => {
                                  e.stopPropagation();
                                  choose(row, choice);
                                }}
                              >
                                {choice === 'left'
                                  ? '选左'
                                  : choice === 'right'
                                    ? '选右'
                                    : choice === 'both'
                                      ? '都保留'
                                      : '都不要'}
                              </button>
                            );
                          })}
                        </div>
                      )}
                      <div className="row-format">
                        <span className="format-summaries">
                          左：{formatSummary(files.left!, row.left)}
                          <br />
                          右：{formatSummary(files.right!, row.right)}
                        </span>
                        <label>
                          段落格式
                          <select
                            aria-label={`第 ${rows.indexOf(row) + 1} 块格式`}
                            value={plan.formats?.[row.id] ?? 'default'}
                            disabled={!!busy || plan.choices[row.id] === 'omit'}
                            onClick={(e) => e.stopPropagation()}
                            onChange={(e) => chooseFormat(row, e.target.value as FormatChoice)}
                          >
                            {(['default', 'left', 'right', 'smart'] as FormatChoice[]).map(
                              (format) => {
                                const issue = formatIssue(
                                  row,
                                  format,
                                  files as Record<Side, WordFile>,
                                  plan.base,
                                );
                                return (
                                  <option
                                    key={format}
                                    value={format}
                                    disabled={!!issue}
                                    title={issue}
                                  >
                                    {format === 'default'
                                      ? '跟随全局'
                                      : format === 'left'
                                        ? '左侧段落格式（整段统一）'
                                        : format === 'right'
                                          ? '右侧段落格式（整段统一）'
                                          : '智能匹配底稿（整段统一）'}
                                  </option>
                                );
                              },
                            )}
                          </select>
                        </label>
                        <span className="format-result">
                          内容：
                          {(plan.choices[row.id] ?? plan.base) === 'both'
                            ? '两侧'
                            : (plan.choices[row.id] ?? plan.base) === 'omit'
                              ? '不保留'
                              : sideLabel((plan.choices[row.id] ?? plan.base) as Side)}{' '}
                          / 格式：
                          {effectiveFormat(row, plan) === 'default'
                            ? `${sideLabel(plan.base)}原格式`
                            : effectiveFormat(row, plan) === 'smart'
                              ? `智能${sideLabel(plan.base)}`
                              : `${sideLabel(effectiveFormat(row, plan) as Side)}统一`}
                        </span>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
            <div className="below-workspace">
              <button className="text-button" onClick={() => setShowWarnings(!showWarnings)}>
                <ShieldCheck size={16} />
                格式与支持范围 <ChevronDown size={14} />
              </button>
              <span>未确认差异保留{sideLabel(plan.base)}底稿内容</span>
              <button className="text-button" onClick={saveReport}>
                下载选择记录
              </button>
            </div>
            {showWarnings && (
              <div className="support-panel">
                {(['left', 'right'] as Side[]).map((side) => (
                  <div key={side}>
                    <strong>{sideLabel(side)}文档</strong>
                    <ul>
                      {files[side]!.warnings.map((w, i) => (
                        <li key={i}>{w}</li>
                      ))}
                    </ul>
                  </div>
                ))}
                <p>
                  新增段落继承附近底稿段落的格式，不继承其自动编号。复杂块在底稿中原样保留。表格只支持同结构单元格的文字修改。
                </p>
              </div>
            )}
          </>
        ) : (
          <section className="welcome">
            <div className="welcome-art">
              <div className="mini-page">
                <i />
                <i />
                <i className="pink" />
                <i />
                <i />
              </div>
              <span>
                <GitMerge size={28} />
              </span>
              <div className="mini-page">
                <i />
                <i className="green" />
                <i />
                <i />
                <i />
              </div>
            </div>
            <h2>把两个版本放在一起</h2>
            <p>
              在上方打开两份 Word 文档，或先体验示例。
              <br />
              从内容对比到格式保留，每一步由你决定。
            </p>
            <button className="button primary" disabled={!!busy} onClick={demo}>
              体验合并示例 <ArrowDown size={15} />
            </button>
            <div className="welcome-features">
              <span>
                <FileDiff size={17} />
                逐字差异
              </span>
              <span>
                <CheckCheck size={17} />
                逐项采纳
              </span>
              <span>
                <ShieldCheck size={17} />
                底稿格式保留
              </span>
            </div>
          </section>
        )}
        <footer>
          <span>
            <ShieldCheck size={14} /> 本地处理 · 无上传 · 无账号
          </span>
          <span>
            Word Merge Studio <span className="version">v0.2</span>
          </span>
        </footer>
      </main>
      {result && (
        <div className="modal-backdrop">
          <section
            className="modal preview-modal"
            role="dialog"
            aria-modal="true"
            aria-label="合并结果预览"
          >
            <div className="modal-header">
              <div>
                <h2>合并结果预览</h2>
                <p>
                  使用{sideLabel(plan.base)}文档底稿；段落格式按独立选择应用。最终分页以 Word 为准。
                </p>
              </div>
              <button className="button primary" onClick={() => exportWord()}>
                <Download size={16} />
                导出 Word
              </button>
              <button
                className="icon-button"
                onClick={() => setResult(undefined)}
                aria-label="关闭预览"
              >
                <X />
              </button>
            </div>
            <Preview bytes={result} title="合并结果 Word 预览" />
          </section>
        </div>
      )}
      {help && (
        <div className="modal-backdrop" onClick={() => setHelp(false)}>
          <section
            className="modal help-modal"
            role="dialog"
            aria-modal="true"
            aria-label="使用说明"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="modal-header">
              <h2>如何合并</h2>
              <button className="icon-button" onClick={() => setHelp(false)} aria-label="关闭说明">
                <X />
              </button>
            </div>
            <ol>
              <li>打开两份 .docx，选择左侧或右侧作为格式基准。</li>
              <li>在差异审阅中逐项选左、选右、都保留或都不要。</li>
              <li>通过原版预览核对文档，通过合并预览检查结果。</li>
              <li>导出 Word，未确认项自动沿用底稿；源文件不会被覆盖。</li>
            </ol>
            <p>
              文档底稿控制页面设置、页眉页脚和样式文件。文字与段落格式可以分别选择，例如先“选右”采纳文字，再选“左侧段落格式”。默认精确保留模式继承对应片段格式；整段统一模式按所选段落的主体格式排版，会清除局部粗体、斜体等差异。
            </p>
            <p>
              智能匹配在本机识别标题、正文与列表；优先使用底稿同类对应段落，没有对应段落时选择同类常用格式，不使用邻近标题替代正文。它是启发式规则，需预览核对。表格仅统一单元格文字格式，几何和底色仍跟随底稿。跨侧自动编号格式暂不支持。
            </p>
            <p>预览是浏览器近似排版，不是 Microsoft Word 引擎。</p>
          </section>
        </div>
      )}
    </div>
  );
}
