import { useState } from 'react';
import { ArrowRight, FileDiff, Files, GitMerge, House, ShieldCheck } from 'lucide-react';
import App from './App';
import ChapterWorkspace from './ChapterWorkspace';
import './chapters.css';

type Page = 'home' | 'compare' | 'chapters';
export default function Workbench() {
  const [page, setPage] = useState<Page>('home');
  const [visited, setVisited] = useState<Partial<Record<Page, boolean>>>({ home: true });
  const open = (next: Page) => {
    setVisited((old) => ({ ...old, [next]: true }));
    setPage(next);
    window.scrollTo(0, 0);
  };
  return (
    <>
      <header className="workbench-header">
        <button
          className="brand workbench-brand"
          onClick={() => open('home')}
          aria-label="Word Merge 主页"
        >
          <span className="brand-mark">
            <GitMerge size={22} />
          </span>
          <strong>Word Merge</strong>
        </button>
        <nav aria-label="工作台导航">
          {(
            [
              ['home', '主页', House],
              ['compare', '对比合并', FileDiff],
              ['chapters', '合并 Word', Files],
            ] as const
          ).map(([key, label, Icon]) => (
            <button
              key={key}
              aria-current={page === key ? 'page' : undefined}
              onClick={() => open(key)}
            >
              <Icon size={16} />
              {label}
            </button>
          ))}
        </nav>
        <span className="workbench-privacy">
          <ShieldCheck size={16} />
          文件只在本机处理
        </span>
      </header>
      {page === 'home' && (
        <main className="home-main">
          <div className="home-intro">
            <div className="eyebrow">YOUR LOCAL DOCUMENT WORKBENCH</div>
            <h1>
              让文档，
              <br />
              有条不紊地合在一起<span>。</span>
            </h1>
            <p>
              从两份稿件的取舍，到多个章节的整理。
              <br />
              保留重要的内容，把格式交给清晰的选择。
            </p>
          </div>
          <div className="home-tools">
            <button className="home-tool" onClick={() => open('compare')}>
              <span className="home-tool-icon">
                <FileDiff size={30} />
              </span>
              <small>01 / COMPARE</small>
              <h2>对比合并</h2>
              <p>左边一份，右边一份。逐块比较差异，分别选择文字和格式，合成最终版本。</p>
              <div className="home-diagram compare-diagram">
                <span>
                  原稿 <i />
                  <i />
                  <i />
                </span>
                <span>
                  修订 <i />
                  <i />
                  <i />
                </span>
              </div>
              <strong className="home-tool-link">
                打开对比工作台 <ArrowRight size={18} />
              </strong>
            </button>
            <button className="home-tool" onClick={() => open('chapters')}>
              <span className="home-tool-icon">
                <Files size={30} />
              </span>
              <small>02 / COMPOSE</small>
              <h2>合并 Word</h2>
              <p>把不同章节排成一本。支持原格式拼接，或识别标题大纲，按层级统一格式。</p>
              <div className="home-diagram chapters-diagram">
                <span>
                  第一章
                  <i />
                  <i />
                </span>
                <span>
                  第二章
                  <i />
                  <i />
                </span>
                <span>
                  第三章
                  <i />
                  <i />
                </span>
              </div>
              <strong className="home-tool-link">
                开始整理章节 <ArrowRight size={18} />
              </strong>
            </button>
          </div>
          <div className="home-footnote">
            <ShieldCheck size={18} />
            <span>无需上传文档 · 无需配置 AI 密钥 · 导出可编辑的 .docx</span>
            <span className="home-version">Word Merge Studio / v0.3</span>
          </div>
        </main>
      )}
      {visited.compare && (
        <div hidden={page !== 'compare'}>
          <App embedded />
        </div>
      )}
      {visited.chapters && (
        <div hidden={page !== 'chapters'}>
          <ChapterWorkspace />
        </div>
      )}
    </>
  );
}
