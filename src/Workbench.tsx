import { t } from './i18n';
import { useState } from 'react';
import { ArrowRight, FileDiff, Files, GitMerge, House, ShieldCheck } from 'lucide-react';
import App from './App';
import ChapterWorkspace from './ChapterWorkspace';
import SectionWorkspace from './SectionWorkspace';
import { useLanguage } from './i18n';
import './chapters.css';

type Page = 'home' | 'compare' | 'chapters' | 'sections';
export default function Workbench() {
  const [language, setLanguage] = useLanguage();
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
          aria-label={t('Word Merge 主页')}
        >
          <span className="brand-mark">
            <GitMerge size={22} />
          </span>
          <strong>Word Merge</strong>
        </button>
        <nav aria-label={t('工作台导航')}>
          {(
            [
              ['home', t('主页'), House],
              ['compare', t('对比合并'), FileDiff],
              ['chapters', t('合并 Word'), Files],
              ['sections', t('按章节融合'), GitMerge],
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
          {t('文件只在本机处理')}
        </span>
        <select
          className="language-select"
          aria-label={t('Language / 语言')}
          value={language}
          onChange={(e) => setLanguage(e.target.value as 'zh' | 'en')}
        >
          <option value="zh">{t('中文')}</option>
          <option value="en">English</option>
        </select>
      </header>
      {page === 'home' && (
        <main className="home-main">
          <div className="home-intro">
            <div className="eyebrow">YOUR LOCAL DOCUMENT WORKBENCH</div>
            <h1>
              {t('让文档，')}
              <br />
              {t('有条不紊地合在一起')}
              <span>{t('。')}</span>
            </h1>
            <p>
              {t('从两份稿件的取舍，到多个章节的整理。')}
              <br />
              {t('保留重要的内容，把格式交给清晰的选择。')}
            </p>
          </div>
          <div className="home-tools">
            <button className="home-tool" onClick={() => open('sections')}>
              <span className="home-tool-icon">
                <GitMerge size={30} />
              </span>
              <small>03 / SECTIONS</small>
              <h2>{t('按章节融合')}</h2>
              <p>
                {t(
                  '把多份相似结构的报告按对应章节归组。移动内容卡片，生成三级导航，按需统一格式。',
                )}
              </p>
              <div className="home-diagram chapters-diagram">
                <span>
                  {t('A / 目标')}
                  <i />
                  <i />
                </span>
                <span>
                  {t('B / 目标')}
                  <i />
                  <i />
                </span>
                <span>
                  {t('C / 目标')}
                  <i />
                  <i />
                </span>
              </div>
              <strong className="home-tool-link">
                {t('打开章节工作室')}
                <ArrowRight size={18} />
              </strong>
            </button>
            <button className="home-tool" onClick={() => open('compare')}>
              <span className="home-tool-icon">
                <FileDiff size={30} />
              </span>
              <small>01 / COMPARE</small>
              <h2>{t('对比合并')}</h2>
              <p>{t('左边一份，右边一份。逐块比较差异，分别选择文字和格式，合成最终版本。')}</p>
              <div className="home-diagram compare-diagram">
                <span>
                  {t('原稿')}
                  <i />
                  <i />
                  <i />
                </span>
                <span>
                  {t('修订')}
                  <i />
                  <i />
                  <i />
                </span>
              </div>
              <strong className="home-tool-link">
                {t('打开对比工作台')}
                <ArrowRight size={18} />
              </strong>
            </button>
            <button className="home-tool" onClick={() => open('chapters')}>
              <span className="home-tool-icon">
                <Files size={30} />
              </span>
              <small>02 / COMPOSE</small>
              <h2>{t('合并 Word')}</h2>
              <p>{t('把不同章节排成一本。支持原格式拼接，或识别标题大纲，按层级统一格式。')}</p>
              <div className="home-diagram chapters-diagram">
                <span>
                  {t('第一章')}
                  <i />
                  <i />
                </span>
                <span>
                  {t('第二章')}
                  <i />
                  <i />
                </span>
                <span>
                  {t('第三章')}
                  <i />
                  <i />
                </span>
              </div>
              <strong className="home-tool-link">
                {t('开始整理章节')}
                <ArrowRight size={18} />
              </strong>
            </button>
          </div>
          <div className="home-footnote">
            <ShieldCheck size={18} />
            <span>{t('无需上传文档 · 无需配置 AI 密钥 · 导出可编辑的 .docx')}</span>
            <span className="home-version">Word Merge Studio / v0.4</span>
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
      {visited.sections && (
        <div hidden={page !== 'sections'}>
          <SectionWorkspace />
        </div>
      )}
    </>
  );
}
