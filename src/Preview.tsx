import { t } from './i18n';
import { useEffect, useRef, useState } from 'react';
import { renderAsync } from 'docx-preview';

export default function Preview({ bytes, title }: { bytes: Uint8Array; title: string }) {
  const ref = useRef<HTMLIFrameElement>(null);
  const [error, setError] = useState('');
  const [ready, setReady] = useState(false);
  useEffect(() => {
    const frame = ref.current;
    if (!frame || !ready) return;
    let cancelled = false;
    let observer: ResizeObserver | undefined;
    const run = async () => {
      try {
        setError('');
        const doc = frame.contentDocument!;
        const host = doc.createElement('div');
        const styles = doc.createElement('div');
        await renderAsync(bytes.slice().buffer, host, styles, {
          className: 'word-preview',
          inWrapper: true,
          ignoreWidth: false,
          ignoreHeight: true,
          breakPages: true,
          useBase64URL: true,
          renderAltChunks: false,
          renderComments: false,
          renderChanges: false,
        });
        if (cancelled) return;
        doc.body.replaceChildren(styles, host);
        const style = doc.createElement('style');
        style.textContent =
          'html,body{margin:0;background:#e9ebe8} .word-preview-wrapper{padding:22px!important;background:#e9ebe8!important} section.word-preview{margin:0 auto 20px!important;box-shadow:0 2px 12px #0001!important} a{pointer-events:none}';
        doc.head.appendChild(style);
        doc.addEventListener('click', (e) => e.preventDefault());
        const pages = Array.from(host.querySelectorAll<HTMLElement>('section.word-preview'));
        const pageWidth = Math.max(0, ...pages.map((page) => page.offsetWidth));
        if (pageWidth) {
          const width = pageWidth + 44;
          host.style.width = `${width}px`;
          host.style.margin = '0 auto';
          const fit = () => {
            host.style.zoom = String(Math.min(1, Math.max(0.1, (frame.clientWidth - 16) / width)));
          };
          fit();
          observer = new ResizeObserver(fit);
          observer.observe(frame);
        }
      } catch (e) {
        if (!cancelled) setError(e instanceof Error ? e.message : t('预览失败'));
      }
    };
    void run();
    return () => {
      cancelled = true;
      observer?.disconnect();
    };
  }, [bytes, ready]);
  return (
    <div className="preview-wrap">
      {error && (
        <div className="error">
          {t('预览不可用：')}
          {error}
          {t('。可继续查看差异列表。')}
        </div>
      )}
      <iframe
        ref={ref}
        title={title}
        sandbox="allow-same-origin"
        onLoad={() => setReady(true)}
        srcDoc={
          "<!doctype html><html><head><meta charset=\"UTF-8\"><meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; style-src 'unsafe-inline'; img-src data: blob:; font-src data: blob:; script-src 'none'; connect-src 'none'; form-action 'none'; base-uri 'none'\"></head><body></body></html>"
        }
      />
    </div>
  );
}
