import { useSyncExternalStore } from 'react';
import en from './locales/en.json';
export type Language = 'zh' | 'en';
let language: Language = 'zh';
try {
  language = localStorage.getItem('word-merge-language') === 'en' ? 'en' : 'zh';
} catch {
  /* Storage can be disabled. */
}
const listeners = new Set<() => void>();
function updateDocumentLanguage() {
  document.documentElement.lang = language === 'zh' ? 'zh-CN' : 'en';
  document.title =
    language === 'zh' ? 'Word Merge · 文档合并工作台' : 'Word Merge · Document workspace';
}
updateDocumentLanguage();
export function setLanguage(next: Language) {
  language = next;
  try {
    localStorage.setItem('word-merge-language', next);
  } catch {
    /* Keep session preference. */
  }
  updateDocumentLanguage();
  listeners.forEach((fn) => fn());
}
export function useLanguage(): [Language, typeof setLanguage] {
  const value = useSyncExternalStore(
    (fn) => {
      listeners.add(fn);
      return () => {
        listeners.delete(fn);
      };
    },
    () => language,
  );
  return [value, setLanguage];
}
export function t(key: string, ...args: unknown[]): string {
  const text = language === 'en' ? ((en as Record<string, string>)[key] ?? key) : key;
  return text.replace(/\{(\d+)\}/g, (_, i) => String(args[Number(i)] ?? ''));
}
