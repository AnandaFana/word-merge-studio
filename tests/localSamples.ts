import { existsSync, readdirSync } from 'node:fs';
import { join } from 'node:path';

const folder = existsSync('examples') ? 'examples' : '.';
// Personal fixtures are opt-in and are never part of public CI.
const names =
  process.env.RUN_LOCAL_SAMPLES === '1'
    ? readdirSync(folder).filter((n) => /\.docx$/i.test(n))
    : [];
const base = names.find((n) => names.includes(n.replace(/\.docx$/i, ' - 2.docx')));
export const compareSamples = base
  ? [join(folder, base), join(folder, base.replace(/\.docx$/i, ' - 2.docx'))]
  : [];
export const chapterSamples = names
  .filter((n) => /^0[123]_/.test(n) && !n.includes(' - '))
  .sort()
  .map((n) => join(folder, n));
