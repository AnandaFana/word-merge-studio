import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { stripTypeScriptTypes, createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';
import { resolve } from 'node:path';

// Native Node 22.13+ type stripping keeps the public examples identical to the in-app demo.
// Optional dependency root is useful for managed runtimes; normal contributors use npm ci.
const root = process.argv[2];
const require = createRequire(
  root ? pathToFileURL(resolve(root, '../package.json')) : import.meta.url,
);
const jszip = pathToFileURL(require.resolve('jszip')).href;
const dataModule = (code) => 'data:text/javascript;base64,' + Buffer.from(code).toString('base64');
let demo = stripTypeScriptTypes(await readFile(new URL('../src/demo.ts', import.meta.url), 'utf8'));
demo = demo.replace("'jszip'", JSON.stringify(jszip));
let sections = stripTypeScriptTypes(
  await readFile(new URL('../src/sectionDemo.ts', import.meta.url), 'utf8'),
);
sections = sections
  .replace("'jszip'", JSON.stringify(jszip))
  .replace("'./demo'", JSON.stringify(dataModule(demo)));
const { createSectionDemo } = await import(dataModule(sections));
const folder = new URL('../examples/public/', import.meta.url);
await mkdir(folder, { recursive: true });
for (const key of ['garden', 'library', 'cafe']) {
  await writeFile(new URL(`${key}.docx`, folder), await createSectionDemo(key));
  console.log(`Generated fictional fixture: examples/public/${key}.docx`);
}
