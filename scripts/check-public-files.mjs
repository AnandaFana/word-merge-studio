import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
const files = execFileSync('git', ['ls-files', '-z'], { encoding: 'utf8' })
  .split('\0')
  .filter(Boolean);
const examples = new Set(['garden', 'library', 'cafe'].map((k) => `examples/public/${k}.docx`));
const forbidden = files.filter(
  (path) =>
    /(^|\/)(tmp|\.qa|node_modules|dist|\.npm-cache)(\/|$)/.test(path) ||
    (/(^|\/)\.env(?:\.|$)/.test(path) && !path.endsWith('.env.example')) ||
    (/\.(pem|key|doc|docm|pdf|docx)$/i.test(path) && !examples.has(path)),
);
const secrets = files
  .filter((path) => !/\.(docx|png|jpg|jpeg|woff2?)$/i.test(path))
  .filter((path) => {
    const content = readFileSync(path, 'utf8');
    return (
      /-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----/.test(content) ||
      /(?:gh[pousr]_[A-Za-z0-9]{30,}|github_pat_[A-Za-z0-9_]{40,})/.test(content)
    );
  });
if (forbidden.length || secrets.length) {
  console.error('Public-file check failed (paths only):', [...new Set([...forbidden, ...secrets])]);
  process.exitCode = 1;
} else
  console.log(
    `Public-file check passed: ${files.length} tracked files; only named fictional DOCX fixtures are allowed.`,
  );
