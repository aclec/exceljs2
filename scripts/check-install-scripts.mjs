// Supply-chain guard (run in CI): the published package must not run code when it is installed,
// and no dependency may be allowed to run its own install scripts.
//
// - package.json must not define preinstall / install / postinstall / prepublish. "prepare" is allowed
//   because npm and bun only run it for the root project of a local/git checkout, never when the
//   package is installed from the registry; here it only installs the husky git hooks (dev-only).
// - "trustedDependencies" must be present and empty: bun runs the lifecycle scripts of the dependencies
//   listed there and, when the field is missing, of its built-in allow-list (which includes esbuild,
//   whose postinstall can download a binary). An explicit [] blocks them all.
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const FORBIDDEN_SCRIPTS = ['preinstall', 'install', 'postinstall', 'prepublish'];
const ALLOWED_PREPARE = 'husky';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
const scripts = pkg.scripts || {};
const problems = [];

for (const name of FORBIDDEN_SCRIPTS) {
  if (name in scripts) {
    problems.push(`package.json defines the "${name}" lifecycle script: ${JSON.stringify(scripts[name])}`);
  }
}
if ('prepare' in scripts && scripts.prepare !== ALLOWED_PREPARE) {
  problems.push(`"prepare" may only run "${ALLOWED_PREPARE}", found ${JSON.stringify(scripts.prepare)}`);
}
const trusted = pkg.trustedDependencies;
if (!Array.isArray(trusted) || trusted.length > 0) {
  problems.push(`"trustedDependencies" must be an empty array, found ${JSON.stringify(trusted)}`);
}

if (problems.length) {
  for (const problem of problems) process.stderr.write(`error: ${problem}\n`);
  process.exit(1);
}
process.stdout.write('ok: no install-time lifecycle scripts, trustedDependencies is []\n');
