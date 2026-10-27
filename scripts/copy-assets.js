// Copies the shared web apps (and relay code) into a platform wrapper before packaging.
// node scripts/copy-assets.js desktop   -> desktop/web, desktop/shared
// node scripts/copy-assets.js android   -> android/www
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const target = process.argv[2];
const copy = (from, to) => { fs.rmSync(to, { recursive: true, force: true }); fs.cpSync(from, to, { recursive: true }); console.log('copied', path.relative(root, from), '->', path.relative(root, to)); };
if (target === 'desktop') { copy(path.join(root, 'web'), path.join(root, 'desktop/web')); copy(path.join(root, 'shared'), path.join(root, 'desktop/shared')); }
else if (target === 'android') copy(path.join(root, 'web'), path.join(root, 'android/www'));
else { console.error('usage: copy-assets.js desktop|android'); process.exit(1); }
