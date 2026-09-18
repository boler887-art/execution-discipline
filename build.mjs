import { cp, mkdir, rm } from 'node:fs/promises';

await rm('dist', { recursive: true, force: true });
await mkdir('dist', { recursive: true });
for (const f of ['index.html', 'app.js', 'styles.css']) await cp(f, `dist/${f}`);
await cp('public', 'dist', { recursive: true });
console.log('Built dist/ for GitHub Pages base /VZHDO/');
