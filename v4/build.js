/*
 * build.js — minify app.js into app.min.js using terser.
 * Run after editing app.js:  npm install && npm run build
 */
const { minify } = require('terser');
const fs = require('fs');
const path = require('path');

(async () => {
  const src = fs.readFileSync(path.join(__dirname, 'app.js'), 'utf8');
  const result = await minify(src, {
    compress: { drop_console: false, passes: 2 },
    mangle: { toplevel: true },
    format: { comments: false },
    sourceMap: false,
  });
  if (result.error) { console.error(result.error); process.exit(1); }
  fs.writeFileSync(path.join(__dirname, 'app.min.js'), result.code);
  const before = src.length;
  const after = result.code.length;
  console.log(`app.js     ${before} bytes`);
  console.log(`app.min.js ${after} bytes  (${Math.round((1 - after/before) * 100)}% smaller)`);
})();
