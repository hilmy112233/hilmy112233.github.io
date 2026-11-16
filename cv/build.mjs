// Renders cv/cv.html to A4 PDFs (EN + ID) via headless Edge. No deps.
// Run: node cv/build.mjs
//
// Fonts are inlined as data URIs before printing: Chromium's print path drops
// file:// woff2 requests and silently falls back to Segoe UI.
import { execFileSync } from 'node:child_process';
import { existsSync, statSync, mkdtempSync, readFileSync, writeFileSync, rmSync, openSync, closeSync } from 'node:fs';
import { resolve, join } from 'node:path';
import { tmpdir } from 'node:os';

const EDGE = [
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
  'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
].find(existsSync);
if (!EDGE) throw new Error('No Edge/Chrome found');

const root = resolve(import.meta.dirname, '..');
const cvDir = join(root, 'cv');
const profile = () => mkdtempSync(join(tmpdir(), 'cv-'));
const FLAGS = ['--headless=new', '--disable-gpu', '--no-first-run',
  '--allow-file-access-from-files', '--font-render-hinting=none'];

function headless(args, redirectTo) {
  const fd = redirectTo ? openSync(redirectTo, 'w') : 'ignore';
  try {
    execFileSync(EDGE, [...FLAGS, '--user-data-dir=' + profile(), ...args],
      { stdio: ['ignore', fd, 'ignore'] });
  } catch {
    // headless exits non-zero on some machines even when the output was written
  } finally {
    if (redirectTo) closeSync(fd);
  }
}

const pageCount = (pdf) =>
  Math.max(...[...readFileSync(pdf).toString('latin1').matchAll(/\/Count\s+(\d+)/g)].map((m) => +m[1]), 0);
const embeddedFonts = (pdf) =>
  [...new Set([...readFileSync(pdf).toString('latin1').matchAll(/\/(?:BaseFont|FontName)\s*\/([^\s/\]>]+)/g)].map((m) => m[1]))];

// 1047 px is the A4 content box (1123 px page − 2 × 10 mm margin) at 96 dpi.
const AVAILABLE = 1047;

function measure(html, lang) {
  const dump = join(profile(), 'dom.html');
  headless(['--virtual-time-budget=6000', '--dump-dom', 'file:///' + html.replace(/\\/g, '/') + '?lang=' + lang + '&measure=1'], dump);
  const dom = readFileSync(dump, 'utf8');
  const fonts = dom.match(/data-fonts="([^"]+)"/)?.[1];
  if (fonts && !fonts.startsWith('true|true')) throw new Error('fonts not applied: ' + fonts);
  const h = dom.match(/data-h="(\d+)"/);
  if (!h) throw new Error('could not measure layout');
  return Number(h[1]);
}

const inlined = readFileSync(join(cvDir, 'cv.html'), 'utf8').replace(
  /url\('fonts\/([\w-]+\.woff2)'\)/g,
  (_, file) => `url(data:font/woff2;base64,${readFileSync(join(cvDir, 'fonts', file)).toString('base64')})`,
);
const tmpHtml = join(cvDir, '_print.html');
writeFileSync(tmpHtml, inlined);

let bad = 0;
try {
  for (const [lang, out] of [
    ['en', 'assets/CV_Hilmy_Abdurrahman_Darmawan.pdf'],
    ['id', 'assets/CV_Hilmy_Abdurrahman_Darmawan_ID.pdf'],
  ]) {
    const dest = join(root, out);
    headless(['--no-pdf-header-footer', '--virtual-time-budget=8000', '--print-to-pdf=' + dest,
      'file:///' + tmpHtml.replace(/\\/g, '/') + '?lang=' + lang]);
    if (!existsSync(dest)) throw new Error('not produced: ' + out);

    const pages = pageCount(dest);
    const fonts = embeddedFonts(dest);
    // Work Sans covers no arrows or Greek, so any fallback face means a glyph
    // silently rendered in Segoe UI instead of the intended font.
    const fallback = fonts.filter((f) => !/Fraunces|WorkSans/.test(f));
    const ok = pages === 1 && fonts.some((f) => /Fraunces|WorkSans/.test(f)) && !fallback.length;
    if (!ok) bad++;
    const kb = (statSync(dest).size / 1024).toFixed(0);
    const fit = measure(tmpHtml, lang);
    if (fit > AVAILABLE) bad++;
    console.log(`${ok && fit <= AVAILABLE ? 'OK  ' : 'FAIL'} ${out}  ${pages} page(s), ${kb} KB, ${fit}/${AVAILABLE} px` +
      (fallback.length ? `, FALLBACK FONT: ${fallback.join(', ')}` : ''));
  }
} finally {
  rmSync(tmpHtml, { force: true });
}
process.exit(bad ? 1 : 0);
