// Render an HTML file to PNG with the pre-installed Chromium.
//   node render.js page.html out.png width height [scale] [transparent]
// width/height are CSS pixels; scale is the device pixel ratio.
const path = require('path');
const { pathToFileURL } = require('url');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');

(async () => {
  const [src, out, w, h, scale = '2', transparent = '0'] = process.argv.slice(2);
  const browser = await chromium.launch(process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE
    ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE }
    : {});
  const page = await browser.newPage({
    viewport: { width: +w, height: +h },
    deviceScaleFactor: +scale,
  });
  await page.goto(pathToFileURL(path.resolve(src)).href);
  await page.evaluate(() => document.fonts.ready);
  await page.screenshot({ path: out, omitBackground: transparent === '1', fullPage: false });
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
