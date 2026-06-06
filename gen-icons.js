// Generates icon-192.png and icon-512.png using pure Node (no canvas dependency)
// Uses SVG → inline data URI written as PNG via sharp-free approach with raw pixel writing
// We'll use the Jimp-free method: write a minimal SVG and use puppeteer to screenshot it

const puppeteer = require('puppeteer-core');
const fs = require('fs');

const SVG = (size) => `
<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 512 512">
  <defs>
    <linearGradient id="bg" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" style="stop-color:#1a2744"/>
      <stop offset="100%" style="stop-color:#253358"/>
    </linearGradient>
    <linearGradient id="teal" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" style="stop-color:#00d4c3"/>
      <stop offset="100%" style="stop-color:#00b8a9"/>
    </linearGradient>
  </defs>
  <!-- 배경 -->
  <rect width="512" height="512" rx="96" fill="url(#bg)"/>
  <!-- 세방 CI: 'S' 형태 곡선 장식 -->
  <path d="M 100 310 Q 100 390 180 390 L 332 390 Q 412 390 412 330 Q 412 270 332 260 L 200 245 Q 140 238 140 185 Q 140 130 200 130 L 330 130 Q 390 130 395 185"
        stroke="url(#teal)" stroke-width="48" fill="none" stroke-linecap="round"/>
  <!-- 하단 텍스트 영역 배경 -->
  <rect x="0" y="390" width="512" height="122" rx="0" fill="rgba(0,184,169,0.12)"/>
  <rect x="0" y="388" width="512" height="4" fill="url(#teal)" rx="2"/>
  <!-- 세방 텍스트 -->
  <text x="256" y="445" text-anchor="middle" font-family="Arial, sans-serif" font-size="52" font-weight="900" fill="#00b8a9" letter-spacing="-1">세방(주)</text>
  <!-- 하불실적관리 텍스트 -->
  <text x="256" y="490" text-anchor="middle" font-family="Arial, sans-serif" font-size="30" font-weight="700" fill="rgba(255,255,255,0.7)" letter-spacing="1">하불실적관리</text>
</svg>`;

async function generate() {
  const browser = await puppeteer.launch({
    executablePath: '/root/.cache/puppeteer/chrome/linux-149.0.7827.22/chrome-linux64/chrome',
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });

  for (const size of [192, 512]) {
    const page = await browser.newPage();
    await page.setViewport({ width: size, height: size, deviceScaleFactor: 1 });
    const svg = SVG(size);
    const dataUrl = 'data:image/svg+xml;base64,' + Buffer.from(svg).toString('base64');
    await page.goto(dataUrl);
    const buf = await page.screenshot({ type: 'png', clip: { x:0, y:0, width:size, height:size } });
    fs.writeFileSync(`icon-${size}.png`, buf);
    console.log(`icon-${size}.png written (${buf.length} bytes)`);
    await page.close();
  }
  await browser.close();
}

generate().catch(e => { console.error(e); process.exit(1); });
