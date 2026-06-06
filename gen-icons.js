const puppeteer = require('puppeteer-core');
const fs = require('fs');

// 웹의 sebang_ci_white.svg 경로에서 직접 읽어 아이콘 생성
const ciSvgRaw = fs.readFileSync('./sebang_ci_white.svg', 'utf8');
// SVG를 base64로 인코딩해 img 태그에 사용
const ciSvgB64 = Buffer.from(ciSvgRaw).toString('base64');

// 아이콘: 네이비 배경 + 세방 CI 로고(흰색) + "실적현황" 텍스트
const HTML = (size) => `<!DOCTYPE html>
<html><head>
<meta charset="UTF-8">
<style>
* { margin:0; padding:0; box-sizing:border-box; }
body { width:${size}px; height:${size}px; overflow:hidden;
  background: linear-gradient(145deg, #1a2744 0%, #253358 100%);
  display:flex; flex-direction:column;
  align-items:center; justify-content:center; gap:${Math.round(size*0.05)}px;
}
.ci-wrap {
  width: ${Math.round(size * 0.72)}px;
  display: flex; align-items: center; justify-content: center;
}
.ci-wrap img { width: 100%; height: auto; }
.label {
  font-family: 'Noto Sans KR', 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif;
  font-size: ${Math.round(size * 0.105)}px;
  font-weight: 800;
  color: #00b8a9;
  letter-spacing: ${Math.round(size * 0.008)}px;
  white-space: nowrap;
}
.sub {
  font-family: 'Noto Sans KR', 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif;
  font-size: ${Math.round(size * 0.072)}px;
  font-weight: 600;
  color: rgba(255,255,255,0.6);
  letter-spacing: ${Math.round(size * 0.004)}px;
  white-space: nowrap;
  margin-top: -${Math.round(size * 0.02)}px;
}
.line {
  width: ${Math.round(size * 0.6)}px;
  height: ${Math.round(size * 0.012)}px;
  background: linear-gradient(90deg, transparent, #00b8a9, transparent);
  border-radius: 4px;
  margin: ${Math.round(size * 0.01)}px 0;
}
</style>
</head><body>
  <div class="ci-wrap">
    <img src="data:image/svg+xml;base64,${ciSvgB64}" />
  </div>
  <div class="line"></div>
  <div class="label">실적현황</div>
  <div class="sub">동부계약관리</div>
</body></html>`;

async function generate() {
  const browser = await puppeteer.launch({
    executablePath: '/root/.cache/puppeteer/chrome/linux-149.0.7827.22/chrome-linux64/chrome',
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });

  for (const size of [192, 512]) {
    const page = await browser.newPage();
    await page.setViewport({ width: size, height: size, deviceScaleFactor: 1 });
    await page.setContent(HTML(size), { waitUntil: 'networkidle0' });
    const buf = await page.screenshot({ type: 'png', clip: { x:0, y:0, width:size, height:size } });
    fs.writeFileSync(`icon-${size}.png`, buf);
    console.log(`icon-${size}.png  ${buf.length} bytes`);
    await page.close();
  }
  await browser.close();
}

generate().catch(e => { console.error(e); process.exit(1); });
