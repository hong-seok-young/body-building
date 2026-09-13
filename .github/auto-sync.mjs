// 앱을 헤드리스 크롬으로 한 번 열어서 inbox.json 을 구글 시트에 반영한다.
//
// 반영 로직(applyPayload)을 여기에 다시 구현하지 않는 것이 핵심이다.
// 두 벌이 되면 언젠가 갈라지고, 갈라지면 시트에 같은 음식이 두 번 들어가거나
// 목표 설정이 조용히 덮어써진다. 실제 앱을 그대로 돌리면 그럴 일이 없다.
//
// 중복 방지도 앱이 이미 하고 있다 — 적용한 id 를 calc.inbox 로 시트에 올리므로,
// 여기서 먼저 반영해도 나중에 폰에서 앱을 열 때 다시 적용되지 않는다.
import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { extname, join, normalize } from "node:path";
import { chromium } from "playwright";

const ROOT = process.cwd();
const PORT = 8173;
const TYPE = {
  ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8", ".css": "text/css; charset=utf-8",
  ".png": "image/png", ".svg": "image/svg+xml",
  ".webmanifest": "application/manifest+json", ".ico": "image/x-icon",
};

// file:// 로는 안 된다 — fetch("./inbox.json") 이 CORS 로 막힌다. 그래서 잠깐 띄운다
const srv = createServer(async (req, res) => {
  const rel = normalize(decodeURIComponent(req.url.split("?")[0])).replace(/^(\.\.[/\\])+/, "");
  const file = join(ROOT, rel === "/" ? "index.html" : rel);
  try {
    const buf = await readFile(file);
    res.writeHead(200, { "content-type": TYPE[extname(file)] || "application/octet-stream",
                         "cache-control": "no-store" });
    res.end(buf);
  } catch { res.writeHead(404).end("not found"); }
});
await new Promise(r => srv.listen(PORT, "127.0.0.1", r));

const browser = await chromium.launch();
const page = await browser.newPage();
const logs = [];
page.on("console", m => logs.push(m.text()));
page.on("pageerror", e => logs.push("[pageerror] " + (e && e.message || e)));

let code = 0;
try {
  await page.goto(`http://127.0.0.1:${PORT}/index.html?auto=1`, { waitUntil: "domcontentloaded" });
  // fullSync → applyInbox 가 다 끝나야 true 가 된다 (index.html 맨 아래)
  await page.waitForFunction(() => window.__autoDone === true, null, { timeout: 180000 });
  await page.waitForTimeout(3000);          // 마지막 setMeta 가 날아갈 시간을 준다
} catch (e) {
  logs.push("[실패] " + (e && e.message || e));
  code = 1;
}
// 시트에 못 닿으면 applyInbox 는 아무 말 없이 한 건도 안 넣고 끝난다
// (markInboxDone 이 setMeta 실패로 false 를 주고 break). 그걸 성공으로 넘기면
// 나는 반영된 줄 알고 있는데 시트는 비어 있다. 그래서 남은 id 를 직접 세어본다
try {
  const left = await page.evaluate(async () => {
    const box = await (await fetch("./inbox.json", { cache: "no-store" })).json();
    const seen = inboxDone();
    return (box && box.items || []).filter(x => x && x.id && seen.indexOf(x.id) < 0)
                                   .map(x => x.id);
  });
  if (left.length) { logs.push("[미반영] " + left.join(", ")); code = 1; }
} catch (e) { logs.push("[확인 실패] " + (e && e.message || e)); code = 1; }

await browser.close();
srv.close();

if (logs.length) console.log(logs.join("\n"));
// applyInbox 는 실패한 건을 "[inbox 실패]" 로 남기고 다음 실행에서 다시 시도한다.
// 조용히 넘어가면 안 되는 신호라 여기서 빨갛게 만든다
if (logs.some(x => x.startsWith("[inbox 실패]") || x.startsWith("[pageerror]"))) code = 1;
console.log(code ? "✗ 반영 중 문제가 있었다" : "✓ 반영 완료");
process.exit(code);
