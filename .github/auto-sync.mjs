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
  // 넉넉하게 준다. 되채우기 지시 하나에 인바디 24건이 들어있으면 시트 왕복만
  // 24번이라 3분으로는 모자란다 — 중간에 잘리면 남은 id 가 전부 미반영으로 뜬다.
  await page.waitForFunction(() => window.__autoDone === true, null, { timeout: 900000 });
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

// id 만 등록되고 음식 줄은 안 들어가는 상태가 있을 수 있다 —
// setMeta 는 닿았는데 그 뒤 add 가 죽으면 위의 "남은 id" 검사는 통과해버린다.
// 그래서 마지막 항목이 건드린 날짜의 실제 시트 기록을 찍어서 눈으로 확인한다.
try {
  const chk = await page.evaluate(async () => {
    const box = await (await fetch("./inbox.json", { cache: "no-store" })).json();
    const items = (box && box.items) || [];
    const last = items[items.length - 1] || {};
    const want = (last.food || []).map(f => ({ d: f.date, n: f.name }));
    // 오늘은 항상 찍는다 — 마지막 항목이 인바디뿐이면 식단이 한 줄도 안 보여서,
    // 되감기 같은 사고가 났을 때 중복이 생겼는지 알 길이 없다
    const ds = [...new Set([todayStr(), ...want.map(w => w.d),
                            ...Object.keys(last.act || {})])].filter(Boolean);
    // 인바디는 백엔드가 헤더에 없는 칸을 조용히 버린다 — 백엔드를 다시 배포하기
    // 전에 새 항목을 넣으면 "반영 완료" 도장만 찍히고 값은 어디에도 안 남는다.
    // 그래서 시트에서 다시 읽어 넣으려던 값이 실제로 거기 있는지 본다.
    const ibMissing = [];
    if ((last.inbody || []).length) {
      // syncInbody 가 아니라 날것으로 읽는다 — 그건 시트가 모르는 칸을 앱이 들고
      // 있던 값으로 메워주므로 버려진 걸 못 본다
      const rows = (await api({ action: "list", type: "inbody", from: "2000-01-01", to: todayStr() })).items || [];
      for (const w of last.inbody) {
        const r = rows.find(x => x.date === w.date) || {};
        for (const k of Object.keys(w)) {
          if (k === "date" || k === "id" || !(+w[k] > 0)) continue;
          if (+r[k] !== +w[k])
            ibMissing.push(`${w.date} · ${k} = ${w[k]} (시트: ${r[k] === undefined ? "칸 자체가 없음" : r[k]})`);
        }
      }
    }
    return {
      id: last.id || "",
      ibMissing,
      missing: want.filter(w => !(cache[w.d] || []).some(x => x.name === w.n))
                   .map(w => `${w.d} · ${w.n}`),
      days: ds.map(d => ({
        d,
        rows: (cache[d] || []).map(x => `${mealOf(x)} · ${x.name} — ${Math.round(+x.kcal || 0)}kcal`),
        kcal: Math.round((totals(cache[d] || []) || {}).kcal || 0),
      })),
    };
  });
  for (const x of chk.days)
    logs.push(`[시트] ${x.d} — ${x.rows.length}줄 · 합계 ${x.kcal}kcal\n  `
              + (x.rows.join("\n  ") || "(비어 있음)"));
  // 마지막 항목은 뒤에서 지우는 지시가 있을 수 없다. 그 음식 줄이 시트에 없으면
  // id 만 등록되고 내용은 빠진 상태다 — 조용히 넘기면 하루치가 통째로 사라진다
  if (chk.missing.length) {
    logs.push(`[빠짐] ${chk.id} 의 음식이 시트에 없다:\n  ` + chk.missing.join("\n  "));
    code = 1;
  }
  if (chk.ibMissing.length) {
    logs.push(`[빠짐] ${chk.id} 의 인바디 값이 시트에 없다 — 백엔드를 다시 배포했는지 확인:\n  `
              + chk.ibMissing.join("\n  "));
    code = 1;
  }
} catch (e) { logs.push("[시트 확인 실패] " + (e && e.message || e)); }

await browser.close();
srv.close();

if (logs.length) console.log(logs.join("\n"));
// applyInbox 는 실패한 건을 "[inbox 실패]" 로 남기고 다음 실행에서 다시 시도한다.
// 조용히 넘어가면 안 되는 신호라 여기서 빨갛게 만든다
if (logs.some(x => x.startsWith("[inbox 실패]") || x.startsWith("[pageerror]"))) code = 1;
console.log(code ? "✗ 반영 중 문제가 있었다" : "✓ 반영 완료");
process.exit(code);
