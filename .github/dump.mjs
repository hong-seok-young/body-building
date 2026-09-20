// 시트를 읽어서 inbox.json 의 지시와 대조한다. 읽기만 한다 — 아무것도 안 쓴다.
//
// 목적: "내가 앱에서 손으로 넣거나 지운 게 있나" 에 답하는 것.
//   · 시트에 있는데 지시에 없는 줄  → 앱에서 손으로 넣은 것
//   · 지시로 넣었는데 시트에 없는 줄 → 앱에서 손으로 지운 것 (또는 안 들어간 것)
import { readFile } from "node:fs/promises";

const html = await readFile("index.html", "utf8");
const url = html.match(/url:\s*"([^"]+)"/)[1];
const secret = html.match(/secret:\s*"([^"]+)"/)[1];

async function api(body) {
  // 구글이 가끔 404 나 doGet 응답을 준다 — 읽기라서 다시 물어봐도 안전하다
  for (let i = 0; i < 3; i++) {
    if (i) await new Promise(r => setTimeout(r, 800 * 3 ** (i - 1)));
    try {
      const r = await fetch(url, { method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ ...body, secret }), redirect: "follow" });
      const t = await r.text();
      const j = JSON.parse(t);
      if (j.alive === true && j.items === undefined) throw new Error("doGet 응답");
      if (!Array.isArray(j.items)) throw new Error("items 없음: " + Object.keys(j).join(","));
      return j;
    } catch (e) { if (i === 2) throw e; console.log(`  (다시 시도 ${i + 1}: ${e.message})`); }
  }
}

const sheet = (await api({ action: "list", from: "2026-01-01", to: "2026-12-31" })).items;

// 재배포가 됐는지부터 본다 — time·src 칸이 없으면 백엔드가 옛 버전이다
const cols = Object.keys(sheet[0] || {});
const ready = cols.indexOf("time") >= 0 && cols.indexOf("src") >= 0;
console.log(`### 백엔드 time·src 칸: ${ready ? "있음 ✅ (재배포 완료)" : "없음 ❌ (아직 옛 버전)"}`);
console.log(`### food 시트 칸: ${cols.join(", ")}`);
try {
  const lg = await api({ action: "list", type: "log", from: "2000-01-01", to: "2100-01-01" });
  console.log(`### log 시트: 있음 ✅ (${lg.items.length}줄)`);
} catch (e) { console.log(`### log 시트: ${e.message}`); }
console.log("");
console.log(`### 시트 식단 줄 수: ${sheet.length}\n`);

// ── 지시가 만든 (날짜|끼니|이름) 집합. del 로 지운 것은 빼준다
const box = JSON.parse(await readFile("inbox.json", "utf8"));
const key = x => `${x.date}|${x.meal}|${x.name}`;
const fromInbox = new Set();
for (const it of box.items) {
  for (const d of it.del || []) if (d && d.name) fromInbox.delete(key(d));
  for (const f of it.food || []) if (f && f.name) fromInbox.add(key(f));
}
console.log(`### 지시가 남긴 줄 수: ${fromInbox.size}\n`);

const inSheet = new Set(sheet.map(key));

// ── 1. 시트에 있는데 지시에 없다 = 앱에서 손으로 넣은 것
const manual = sheet.filter(x => !fromInbox.has(key(x)))
                    .sort((a, b) => (a.date + a.meal).localeCompare(b.date + b.meal));
console.log(`## 앱에서 손으로 넣은 줄 — ${manual.length}건`);
for (const x of manual)
  console.log(`  + ${x.date} ${x.meal} · ${x.name} — ${Math.round(+x.kcal || 0)}kcal`);

// ── 2. 지시로 넣었는데 시트에 없다 = 손으로 지웠거나 안 들어간 것
const gone = [...fromInbox].filter(k => !inSheet.has(k)).sort();
console.log(`\n## 지시로 넣었는데 시트에 없는 줄 — ${gone.length}건`);
for (const k of gone) console.log(`  - ${k.split("|").join(" ")}`);

// ── 3. 같은 날 같은 끼니에 같은 이름이 둘 이상 (중복 점검)
const cnt = {};
sheet.forEach(x => { cnt[key(x)] = (cnt[key(x)] || 0) + 1; });
const dups = Object.entries(cnt).filter(([, n]) => n > 1);
console.log(`\n## 중복된 줄 — ${dups.length}건`);
for (const [k, n] of dups.sort()) console.log(`  ! ${k.split("|").join(" ")} × ${n}`);

// ── 4. 날짜별 요약
console.log(`\n## 날짜별`);
const byDate = {};
for (const x of sheet) (byDate[x.date] = byDate[x.date] || []).push(x);
for (const d of Object.keys(byDate).sort()) {
  const rows = byDate[d];
  const kcal = Math.round(rows.reduce((s, r) => s + (+r.kcal || 0), 0));
  const hand = rows.filter(x => !fromInbox.has(key(x))).length;
  console.log(`  ${d}  ${String(rows.length).padStart(2)}줄  ${String(kcal).padStart(5)}kcal`
              + (hand ? `   ← 손으로 넣은 것 ${hand}건` : ""));
}
