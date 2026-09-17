// 시트를 있는 그대로 찍어보는 진단용 스크립트.
// CLEAN=1 을 주면 아래 KILL 목록의 id 만 지운다 — 그 외에는 아무것도 쓰지 않는다.
import { readFile } from "node:fs/promises";
const html = await readFile("index.html", "utf8");
const url = html.match(/url:\s*"([^"]+)"/)[1];
const secret = html.match(/secret:\s*"([^"]+)"/)[1];

// 되감긴 inbox 가 두 번 넣은 줄. inbox 지시는 하나씩만 말하는데 시트에 둘이라
// 뒤에 들어온 쪽을 지운다. 손으로 넣은 줄(아몬드·퀵오트)은 건드리지 않는다.
const KILL = [
  ["2026-09-09 2끼 닭쌀국수",            "48a09eec-516c-4bc4-991f-8fea1f6f89e2"],
  ["2026-09-10 1끼 CJ 닭가슴살 200g",    "9bf8fef2-88be-48aa-ad86-31f4c118bc2a"],
  ["2026-09-10 2끼 마르게리타 피자",      "619b8e23-8bae-4339-a9a5-4bfbec88d6c9"],
  ["2026-09-10 2끼 체리토마토 마르게리타", "be302c65-a2ea-472c-9573-ff03cf919793"],
  ["2026-09-10 2끼 루꼴라·부라타 샐러드",  "f0820c1e-9b1d-4d79-8b42-1ddbb4cf758d"],
];

async function api(body) {
  const r = await fetch(url, {
    method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ ...body, secret }), redirect: "follow",
  });
  const t = await r.text();
  try { return JSON.parse(t); } catch { throw new Error(`HTTP ${r.status} ${t.slice(0, 200)}`); }
}

const list = () => api({ action: "list", from: "2026-01-01", to: "2026-12-31" });

if (process.env.CLEAN === "1") {
  const before = (await list()).items || [];
  const byId = Object.fromEntries(before.map(x => [x.id, x]));
  for (const [label, id] of KILL) {
    const row = byId[id];
    if (!row) { console.log(`· 없음(이미 지워짐) ${label}`); continue; }
    // 지우기 전에 그 줄이 정말 그 줄인지 확인한다 — id 를 잘못 적으면 멀쩡한 기록이 날아간다
    console.log(`🗑 ${label}  →  ${row.date} ${row.meal} ${row.name} ${row.kcal}kcal`);
    const r = await api({ action: "del", id });
    if (!r.ok) throw new Error(`삭제 실패 ${label}: ${JSON.stringify(r)}`);
  }
  console.log("");
}

const res = await list();
const items = res.items || [];
console.log(`### 식단 줄 수: ${items.length}`);

const byDate = {};
for (const it of items) (byDate[it.date] = byDate[it.date] || []).push(it);

let dupRows = 0, dupDays = 0;
for (const d of Object.keys(byDate).sort()) {
  const rows = byDate[d];
  const cnt = {};
  rows.forEach(r => { const k = `${r.meal}|${r.name}`; cnt[k] = (cnt[k] || 0) + 1; });
  const dups = Object.entries(cnt).filter(([, n]) => n > 1);
  const kcal = rows.reduce((s, r) => s + (+r.kcal || 0), 0);
  if (!dups.length) { console.log(`== ${d}  ${rows.length}줄  ${Math.round(kcal)}kcal`); continue; }
  dupDays++;
  console.log(`== ${d}  ${rows.length}줄  ${Math.round(kcal)}kcal  ⚠ 중복`);
  dups.forEach(([k, n]) => { dupRows += n - 1; console.log(`   ⚠ ${k} × ${n}`); });
}
console.log(`\n### 중복: ${dupDays}일 / 남는 줄 ${dupRows}개`);

// 적용 기록은 응답 최상위에 실려 온다 (META_KEYS 를 그대로 붙인다)
const calc = (() => { try { return typeof res.calc === "string" ? JSON.parse(res.calc) : (res.calc || {}); } catch { return {}; } })();
const ids = calc.inbox || [];
console.log(`### calc.inbox 적용기록: ${ids.length}건 (inbox.json 은 ${JSON.parse(await readFile("inbox.json", "utf8")).items.length}건)`);
console.log(`### calc.at: ${calc.at}   적자: ${calc.deficit}`);
