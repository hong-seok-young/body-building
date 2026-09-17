// 시트를 있는 그대로 찍어보는 진단용 스크립트. 아무것도 쓰지 않는다.
const m = (await import("node:fs/promises")).readFile;
const html = await m("index.html", "utf8");
const url = html.match(/url:\s*"([^"]+)"/)[1];
const secret = html.match(/secret:\s*"([^"]+)"/)[1];

async function api(body) {
  const r = await fetch(url, {
    method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ ...body, secret }), redirect: "follow",
  });
  const t = await r.text();
  try { return JSON.parse(t); } catch { throw new Error(`HTTP ${r.status} ${t.slice(0, 200)}`); }
}

const res = await api({ action: "list", from: "2026-01-01", to: "2026-12-31" });
const items = res.items || [];
console.log(`### 식단 줄 수: ${items.length}`);

// 날짜별로 모아서, 같은 끼니·이름이 두 번 이상이면 표시한다
const byDate = {};
for (const it of items) (byDate[it.date] = byDate[it.date] || []).push(it);

let dupRows = 0, dupDays = 0;
for (const d of Object.keys(byDate).sort()) {
  const rows = byDate[d];
  const cnt = {};
  rows.forEach(r => { const k = `${r.meal}|${r.name}`; cnt[k] = (cnt[k] || 0) + 1; });
  const dups = Object.entries(cnt).filter(([, n]) => n > 1);
  const kcal = rows.reduce((s, r) => s + (+r.kcal || 0), 0);
  console.log(`\n== ${d}  (${rows.length}줄, ${Math.round(kcal)}kcal)${dups.length ? "  ⚠ 중복" : ""}`);
  rows.forEach(r => console.log(`   ${r.meal}\t${r.name}\t${r.kcal}\t[${r.id}]`));
  if (dups.length) {
    dupDays++;
    dups.forEach(([k, n]) => { dupRows += n - 1; console.log(`   ⚠ ${k} × ${n}`); });
  }
}
console.log(`\n### 중복: ${dupDays}일 / 남는 줄 ${dupRows}개`);

// 적용 기록이 살아있는지 — 이게 비면 inbox 가 처음부터 다시 적용된다
const meta = res.meta || {};
const calc = (() => { try { return JSON.parse(meta.calc || "{}"); } catch { return {}; } })();
const ids = calc.inbox || [];
console.log(`### calc.inbox 적용기록: ${ids.length}건`);
console.log(`### calc.at: ${calc.at}  deficit: ${calc.deficit}`);
console.log(ids.join(" "));
