/**
 * 바디빌딩 트래커 — Google Apps Script 백엔드
 *
 * 설치 방법은 SETUP.md 참고. 요약:
 *   1) 구글 시트 새로 만들기 → 확장 프로그램 → Apps Script
 *   2) 이 파일 내용을 전부 붙여넣기
 *   3) 아래 SECRET 을 본인이 정한 긴 문자열로 바꾸기
 *   4) 배포 → 새 배포 → 유형 "웹 앱" / 실행 계정 "나" / 액세스 "모든 사용자"
 *   5) 나온 URL 과 SECRET 을 앱 ⚙️ 설정에 입력
 *
 * ⚠️ 액세스가 "모든 사용자"라서 URL 을 아는 사람은 누구나 요청을 보낼 수 있습니다.
 *    실제 방어는 SECRET 하나뿐이니, 길고 추측 불가능한 값으로 바꾸고
 *    URL·SECRET 을 공개 저장소나 채팅에 올리지 마세요.
 *    (이 파일은 템플릿이므로 아래 값을 그대로 쓰면 안 됩니다.)
 */
const SECRET = "여기에-본인이-정한-긴-비밀키-넣기";

// 시트 이름과 컬럼. 컬럼을 추가하고 싶으면 여기에 이름만 넣으면 되고,
// 읽기/쓰기는 헤더 기준으로 자동 매핑되므로 아래 코드는 고칠 필요가 없습니다.
const SHEETS = {
  food:   ["id", "date", "meal", "name", "kcal", "p", "c", "f", "na", "fib", "sug"],
  inbody: ["id", "date", "weight", "smm", "bfm", "pbf", "bmi", "bmr", "score", "note"],
  meta:   ["key", "value"],
};
// list 응답에 함께 실어 보낼 meta 키 (앱이 기대하는 것)
const META_KEYS = ["goals", "presets", "plan", "calc", "act"];

// ───────────────────────── 엔트리 ─────────────────────────

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.secret !== SECRET) return out({ ok: false, error: "bad-secret" });

    const type = req.type === "inbody" ? "inbody" : "food";
    switch (req.action) {
      case "list":    return out(Object.assign({ ok: true }, listItems(type, req.from, req.to)));
      case "add":     return out({ ok: true, item: addItem(type, req.item) });
      case "del":     return out({ ok: true, id: delItem(type, req.id) });
      case "setMeta": return out({ ok: true, key: req.key, value: setMeta(req.key, req.value) });
      default:        return out({ ok: false, error: "unknown-action: " + req.action });
    }
  } catch (err) {
    return out({ ok: false, error: String(err && err.message || err) });
  }
}

// 브라우저에서 URL 을 그냥 열었을 때 살아있는지 확인용 (데이터는 주지 않음)
function doGet() {
  return out({ ok: true, alive: true, note: "POST 로 요청하세요" });
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

// ───────────────────────── 시트 도우미 ─────────────────────────

function sheetOf(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]);
    sh.setFrozenRows(1);
    // 날짜가 Date 객체로 변환되지 않게 텍스트 서식
    const di = SHEETS[name].indexOf("date");
    if (di >= 0) sh.getRange(2, di + 1, sh.getMaxRows() - 1, 1).setNumberFormat("@");
  }
  // 헤더가 비어 있으면(사용자가 지웠다면) 다시 채운다
  if (!String(sh.getRange(1, 1).getValue())) {
    sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]);
  }
  return sh;
}

function headers(sh) {
  const w = Math.max(1, sh.getLastColumn());
  return sh.getRange(1, 1, 1, w).getValues()[0].map(String);
}

function asDateStr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(v).slice(0, 10);
}

function rowToObj(hdr, row) {
  const o = {};
  hdr.forEach(function (h, i) {
    if (!h) return;
    o[h] = h === "date" ? asDateStr(row[i]) : row[i];
  });
  return o;
}

function objToRow(hdr, obj) {
  return hdr.map(function (h) {
    if (!h) return "";
    return obj[h] === undefined || obj[h] === null ? "" : obj[h];
  });
}

// ───────────────────────── 기록 ─────────────────────────

function listItems(type, from, to) {
  const sh = sheetOf(type);
  const hdr = headers(sh);
  const last = sh.getLastRow();
  let items = [];
  if (last > 1) {
    items = sh.getRange(2, 1, last - 1, hdr.length).getValues()
      .map(function (r) { return rowToObj(hdr, r); })
      .filter(function (o) {
        if (!o.id) return false;
        if (from && o.date < from) return false;
        if (to && o.date > to) return false;
        return true;
      });
  }
  const res = { items: items };
  // 식단 목록을 받아갈 때 설정값도 같이 실어 보낸다 (앱이 그렇게 기대함)
  if (type === "food") {
    const meta = getAllMeta();
    META_KEYS.forEach(function (k) { if (meta[k] !== undefined) res[k] = meta[k]; });
  }
  return res;
}

function addItem(type, item) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheetOf(type);
    const hdr = headers(sh);
    const rec = Object.assign({}, item, { id: Utilities.getUuid() });
    rec.date = asDateStr(rec.date);
    sh.appendRow(objToRow(hdr, rec));
    return rec;
  } finally {
    lock.releaseLock();
  }
}

function delItem(type, id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheetOf(type);
    const hdr = headers(sh);
    const col = hdr.indexOf("id") + 1;
    const last = sh.getLastRow();
    if (col < 1 || last < 2) return id;
    const ids = sh.getRange(2, col, last - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) {
      if (String(ids[i][0]) === String(id)) { sh.deleteRow(i + 2); break; }
    }
    return id;
  } finally {
    lock.releaseLock();
  }
}

// ───────────────────────── 설정(meta) ─────────────────────────
// key 하나에 JSON 한 덩어리. 시트 셀 한도(50,000자)를 넘으면 저장하지 않고 에러를 낸다.

function getAllMeta() {
  const sh = sheetOf("meta");
  const last = sh.getLastRow();
  const o = {};
  if (last > 1) {
    sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
      if (!r[0]) return;
      try { o[String(r[0])] = JSON.parse(r[1]); } catch (e) { /* 깨진 값은 무시 */ }
    });
  }
  return o;
}

function setMeta(key, value) {
  const json = JSON.stringify(value);
  if (json.length > 49000) throw new Error("meta '" + key + "' 가 너무 큽니다 (" + json.length + "자)");
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheetOf("meta");
    const last = sh.getLastRow();
    if (last > 1) {
      const keys = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < keys.length; i++) {
        if (String(keys[i][0]) === String(key)) {
          sh.getRange(i + 2, 2).setValue(json);
          return value;
        }
      }
    }
    sh.appendRow([key, json]);
    return value;
  } finally {
    lock.releaseLock();
  }
}

// ───────────────────────── 점검용 ─────────────────────────
// Apps Script 편집기에서 이 함수를 직접 실행하면 시트 3개가 만들어지고
// 실행 로그에 결과가 찍힙니다. 배포 전에 한 번 돌려보세요.
function setupAndTest() {
  ["food", "inbody", "meta"].forEach(sheetOf);
  const a = addItem("food", { date: "2000-01-01", meal: "간식", name: "__테스트__", kcal: 1, p: 0, c: 0, f: 0, na: 0, fib: 0, sug: 0 });
  const found = listItems("food", "2000-01-01", "2000-01-01").items.length;
  delItem("food", a.id);
  setMeta("__test__", { ok: 1 });
  Logger.log("시트 생성 완료 / 추가·조회·삭제 정상: " + (found === 1 ? "OK" : "실패"));
  Logger.log("SECRET 을 바꿨는지 확인: " + (SECRET.indexOf("여기에") === 0 ? "❌ 아직 기본값입니다" : "✅ 변경됨"));
}
