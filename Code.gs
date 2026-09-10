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
  // 한국 영양성분표 의무 표시 9종 + 식이섬유
  //  kcal 열량 / p 단백질 / c 탄수화물 / f 지방 / na 나트륨 / sug 당류
  //  sat 포화지방 / trans 트랜스지방 / chol 콜레스테롤 / fib 식이섬유
  food:   ["id", "date", "meal", "name", "kcal", "p", "c", "f", "na", "fib", "sug", "sat", "trans", "chol"],
  inbody: ["id", "date", "weight", "smm", "bfm", "pbf", "bmi", "bmr", "score", "note"],
  meta:   ["key", "value"],
};
// list 응답에 함께 실어 보낼 meta 키 (앱이 기대하는 것)
const META_KEYS = ["goals", "presets", "plan", "calc", "act", "sets"];

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
      case "foodSearch": return out({ ok: true, items: foodSearch(req.q, req.debug) });
      case "barcode":    return out({ ok: true, data: barcodeLookup(req.code) });
      case "dbSearch":   return out({ ok: true, items: localDbSearch(req.q) });
      case "dbCount":    return out({ ok: true, count: dbRowCount() });
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
  // 코드에 컬럼이 추가됐는데 시트에 없으면 뒤에 덧붙인다.
  // 기존 행은 그 칸이 빈 값이 되고, 앱이 빈 값을 0 으로 처리하므로 그대로 열린다
  const have = headers(sh);
  const missing = SHEETS[name].filter(function (h) { return have.indexOf(h) < 0; });
  if (missing.length) {
    sh.getRange(1, have.length + 1, 1, missing.length).setValues([missing]);
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
    // 인바디는 하루 한 줄이다. 같은 날짜가 이미 있으면 덧붙이지 않고 갈아끼운다 —
    // 여러 기기가 같은 inbox 지시를 각자 적용하면 같은 검사가 두 줄로 들어간다.
    if (type === "inbody") {
      const row = rowOfDate(sh, hdr, rec.date);
      if (row > 0) {
        const idCol = hdr.indexOf("id") + 1;
        const keep = idCol > 0 ? sh.getRange(row, idCol).getValue() : "";
        if (keep) rec.id = String(keep);          // id 는 유지해야 앱의 삭제가 계속 맞는다
        sh.getRange(row, 1, 1, hdr.length).setValues([objToRow(hdr, rec)]);
        return rec;
      }
    }
    sh.appendRow(objToRow(hdr, rec));
    return rec;
  } finally {
    lock.releaseLock();
  }
}

// 날짜 열에서 그 날짜가 있는 행 번호를 찾는다 (없으면 0)
function rowOfDate(sh, hdr, ds) {
  const col = hdr.indexOf("date") + 1;
  const last = sh.getLastRow();
  if (col < 1 || last < 2) return 0;
  const vals = sh.getRange(2, col, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (asDateStr(vals[i][0]) === ds) return i + 2;
  return 0;
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

// ───────────────────────── 식약처 식품영양성분 DB 검색 ─────────────────────────
/**
 * 식품안전나라 OpenAPI(I2790) 로 시중 제품·식품을 검색합니다.
 * 브라우저에서 직접 부르면 CORS 에 막히지만, 여기(구글 서버)에서 부르면 됩니다.
 * 키도 앱이 아니라 여기 남으므로 노출되지 않습니다.
 *
 * 설정
 *   1) https://www.foodsafetykorea.go.kr/api/openApiInfo.do 에서 무료 인증키 발급
 *   2) Apps Script ⚙️ 프로젝트 설정 → 스크립트 속성
 *        MFDS_KEY = 발급받은 인증키
 *   3) 배포 → 배포 관리 → ✏️ → 새 버전 → 배포
 */
// 공공데이터포털(data.go.kr) 의 "식품의약품안전처_식품영양성분DB정보" 를 부른다.
// 예전에 쓰던 식품안전나라 I2790 은 인증키 발급처도 응답 규격도 다른 별개 서비스다.
//   신청: data.go.kr → "식품영양성분DB정보" → 활용신청(자동승인)
//   키:   마이페이지 → 개발계정 → 일반 인증키
const MFDS_URL = "https://apis.data.go.kr/1471000/FoodNtrCpntDbInfo02/getFoodNtrCpntDbInq02";

function foodSearch(q, debug) {
  const key = PropertiesService.getScriptProperties().getProperty("MFDS_KEY");
  if (!key) throw new Error("MFDS_KEY 가 없습니다. Apps Script → 프로젝트 설정 → 스크립트 속성에 추가하세요");
  if (!q || !String(q).trim()) return [];

  const url = MFDS_URL + "?serviceKey=" + encodeURIComponent(key)
            + "&type=json&pageNo=1&numOfRows=30"
            + "&FOOD_NM_KR=" + encodeURIComponent(String(q).trim());
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const body = res.getContentText();
  if (res.getResponseCode() !== 200) throw new Error("식약처 API " + res.getResponseCode() + ": " + body.slice(0, 200));

  let j;
  try { j = JSON.parse(body); }
  catch (e) { throw new Error("응답이 JSON 이 아닙니다: " + body.slice(0, 200)); }

  // 껍데기가 {header,body} 일 때도 {response:{header,body}} 일 때도 있다
  const head = j.header || (j.response && j.response.header) || {};
  const rc = String(head.resultCode || "");
  if (rc && rc !== "00" && rc !== "0" && rc !== "INFO-000")
    throw new Error("식약처: " + rc + " " + (head.resultMsg || ""));

  const rows = mfdsItems(j);
  if (debug) return rows.slice(0, 1);        // 필드명 확인용 — 원본 한 건을 그대로 돌려준다

  const hits = rows.map(function (r) { return mapNutrient(r); }).filter(function (x) { return x.name; });
  // 식약처에 없으면 Open Food Facts 로 보완
  // (전세계 오픈 DB, 인증키 불필요. 한국 제품은 바코드로 등록된 것들이 잡힌다)
  if (!hits.length) return offSearch(q);
  return hits;
}

// items 가 배열일 때도, {item:[…]} 일 때도, 한 건만 객체로 올 때도 있다
function mfdsItems(j) {
  const b = j.body || (j.response && j.response.body) || {};
  let it = b.items;
  if (it && it.item) it = it.item;
  if (!it) return [];
  return Array.isArray(it) ? it : [it];
}

// ───────────────────────── Open Food Facts (키 불필요) ─────────────────────────

function offNum(n, mul) { const v = parseFloat(n); return isNaN(v) ? 0 : Math.round(v * (mul || 1) * 10) / 10; }

// nutriments 는 100g 기준(_100g)이다. 이름에 100g 을 박아 오해가 없게 한다
function offMap(pr) {
  const n = pr.nutriments || {};
  const brand = String(pr.brands || "").split(",")[0].trim();
  const name = String(pr.product_name_ko || pr.product_name || "").trim();
  if (!name) return null;
  return {
    name: (brand && name.indexOf(brand) < 0 ? brand + " " : "") + name,
    serving: "100g",
    kcal: offNum(n["energy-kcal_100g"]),
    p:    offNum(n.proteins_100g),
    c:    offNum(n.carbohydrates_100g),
    f:    offNum(n.fat_100g),
    sug:  offNum(n.sugars_100g),
    na:   offNum(n.sodium_100g, 1000),      // OFF 는 g 단위 → mg
    fib:  offNum(n.fiber_100g),
    src:  "OFF",
  };
}

function offSearch(q) {
  const url = "https://world.openfoodfacts.org/api/v2/search"
            + "?search_terms=" + encodeURIComponent(q)
            + "&fields=product_name,product_name_ko,brands,nutriments&page_size=20";
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true,
    headers: { "User-Agent": "bodybuilding-tracker/1.0 (personal use)" } });
  if (res.getResponseCode() !== 200) return [];
  let j; try { j = JSON.parse(res.getContentText()); } catch (e) { return []; }
  return (j.products || []).map(offMap).filter(function (x) { return x && x.kcal; });
}

// 바코드 한 건 조회 — 시중 제품은 이게 가장 확실하다
function barcodeLookup(code) {
  const c = String(code || "").replace(/[^0-9]/g, "");
  if (c.length < 8) throw new Error("바코드 숫자가 짧습니다: " + code);
  const url = "https://world.openfoodfacts.org/api/v2/product/" + c
            + ".json?fields=product_name,product_name_ko,brands,nutriments";
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true,
    headers: { "User-Agent": "bodybuilding-tracker/1.0 (personal use)" } });
  if (res.getResponseCode() === 404) throw new Error("등록되지 않은 바코드입니다 (" + c + ") — 라벨 촬영을 쓰세요");
  if (res.getResponseCode() !== 200) throw new Error("Open Food Facts " + res.getResponseCode());
  let j; try { j = JSON.parse(res.getContentText()); } catch (e) { throw new Error("응답 해석 실패"); }
  if (j.status !== 1 || !j.product) throw new Error("등록되지 않은 바코드입니다 (" + c + ") — 라벨 촬영을 쓰세요");
  const m = offMap(j.product);
  if (!m) throw new Error("제품명이 비어 있습니다 (" + c + ") — 라벨 촬영을 쓰세요");
  m.barcode = c;
  return m;
}

// 필드명이 서비스·개정에 따라 달라지므로 여러 후보를 훑는다.
// 못 찾으면 0 이 되고 앱에서는 빈칸으로 보이니, 그럴듯한 오답이 들어가지는 않는다.
function pick(r, names) {
  for (let i = 0; i < names.length; i++) {
    const v = r[names[i]];
    if (v !== undefined && v !== null && String(v).trim() !== "") return v;
  }
  return "";
}
function numOf(v) {
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? 0 : n;
}
function mapNutrient(r) {
  const name = String(pick(r, ["DESC_KOR", "FOOD_NM_KR", "PRDLST_NM", "food_name"])).trim();
  const maker = String(pick(r, ["MAKER_NM", "MAKER_NAME", "BSSH_NM", "maker_name"])).trim();
  const serv = String(pick(r, ["SERVING_SIZE", "SERVING_WT", "serving_size"])).trim();
  return {
    name: (maker && name.indexOf(maker) < 0 ? maker + " " : "") + name,
    // 숫자만 오면 g 를 붙이고, "100g" 처럼 단위가 이미 있으면 그대로 둔다
    serving: !serv ? "100g" : (/^[0-9.]+$/.test(serv) ? serv + "g" : serv),
    // AMT_NUM 번호는 data.go.kr 명세 기준이다. 예전 코드가 탄수 자리에 당류(7),
    // 당류 자리에 식이섬유(8), 나트륨 자리에 9 번을 읽고 있어 값이 서로 밀려 있었다.
    kcal: numOf(pick(r, ["AMT_NUM1", "NUTR_CONT1", "ENERC"])),
    c:    numOf(pick(r, ["AMT_NUM6", "NUTR_CONT2", "CHOCDF"])),
    p:    numOf(pick(r, ["AMT_NUM3", "NUTR_CONT3", "PROCNT"])),
    f:    numOf(pick(r, ["AMT_NUM4", "NUTR_CONT4", "FATCE"])),
    sug:  numOf(pick(r, ["AMT_NUM7", "NUTR_CONT5", "SUGAR"])),
    na:   numOf(pick(r, ["AMT_NUM13", "NUTR_CONT6", "NAT"])),
    fib:  numOf(pick(r, ["AMT_NUM8", "NUTR_CONT7", "FIBTG"])),
    chol: numOf(pick(r, ["AMT_NUM23", "CHOLE"])),
    sat:  numOf(pick(r, ["AMT_NUM24", "FASAT"])),
    trans: numOf(pick(r, ["AMT_NUM25", "FATRN"])),
  };
}

// ═══════════ 식약처 통합식품영양성분DB 일괄 적재 ═══════════
/**
 * 식약처가 파일로 공개하는 영양성분 DB 전체를 시트에 넣습니다.
 * 한 번 넣어두면 인증키·인터넷 없이 수만 건을 검색할 수 있습니다.
 *
 * 【사용법】
 *  1) 아래에서 파일을 받습니다 (셋 중 아무거나, 여러 개 넣어도 됩니다)
 *     · 공공데이터포털 data.go.kr → "식품영양성분" 검색 → 파일데이터 → CSV 다운로드
 *     · 식품안전나라 → 전문정보 → 식품영양성분DB → 다운로드
 *     · 「통합식품영양성분DB」 — 음식/가공식품/원재료성식품 3종
 *  2) 받은 파일을 구글 드라이브에 올립니다 (드래그하면 끝)
 *     · CSV 그대로 올려도 되고, 엑셀이면 드라이브에서 "Google 스프레드시트로 열기" 후 저장
 *  3) 아래 FILE_NAME 을 그 파일 이름으로 바꾸고 importFoodDb 를 실행합니다
 *
 * 컬럼 이름은 데이터셋마다 달라서(식품명/제품명/DESC_KOR …) 한국어 키워드로
 * 자동 매칭합니다. 못 찾은 항목은 0 이 되고, 실행 로그에 매칭 결과가 찍힙니다.
 */
function importFoodDb() {
  const FILE_NAME = "여기에-드라이브에-올린-파일이름";   // 예: "통합식품영양성분DB.csv"
  const files = DriveApp.getFilesByName(FILE_NAME);
  if (!files.hasNext()) throw new Error("드라이브에 '" + FILE_NAME + "' 파일이 없습니다");
  const file = files.next();

  let rows;
  const mime = file.getMimeType();
  if (mime === MimeType.GOOGLE_SHEETS) {
    const sh = SpreadsheetApp.openById(file.getId()).getSheets()[0];
    rows = sh.getDataRange().getValues();
  } else {
    // CSV. 한글 인코딩이 깨지면 EUC-KR 로 다시 시도한다
    let txt = file.getBlob().getDataAsString("UTF-8");
    if (txt.indexOf("\uFFFD") >= 0) txt = file.getBlob().getDataAsString("EUC-KR");
    rows = Utilities.parseCsv(txt);
  }
  if (!rows || rows.length < 2) throw new Error("데이터가 비었습니다");

  const map = matchColumns(rows[0]);
  Logger.log("컬럼 매칭 결과:");
  Object.keys(map).forEach(function (k) {
    Logger.log("  " + k + " → " + (map[k] >= 0 ? "'" + rows[0][map[k]] + "'" : "❌ 못 찾음"));
  });
  if (map.name < 0) throw new Error("식품명 컬럼을 찾지 못했습니다. 헤더: " + rows[0].slice(0, 15).join(" | "));

  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const nm = String(r[map.name] || "").trim();
    if (!nm) continue;
    const maker = map.maker >= 0 ? String(r[map.maker] || "").trim() : "";
    out.push([
      (maker && nm.indexOf(maker) < 0 ? maker + " " : "") + nm,
      map.serving >= 0 ? String(r[map.serving] || "").trim() : "",
      cell(r, map.kcal), cell(r, map.p), cell(r, map.c), cell(r, map.f),
      cell(r, map.na), cell(r, map.sug), cell(r, map.sat), cell(r, map.trans),
      cell(r, map.chol), cell(r, map.fib),
    ]);
  }
  if (!out.length) throw new Error("옮길 행이 없습니다");

  const sh = dbSheet();
  sh.getRange(2, 1, Math.max(1, sh.getMaxRows() - 1), DB_COLS.length).clearContent();
  if (sh.getMaxRows() < out.length + 1) sh.insertRowsAfter(sh.getMaxRows(), out.length + 1 - sh.getMaxRows());
  // 한 번에 다 쓰면 시간초과가 나므로 나눠서 넣는다
  const CH = 2000;
  for (let i = 0; i < out.length; i += CH) {
    const part = out.slice(i, i + CH);
    sh.getRange(2 + i, 1, part.length, DB_COLS.length).setValues(part);
  }
  Logger.log("✅ " + out.length + "건을 fooddb 시트에 넣었습니다");
  return out.length;
}

const DB_COLS = ["name", "serving", "kcal", "p", "c", "f", "na", "sug", "sat", "trans", "chol", "fib"];

function dbSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("fooddb");
  if (!sh) {
    sh = ss.insertSheet("fooddb");
    sh.getRange(1, 1, 1, DB_COLS.length).setValues([DB_COLS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function cell(row, idx) {
  if (idx < 0) return 0;
  const n = parseFloat(String(row[idx]).replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? 0 : n;
}

// 헤더 이름을 키워드로 매칭. 데이터셋마다 이름이 달라도 붙는다
function matchColumns(hdr) {
  const H = hdr.map(function (h) { return String(h).replace(/\s|\(|\)|_/g, ""); });
  function find(cands, avoid) {
    for (let c = 0; c < cands.length; c++) {
      for (let i = 0; i < H.length; i++) {
        if (H[i].indexOf(cands[c]) < 0) continue;
        if (avoid && avoid.some(function (a) { return H[i].indexOf(a) >= 0; })) continue;
        return i;
      }
    }
    return -1;
  }
  return {
    name:    find(["식품명", "제품명", "DESCKOR", "FOODNM", "품목명"]),
    maker:   find(["제조사", "업체명", "MAKERNAME", "BSSHNM", "제조업소"]),
    serving: find(["1회제공량", "영양성분함량기준량", "SERVINGWT", "기준량", "내용량"]),
    kcal:    find(["에너지", "열량", "칼로리", "kcal", "NUTRCONT1"]),
    p:       find(["단백질", "PROCNT", "NUTRCONT3"]),
    c:       find(["탄수화물", "CHOCDF", "NUTRCONT2"]),
    f:       find(["지방", "FATCE", "NUTRCONT4"], ["포화", "트랜스"]),
    na:      find(["나트륨", "NAT", "NUTRCONT6"]),
    sug:     find(["당류", "SUGAR", "NUTRCONT5"]),
    sat:     find(["포화지방", "FASAT"]),
    trans:   find(["트랜스지방", "FATRN"]),
    chol:    find(["콜레스테롤", "CHOLE"]),
    fib:     find(["식이섬유", "FIBTG"]),
  };
}

function dbRowCount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("fooddb");
  return sh ? Math.max(0, sh.getLastRow() - 1) : 0;
}

// 시트에 적재된 DB 검색. 이름 컬럼만 먼저 훑어 후보 행을 찾고 그 행만 상세 조회한다
// (수만 행 전체를 읽으면 느리다)
function localDbSearch(q) {
  const key = String(q || "").trim();
  if (!key) return [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("fooddb");
  const last = sh ? sh.getLastRow() : 0;
  if (!sh || last < 2) return [];

  const names = sh.getRange(2, 1, last - 1, 1).getValues();
  const hitRows = [];
  for (let i = 0; i < names.length && hitRows.length < 40; i++) {
    if (String(names[i][0]).indexOf(key) >= 0) hitRows.push(i + 2);
  }
  return hitRows.map(function (r) {
    const v = sh.getRange(r, 1, 1, DB_COLS.length).getValues()[0];
    const o = { src: "식약처" };
    DB_COLS.forEach(function (c, i) { o[c] = i < 2 ? v[i] : (+v[i] || 0); });
    if (!o.serving) o.serving = "100g";
    return o;
  });
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
  const props = PropertiesService.getScriptProperties();
  Logger.log("제품 검색(MFDS_KEY): " + (props.getProperty("MFDS_KEY") ? "✅ 설정됨" : "⚠️ 없음 — API 검색만 안 됨"));
  Logger.log("내장 식품DB(fooddb 시트): " + (dbRowCount() ? "✅ " + dbRowCount() + "건" : "⚠️ 비어 있음 — importFoodDb 실행하면 채워집니다"));
}

// 식약처 검색이 되는지, 필드명이 예상과 맞는지 확인. 실행 로그를 저에게 보여주시면
// 필드가 다를 경우 매핑을 정확히 고칠 수 있습니다.
function testFoodSearch() {
  const Q = "닭가슴살";        // ← 여기를 "썬칩" 등으로 바꿔서 실제 커버리지를 확인하세요
  Logger.log("── 검색어: " + Q + " ──");
  Logger.log("[매핑된 결과]");
  Logger.log(JSON.stringify(foodSearch(Q).slice(0, 5), null, 2));
  Logger.log("[식약처 원본 필드명]");
  Logger.log(JSON.stringify(foodSearch(Q, true), null, 2));
}

// 바코드 조회 확인. 제품 뒷면 바코드 숫자를 그대로 넣고 실행하세요
function testBarcode() {
  const CODE = "8801117370987";     // ← 여기를 실제 바코드로 바꾸세요
  try { Logger.log(JSON.stringify(barcodeLookup(CODE), null, 2)); }
  catch (e) { Logger.log("실패: " + e.message); }
}
