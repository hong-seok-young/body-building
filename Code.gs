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
      case "ocr":     return out({ ok: true, data: ocr(req.image, req.kind) });
      case "foodSearch": return out({ ok: true, items: foodSearch(req.q, req.debug) });
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

// ───────────────────────── 사진 판독 (OCR) ─────────────────────────
/**
 * 인바디 결과지·식품 영양정보표 사진을 읽어 숫자로 돌려줍니다.
 * Gemini 를 씁니다 — 표 레이아웃을 알아서 이해하므로 인바디 기종이 달라도 됩니다.
 *
 * 설정 (키를 코드에 넣지 않는 이유: 배포 버전에 그대로 박혀 남습니다)
 *   1) https://aistudio.google.com/apikey 에서 무료 API 키 발급
 *   2) Apps Script 좌측 ⚙️ 프로젝트 설정 → 스크립트 속성 → 속성 추가
 *        GEMINI_KEY   = 발급받은 키
 *        GEMINI_MODEL = (선택) 기본값은 아래 DEFAULT_MODEL. 모델을 못 찾는다는
 *                       에러가 나면 AI Studio 에서 쓸 수 있는 모델명으로 바꾸세요
 */
const DEFAULT_MODEL = "gemini-2.5-flash";

const OCR_PROMPT = {
  inbody:
    "이 사진은 인바디(InBody) 체성분 분석 결과지다. 아래 스키마의 JSON 만 출력해라.\n" +
    "숫자만 넣고 단위·쉼표는 빼라. 사진에서 확실히 찾을 수 없는 항목은 null 로 둬라.\n" +
    "추측하지 마라 — 안 보이면 null 이다.\n" +
    '{"date":"검사일시의 날짜를 YYYY-MM-DD 로","weight":체중kg,"smm":골격근량kg,' +
    '"bfm":체지방량kg,"pbf":체지방률퍼센트,"bmi":BMI,"bmr":기초대사량kcal,"score":인바디점수}',
  label:
    "이 사진은 식품 포장의 영양정보표다. 아래 스키마의 JSON 만 출력해라.\n" +
    "name 은 브랜드+제품명으로 30자 이내. serving 은 기준 표기(예: '100g', '1스쿱(38g)').\n" +
    "영양성분은 그 serving 1회분 기준으로 맞춰라. 표가 100g당으로만 적혀 있고 총 내용량이 다르면\n" +
    "총 내용량 기준으로 환산해서 넣고 serving 에 총 내용량을 적어라.\n" +
    "없는 항목은 null. 추측하지 마라.\n" +
    '{"name":"","serving":"","kcal":,"p":단백질g,"c":탄수화물g,"f":지방g,' +
    '"na":나트륨mg,"fib":식이섬유g,"sug":당류g}',
};

function ocr(imageB64, kind) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty("GEMINI_KEY");
  if (!key) throw new Error("GEMINI_KEY 가 없습니다. Apps Script → 프로젝트 설정 → 스크립트 속성에 추가하세요");
  if (!imageB64) throw new Error("이미지가 비었습니다");
  const prompt = OCR_PROMPT[kind === "inbody" ? "inbody" : "label"];
  const model = props.getProperty("GEMINI_MODEL") || DEFAULT_MODEL;

  const res = UrlFetchApp.fetch(
    "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent",
    {
      method: "post",
      contentType: "application/json",
      headers: { "x-goog-api-key": key },     // 키를 URL 에 넣지 않는다 (로그에 남음)
      muteHttpExceptions: true,
      payload: JSON.stringify({
        contents: [{ parts: [
          { text: prompt },
          { inline_data: { mime_type: "image/jpeg", data: imageB64 } },
        ] }],
        generationConfig: { temperature: 0, responseMimeType: "application/json" },
      }),
    });

  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code !== 200) {
    let msg = body.slice(0, 300);
    try { msg = JSON.parse(body).error.message; } catch (e) {}
    throw new Error("Gemini " + code + ": " + msg);
  }
  let text;
  try {
    text = JSON.parse(body).candidates[0].content.parts[0].text;
  } catch (e) {
    throw new Error("응답을 해석할 수 없습니다: " + body.slice(0, 200));
  }
  let data;
  try {
    data = JSON.parse(text);
  } catch (e) {
    // 혹시 코드펜스나 설명이 섞여 오면 첫 { ... } 만 건져낸다
    const m = text.match(/\{[\s\S]*\}/);
    if (!m) throw new Error("JSON 이 아닌 응답: " + text.slice(0, 200));
    data = JSON.parse(m[0]);
  }
  return data;
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
function foodSearch(q, debug) {
  const key = PropertiesService.getScriptProperties().getProperty("MFDS_KEY");
  if (!key) throw new Error("MFDS_KEY 가 없습니다. Apps Script → 프로젝트 설정 → 스크립트 속성에 추가하세요");
  if (!q || !String(q).trim()) return [];

  const url = "https://openapi.foodsafetykorea.go.kr/api/" + encodeURIComponent(key)
            + "/I2790/json/1/30/DESC_KOR=" + encodeURIComponent(String(q).trim());
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error("식약처 API " + res.getResponseCode());

  let j;
  try { j = JSON.parse(res.getContentText()); }
  catch (e) { throw new Error("응답이 JSON 이 아닙니다: " + res.getContentText().slice(0, 200)); }

  // 응답 껍데기가 서비스마다 달라서 첫 번째 객체에서 row 를 찾아낸다
  let box = j.I2790 || null;
  if (!box) for (const k in j) { if (j[k] && (j[k].row || j[k].RESULT)) { box = j[k]; break; } }
  if (!box) throw new Error("예상과 다른 응답: " + JSON.stringify(j).slice(0, 200));

  const code = box.RESULT && box.RESULT.CODE;
  if (code && code !== "INFO-000") {
    if (String(code).indexOf("INFO-200") === 0) return [];        // 검색 결과 없음
    throw new Error("식약처: " + code + " " + (box.RESULT.MSG || ""));
  }
  const rows = box.row || [];
  if (debug) return rows.slice(0, 1);        // 필드명 확인용 — 원본 한 건을 그대로 돌려준다

  return rows.map(function (r) { return mapNutrient(r); }).filter(function (x) { return x.name; });
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
  const maker = String(pick(r, ["MAKER_NAME", "BSSH_NM", "maker_name"])).trim();
  const serv = String(pick(r, ["SERVING_WT", "SERVING_SIZE", "serving_size"])).trim();
  return {
    name: (maker && name.indexOf(maker) < 0 ? maker + " " : "") + name,
    serving: serv ? serv + "g" : "100g",
    kcal: numOf(pick(r, ["NUTR_CONT1", "AMT_NUM1", "ENERC"])),
    c:    numOf(pick(r, ["NUTR_CONT2", "AMT_NUM7", "CHOCDF"])),
    p:    numOf(pick(r, ["NUTR_CONT3", "AMT_NUM3", "PROCNT"])),
    f:    numOf(pick(r, ["NUTR_CONT4", "AMT_NUM4", "FATCE"])),
    sug:  numOf(pick(r, ["NUTR_CONT5", "AMT_NUM8", "SUGAR"])),
    na:   numOf(pick(r, ["NUTR_CONT6", "AMT_NUM9", "NAT"])),
    fib:  numOf(pick(r, ["NUTR_CONT7", "AMT_NUM10", "FIBTG"])),
  };
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
  Logger.log("사진 판독(GEMINI_KEY): " + (props.getProperty("GEMINI_KEY") ? "✅ 설정됨" : "⚠️ 없음 — 사진 판독만 안 됨"));
  Logger.log("제품 검색(MFDS_KEY): " + (props.getProperty("MFDS_KEY") ? "✅ 설정됨" : "⚠️ 없음 — 제품 검색만 안 됨"));
}

// 식약처 검색이 되는지, 필드명이 예상과 맞는지 확인. 실행 로그를 저에게 보여주시면
// 필드가 다를 경우 매핑을 정확히 고칠 수 있습니다.
function testFoodSearch() {
  Logger.log("── 매핑된 결과 ──");
  Logger.log(JSON.stringify(foodSearch("닭가슴살").slice(0, 3), null, 2));
  Logger.log("── 원본 필드명 확인용 ──");
  Logger.log(JSON.stringify(foodSearch("닭가슴살", true), null, 2));
}

// 사진 판독이 되는지 실제로 한 번 확인. 인바디 사진을 구글 드라이브에 올리고
// 파일 ID 를 넣어 실행하면 판독 결과가 실행 로그에 찍힙니다.
function testOcr() {
  const FILE_ID = "여기에-드라이브-이미지-파일-ID";
  const blob = DriveApp.getFileById(FILE_ID).getBlob();
  Logger.log(JSON.stringify(ocr(Utilities.base64Encode(blob.getBytes()), "inbody"), null, 2));
}
