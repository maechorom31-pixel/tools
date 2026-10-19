/* 면접 연습 수집 — 생기부 블라인드 '면접 준비'의 학생 배포용 Apps Script
 * 시트: 설정(관리 비밀번호) · 질문세트(토큰별 질문지) · 답변(학생 제출)
 * 학생은 토큰이 든 링크(QR)로만 들어오고, 이름은 어디에도 적지 않는다.
 * 설치: 새 스프레드시트 → 확장 프로그램 > Apps Script → 이 파일과 index.html 붙여넣기
 *      → setup 실행(권한 승인) → 설정 시트 B2에 관리 비밀번호 적기
 *      → 배포 > 새 배포 > 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → 웹 앱 URL을 허브 '학생 배포' 설정에 넣기 */

const SHEETS = {
  settings: { name: "설정", header: ["항목", "값"] },
  sets: { name: "질문세트", header: ["토큰", "라벨", "제목", "면접일", "질문JSON", "만든때", "상태"] },
  answers: { name: "답변", header: ["토큰", "회차", "번호", "섹션", "질문", "답변", "소요초", "제출때"] },
};
const STATUS = { on: "배포중", off: "중지" };

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(k => {
    const def = SHEETS[k];
    let sh = ss.getSheetByName(def.name);
    if (!sh) { sh = ss.insertSheet(def.name); sh.appendRow(def.header); sh.setFrozenRows(1); }
  });
  const st = ss.getSheetByName(SHEETS.settings.name);
  if (st.getLastRow() < 2) st.appendRow(["관리 비밀번호", ""]);
  if (st.getLastRow() < 3) st.appendRow(["앱 제목", "면접 연습"]);
  return "시트를 만들었습니다. 설정 시트 B2에 관리 비밀번호를 적어 주세요.";
}
function sheet_(k) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS[k].name);
  if (!sh) throw new Error("시트 '" + SHEETS[k].name + "'이 없습니다. setup을 먼저 실행해 주세요.");
  return sh;
}
/* 설정 시트를 헤더 라벨로 읽는다 (칸 위치에 기대지 않음) */
function setting_(label) {
  const rows = sheet_("settings").getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (String(rows[i][0]).trim() === label) return String(rows[i][1]).trim();
  return "";
}
function col_(sh, label) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const i = head.indexOf(label);
  if (i < 0) throw new Error("시트 '" + sh.getName() + "'에 '" + label + "' 열이 없습니다.");
  return i;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- 교사 허브 → 여기 (POST, text/plain JSON) ---------- */
function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: "요청을 읽지 못했습니다." }); }
  const pw = setting_("관리 비밀번호");
  if (!pw) return json_({ ok: false, error: "설정 시트 B2에 관리 비밀번호가 비어 있습니다." });
  if (String(req.pw || "") !== pw) return json_({ ok: false, error: "관리 비밀번호가 다릅니다." });
  try {
    if (req.action === "ping") return json_({ ok: true, title: setting_("앱 제목") || "면접 연습" });
    if (req.action === "publish") return json_(publish_(req));
    if (req.action === "stop") return json_(setStatus_(req.token, STATUS.off));
    if (req.action === "resume") return json_(setStatus_(req.token, STATUS.on));
    if (req.action === "answers") return json_({ ok: true, rounds: answersFor_(req.token) });
    if (req.action === "list") return json_({ ok: true, sets: listSets_() });
    return json_({ ok: false, error: "모르는 요청: " + req.action });
  } catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); }
}
function publish_(req) {
  if (!/^[A-Za-z0-9_-]{16,64}$/.test(String(req.token || ""))) throw new Error("토큰 형식이 맞지 않습니다.");
  const qs = Array.isArray(req.questions) ? req.questions : [];
  if (!qs.length) throw new Error("질문이 없습니다.");
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = sheet_("sets"), cT = col_(sh, "토큰");
    const row = [req.token, String(req.alias || ""), String(req.title || ""), String(req.date || ""), JSON.stringify(qs), new Date(), STATUS.on];
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) if (String(rows[i][cT]) === req.token) { sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return { ok: true, updated: true }; }
    sh.appendRow(row);
    return { ok: true, updated: false };
  } finally { lock.releaseLock(); }
}
function setStatus_(token, status) {
  const sh = sheet_("sets"), cT = col_(sh, "토큰"), cS = col_(sh, "상태");
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (String(rows[i][cT]) === token) { sh.getRange(i + 1, cS + 1).setValue(status); return { ok: true, status }; }
  throw new Error("그 토큰의 질문지가 없습니다.");
}
function listSets_() {
  const sh = sheet_("sets"), rows = sh.getDataRange().getValues(), head = rows[0].map(String);
  const out = [];
  for (let i = 1; i < rows.length; i++) { const r = {}; head.forEach((h, j) => { if (h !== "질문JSON") r[h] = rows[i][j] instanceof Date ? rows[i][j].toISOString() : rows[i][j]; }); out.push(r); }
  return out;
}
function answersFor_(token) {
  const sh = sheet_("answers"), rows = sh.getDataRange().getValues();
  if (rows.length < 2) return [];
  const c = {}; ["토큰", "회차", "번호", "섹션", "질문", "답변", "소요초", "제출때"].forEach(h => c[h] = col_(sh, h));
  const byRound = {};
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i]; if (String(r[c["토큰"]]) !== token) continue;
    const k = String(r[c["회차"]]);
    if (!byRound[k]) byRound[k] = { round: Number(k), submittedAt: r[c["제출때"]] instanceof Date ? r[c["제출때"]].toISOString() : String(r[c["제출때"]]), items: [] };
    byRound[k].items.push({ n: Number(r[c["번호"]]), section: String(r[c["섹션"]]), q: String(r[c["질문"]]), answer: String(r[c["답변"]]), seconds: Number(r[c["소요초"]]) || 0 });
  }
  return Object.keys(byRound).map(k => byRound[k]).sort((a, b) => a.round - b.round).map(r => { r.items.sort((a, b) => a.n - b.n); return r; });
}

/* ---------- 학생 페이지 (GET ?t=토큰) ---------- */
function doGet(e) {
  const token = String((e && e.parameter && e.parameter.t) || "").trim();
  const t = HtmlService.createTemplateFromFile("index");
  t.token = /^[A-Za-z0-9_-]{16,64}$/.test(token) ? token : "";
  t.appTitle = setting_("앱 제목") || "면접 연습";
  return t.evaluate().setTitle(t.appTitle).addMetaTag("viewport", "width=device-width, initial-scale=1").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
/* 학생 화면이 부르는 함수 — 토큰이 맞고 배포중일 때만 질문을 준다. 라벨(반·번호)은 주지 않는다 */
function getSet(token) {
  const sh = sheet_("sets"), rows = sh.getDataRange().getValues();
  const cT = col_(sh, "토큰"), cTi = col_(sh, "제목"), cD = col_(sh, "면접일"), cQ = col_(sh, "질문JSON"), cS = col_(sh, "상태");
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][cT]) !== token) continue;
    if (String(rows[i][cS]) !== STATUS.on) return { ok: false, error: "이 연습 링크는 닫혔습니다. 선생님께 여쭤보세요." };
    const date = rows[i][cD] instanceof Date ? Utilities.formatDate(rows[i][cD], "Asia/Seoul", "yyyy-MM-dd") : String(rows[i][cD] || "");
    return { ok: true, title: String(rows[i][cTi]), date, questions: JSON.parse(rows[i][cQ]), nextRound: nextRound_(token) };
  }
  return { ok: false, error: "연습 링크를 찾지 못했습니다. 주소가 끊기지 않았는지 확인해 주세요." };
}
function nextRound_(token) {
  const sh = sheet_("answers"), rows = sh.getDataRange().getValues();
  if (rows.length < 2) return 1;
  const cT = col_(sh, "토큰"), cR = col_(sh, "회차");
  let max = 0;
  for (let i = 1; i < rows.length; i++) if (String(rows[i][cT]) === token) max = Math.max(max, Number(rows[i][cR]) || 0);
  return max + 1;
}
function saveAnswers(token, answers) {
  if (!/^[A-Za-z0-9_-]{16,64}$/.test(String(token || ""))) throw new Error("토큰이 맞지 않습니다.");
  const set = getSet(token); if (!set.ok) throw new Error(set.error);
  const list = Array.isArray(answers) ? answers.filter(a => a && (String(a.answer || "").trim() || Number(a.seconds) > 0)) : [];
  if (!list.length) throw new Error("적은 답변이 없습니다.");
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const round = nextRound_(token), sh = sheet_("answers"), now = new Date();
    const rows = list.map(a => [token, round, Number(a.n) || 0, String(a.section || ""), String(a.q || "").slice(0, 500), String(a.answer || "").slice(0, 5000), Number(a.seconds) || 0, now]);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return { ok: true, round, count: rows.length };
  } finally { lock.releaseLock(); }
}
