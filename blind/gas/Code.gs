/* 면접 연습 수집 — 생기부 블라인드 '면접 준비'의 학생 배포용 Apps Script
 * 시트: 설정(관리 비밀번호) · 질문세트(토큰별 질문지) · 답변(학생 제출)
 * 학생은 토큰이 든 링크(QR)로만 들어오고, 이름은 어디에도 적지 않는다.
 * 설치: 새 스프레드시트 → 확장 프로그램 > Apps Script → 이 파일과 index.html 붙여넣기
 *      → setup 실행(권한 승인) → 설정 시트 B2에 관리 비밀번호 적기
 *      → 배포 > 새 배포 > 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → 웹 앱 URL을 허브 '학생 배포' 설정에 넣기 */

const SHEETS = {
  settings: { name: "설정", header: ["항목", "값"] },
  sets: { name: "질문세트", header: ["토큰", "라벨", "제목", "면접일", "질문JSON", "만든때", "상태", "한마디"] },
  answers: { name: "답변", header: ["토큰", "회차", "번호", "섹션", "질문", "답변", "소요초", "제출때", "문항키"] },
};
const STATUS = { on: "배포중", off: "중지" };
/* 녹음 전사용 Gemini 키 — 설정 시트의 'Gemini 키' 칸이 우선이고, 비어 있으면 아래 따옴표 안의 값을 씁니다.
 * 이 파일을 Apps Script 편집기(선생님만 보는 곳)에 붙여 넣은 뒤 거기서만 채우세요. 깃허브에 올리는 파일에는 비워 둡니다. */
const GEMINI_KEY_FALLBACK = "";

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
  const labels = st.getDataRange().getValues().map(r => String(r[0]).trim());
  if (labels.indexOf("Gemini 키") < 0) st.appendRow(["Gemini 키", ""]);
  if (labels.indexOf("전사 모델") < 0) st.appendRow(["전사 모델", "gemini-2.5-flash"]);
  ensureCols_();
  return "시트를 만들었습니다. 설정 시트 B2에 관리 비밀번호를 적어 주세요.";
}
/* 예전 판에서 만든 시트에 새 열(한마디·문항키)이 없으면 머리글 끝에 붙인다 */
function ensureCols_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(k => {
    const sh = ss.getSheetByName(SHEETS[k].name); if (!sh) return;
    const head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
    SHEETS[k].header.forEach(h => { if (head.indexOf(h) < 0) { head.push(h); sh.getRange(1, head.length).setValue(h); } });
  });
}
/* 시트 한 칸은 50,000자까지. 긴 JSON은 45,000자씩 '질문JSON', '질문JSON2', '질문JSON3'… 칸에 나눠 쓰고 읽을 때 이어 붙인다 */
const CHUNK = 45000;
function chunkCols_(sh, n) {   // n개의 조각 칸이 있게 하고 열 번호 배열을 돌려준다
  let head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const cols = [];
  for (let i = 1; i <= n; i++) {
    const label = i === 1 ? "질문JSON" : "질문JSON" + i;
    let c = head.indexOf(label);
    if (c < 0) { head.push(label); sh.getRange(1, head.length).setValue(label); c = head.length - 1; }
    cols.push(c);
  }
  return cols;
}
function readJson_(sh, head, row) {
  let out = String(row[head.indexOf("질문JSON")] || "");
  for (let i = 2; ; i++) { const c = head.indexOf("질문JSON" + i); if (c < 0 || !row[c]) break; out += String(row[c]); }
  return out;
}
function colOpt_(sh, label) { const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String); return head.indexOf(label); }
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
    if (req.action === "publish") { ensureCols_(); return json_(publish_(req)); }
    if (req.action === "stop") return json_(setStatus_(req.token, STATUS.off));
    if (req.action === "resume") return json_(setStatus_(req.token, STATUS.on));
    if (req.action === "answers") return json_({ ok: true, rounds: answersFor_(req.token) });
    if (req.action === "list") return json_({ ok: true, sets: listSets_() });
    if (req.action === "transcribe") return json_({ ok: true, result: transcribe_(req.audio, req.mime, "interview", req.hint) });
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
    const json = JSON.stringify(qs), parts = []; for (let i = 0; i < json.length; i += CHUNK) parts.push(json.slice(i, i + CHUNK));
    chunkCols_(sh, Math.max(parts.length, 1));
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    const vals = { "토큰": req.token, "라벨": String(req.alias || "").replace(/\s*[가-힣]{2,4}$/, m => /반|번/.test(m) ? m : ""), "제목": String(req.title || ""), "면접일": String(req.date || ""), "만든때": new Date(), "상태": STATUS.on, "한마디": String(req.note || "").slice(0, 500) };
    head.forEach((h, i) => { const m = h.match(/^질문JSON(\d*)$/); if (m) vals[h] = parts[m[1] ? Number(m[1]) - 1 : 0] || ""; });
    const row = head.map(h => (h in vals) ? vals[h] : "");
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) if (String(rows[i][cT]) === req.token) { sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return { ok: true, updated: true, chunks: parts.length }; }
    sh.appendRow(row);
    return { ok: true, updated: false, chunks: parts.length };
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
  for (let i = 1; i < rows.length; i++) { const r = {}; head.forEach((h, j) => { if (!/^질문JSON/.test(h)) r[h] = rows[i][j] instanceof Date ? rows[i][j].toISOString() : rows[i][j]; }); out.push(r); }
  return out;
}
function answersFor_(token) {
  const sh = sheet_("answers"), rows = sh.getDataRange().getValues();
  if (rows.length < 2) return [];
  const c = {}; ["토큰", "회차", "번호", "섹션", "질문", "답변", "소요초", "제출때"].forEach(h => c[h] = col_(sh, h));
  const cK = colOpt_(sh, "문항키");
  const byRound = {};
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i]; if (String(r[c["토큰"]]) !== token) continue;
    const k = String(r[c["회차"]]);
    if (!byRound[k]) byRound[k] = { round: Number(k), submittedAt: r[c["제출때"]] instanceof Date ? r[c["제출때"]].toISOString() : String(r[c["제출때"]]), items: [] };
    byRound[k].items.push({ n: Number(r[c["번호"]]), k: cK >= 0 ? String(r[cK] || "") : "", section: String(r[c["섹션"]]), q: String(r[c["질문"]]), answer: String(r[c["답변"]]), seconds: Number(r[c["소요초"]]) || 0 });
  }
  return Object.keys(byRound).map(k => byRound[k]).sort((a, b) => a.round - b.round).map(r => { r.items.sort((a, b) => a.n - b.n); return r; });
}

/* ---------- 녹음 전사 — 설정 시트의 Gemini 키로 이 스크립트가 Gemini를 부른다. 키는 시트 밖으로 나가지 않는다 ---------- */
function geminiKey_() { return setting_("Gemini 키") || GEMINI_KEY_FALLBACK; }
function transcribe_(b64, mime, mode, hint) {
  const key = geminiKey_(); if (!key) throw new Error("설정 시트에 'Gemini 키'가 비어 있습니다.");
  if (!b64 || String(b64).length < 100) throw new Error("오디오 자료가 비어 있습니다.");
  if (String(b64).length > 26 * 1024 * 1024) throw new Error("녹음 파일이 너무 큽니다(약 19MB까지). 나눠서 올려 주세요.");
  const model = setting_("전사 모델") || "gemini-2.5-flash";
  const rules = mode === "interview"
    ? "이 녹음은 고등학교 모의면접(면접관 교사와 학생)입니다. 화자를 면접관/학생으로 구분하고 확신이 없으면 '모름'으로 적습니다."
    : "이 녹음은 학생이 혼자 면접 답변을 연습한 것입니다. 화자는 모두 학생입니다.";
  const prompt = rules + " 들리는 대로 한국어로 전사합니다. 추임새('어', '음', '그게', '진짜')와 말 더듬, 끊긴 말을 지우지 않고 그대로 적습니다 — 피드백 재료입니다. 발화가 바뀌거나 30초가 지나면 새 segment를 만들고 t에는 시작 시각을 mm:ss로 적습니다. 들리지 않는 곳은 (불명)으로 적습니다. 사람 이름·학교명이 들리면 [이름]·[학교]로 바꿉니다. notes에는 녹음 상태(잡음·끊김)만 적습니다. 출력은 JSON만." + (hint ? " 참고: " + String(hint).slice(0, 300) : "");
  const schema = { type: "object", properties: { segments: { type: "array", items: { type: "object", properties: { t: { type: "string" }, who: { type: "string" }, text: { type: "string" } }, required: ["t", "who", "text"] } }, notes: { type: "array", items: { type: "string" } } }, required: ["segments", "notes"] };
  const body = { contents: [{ parts: [{ inlineData: { mimeType: mime || "audio/mp4", data: b64 } }, { text: prompt }] }], generationConfig: { responseMimeType: "application/json", responseSchema: schema, temperature: 0.2, maxOutputTokens: 16384 } };
  const res = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/models/" + encodeURIComponent(model) + ":generateContent?key=" + encodeURIComponent(key), { method: "post", contentType: "application/json", payload: JSON.stringify(body), muteHttpExceptions: true });
  const code = res.getResponseCode(), txt = res.getContentText();
  if (code !== 200) { let m = ""; try { m = JSON.parse(txt).error.message; } catch (e) { m = txt.slice(0, 200); } throw new Error("Gemini 응답 " + code + ": " + m); }
  const j = JSON.parse(txt);
  const out = ((j.candidates || [])[0] || {}).content; const text = out && out.parts ? out.parts.map(p => p.text || "").join("") : "";
  if (!text) throw new Error("전사 결과가 비어 있습니다" + (((j.candidates || [])[0] || {}).finishReason ? " (" + j.candidates[0].finishReason + ")" : "") + ".");
  let d; try { d = JSON.parse(text); } catch (e) { const i = text.indexOf("{"), k = text.lastIndexOf("}"); d = JSON.parse(text.slice(i, k + 1)); }
  const segs = (d.segments || []).map(x => ({ t: String(x.t || ""), who: String(x.who || ""), text: String(x.text || "") })).filter(x => x.text);
  return { model, segments: segs, notes: d.notes || [], text: segs.map(x => (x.t ? "[" + x.t + "] " : "") + (x.who && mode === "interview" ? x.who + ": " : "") + x.text).join("\n") };
}
/* 학생 화면이 부르는 전사 — 토큰이 맞고 배포 중일 때만 */
function transcribeAnswer(token, b64, mime) {
  const set = getSet(token); if (!set.ok) throw new Error(set.error);
  return transcribe_(b64, mime, "self", "");
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
  const head = rows[0].map(String);
  const cT = col_(sh, "토큰"), cTi = col_(sh, "제목"), cD = col_(sh, "면접일"), cS = col_(sh, "상태"), cN = colOpt_(sh, "한마디");
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][cT]) !== token) continue;
    if (String(rows[i][cS]) !== STATUS.on) return { ok: false, error: "이 연습 링크는 닫혔습니다. 선생님께 여쭤보세요." };
    const date = rows[i][cD] instanceof Date ? Utilities.formatDate(rows[i][cD], "Asia/Seoul", "yyyy-MM-dd") : String(rows[i][cD] || "");
    const questions = JSON.parse(readJson_(sh, head, rows[i]));
    questions.forEach((q, j) => { if (q && !q.k) q.k = "q" + j; });   // 예전 판으로 올린 세트에는 문항키가 없다
    return { ok: true, title: String(rows[i][cTi]), date, questions, note: cN >= 0 ? String(rows[i][cN] || "") : "", nextRound: nextRound_(token), done: doneKeys_(token, questions), canTranscribe: !!geminiKey_() };
  }
  return { ok: false, error: "연습 링크를 찾지 못했습니다. 주소가 끊기지 않았는지 확인해 주세요." };
}
/* 지난 회차에 답한 문항(적은 답이 있거나 시간을 잰 것)의 문항키. 예전 답변 행에 문항키가 없으면 질문 글로 맞춘다 */
function doneKeys_(token, questions) {
  const sh = sheet_("answers"), rows = sh.getDataRange().getValues();
  if (rows.length < 2) return [];
  const cT = col_(sh, "토큰"), cQ = col_(sh, "질문"), cA = col_(sh, "답변"), cS = col_(sh, "소요초"), cK = colOpt_(sh, "문항키");
  const byQ = {}; (questions || []).forEach(q => { if (q && q.k) byQ[String(q.q || "").trim()] = q.k; });
  const out = {};
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i]; if (String(r[cT]) !== token) continue;
    if (!String(r[cA] || "").trim() && !(Number(r[cS]) > 0)) continue;
    const k = cK >= 0 && r[cK] ? String(r[cK]) : byQ[String(r[cQ] || "").trim()];
    if (k) out[k] = true;
  }
  return Object.keys(out);
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
  ensureCols_();   // 예전 판 답변 시트에 '문항키' 열이 없으면 붙인다
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const round = nextRound_(token), sh = sheet_("answers"), now = new Date();
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    const rows = list.map(a => { const v = { "토큰": token, "회차": round, "번호": Number(a.n) || 0, "섹션": String(a.section || ""), "질문": String(a.q || "").slice(0, 500), "답변": String(a.answer || "").slice(0, 5000), "소요초": Number(a.seconds) || 0, "제출때": now, "문항키": String(a.k || "").slice(0, 40) }; return head.map(h => (h in v) ? v[h] : ""); });
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return { ok: true, round, count: rows.length };
  } finally { lock.releaseLock(); }
}
