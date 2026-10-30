/**
 * 연수 설문 관리 — 구글 시트 + Apps Script 웹앱
 * Prompted by kidsreturn_79 · © 2026.10
 *
 * 이 스프레드시트(관리용)
 *   설문목록 : 만든 설문 목록 (설문ID, 제목, 상태, 응답 수, 설문 시트 주소…)
 *
 * 설문마다 따로 만들어지는 스프레드시트 (드라이브 「연수 설문」 폴더)
 *   설정   : 설문 설정(JSON)
 *   응답자 : 한 줄 = 응답자 1명 (기본정보, 분반 선택, 점수 요약, 서술형)
 *   점수   : 한 줄 = 점수 1개 (피벗/통계용)
 *
 * 웹앱 주소
 *   .../exec?s=설문ID    → 그 설문의 응답 화면
 *   .../exec             → 진행 중인 설문 목록
 *   .../exec?page=admin  → 설문 관리(관리자 비밀번호 필요)
 *
 * 다른 관리자에게 줄 때: 이 스프레드시트의 사본 링크(.../copy)를 주면
 * 그 사람 드라이브에 통째로 복사되고, ① 초기 설정 + 배포만 하면 자기 시트로 운영한다.
 */

var SHEET_CFG = '설정';
var SHEET_RESP = '응답자';
var SHEET_SCORE = '점수';
var SHEET_REG = '설문목록';
var FOLDER_NAME = '연수 설문';
var REG_HEADERS = ['설문ID', '제목', '과정', '기수', '상태', '응답 수', '만든 날짜', '마지막 응답', '시트 ID', '시트 주소', '마감 예정'];
var STATUS_OPEN = '진행 중';
var STATUS_CLOSED = '마감';
var EDIT_KEY_HEADER = '수정키'; // 응답 수정용 열쇠. 관리 화면·엑셀에는 내보내지 않는다.

var SCORE_BASE_HEADERS = ['응답ID', '제출시각', '기수', '영역', '일자', '시간', '과목', '강사', '분반수업', '문항', '점수', '응답'];

/* ───────── 시트 메뉴 ───────── */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('설문 도구')
    .addItem('① 초기 설정', 'setup')
    .addItem('관리자 비밀번호 변경', 'changeAdminKey')
    .addItem('웹앱 주소 보기', 'showUrls')
    .addItem('다른 관리자에게 공유하는 방법', 'showShareHelp')
    .addToUi();
}

/**
 * 처음 한 번 실행. 사본(다른 관리자)에서도 이것만 실행하면 된다.
 * 원본에서 복사된 설문 목록·비밀번호는 이 설치에서 쓰지 않도록 초기화한다.
 */
function setup() {
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { /* 편집기에서 실행: 비밀번호는 메뉴에서 */ }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getScriptProperties();
  var thisScript = ScriptApp.getScriptId();
  var prevScript = props.getProperty('INSTALL_SCRIPT');
  // 다른 시트에서 복사돼 온 설정이면 새로 설치. (같은 시트에서 예전 버전을 업그레이드할 때는 비밀번호 유지)
  var freshInstall = props.getProperty('SHEET_ID') !== ss.getId() || (!!prevScript && prevScript !== thisScript);

  if (freshInstall) {
    props.deleteProperty('ADMIN_KEY');
    props.deleteProperty('FOLDER_ID');
  }
  props.setProperty('SHEET_ID', ss.getId());
  props.setProperty('INSTALL_SCRIPT', thisScript);

  // 설문목록 시트
  var reg = ss.getSheetByName(SHEET_REG);
  if (!reg) {
    var blank = ss.getSheets().filter(function (s) { return s.getLastRow() === 0 && s.getLastColumn() === 0; })[0];
    reg = blank ? blank.setName(SHEET_REG) : ss.insertSheet(SHEET_REG, 0);
  }
  if (freshInstall && reg.getLastRow() > 1) reg.getRange(2, 1, reg.getLastRow() - 1, reg.getMaxColumns()).clearContent();
  reg.getRange(1, 1, 1, REG_HEADERS.length).setValues([REG_HEADERS]).setFontWeight('bold');
  reg.setFrozenRows(1);

  // 예전(설문 1개) 방식에서 쓰던 빈 시트 정리
  [SHEET_CFG, SHEET_RESP, SHEET_SCORE].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (sh && ss.getSheets().length > 1 && (n === SHEET_CFG || sh.getLastRow() <= 1)) ss.deleteSheet(sh);
  });

  // 폴더 만들고 관리용 시트도 그 안으로
  var folder = folder_();
  try { DriveApp.getFileById(ss.getId()).moveTo(folder); } catch (e) { /* 이동 권한이 없으면 그대로 둔다 */ }

  if (!ui) { Logger.log('준비 완료. 비밀번호는 시트 메뉴 [설문 도구 → 관리자 비밀번호 변경]에서 설정하세요.'); return; }
  if (!props.getProperty('ADMIN_KEY')) {
    var r = ui.prompt('관리자 비밀번호 설정', '설문 관리 화면에 들어갈 때 쓸 비밀번호를 입력하세요.', ui.ButtonSet.OK);
    var key = (r.getResponseText() || '').trim();
    if (!key) { ui.alert('비밀번호가 비어 있어 설정하지 않았어요. 다시 실행해 주세요.'); return; }
    props.setProperty('ADMIN_KEY', key);
  }
  ui.alert('초기 설정 완료!\n\n드라이브에 「' + FOLDER_NAME + '」 폴더를 만들었어요.\n\n' +
    '다음 단계\n1) Apps Script에서 [배포 → 새 배포 → 웹 앱] (실행 사용자: 나 / 액세스 권한: 모든 사용자)\n' +
    '2) 시트 메뉴 [설문 도구 → 웹앱 주소 보기]로 내 관리자 페이지 주소 확인');
}

function changeAdminKey() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('관리자 비밀번호 변경', '새 비밀번호를 입력하세요.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var key = (r.getResponseText() || '').trim();
  if (!key) return;
  var props = PropertiesService.getScriptProperties();
  props.setProperty('ADMIN_KEY', key);
  props.setProperty('INSTALL_SCRIPT', ScriptApp.getScriptId());
  ui.alert('변경했어요.');
}

// ★ 응답·관리 화면(GitHub Pages) 주소. 내 GitHub Pages 주소로 바꾸세요. (끝에 / 포함)
//   모든 관리자가 같은 화면 파일을 쓰고, 링크의 api=배포ID로 각자 서버(시트)에 연결된다.
var PAGES_BASE = 'https://ttorwai-stack.github.io/yeonsu-survey/';

/** PAGES_BASE를 아직 안 바꿨는지 */
function pagesBaseUnset_() {
  return !/^https:\/\/[^\/]+\.github\.io\/[^\/]+\/$/.test(PAGES_BASE) || /YOUR-/i.test(PAGES_BASE);
}

/** 이 설치의 웹앱 배포 ID (배포 전이면 빈 값) */
function deploymentId_() {
  var m = String(ScriptApp.getService().getUrl() || '').match(/\/s\/([\w-]+)\/(?:exec|dev)/);
  return m ? m[1] : '';
}

function showUrls() {
  var ui = SpreadsheetApp.getUi();
  var id = deploymentId_();
  if (!id) { ui.alert('아직 웹앱으로 배포하지 않았어요.\nApps Script에서 [배포 → 새 배포 → 웹 앱]을 먼저 해 주세요.\n(실행 사용자: 나 / 액세스 권한: 모든 사용자)'); return; }
  if (pagesBaseUnset_()) { ui.alert('코드 맨 위의 PAGES_BASE를 내 GitHub Pages 주소로 바꿔 주세요.\n(예: https://내아이디.github.io/저장소이름/)\n바꾼 뒤 저장하고 이 메뉴를 다시 실행하세요.'); return; }
  ui.alert('내 설문 관리 주소',
    '관리자 페이지 (이 주소로 들어가 비밀번호를 넣으세요. 즐겨찾기 해 두면 편해요)\n' +
    PAGES_BASE + 'admin.html?api=' + id + '\n\n' +
    '설문별 응답 링크는 관리자 페이지의 [링크 복사]로 받아 연수생에게 보내면 돼요.\n\n' +
    '내 웹앱 주소 (참고)\n' + ScriptApp.getService().getUrl(),
    ui.ButtonSet.OK);
}

function showShareHelp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.getUi().alert('다른 관리자에게 공유하기',
    '아래 사본 링크를 보내 주세요. 받은 사람 드라이브에 이 도구가 통째로 복사돼요.\n\n' +
    'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/copy\n\n' +
    '받은 사람이 할 일\n' +
    '1) 사본 만들기 → 시트 새로고침 → [설문 도구 → ① 초기 설정] (권한 허용, 비밀번호 설정)\n' +
    '2) 확장 프로그램 → Apps Script → [배포 → 새 배포 → 웹 앱] (실행: 나 / 액세스: 모든 사용자)\n' +
    '3) 시트 메뉴 [설문 도구 → 웹앱 주소 보기] → 나오는 관리자 페이지 주소로 접속\n\n' +
    '내 설문 목록·응답·비밀번호는 복사되지 않아요. 상대방 설문은 상대방 드라이브에만 저장돼요.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ───────── 웹앱 ───────── */

function doGet(e) {
  var p = (e && e.parameter) || {};
  var admin = p.page === 'admin';
  var title = admin ? '설문 관리' : '연수 설문';
  if (!admin && p.s) { var m = findSurvey_(p.s); if (m) title = m.title; }
  var out;
  try {
    out = HtmlService.createHtmlOutputFromFile(admin ? 'Admin' : 'Index');
  } catch (err) {
    // Apps Script 안에 화면 파일(Index/Admin)을 넣지 않은 설치: 화면은 GitHub Pages에 있으니 그리로 안내한다.
    var id = deploymentId_();
    var dest = (pagesBaseUnset_() || !id) ? '' :
      PAGES_BASE + (admin ? 'admin.html' : 'index.html') + '?api=' + id + (!admin && p.s ? '&s=' + encodeURIComponent(p.s) : '');
    out = HtmlService.createHtmlOutput(
      '<meta name="viewport" content="width=device-width, initial-scale=1">' +
      '<p style="font:16px sans-serif;padding:24px">설문 화면은 별도 주소에서 열려요.' +
      (dest ? '<br><br><a target="_top" href="' + dest.replace(/"/g, '&quot;') + '">여기를 눌러 이동하기</a>' : '<br>(PAGES_BASE 설정이 필요해요.)') + '</p>');
  }
  return out.setTitle(title).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * GitHub Pages 등 외부에 올린 화면에서 쓰는 데이터 창구.
 * 요청: POST, 본문 = {"action": "함수 이름", "args": [...]} (Content-Type: text/plain)
 * 응답: {"ok": true, "data": ...} 또는 {"ok": false, "error": "메시지"}
 */
var API_ACTIONS = {
  listOpenSurveys: listOpenSurveys, getSurvey: getSurvey, submitResponse: submitResponse,
  verifyKey: verifyKey, listSurveys: listSurveys, createSurvey: createSurvey,
  setSurveyStatus: setSurveyStatus, getResults: getResults, saveConfig: saveConfig, setCloseAt: setCloseAt, deleteSurvey: deleteSurvey
};

function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var fn = API_ACTIONS[req.action];
    if (!fn) throw new Error('알 수 없는 요청이에요.');
    out = { ok: true, data: fn.apply(null, req.args || []) };
  } catch (err) {
    var msg = (err && err.message) || String(err);
    // 구글 쪽 일시적 혼잡 오류는 BUSY로 바꿔 화면이 자동으로 다시 시도하게 한다.
    if (!/^BUSY/.test(msg) && /too many|너무 많|invoked|Exceeded|lock|잠금|timed out|시간 초과|unavailable|사용할 수 없|internal error|내부 오류|try again|다시 시도/i.test(msg)) msg = 'BUSY: ' + msg;
    out = { ok: false, error: msg };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* 응답자용 (비밀번호 없음) */

function listOpenSurveys() {
  return listSurveys_().filter(function (m) { return m.status === STATUS_OPEN; })
    .map(function (m) { return { id: m.id, title: m.title, course: m.course, cohort: m.cohort }; });
}

function getSurvey(id) {
  var m = findSurvey_(id);
  if (!m) throw new Error('설문을 찾을 수 없어요. 링크를 다시 확인해 주세요.');
  return { id: m.id, status: m.status, config: readConfig_(m.sheetId) };
}

/**
 * 응답 저장. edit = {id, token}이면 그 응답을 새 내용으로 바꾼다(응답 수는 그대로).
 * token(수정 열쇠)은 처음 제출한 기기에만 저장되므로 다른 사람의 응답은 고칠 수 없다.
 */
function submitResponse(surveyId, payload, edit, client) {
  var m = findSurvey_(surveyId);
  if (!m) throw new Error('설문을 찾을 수 없어요.');
  if (m.status !== STATUS_OPEN) throw new Error('마감된 설문이에요.');
  var cfg = readConfig_(m.sheetId);
  var err = validatePayload_(cfg, payload);
  if (err) throw new Error(err);
  var ss = SpreadsheetApp.openById(m.sheetId);
  var now = new Date();
  var ts = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  var lock = LockService.getScriptLock();

  // 수정: 드물게 일어나므로 차례로 처리한다. 먼저 수신함을 정리해야 방금 낸 응답도 찾을 수 있다.
  if (edit && edit.id) {
    if (!lock.tryLock(30000)) throw new Error('BUSY: 지금 제출이 몰리고 있어요. 잠시 후 자동으로 다시 시도해요.');
    try {
      processQueue_(ss, m);
      var respSh = ss.getSheetByName(SHEET_RESP), scoreSh = ss.getSheetByName(SHEET_SCORE);
      if (!removeResponse_(respSh, scoreSh, edit.id, edit.token)) throw new Error('수정할 응답을 찾지 못했어요.');
      writeResponses_(ss, cfg, [{ id: String(edit.id), token: String(edit.token), ts: ts, payload: payload }]);
      return { ok: true, id: String(edit.id), token: String(edit.token) };
    } finally {
      lock.releaseLock();
    }
  }

  // 새 응답: 수신함에 한 줄만 적고 바로 끝낸다.
  // 화면이 만든 응답ID·수정키를 쓰면, 같은 제출이 다시 와도(재시도) 수신함 정리 때 하나만 남는다.
  var ok = client && /^R[0-9a-z]{10,40}$/i.test(client.id || '') && /^[0-9a-z-]{16,64}$/i.test(client.token || '');
  var id = ok ? client.id : 'R' + Utilities.formatDate(now, 'Asia/Seoul', 'yyMMddHHmmss') + Math.floor(Math.random() * 9000 + 1000);
  var token = ok ? client.token : Utilities.getUuid();
  // 동시에 여러 줄을 붙이면 서로 덮어쓸 수 있어(실측) 짧은 순번 잠금으로 한 줄씩 붙인다.
  // 정리(processQueue_)와는 다른 잠금을 써서, 정리 중에도 접수는 막히지 않는다.
  var q = queueSheet_(ss);
  var alock = LockService.getDocumentLock() || LockService.getUserLock();
  if (!alock.tryLock(30000)) throw new Error('BUSY: 지금 제출이 몰리고 있어요. 잠시 후 자동으로 다시 시도해요.');
  try {
    q.appendRow([ts, id, token, JSON.stringify(payload), '']);
    SpreadsheetApp.flush();
  } finally {
    alock.releaseLock();
  }
  // 정리 담당이 비어 있으면 이번 요청이 쌓인 응답을 한꺼번에 시트로 옮긴다. (실패해도 다음 요청·결과 화면에서 다시 정리)
  try {
    if (lock.tryLock(0)) { try { processQueue_(ss, m); } finally { lock.releaseLock(); } }
  } catch (e) { /* 저장은 이미 끝남 */ }
  return { ok: true, id: id, token: token };
}

var SHEET_QUEUE = '수신';
var QUEUE_HEADERS = ['수신시각', '응답ID', '수정키', '응답(원본)', '처리'];

/** 수신함 시트 (없으면 만든다). */
function queueSheet_(ss) {
  var q = ss.getSheetByName(SHEET_QUEUE);
  if (q) return q;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    q = ss.getSheetByName(SHEET_QUEUE);
    if (!q) {
      q = ss.insertSheet(SHEET_QUEUE);
      q.getRange(1, 1, 1, QUEUE_HEADERS.length).setValues([QUEUE_HEADERS]).setFontWeight('bold');
      q.setFrozenRows(1);
    }
    return q;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 수신함에서 아직 처리 안 된 응답을 응답자·점수 시트로 한꺼번에 옮긴다. (호출하는 쪽이 잠금을 잡고 있어야 함)
 * 이미 응답자 시트에 있는 응답ID는 건너뛰어, 중간에 실패해 다시 처리해도 중복되지 않는다.
 */
function processQueue_(ss, m) {
  var q = ss.getSheetByName(SHEET_QUEUE);
  if (!q || q.getLastRow() < 2) return 0;
  var vals = q.getRange(2, 1, q.getLastRow() - 1, QUEUE_HEADERS.length).getValues();
  var todo = [];
  vals.forEach(function (r, i) { if (r[1] && r[4] !== 'Y') todo.push({ row: i + 2, ts: String(r[0]), id: String(r[1]), token: String(r[2]), json: r[3] }); });
  if (!todo.length) return 0;
  var respSh = ss.getSheetByName(SHEET_RESP);
  var have = {};
  if (respSh.getLastRow() > 1) {
    var H = respSh.getRange(1, 1, 1, respSh.getLastColumn()).getValues()[0], ci = H.indexOf('응답ID');
    if (ci >= 0) respSh.getRange(2, ci + 1, respSh.getLastRow() - 1, 1).getValues().forEach(function (r) { have[String(r[0])] = 1; });
  }
  var cfg = normalizeConfig_(JSON.parse(ss.getSheetByName(SHEET_CFG).getRange('A2').getValue()));
  var items = [];
  todo.forEach(function (t) {
    if (have[t.id]) return; // 이미 옮긴 응답이거나 같은 제출의 재전송
    have[t.id] = 1;
    try { items.push({ id: t.id, token: t.token, ts: t.ts, payload: JSON.parse(t.json) }); } catch (e) { /* 깨진 줄은 건너뜀(원본은 수신함에 남음) */ }
  });
  if (items.length) writeResponses_(ss, cfg, items);
  // 처리 표시: 대부분 아래쪽에 연속으로 모여 있으므로 연속 구간 단위로 한 번에 표시
  todo.sort(function (a, b) { return a.row - b.row; });
  var start = todo[0].row, prev = start;
  for (var i = 1; i <= todo.length; i++) {
    var r = i < todo.length ? todo[i].row : -1;
    if (r !== prev + 1) {
      var n = prev - start + 1, ys = [];
      for (var k = 0; k < n; k++) ys.push(['Y']);
      q.getRange(start, 5, n, 1).setValues(ys);
      start = r;
    }
    prev = r;
  }
  try { updateSurvey_(m.id, { '응답 수': respSh.getLastRow() - 1, '마지막 응답': todo[todo.length - 1].ts }); } catch (e) { /* 숫자만 늦게 맞춰짐 */ }
  return items.length;
}

/** 응답 여러 건을 응답자·점수 시트에 한 번에 쓴다. */
function writeResponses_(ss, cfg, items) {
  var respHeaders = [], scoreHeaders = [], respRows = [], scoreRows = [];
  items.forEach(function (it) {
    var b = buildRows_(cfg, it.payload, it.id, it.ts);
    b.respHeaders.concat([EDIT_KEY_HEADER]).forEach(function (h) { if (respHeaders.indexOf(h) < 0) respHeaders.push(h); });
    b.scoreHeaders.forEach(function (h) { if (scoreHeaders.indexOf(h) < 0) scoreHeaders.push(h); });
    b.respRow[EDIT_KEY_HEADER] = it.token;
    respRows.push(b.respRow);
    scoreRows = scoreRows.concat(b.scoreRows);
  });
  writeObjects_(ss.getSheetByName(SHEET_RESP), respHeaders, respRows);
  writeObjects_(ss.getSheetByName(SHEET_SCORE), scoreHeaders, scoreRows);
}

/** 응답ID + 수정 열쇠가 맞을 때만 그 응답의 행들을 지운다. */
function removeResponse_(respSh, scoreSh, id, token) {
  var v = respSh.getDataRange().getValues();
  var ci = v[0].indexOf('응답ID'), ki = v[0].indexOf(EDIT_KEY_HEADER);
  if (ci < 0 || ki < 0 || !token) return false;
  var row = -1;
  for (var r = 1; r < v.length; r++) {
    if (String(v[r][ci]) === String(id)) { if (String(v[r][ki]) !== String(token)) return false; row = r + 1; break; }
  }
  if (row < 0) return false;
  respSh.deleteRow(row);
  if (scoreSh.getLastRow() > 1) {
    var ids = scoreSh.getRange(2, 1, scoreSh.getLastRow() - 1, 1).getValues();
    // 아래에서부터 연속 구간 단위로 지운다
    for (var i = ids.length - 1; i >= 0; i--) {
      if (String(ids[i][0]) !== String(id)) continue;
      var end = i;
      while (i - 1 >= 0 && String(ids[i - 1][0]) === String(id)) i--;
      scoreSh.deleteRows(i + 2, end - i + 1);
    }
  }
  return true;
}

/* 관리자용 */

function verifyKey(key) {
  return checkKey_(key);
}

function listSurveys(key) {
  requireKey_(key);
  return listSurveys_(true).reverse();
}

/** 새 설문. srcId가 있으면 그 설문 설정을 복사하고, 없으면 tpl(post/pre/blank) 틀로 만든다. */
function createSurvey(key, srcId, tpl) {
  requireKey_(key);
  var cfg;
  if (srcId) {
    var src = findSurvey_(srcId);
    if (!src) throw new Error('복사할 설문을 찾을 수 없어요.');
    cfg = readConfig_(src.sheetId);
    cfg.title = '(복사) ' + cfg.title;
  } else {
    cfg = templateConfig_(tpl);
  }
  return createSurveyFromConfig_(cfg);
}

/** 설정으로 설문 스프레드시트를 만들고 목록에 등록한다. (편집기에서 직접 실행할 때도 사용) */
function createSurveyFromConfig_(cfg) {
  var err = validateConfig_(cfg);
  if (err) throw new Error(err);
  var ss = SpreadsheetApp.create(cfg.title);
  var first = ss.getSheets()[0].setName(SHEET_CFG);
  first.getRange('A1').setValue('설문 설정 (관리 화면에서 수정하세요. 직접 고치지 마세요)');
  first.getRange('A2').setValue(JSON.stringify(cfg));
  ss.insertSheet(SHEET_RESP).setFrozenRows(1);
  ss.insertSheet(SHEET_SCORE).setFrozenRows(1);
  var qs = ss.insertSheet(SHEET_QUEUE);
  qs.getRange(1, 1, 1, QUEUE_HEADERS.length).setValues([QUEUE_HEADERS]).setFontWeight('bold');
  qs.setFrozenRows(1);
  DriveApp.getFileById(ss.getId()).moveTo(folder_());

  var id = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
  var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  registry_().appendRow([id, cfg.title, cfg.course, cfg.cohort, STATUS_OPEN, 0, now, '', ss.getId(), ss.getUrl()]);
  cacheDel_(CACHE_LIST_KEY);
  return { id: id };
}

function setSurveyStatus(key, id, status) {
  requireKey_(key);
  if (status !== STATUS_OPEN && status !== STATUS_CLOSED) throw new Error('잘못된 상태예요.');
  var fields = { '상태': status };
  // 예약 시각이 이미 지난 설문을 다시 열면, 바로 또 닫히지 않도록 예약을 해제한다
  var m = findSurvey_(id);
  if (status === STATUS_OPEN && m && isDue_(m.closeAt)) fields['마감 예정'] = '';
  updateSurvey_(id, fields);
  return { ok: true };
}

/**
 * 설문 삭제. 마감된 설문만 지울 수 있다(진행 중인 설문을 실수로 지우지 않게).
 * 결과 시트는 드라이브 휴지통으로 옮기므로 30일 안에는 복구할 수 있다. 목록(설문목록 시트)에서는 줄을 지운다.
 */
function deleteSurvey(key, id) {
  requireKey_(key);
  var m = findSurvey_(id);
  if (!m) throw new Error('설문을 찾을 수 없어요.');
  if (m.status !== STATUS_CLOSED) throw new Error('진행 중인 설문은 삭제할 수 없어요. 먼저 [마감하기]를 눌러 주세요.');
  var trashed = false;
  try {
    DriveApp.getFileById(m.sheetId).setTrashed(true);
    trashed = true;
  } catch (e) {
    // 시트를 이미 지웠거나 못 찾는 경우는 목록에서만 지운다. 그 밖의 오류(권한 등)는 멈춘다.
    if (!/not found|찾을 수 없|No item|존재하지/i.test(String((e && e.message) || e))) throw new Error('결과 시트를 드라이브 휴지통으로 옮기지 못했어요: ' + ((e && e.message) || e));
  }
  var sh = registry_();
  var ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(function (r) { return String(r[0]); });
  var row = ids.indexOf(String(id)) + 1;
  if (row >= 2) sh.deleteRow(row);
  cacheDel_(CACHE_LIST_KEY);
  cacheDel_('svy-cfg-' + m.sheetId);
  return { ok: true, trashed: trashed };
}

/** 자동 마감 시각 예약. closeAt = "yyyy-MM-dd HH:mm"(한국 시간). 빈 값이면 예약 해제. */
function setCloseAt(key, id, closeAt) {
  requireKey_(key);
  closeAt = String(closeAt || '').trim();
  if (closeAt) {
    if (!/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(closeAt)) throw new Error('마감 시각 형식이 맞지 않아요. 예: 2026-11-01 23:59');
    if (closeAt <= nowKst_()) throw new Error('지금보다 이후 시각으로 정해 주세요.');
  }
  if (!findSurvey_(id)) throw new Error('설문을 찾을 수 없어요.');
  updateSurvey_(id, { '마감 예정': closeAt });
  return { ok: true, closeAt: closeAt };
}

function getResults(key, id) {
  requireKey_(key);
  var m = listSurveys_(true).filter(function (x) { return x.id === String(id); })[0];
  if (!m) throw new Error('설문을 찾을 수 없어요.');
  var ss = SpreadsheetApp.openById(m.sheetId);
  // 제출이 몰리는 중이면 정리를 기다리지 않는다(정리 담당이 비어 있을 때만 잠깐 정리).
  var lock = LockService.getScriptLock();
  if (lock.tryLock(500)) { try { processQueue_(ss, m); } catch (e) { /* 아래에서 수신함을 직접 합친다 */ } finally { lock.releaseLock(); } }
  var cfg = readConfig_(m.sheetId);
  var resp = readTable_(ss.getSheetByName(SHEET_RESP));
  var scores = readTable_(ss.getSheetByName(SHEET_SCORE));
  // 아직 시트로 옮겨지지 않은 응답은 수신함에서 바로 읽어 결과에 합친다 → 기다리지 않아도 결과는 항상 전부
  mergeQueue_(ss, cfg, resp, scores);
  var k = resp.headers.indexOf(EDIT_KEY_HEADER);
  if (k >= 0) { resp.headers.splice(k, 1); resp.rows.forEach(function (r) { r.splice(k, 1); }); }
  return { meta: m, config: cfg, resp: resp, scores: scores };
}

/** 수신함의 미처리 응답(응답자 시트에 아직 없는 응답ID)을 표(headers/rows)에 메모리로만 더한다. */
function mergeQueue_(ss, cfg, resp, scores) {
  var q = ss.getSheetByName(SHEET_QUEUE);
  if (!q || q.getLastRow() < 2) return;
  var have = {}, ci = resp.headers.indexOf('응답ID');
  if (ci >= 0) resp.rows.forEach(function (r) { have[String(r[ci])] = 1; });
  var items = [];
  q.getRange(2, 1, q.getLastRow() - 1, QUEUE_HEADERS.length).getDisplayValues().forEach(function (r) {
    var id = String(r[1]);
    if (!id || r[4] === 'Y' || have[id]) return;
    have[id] = 1;
    try { items.push({ id: id, token: r[2], ts: r[0], payload: JSON.parse(r[3]) }); } catch (e) { /* 깨진 줄은 건너뜀 */ }
  });
  items.forEach(function (it) {
    var b = buildRows_(cfg, it.payload, it.id, it.ts);
    b.respRow[EDIT_KEY_HEADER] = it.token;
    appendToTable_(resp, b.respHeaders.concat([EDIT_KEY_HEADER]), [b.respRow]);
    appendToTable_(scores, b.scoreHeaders, b.scoreRows);
  });
}

function appendToTable_(t, headers, objs) {
  headers.forEach(function (h) { if (t.headers.indexOf(h) < 0) { t.headers.push(h); t.rows.forEach(function (r) { r.push(''); }); } });
  objs.forEach(function (o) { t.rows.push(t.headers.map(function (h) { return o[h] === undefined ? '' : String(o[h]); })); });
}

function saveConfig(key, id, cfg) {
  requireKey_(key);
  var err = validateConfig_(cfg);
  if (err) throw new Error(err);
  var m = findSurvey_(id);
  if (!m) throw new Error('설문을 찾을 수 없어요.');
  var ss = SpreadsheetApp.openById(m.sheetId);
  ss.getSheetByName(SHEET_CFG).getRange('A2').setValue(JSON.stringify(cfg));
  cachePut_('svy-cfg-' + m.sheetId, cfg, 300);
  if (ss.getName() !== cfg.title) ss.rename(cfg.title);
  updateSurvey_(id, { '제목': cfg.title, '과정': cfg.course, '기수': cfg.cohort });
  return { ok: true };
}

/* ───────── 시트 입출력 ───────── */

function installed_() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty('INSTALL_SCRIPT') === ScriptApp.getScriptId() && !!props.getProperty('SHEET_ID');
}

function registry_() {
  if (!installed_()) throw new Error('초기 설정이 필요해요. 관리용 시트에서 [설문 도구 → ① 초기 설정]을 실행해 주세요.');
  var ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  return ss.getSheetByName(SHEET_REG);
}

function folder_() {
  var props = PropertiesService.getScriptProperties();
  var fid = props.getProperty('FOLDER_ID');
  if (fid) { try { var f = DriveApp.getFolderById(fid); if (!f.isTrashed()) return f; } catch (e) { /* 새로 만든다 */ } }
  var folder = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', folder.getId());
  return folder;
}

/*
 * 동시 접속이 몰려도 시트를 덜 읽도록 설문 목록·설정을 잠시(캐시) 보관한다.
 * 상태·제목·설정이 바뀌면 바로 지운다. (응답 수 숫자만 최대 1분 늦게 보일 수 있음)
 */
var CACHE_LIST_KEY = 'svy-list-v1';
function cache_() { return CacheService.getScriptCache(); }
function cacheGet_(k) { try { var v = cache_().get(k); return v ? JSON.parse(v) : null; } catch (e) { return null; } }
function cachePut_(k, v, sec) { try { cache_().put(k, JSON.stringify(v), sec); } catch (e) { /* 캐시는 없어도 동작 */ } }
function cacheDel_(k) { try { cache_().remove(k); } catch (e) { /* 무시 */ } }

function nowKst_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm'); }

/** 예약 시각("yyyy-MM-dd HH:mm", 한국 시간)이 지났는지. 같은 형식 문자열끼리라 글자 순서 비교로 충분하다. */
function isDue_(closeAt) {
  return /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(String(closeAt || '')) && String(closeAt) <= nowKst_();
}

/** 진행 중인 설문 중 예약 시각이 지난 것은 마감으로 본다. (읽는 순간 계산하므로 응답 접수도 정확히 그 시각에 멈춘다) */
function applySchedule_(list) {
  return list.map(function (m) {
    if (m.status !== STATUS_OPEN || !isDue_(m.closeAt)) return m;
    var c = {};
    Object.keys(m).forEach(function (k) { c[k] = m[k]; });
    c.status = STATUS_CLOSED;
    c.autoClosed = true;
    return c;
  });
}

function listSurveys_(fresh) {
  if (!fresh) { var c = cacheGet_(CACHE_LIST_KEY); if (c) return applySchedule_(c); }
  var sh = registry_();
  var rows = sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, REG_HEADERS.length).getDisplayValues();
  var list = [], rowOf = [];
  rows.forEach(function (r, i) {
    if (!r[0]) return;
    list.push({ id: r[0], title: r[1], course: r[2], cohort: r[3], status: r[4], count: +r[5] || 0, created: r[6], last: r[7], sheetId: r[8], sheetUrl: r[9], closeAt: r[10] || '' });
    rowOf.push(i + 2);
  });
  list = applySchedule_(list);
  // 예약 시각이 지난 설문은 시트의 상태 칸도 마감으로 맞춰 둔다
  list.forEach(function (m, i) {
    if (!m.autoClosed) return;
    try { sh.getRange(rowOf[i], REG_HEADERS.indexOf('상태') + 1).setValue(STATUS_CLOSED); } catch (e) { /* 다음에 다시 맞춰짐 */ }
  });
  cachePut_(CACHE_LIST_KEY, list, 60);
  return list;
}

function findSurvey_(id) {
  if (!id) return null;
  return listSurveys_().filter(function (m) { return m.id === String(id); })[0] || null;
}

function updateSurvey_(id, fields) {
  var sh = registry_();
  var ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(function (r) { return String(r[0]); });
  var row = ids.indexOf(String(id)) + 1;
  if (row < 2) throw new Error('설문을 찾을 수 없어요.');
  Object.keys(fields).forEach(function (h) {
    var col = REG_HEADERS.indexOf(h) + 1;
    if (col < 1) return;
    var cell = sh.getRange(row, col);
    if (h === '마감 예정') {
      if (!String(sh.getRange(1, col).getValue() || '')) sh.getRange(1, col).setValue(h).setFontWeight('bold');
      cell.setNumberFormat('@'); // 2026-11-01 23:59 가 날짜로 바뀌지 않게 글자로 저장
    }
    cell.setValue(fields[h]);
  });
  // 응답 수·마지막 응답만 바뀐 경우가 아니면(상태·제목 등) 목록 캐시를 바로 지운다
  if (Object.keys(fields).some(function (h) { return h !== '응답 수' && h !== '마지막 응답'; })) cacheDel_(CACHE_LIST_KEY);
}

function readConfig_(sheetId) {
  var key = 'svy-cfg-' + sheetId;
  var c = cacheGet_(key);
  if (c) return normalizeConfig_(c);
  var cfg = JSON.parse(SpreadsheetApp.openById(sheetId).getSheetByName(SHEET_CFG).getRange('A2').getValue());
  cachePut_(key, cfg, 300);
  return normalizeConfig_(cfg);
}

function checkKey_(key) {
  var props = PropertiesService.getScriptProperties();
  var real = props.getProperty('ADMIN_KEY');
  return installed_() && !!real && String(key || '') === real;
}

function requireKey_(key) {
  if (!checkKey_(key)) throw new Error('비밀번호가 맞지 않아요.');
}

/** 헤더가 없으면 오른쪽에 추가하고, 객체 배열을 헤더 순서에 맞춰 이어 붙인다. */
function writeObjects_(sheet, headers, objs) {
  if (!objs.length) return;
  var lastCol = sheet.getLastColumn();
  var current = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var missing = headers.filter(function (h) { return current.indexOf(h) < 0; });
  if (missing.length) {
    sheet.getRange(1, current.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
    current = current.concat(missing);
  }
  var rows = objs.map(function (o) {
    return current.map(function (h) { return o[h] === undefined ? '' : o[h]; });
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, current.length).setValues(rows);
}

function readTable_(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return { headers: [], rows: [] };
  var v = sheet.getDataRange().getDisplayValues();
  return { headers: v[0], rows: v.slice(1) };
}

/* ───────── 아래는 순수 함수 (로컬 미리보기에서도 그대로 사용) ───────── */

/*
 * 설문 구조 (v2)
 *   cfg = { v: 2, title, course, cohort, period, intro, doneMsg, sections: [섹션] }
 *   섹션 = { id, title, desc, blocks: [블록] }   → 응답 화면의 한 페이지
 *   블록
 *     choice    객관식        { id, label, options[], other, placeholder, mode: single|multi|rank, max }
 *     scale     척도 묶음      { id, title, area, prefix, labels[3~7](1점→맨 끝 점수), posName, items[{short, text}] }
 *     timetable 시간표 과목평가 { id, area, labels[3~7], items[{key, text}], slots[], splitHint }  (설문당 1개)
 *     text      서술형        { id, short, text, required }
 * 예전 형식(기본정보→만족도→시간표→서술형 고정 4단계)은 normalizeConfig_가 같은 ID·같은 열 이름으로 바꿔 준다
 * → 이미 쌓인 응답과 그대로 이어진다.
 */

var DEFAULT_LABELS = {
  agree: ['전혀 아니다', '그렇지 않다', '보통이다', '그렇다', '매우 그렇다'],
  quality: ['매우미흡', '미흡', '보통', '우수', '매우우수']
};

function normalizeConfig_(c) {
  if (!c) return c;
  if (c.v === 2 && Array.isArray(c.sections)) return c;
  var T = {}, d = defaultTexts_(), tx = c.texts || {};
  Object.keys(d).forEach(function (k) { T[k] = tx[k] != null && String(tx[k]).trim() !== '' ? String(tx[k]) : d[k]; });
  var scales = c.scales || {}, sec = [];
  var demo = c.demographics || [], ov = (c.overall && c.overall.items) || [], slots = c.slots || [], si = c.sessionItems || [], open = c.open || [];
  if (demo.length) sec.push({ id: 'demo', title: T.demoTitle, desc: T.demoDesc, blocks: demo.map(function (x) {
    var b = { type: 'choice', id: x.id, label: x.label, options: (x.options || []).slice(), other: !!x.other, placeholder: x.placeholder || '', mode: x.type === 'multi' || x.type === 'rank' ? x.type : 'single' };
    if (x.max) b.max = x.max;
    return b;
  }) });
  if (ov.length) sec.push({ id: 'overall', title: T.overallTitle, desc: T.overallDesc, blocks: [{
    type: 'scale', id: 'overall', title: '', area: '과정만족도', prefix: '과정',
    labels: (scales[(c.overall && c.overall.scale) || 'agree'] || DEFAULT_LABELS.agree).slice(), posName: T.posName,
    items: ov.map(function (i) { return { short: i.short, text: i.text }; })
  }] });
  if (slots.length) sec.push({ id: 'slots', title: T.slotsTitle, desc: T.slotsDesc, blocks: [{
    type: 'timetable', id: 'slots', area: '과목평가',
    labels: (scales[(si[0] && si[0].scale) || 'quality'] || DEFAULT_LABELS.quality).slice(),
    items: si.map(function (i) { return { key: i.key, text: i.text }; }), slots: slots, splitHint: T.splitHint
  }] });
  if (open.length) sec.push({ id: 'open', title: T.openTitle, desc: T.openDesc, blocks: open.map(function (q) {
    return { type: 'text', id: q.id, short: q.short, text: q.text, required: !!q.required };
  }) });
  return { v: 2, title: c.title, course: c.course || '', cohort: c.cohort || '', period: c.period || '', intro: c.intro || '', doneMsg: T.doneMsg, sections: sec };
}

function allBlocks_(cfg) {
  var out = [];
  (cfg.sections || []).forEach(function (s) { (s.blocks || []).forEach(function (b) { out.push(b); }); });
  return out;
}

/** 여러 개 고르기의 최대 개수 / 순위 매기기의 순위 개수 (비우면 전부) */
function pickCount_(b) {
  var all = b.options.length + (b.other ? 1 : 0), m = parseInt(b.max, 10);
  return m >= 1 ? Math.min(m, all) : all;
}

function hasOther_(v) { return Array.isArray(v) ? v.indexOf('기타') >= 0 : v === '기타'; }

/** 시트에 쓰는 값: 여러 개는 " | "로, 순위는 1순위부터 " > "로 잇는다. (결과 화면이 같은 기호로 다시 나눔) */
var CHOICE_SEP = { multi: ' | ', rank: ' > ' };
function choiceCell_(b, v) { return Array.isArray(v) ? v.join(CHOICE_SEP[b.mode] || ' | ') : String(v == null ? '' : v).trim(); }

/** 척도 묶음의 응답자 시트 열 머리: [접두어] 짧은 이름 */
function scalePrefix_(b) { return b.prefix || b.title || b.id; }

/** 응답 → {ans: {블록ID: 값}, other: {블록ID: 기타 내용}}. 예전 형식(demo/overall/slots/open) 응답도 받는다. */
function payloadAnswers_(p) {
  p = p || {};
  if (p.v === 2) return { ans: p.ans || {}, other: p.other || {} };
  var ans = {};
  Object.keys(p.demo || {}).forEach(function (k) { ans[k] = p.demo[k]; });
  if (p.overall) ans.overall = p.overall;
  if (p.slots) ans.slots = p.slots;
  Object.keys(p.open || {}).forEach(function (k) { ans[k] = p.open[k]; });
  return { ans: ans, other: p.other || {} };
}

function slotLabel_(slot) {
  if (slot.options.length > 1) return slot.date + ' ' + slot.time + ' 분반수업';
  var o = slot.options[0];
  return slot.date + ' ' + slot.time + ' ' + o.title + '(' + lastName_(o.instructor) + ')';
}

/** "시흥능곡초 이요한" → "이요한" */
function lastName_(instructor) {
  var parts = String(instructor || '').trim().split(/\s+/);
  return parts[parts.length - 1] || '';
}

function validatePayload_(cfg, p) {
  if (!p) return '응답이 비어 있어요.';
  cfg = normalizeConfig_(cfg);
  var P = payloadAnswers_(p), bl = allBlocks_(cfg);
  for (var i = 0; i < bl.length; i++) {
    var e = validateBlock_(bl[i], P.ans[bl[i].id], P.other[bl[i].id]);
    if (e) return e;
  }
  return '';
}

function validateBlock_(b, v, other) {
  var i;
  if (b.type === 'choice') {
    if (b.mode === 'multi' || b.mode === 'rank') {
      if (!Array.isArray(v) || !v.length) return '"' + b.label + '"에 응답해 주세요.';
      var k = pickCount_(b), seen = {};
      if (b.mode === 'rank' && v.length !== k) return '"' + b.label + '"은(는) ' + k + '순위까지 매겨 주세요.';
      if (v.length > k) return '"' + b.label + '"은(는) ' + k + '개까지 고를 수 있어요.';
      for (i = 0; i < v.length; i++) {
        if (seen[v[i]] || (b.options.indexOf(v[i]) < 0 && !(b.other && v[i] === '기타'))) return '"' + b.label + '" 선택지가 바뀌었어요. 새로고침한 뒤 다시 골라 주세요.';
        seen[v[i]] = 1;
      }
    } else if (!String(v || '').trim()) return '"' + b.label + '"에 응답해 주세요.';
    if (b.other && hasOther_(v) && !String(other || '').trim()) return '"' + b.label + '" 기타 내용을 적어 주세요.';
  } else if (b.type === 'scale') {
    for (i = 0; i < b.items.length; i++) {
      if (!validScore_(v && v[i], b.labels.length)) return '"' + b.items[i].short + '" 문항에 응답해 주세요.';
    }
  } else if (b.type === 'timetable') {
    for (i = 0; i < b.slots.length; i++) {
      var s = b.slots[i], a = v && v[s.id];
      if (!a || a.choice == null || !s.options[a.choice]) return slotLabel_(s) + ': 수강한 분반을 선택해 주세요.';
      for (var j = 0; j < b.items.length; j++) {
        if (!validScore_(a.scores && a.scores[b.items[j].key], b.labels.length)) return slotLabel_(s) + ': 모든 문항에 응답해 주세요.';
      }
    }
  } else if (b.type === 'text') {
    if (b.required && !String(v || '').trim()) return '"' + b.short + '"에 응답해 주세요.';
  }
  return '';
}

/** 1~n점 정수인지 (n = 그 척도의 점수 개수, 기본 5) */
function validScore_(v, n) {
  return typeof v === 'number' && v % 1 === 0 && v >= 1 && v <= (n || 5);
}

/** 척도 이름(점수 개수) 점검: 3~7개, 빈 이름 없음 */
function scaleLabelsBad_(l) {
  return !Array.isArray(l) || l.length < 3 || l.length > 7 || l.some(function (x) { return !String(x == null ? '' : x).trim(); });
}

/** 저장 전 점검. 비어 있는 값(영역 이름 등)은 여기서 채운다. */
function validateConfig_(cfg) {
  if (!cfg || !cfg.title) return '설문 제목을 입력해 주세요.'; // 기수는 선택 (한 번만 운영하는 과정은 비워 둠)
  cfg.cohort = String(cfg.cohort || '').trim();
  if (cfg.v !== 2 || !Array.isArray(cfg.sections)) return '설문 구조가 올바르지 않아요. 화면을 새로고침해 주세요.';
  var ids = {}, cols = {}, areas = {}, nT = 0, n = 0;
  function col(name) { if (cols[name]) return '"' + name + '" 이름이 두 번 쓰였어요. 결과 표의 열 이름이 겹치지 않게 바꿔 주세요.'; cols[name] = 1; return ''; }
  for (var si = 0; si < cfg.sections.length; si++) {
    var sec = cfg.sections[si], e;
    if (!sec.title) return (si + 1) + '번째 섹션의 이름을 적어 주세요.';
    if (!sec.blocks || !sec.blocks.length) return '"' + sec.title + '" 섹션에 문항이 없어요.';
    for (var bi = 0; bi < sec.blocks.length; bi++) {
      var b = sec.blocks[bi];
      if (!b.id || ids[b.id]) return '"' + sec.title + '" 섹션에 문항 ID가 겹치는 문항이 있어요.';
      ids[b.id] = 1; n++;
      if (b.type === 'choice') {
        if (!b.label) return '"' + sec.title + '" 섹션에 이름이 빈 객관식 문항이 있어요.';
        if (!b.options.length && !b.other) return '"' + b.label + '"에 선택지가 없어요.';
        if (['single', 'multi', 'rank'].indexOf(b.mode) < 0) b.mode = 'single';
        if ((e = col(b.label))) return e;
      } else if (b.type === 'scale') {
        if (!b.items || !b.items.length) return '"' + sec.title + '" 섹션의 척도 묶음에 문항이 없어요.';
        if (scaleLabelsBad_(b.labels)) return '척도 이름은 3~7개를 쉼표로 적어 주세요. (5개면 5점 척도)';
        if (!b.prefix) b.prefix = b.title || sec.title;
        if (!b.area) b.area = b.prefix;
        if (areas[b.area]) return '척도 묶음 이름 "' + b.area + '"이(가) 겹쳐요. 묶음 이름을 서로 다르게 해 주세요.';
        areas[b.area] = 1;
        for (var k = 0; k < b.items.length; k++) {
          if (!b.items[k].short || !b.items[k].text) return '"' + sec.title + '" 섹션의 척도 문항에 빈 줄이 있어요.';
          if ((e = col('[' + b.prefix + '] ' + b.items[k].short))) return e;
        }
      } else if (b.type === 'timetable') {
        if (++nT > 1) return '시간표 과목평가는 설문에 하나만 넣을 수 있어요.';
        if (!b.items || !b.items.length) return '시간표가 있으면 과목마다 묻는 문항이 1개 이상 필요해요.';
        if (!b.slots || !b.slots.length) return '시간표에 과목을 한 줄 이상 넣어 주세요.';
        if (scaleLabelsBad_(b.labels)) return '척도 이름은 3~7개를 쉼표로 적어 주세요. (5개면 5점 척도)';
        if (!b.area) b.area = '과목평가';
        for (var t = 0; t < b.slots.length; t++) {
          if (!b.slots[t].options.length) return '시간표 ' + (t + 1) + '번째 줄에 과목이 없어요.';
          for (var o = 0; o < b.slots[t].options.length; o++) if (!b.slots[t].options[o].title) return '시간표에 과목명이 빈 줄이 있어요.';
        }
      } else if (b.type === 'text') {
        if (!b.short || !b.text) return '"' + sec.title + '" 섹션의 서술형 문항에 짧은 이름이나 질문이 비어 있어요.';
        if ((e = col(b.short))) return e;
      } else return '알 수 없는 문항 종류가 있어요.';
    }
  }
  if (!n) return '문항이 하나도 없어요.';
  return '';
}

/**
 * 응답 하나를 시트 행으로 변환한다.
 * 응답자 행: 사람 1명 = 1줄 (분반은 "선택" 열 + 점수 열로 표현)
 * 점수 행  : 점수 1개 = 1줄 (객관식 응답을 함께 붙여 피벗으로 나눠 볼 수 있게)
 */
function buildRows_(cfg, p, id, ts) {
  cfg = normalizeConfig_(cfg);
  var P = payloadAnswers_(p), bl = allBlocks_(cfg);
  var choices = bl.filter(function (b) { return b.type === 'choice'; });
  var demoVals = {};
  choices.forEach(function (b) { demoVals[b.label] = choiceCell_(b, P.ans[b.id]); });

  var resp = { '응답ID': id, '제출시각': ts, '기수': cfg.cohort };
  var respHeaders = ['응답ID', '제출시각', '기수'];
  var scoreRows = [];
  function scoreRow(area, date, time, title, inst, split, item, v, labels) {
    var r = {
      '응답ID': id, '제출시각': ts, '기수': cfg.cohort, '영역': area, '일자': date, '시간': time,
      '과목': title, '강사': inst, '분반수업': split ? 'Y' : 'N', '문항': item, '점수': v, '응답': labels[v - 1]
    };
    Object.keys(demoVals).forEach(function (k) { r[k] = demoVals[k]; });
    scoreRows.push(r);
  }
  function put(h, v) { respHeaders.push(h); resp[h] = v; }

  bl.forEach(function (b) {
    var v = P.ans[b.id];
    if (b.type === 'choice') {
      put(b.label, demoVals[b.label]);
      // 순위 문항은 순위별 열도 둔다 (엑셀에서 바로 집계: 희망 시기(1순위) = 3월)
      if (b.mode === 'rank') for (var r = 0; r < pickCount_(b); r++) put(b.label + '(' + (r + 1) + '순위)', (v && v[r]) || '');
      // "기타" 직접 입력 내용은 집계용 열과 분리 (예: 직렬 = 기타, 직렬(기타 입력) = 건축)
      if (b.other) put(b.label + '(기타 입력)', hasOther_(v) ? String(P.other[b.id] || '').trim() : '');
    } else if (b.type === 'scale') {
      b.items.forEach(function (it, i) {
        put('[' + scalePrefix_(b) + '] ' + it.short, v[i]);
        scoreRow(b.area, '', '', '', '', false, it.short, v[i], b.labels);
      });
    } else if (b.type === 'timetable') {
      b.slots.forEach(function (s) {
        var a = v[s.id], o = s.options[a.choice], split = s.options.length > 1, base = slotLabel_(s);
        if (split) put(base + ' - 선택', o.title + '(' + lastName_(o.instructor) + ')');
        b.items.forEach(function (it) {
          put(base + ' - ' + it.key, a.scores[it.key]);
          scoreRow(b.area || '과목평가', s.date, s.time, o.title, o.instructor, split, it.key, a.scores[it.key], b.labels);
        });
      });
    } else if (b.type === 'text') {
      put(b.short, String(v || '').trim());
    }
  });

  var scoreHeaders = SCORE_BASE_HEADERS.concat(choices.map(function (b) { return b.label; }));
  return { respHeaders: respHeaders, respRow: resp, scoreHeaders: scoreHeaders, scoreRows: scoreRows };
}

/* ───────── 새 설문 틀 ───────── */

/** tpl: post(사후설문 기본형, 기본값) / pre(사전설문 기본형) / blank(빈 설문) */
function templateConfig_(tpl) {
  if (tpl === 'blank') {
    return { v: 2, title: '새 설문', course: '', cohort: '', period: '', intro: '', doneMsg: '소중한 의견 감사합니다.',
      sections: [{ id: 'sec1', title: '질문', desc: '', blocks: [{ type: 'text', id: 'q1', short: '의견', text: '의견을 자유롭게 적어 주세요.', required: false }] }] };
  }
  var c = defaultConfig_();
  c.course = '과정명';
  c.cohort = '';
  c.period = '';
  if (tpl === 'pre') {
    c.title = '[과정명] 사전설문';
    c.intro = '연수 시작 전, 기대와 준비 정도를 여쭙습니다. 연수 운영에 참고하겠습니다. (약 2분 소요)';
    c.texts = defaultTexts_();
    c.texts.overallTitle = '연수 기대';
    c.texts.overallDesc = '연수를 시작하기 전 생각을 골라 주세요.';
    c.texts.posName = '긍정 비율';
    c.overall.items = [
      { short: '직무 연관성', text: '본 교육과정은 본인이 수행하고 있는 직무와 관련성이 높다.' },
      { short: '학습 준비도', text: '나는 과정 수강 전, 본 교육과정의 학습목표, 학습내용을 확인해 보았다.' },
      { short: '현업 활용 기대', text: '본 교육과정의 교육내용은 현업에서 활용이 가능할 것으로 기대한다.' },
      { short: '역량 향상 기대', text: '본 교육과정의 교육을 통해 교육과정이 목표로 하는 역량이 향상될 것으로 기대한다.' },
      { short: '업무성과 기대', text: '본 교육과정의 교육을 통해 교육과정이 목표로 하는 업무성과가 향상될 것으로 기대한다.' }
    ];
    c.slots = []; c.sessionItems = [];
    c.open = [{ id: 'expect', short: '기대하는 점', text: '이번 연수에서 특히 기대하거나 배우고 싶은 내용을 적어 주세요.' }];
    return normalizeConfig_(c);
  }
  c.title = '[과정명] 사후설문';
  c.slots = [
    { id: 's1', date: '1일차', time: '오전', options: [{ title: '과목명', instructor: '소속 강사명' }] },
    { id: 's2', date: '1일차', time: '오후', options: [{ title: '과목명', instructor: '소속 강사명' }] }
  ];
  return normalizeConfig_(c);
}

/** 예전 형식의 화면 문구 기본값 (예전 설정을 바꿀 때만 쓰임) */
function defaultTexts_() {
  return {
    demoTitle: '기본 정보', demoDesc: '통계 분석에만 쓰이며, 개인을 식별하지 않습니다.',
    overallTitle: '과정 만족도', overallDesc: '교육과정과 연수 운영 전반에 대해 평가해 주세요.',
    slotsTitle: '과목 평가', slotsDesc: '시간표 순서대로 과목마다 평가해 주세요. 분반 수업은 들은 수업만 고르면 됩니다.',
    splitHint: '내가 들은 수업을 고르세요',
    openTitle: '의견', openDesc: '구체적으로 적어 주시면 다음 연수에 꼭 반영하겠습니다.',
    doneMsg: '소중한 의견 감사합니다.',
    posName: '만족 비율'
  };
}

function defaultConfig_() {
  return {
    texts: defaultTexts_(),
    title: '[6급 승진자과정 3기] 사후설문',
    course: '6급 승진자과정',
    cohort: '3기',
    period: '2026. 9. 14.(월) ~ 9. 18.(금)',
    intro: '연수에 참여해 주셔서 감사합니다. 더 나은 연수를 위해 솔직한 의견을 부탁드립니다. (약 3분 소요)',
    scales: {
      agree: ['전혀 아니다', '그렇지 않다', '보통이다', '그렇다', '매우 그렇다'],
      quality: ['매우미흡', '미흡', '보통', '우수', '매우우수']
    },
    demographics: [
      { id: 'gender', label: '성별', options: ['남', '여'] },
      { id: 'series', label: '직렬', options: ['교육행정'], other: true, placeholder: '직렬을 적어 주세요 (예: 시설관리, 사서)' },
      { id: 'grade', label: '직급', options: ['6급'], other: true, placeholder: '직급을 적어 주세요 (예: 7급)' },
      { id: 'career', label: '경력', options: ['5년미만', '5년이상 ~ 10년미만', '10년이상 ~ 20년미만', '20년이상'] },
      { id: 'org', label: '소속', options: ['본청 및 교육지원청', '직속기관', '유,초등학교', '중학교', '고등학교'], other: true },
      { id: 'region', label: '지역', options: ['경기남부', '경기북부'] }
    ],
    overall: {
      scale: 'agree',
      items: [
        { short: '전반 만족', text: '본 교육과정은 전반적으로 만족한다.' },
        { short: '추천 의향', text: '본 교육과정을 내 동료에게 추천할 의사가 있다.' },
        { short: '목적 달성', text: '본 교육과정 신청 시 안내된 교육목적이 달성되었다고 생각한다.' },
        { short: '구성 적합성', text: '본 교육과정은 교육 목표에 적합하게 구성되었다.' },
        { short: '교육방법·교재', text: '본 교육과정 교육방법(강의, 토론 등) 및 교재는 교육내용을 배우기 위해 효과적이다.' },
        { short: '운영·지원', text: '본 교육과정의 교육 운영과 지원은 원활하였다.' },
        { short: '이해도 향상', text: '본 교육과정의 교육내용에 대한 이해가 교육받기 전보다 높아졌다.' },
        { short: '현업 도움', text: '본 교육내용은 현업에서 업무 수행이나 자기개발에 도움이 될 것이다.' },
        { short: '연수 환경', text: '본 교육과정 연수 환경(비대면)은 이용에 만족스러웠다.' }
      ]
    },
    sessionItems: [
      { key: '교과편성', text: '이 교과목은 업무능력 향상이나 역량개발에 도움이 될 것이다.', scale: 'quality' },
      { key: '강사', text: '강사는 강의 내용을 이해하기 쉽게 전달하였고, 강의 스킬은 우수하였다.', scale: 'quality' }
    ],
    slots: [
      { id: 's1', date: '9.14.(월)', time: '오전', options: [{ title: '경기교육.Pro', instructor: '시흥능곡초 이요한' }] },
      { id: 's2', date: '9.14.(월)', time: '오전', options: [{ title: '경기교육.Pro', instructor: '조달역량개발원 오연칠' }] },
      { id: 's3', date: '9.14.(월)', time: '오후', options: [
        { title: '급여결재전체크사항', instructor: '성지고 피혜선' },
        { title: '원페이퍼보고서 전략', instructor: '정책기획관 신영민' }] },
      { id: 's4', date: '9.15.(화)', time: '오전', options: [
        { title: '학교시설관리사례', instructor: '학교설립과 임태윤' },
        { title: '특이민원대응', instructor: '로젠탈콘텐츠랩 김윤진' }] },
      { id: 's5', date: '9.15.(화)', time: '오후', options: [{ title: 'A.I.라는 키', instructor: 'flowDX 이애진' }] },
      { id: 's6', date: '9.16.(수)', time: '오전', options: [{ title: '6급브릿지리더십', instructor: '퍼스트교육컨설팅 정재일' }] },
      { id: 's7', date: '9.16.(수)', time: '오후', options: [{ title: '그렇게 행정실장이 된다', instructor: '포곡고 김영무' }] },
      { id: 's8', date: '9.16.(수)', time: '오후', options: [{ title: '승진자 마인드셋', instructor: '선배공직자' }] },
      { id: 's9', date: '9.17.(목)', time: '오전', options: [{ title: '공사계약 집대성', instructor: '포일초 이규상' }] },
      { id: 's10', date: '9.17.(목)', time: '오후', options: [{ title: '실전 감사 인사이트', instructor: '도림초 윤종철' }] },
      { id: 's11', date: '9.18.(금)', time: '오전', options: [{ title: '심신 리부트', instructor: '이강화/윤우람' }] },
      { id: 's12', date: '9.18.(금)', time: '오후', options: [{ title: '미식 인문학', instructor: '팜파티아 김은영' }] }
    ],
    open: [
      { id: 'good', short: '좋았던 점', text: '좋았던 점(연수기획, 교과편성, 강사, 운영방식, 연수 환경 등)을 구체적으로 작성해주세요.' },
      { id: 'improve', short: '개선할 점', text: '개선할 점(연수기획, 교과편성, 강사, 운영방식, 연수 환경 등)이 있다면 구체적으로 작성해주세요.' },
      { id: 'wish', short: '신설 희망 강좌·강사', text: '추후 동일 연수에서 신설하기를 희망하는 강좌나 강사를 추천해주세요.' },
      { id: 'free', short: '자유 의견', text: '이번 연수와 관련하여 자유롭게 의견을 말씀해주세요.' }
    ]
  };
}
