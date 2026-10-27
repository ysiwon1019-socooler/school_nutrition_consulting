/**
 * 우리 아이 식생활 유형 진단 — 구글 시트 저장소
 *
 * 사용 방법
 * 1) 구글 시트를 새로 만들고 [확장 프로그램] → [Apps Script]를 엽니다.
 * 2) 기본 코드를 모두 지우고 이 파일 내용을 붙여 넣습니다.
 * 3) 아래 ACCESS_CODE를 나만 아는 접근 코드로 바꾼 뒤 저장합니다.
 * 4) [배포] → [새 배포] → 유형: 웹 앱
 *      - 실행 사용자: 나
 *      - 액세스 권한: 모든 사용자
 *    → [배포] 후 권한 승인, 표시되는 '웹 앱 URL'(…/exec)을 복사합니다.
 * 5) 진단 웹앱의 [저장 설정]에 웹 앱 URL과 접근 코드를 입력합니다.
 *
 * 코드를 수정했다면 [배포] → [배포 관리] → 연필(수정) → 버전: 새 버전 → [배포] 해야 반영됩니다.
 */

var ACCESS_CODE = '여기에-접근코드-입력';   // 예: 'nutri2026!' — 진단 웹앱에 같은 코드를 입력합니다.
var SHEET_NAME = '진단기록';

var HEADERS = ['ID', '저장일시', '이름(별칭)', '학년', '응답자', '보호자 걱정', '진단 유형',
  '감각예민형', '정서형', '환경형', '생활습관형', '실천 약속', '상담 메모', '문항 응답(1~20)'];

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#EAF3DE');
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function ok_(code) {
  return ACCESS_CODE && ACCESS_CODE !== '여기에-접근코드-입력' && String(code) === ACCESS_CODE;
}

// 시트 수식으로 해석되지 않도록 = + - @ 로 시작하는 글자 앞에 ' 를 붙입니다.
function safe_(v) {
  v = v == null ? '' : String(v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

/** 기록 조회: GET ?code=…&action=list | ping */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!ok_(p.code)) return json_({ ok: false, error: '접근 코드가 올바르지 않습니다.' });
  if (p.action === 'ping') return json_({ ok: true });

  var sh = sheet_();
  var n = sh.getLastRow() - 1;
  if (n < 1) return json_({ ok: true, records: [] });
  var rows = sh.getRange(2, 1, n, HEADERS.length).getDisplayValues();
  var records = rows.filter(function (r) { return r[0]; }).map(function (r) {
    return { id: r[0], t: r[1], nm: r[2], grade: r[3], resp: r[4], concern: r[5], type: r[6],
      S: +r[7], E: +r[8], V: +r[9], L: +r[10], promise: r[11], memo: r[12], ans: r[13] };
  });
  return json_({ ok: true, records: records });
}

/** 기록 저장·삭제: POST (본문 JSON) */
function doPost(e) {
  var d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: '잘못된 요청' }); }
  if (!ok_(d.code)) return json_({ ok: false, error: '접근 코드가 올바르지 않습니다.' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = sheet_();
    if (d.action === 'delete') {
      var ids = sh.getRange(1, 1, sh.getLastRow(), 1).getDisplayValues();
      for (var i = ids.length - 1; i >= 1; i--) {
        if (ids[i][0] === String(d.id)) { sh.deleteRow(i + 1); return json_({ ok: true }); }
      }
      return json_({ ok: false, error: '기록을 찾지 못했습니다.' });
    }
    var r = d.record || {};
    var id = Utilities.getUuid().slice(0, 8);
    var when = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
    sh.appendRow([id, when, safe_(r.nm), safe_(r.grade), safe_(r.resp), safe_(r.concern), safe_(r.type),
      +r.S || 0, +r.E || 0, +r.V || 0, +r.L || 0, safe_(r.promise), safe_(r.memo), "'" + String(r.ans || '')]);
    return json_({ ok: true, id: id, t: when });
  } finally {
    lock.releaseLock();
  }
}
