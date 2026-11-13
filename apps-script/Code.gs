/**
 * 온앤쿨케어 대상자 판정 기록 — 구글 시트 수신 스크립트
 * ------------------------------------------------------------
 * 사용법 (README.md의 '구글 시트 연결' 참고)
 *  1) 빈 구글 시트 → 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기 → 저장
 *  2) 위쪽 함수 선택에서 setup 선택 → 실행 → 권한 허용  ('판정기록', '집계' 시트가 만들어짐)
 *  3) 배포 → 새 배포 → 유형: 웹 앱
 *       - 다음 사용자 인증 정보로 실행: 나
 *       - 액세스 권한이 있는 사용자: 모든 사용자
 *  4) 나온 웹 앱 URL(…/exec)을 GitHub의 config.js → endpoint 에 붙여넣기
 *
 * 같은 '기관명 + 대상자명'으로 다시 기록하면 새 줄을 만들지 않고 최신 판정으로 덮어씁니다.
 */

const TOKEN = 'onc-a486d4ee6e7c';   // config.js의 token과 똑같아야 합니다. 비우면 검사 안 함
const SHEET_NAME = '판정기록';
const SUMMARY_NAME = '집계';

// 시트 열 순서 (A~V)
const HEADERS = [
  '최초기록일시', '최종수정일시', '기관명', '담당자', '대상자명', '세대주',
  '가구원수', '소득자격', '건보가입형태', '건보료평균(원)', '건보료기준(원)', '소득판정',
  '유형', '세대원요건', '하절기(원)', '동절기(원)', '연간합계(원)', '서류구비',
  '미구비서류', '판정의견', '메모', '기록키'
];
const KEY_COL = HEADERS.indexOf('기록키') + 1;   // 22 (V열)

/** 최초 1회 실행: 시트·머리글·집계표 만들기 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME, 0);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#1f2933').setFontColor('#ffffff').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  sh.getRange('A:B').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange('J:K').setNumberFormat('#,##0');
  sh.getRange('O:Q').setNumberFormat('#,##0');
  sh.setColumnWidths(1, HEADERS.length, 110);
  sh.setColumnWidth(19, 260);   // 미구비서류
  sh.hideColumns(KEY_COL);      // 기록키는 숨김

  // 판정의견 색상 표시
  const rng = sh.getRange('T2:T');
  const rules = [
    ['적합', '#ebfbee', '#2b8a3e'], ['보류', '#fff9db', '#b06d00'], ['부적합', '#fff5f5', '#c92a2a']
  ].map(([txt, bg, fg]) => SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith(txt).setBackground(bg).setFontColor(fg).setRanges([rng]).build());
  sh.setConditionalFormatRules(rules);

  // 집계 시트 (QUERY 배열 함수 — 기록이 늘면 자동으로 따라 늘어남)
  const sm = ss.getSheetByName(SUMMARY_NAME) || ss.insertSheet(SUMMARY_NAME);
  sm.clear();
  const src = `'${SHEET_NAME}'!A1:V`;
  sm.getRange('A1').setValue('① 유형 × 판정의견 (가구 수)').setFontWeight('bold');
  sm.getRange('A2').setFormula(`=IFERROR(QUERY(${src},"select M, count(E) where M is not null group by M pivot T",1),"기록 없음")`);
  sm.getRange('A10').setValue('② 적합 판정 가구 — 유형별 가구 수·냉난방비').setFontWeight('bold');
  sm.getRange('A11').setFormula(`=IFERROR(QUERY(${src},"select M, count(E), sum(Q) where T starts with '적합' group by M label count(E) '가구 수', sum(Q) '연간 냉·난방비'",1),"기록 없음")`);
  sm.getRange('A19').setValue('③ 기관·담당자별 기록 수').setFontWeight('bold');
  sm.getRange('A20').setFormula(`=IFERROR(QUERY(${src},"select C, D, count(E) where E is not null group by C, D label count(E) '기록 수'",1),"기록 없음")`);
  sm.getRange('E1').setValue('④ 서류 미구비 가구').setFontWeight('bold');
  sm.getRange('E2').setFormula(`=IFERROR(QUERY(${src},"select C, E, M, S where S is not null and S <> '' order by B desc",1),"없음")`);
  sm.getRange('B11:C17').setNumberFormat('#,##0');
  sm.setColumnWidth(1, 220); sm.setColumnWidth(8, 360);
  SpreadsheetApp.flush();
}

/** 페이지에서 보낸 판정 결과 받기 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const d = JSON.parse(e.postData.contents || '{}');
    if (TOKEN && d.token !== TOKEN) return out_({ ok: false, error: 'token' });
    if (!d.org || !d.caseId) return out_({ ok: false, error: '기관명과 대상자 이름이 필요합니다.' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) { setup(); sh = ss.getSheetByName(SHEET_NAME); }

    const now = new Date();
    const key = (String(d.org).trim() + '|' + String(d.caseId).trim()).toLowerCase();
    const row = [
      now, now, d.org, d.staff, d.caseId, d.head,
      num_(d.hh), d.qual, d.ins, num_(d.insAvg), num_(d.insLimit), d.incomeResult,
      d.type, d.members, num_(d.summer), num_(d.winter), num_(d.total), d.docs,
      d.missing, d.opinion, d.memo, key
    ].map(safe_);

    // 같은 기관+관리번호가 있으면 덮어쓰기(최초기록일시는 유지)
    const last = sh.getLastRow();
    if (last > 1) {
      const hit = sh.getRange(2, KEY_COL, last - 1, 1).createTextFinder(key).matchEntireCell(true).findNext();
      if (hit) {
        const r = hit.getRow();
        row[0] = sh.getRange(r, 1).getValue() || now;
        sh.getRange(r, 1, 1, row.length).setValues([row]);
        return out_({ ok: true, updated: true, row: r });
      }
    }
    sh.appendRow(row);
    return out_({ ok: true, updated: false, row: sh.getLastRow() });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/** 브라우저에서 웹 앱 주소를 열었을 때 연결 확인용 */
function doGet() {
  return out_({ ok: true, service: '온앤쿨케어 판정 기록', time: new Date().toISOString() });
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function num_(v) { const n = Number(v); return (v === '' || v == null || isNaN(n)) ? '' : n; }
// 수식 주입 방지: = + - @ 로 시작하는 글자는 앞에 ' 를 붙여 문자로 저장
function safe_(v) {
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v == null ? '' : v;
}
