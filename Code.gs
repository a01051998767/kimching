/**
 * 함께 담그는 비건 김장 — 참가신청 저장용 Apps Script
 *
 * 쓰는 법 (자세한 건 README.md):
 *  1) 구글 시트에서 [확장 프로그램 → Apps Script]를 열고 이 파일 내용을 통째로 붙여넣기
 *  2) 위쪽 함수 목록에서 setup 을 골라 ▶실행 (처음 한 번, "신청"·"요약" 탭이 만들어져요)
 *  3) [배포 → 새 배포 → 웹 앱]으로 배포하고 URL을 index.html 의 APPS_SCRIPT_URL 에 넣기
 *
 * 금액·마감일을 바꾸면: 아래 CONFIG 수정 → setup 다시 실행 → 배포 새 버전으로 업데이트
 * (index.html 맨 위 CONFIG도 똑같이 바꿔 주세요)
 */
const CONFIG = {
  BASE_FEE: 3000,            // 1인 점심값 (전·라면 재료비)
  PER_HEAD: 17800,           // 1포기당 참가비 (절인배추 6,800 + 양념 약 11,000)
  HEADS_PER_BOX: 5,          // 절인배추 10kg 1박스 포기 수 (4~6포기의 가운데 값)
  DEADLINE: '2026-11-24',    // 신청 마감일 (이 날 밤 12시까지, 한국 시간)

  BOX_PRICE: 34000,          // 절인배추 10kg 1박스 가격 (예상 총 재료비 계산용)
  SEASONING_PER_HEAD: 11000, // 1포기당 양념 재료비 (예상 총 재료비 계산용)

  SHEET: '신청',
  SUMMARY: '요약'
};

const HEADERS = ['신청시각', '이름', '연락처', '포기 수', '하고 싶은 말', '예상 참가비', '수정시각'];

/** 처음 한 번 실행: 탭·제목줄·요약 수식을 만들어요. 다시 실행해도 신청 내역은 지워지지 않아요. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Seoul');

  let sh = ss.getSheetByName(CONFIG.SHEET);
  if (!sh) sh = ss.insertSheet(CONFIG.SHEET, 0);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#e6f2e1');
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange('C:C').setNumberFormat('@');           // 연락처는 글자로 (앞자리 0이 사라지지 않게)
  sh.getRange('F:F').setNumberFormat('#,##0"원"');
  sh.getRange('G:G').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange('A1:G1').setNumberFormat('@');
  sh.setColumnWidth(1, 140); sh.setColumnWidth(3, 130); sh.setColumnWidth(5, 260); sh.setColumnWidth(7, 140);

  let sm = ss.getSheetByName(CONFIG.SUMMARY);
  if (!sm) sm = ss.insertSheet(CONFIG.SUMMARY);
  sm.clear();
  const s = "'" + CONFIG.SHEET + "'!";
  sm.getRange('A1:B15').setValues([
    ['항목', '값'],
    ['총 인원', '=COUNTA(' + s + 'B2:B)'],
    ['총 포기 수', '=SUM(' + s + 'D2:D)'],
    ['절인배추 박스 수 (10kg)', '=ROUNDUP(B3/B12,0)'],
    ['예상 총 재료비 (박스×박스값 + 포기×양념비)', '=B4*B13+B3*B14'],
    ['예상 참가비 총액', '=SUM(' + s + 'F2:F)'],
    ['차액 (참가비 − 재료비 = 점심값 + 여유분)', '=B6-B5'],
    ['', ''],
    ['▼ 설정값 (Code.gs의 CONFIG에서 고친 뒤 setup을 다시 실행하세요)', ''],
    ['1인 점심값', CONFIG.BASE_FEE],
    ['1포기당 참가비', CONFIG.PER_HEAD],
    ['절인배추 1박스 포기 수', CONFIG.HEADS_PER_BOX],
    ['절인배추 1박스 가격', CONFIG.BOX_PRICE],
    ['1포기당 양념 재료비', CONFIG.SEASONING_PER_HEAD],
    ['신청 마감일', CONFIG.DEADLINE]
  ]);
  sm.getRange('A1:B1').setFontWeight('bold').setBackground('#e6f2e1');
  sm.getRange('A2:A7').setFontWeight('bold');
  sm.getRange('B2:B4').setNumberFormat('#,##0');
  sm.getRange('B5:B7').setNumberFormat('#,##0"원"');
  sm.getRange('B10:B11').setNumberFormat('#,##0"원"');
  sm.getRange('B13:B14').setNumberFormat('#,##0"원"');
  sm.getRange('B15').setNumberFormat('@');
  sm.setColumnWidth(1, 380); sm.setColumnWidth(2, 140);

  // 이미 있는 신청 행의 예상 참가비 수식도 다시 채워요.
  const last = sh.getLastRow();
  for (let r = 2; r <= last; r++) sh.getRange(r, 6).setFormula(feeFormula_(r));
}

/** 현황 조회: 숫자만 돌려줘요 (이름·연락처는 절대 내보내지 않음). */
function doGet(e) {
  return json_(Object.assign({ ok: true, closed: isClosed_() }, summary_()));
}

/** 신청 저장 / 수정 */
function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, status: 'error', message: '요청을 읽지 못했어요.' });
  }
  if (isClosed_()) return json_({ ok: false, status: 'closed' });

  const name = clean_(data.name, 30);
  const digits = String(data.phone || '').replace(/\D/g, '');
  const heads = Number(data.heads);
  const message = clean_(data.message, 100);

  if (!name) return json_({ ok: false, status: 'error', message: '이름을 적어 주세요.' });
  if (!/^01[016789]\d{7,8}$/.test(digits)) return json_({ ok: false, status: 'error', message: '휴대폰 번호를 확인해 주세요.' });
  if (!Number.isInteger(heads) || heads < 1 || heads > 100000) return json_({ ok: false, status: 'error', message: '포기 수를 확인해 주세요.' });

  const phone = digits.length === 10
    ? digits.slice(0, 3) + '-' + digits.slice(3, 6) + '-' + digits.slice(6)
    : digits.slice(0, 3) + '-' + digits.slice(3, 7) + '-' + digits.slice(7);

  // 동시에 여러 명이 신청해도 줄이 겹치지 않게 잠금
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return json_({ ok: false, status: 'error', message: '지금 신청이 몰리고 있어요. 잠시 뒤에 다시 시도해 주세요.' });
  }
  try {
    const sh = sheet_();
    const last = sh.getLastRow();
    let row = 0;
    if (last >= 2) {
      const phones = sh.getRange(2, 3, last - 1, 1).getDisplayValues();
      for (let i = 0; i < phones.length; i++) {
        if (String(phones[i][0]).replace(/\D/g, '') === digits) { row = i + 2; break; }
      }
    }
    const now = new Date();
    let status;
    if (row) {
      if (!data.update) return json_({ ok: false, status: 'duplicate' });
      sh.getRange(row, 2, 1, 4).setValues([[name, phone, heads, message]]);
      sh.getRange(row, 7).setValue(now);
      status = 'updated';
    } else {
      row = last + 1;
      sh.getRange(row, 3).setNumberFormat('@');
      sh.getRange(row, 1, 1, 5).setValues([[now, name, phone, heads, message]]);
      status = 'created';
    }
    sh.getRange(row, 6).setFormula(feeFormula_(row));
    SpreadsheetApp.flush();
    return json_(Object.assign({ ok: true, status: status }, summary_()));
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 내부 함수 ---------- */

function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET);
  if (!sh) throw new Error('"' + CONFIG.SHEET + '" 탭이 없어요. setup 함수를 먼저 실행해 주세요.');
  return sh;
}

function summary_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  let count = 0, heads = 0;
  if (last >= 2) {
    sh.getRange(2, 2, last - 1, 3).getValues().forEach(function (r) {
      if (String(r[0]).trim() === '') return;
      count++;
      heads += Number(r[2]) || 0;
    });
  }
  return { count: count, heads: heads };
}

function feeFormula_(row) {
  const sm = "'" + CONFIG.SUMMARY + "'!";
  return '=IF(D' + row + '="","",' + sm + '$B$10+' + sm + '$B$11*D' + row + ')';
}

function isClosed_() {
  return Date.now() > new Date(CONFIG.DEADLINE + 'T23:59:59+09:00').getTime();
}

/** 글자 수를 자르고, 시트에서 수식으로 실행될 수 있는 첫 글자(= + - @)를 막아요. */
function clean_(v, max) {
  let s = String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
