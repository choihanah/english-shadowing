/**
 * 스마트 영어 패턴 섀도잉 - 망각곡선(SRS) 구글 시트 연동 Apps Script
 * 
 * [설치 방법]
 * 1. 구글 시트 상단 메뉴 > 확장 프로그램 > Apps Script 클릭
 * 2. 기존 코드를 모두 지우고 아래 코드를 붙여넣은 뒤 저장 (Ctrl+S)
 * 3. 우측 상단 [배포] > [새 배포] 클릭
 * 4. 유형 선택: [웹 앱]
 *    - 다음 사용자로 실행: 나
 *    - 액세스 권한이 있는 사용자: 모든 사용자
 * 5. [배포] 클릭 후 생성된 웹 앱 URL(https://script.google.com/macros/s/.../exec)을 복사
 * 6. 섀도잉 웹앱 우측 상단 ⚙️(설정) 버튼에 URL을 붙여넣고 저장하면 연동 완료!
 */

const LOG_SHEET_NAME = '복습로그';
const HEADERS = ['표현', '뜻', '암기단계', '다음복습일', '맞힘', '틀림', '최근학습일'];

function getOrCreateLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#eef2ff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doGet(e) {
  try {
    const sheet = getOrCreateLogSheet();
    const rows = sheet.getDataRange().getValues();
    const records = {};

    for (let i = 1; i < rows.length; i++) {
      const phrase = String(rows[i][0] || '').trim();
      if (!phrase) continue;
      const key = phrase.toLowerCase();
      records[key] = {
        phrase: phrase,
        meaning: String(rows[i][1] || '').trim(),
        level: Number(rows[i][2]) || 0,
        nextReviewDate: formatDateVal(rows[i][3]),
        correctCount: Number(rows[i][4]) || 0,
        wrongCount: Number(rows[i][5]) || 0,
        lastReviewed: formatDateVal(rows[i][6])
      };
    }

    const payload = JSON.stringify({ status: 'ok', records: records });
    const callback = e && e.parameter && e.parameter.callback;
    if (callback) {
      return ContentService.createTextOutput(`${callback}(${payload})`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(payload)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : null;
    if (!body || !body.phrase) {
      throw new Error('유효한 데이터가 아닙니다.');
    }

    const sheet = getOrCreateLogSheet();
    const rows = sheet.getDataRange().getValues();
    const targetKey = String(body.phrase).trim().toLowerCase();

    const rowData = [
      String(body.phrase).trim(),
      String(body.meaning || '').trim(),
      Number(body.level) || 0,
      String(body.nextReviewDate || ''),
      Number(body.correctCount) || 0,
      Number(body.wrongCount) || 0,
      String(body.lastReviewed || '')
    ];

    let foundRow = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0] || '').trim().toLowerCase() === targetKey) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow > 0) {
      sheet.getRange(foundRow, 1, 1, HEADERS.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function formatDateVal(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val).trim();
}
