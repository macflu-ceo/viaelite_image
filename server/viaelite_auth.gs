/**
 * VIA ELITE 컨시어지 로그인 인증 시스템 (v2)
 * Google Apps Script
 *
 * 셋업:
 *  1. 이 코드를 시트의 "확장 프로그램 → Apps Script"에 통째로 붙여넣기
 *  2. 아래 SHEET_ID 값을 본인 시트 ID로 변경
 *  3. 저장 후 "배포 → 새 배포 → 웹 앱" 으로 게시
 *     - 다음 사용자 인증: 나
 *     - 액세스 권한: 모든 사용자
 *
 * 시트 컬럼 (1행 헤더, 순서 중요):
 *  A: ID                       (수동)
 *  B: 비밀번호                  (수동)
 *  C: 이름                      (수동)
 *  D: 활성                      (수동, TRUE/FALSE)
 *  E: 가입일                    (수동)
 *  F: 마지막 로그인              (자동)
 *  G: 로그인 횟수                (자동)
 *  H: 이미지 제작 횟수            (자동)
 *  I: 마지막 로그인 기기 (mo/pc)  (자동)
 *  J: AI 사용 횟수                (자동)  ← 신규 (AI모델컷 생성 횟수)
 */

// === 본인 시트 ID로 변경하세요 ===
const SHEET_ID = '여기에_본인_시트_ID_붙여넣기';
const SHEET_NAME = '컨시어지명단';

// ============================================================
// 엔트리 포인트
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'login') {
      return jsonResponse(handleLogin(data.id, data.password, data.device));
    }
    if (data.action === 'card_created') {
      return jsonResponse(handleCardCreated(data.id));
    }
    if (data.action === 'ai_used') {
      return jsonResponse(handleAiUsed(data.id));
    }
    return jsonResponse({ success: false, message: '알 수 없는 요청입니다.' });
  } catch (err) {
    return jsonResponse({ success: false, message: '서버 오류: ' + err.message });
  }
}

// ============================================================
// 로그인 처리
// ============================================================
function handleLogin(id, password, device) {
  if (!id || !password) {
    return { success: false, message: 'ID와 비밀번호를 모두 입력해주세요.' };
  }

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) {
    return { success: false, message: '시트를 찾을 수 없습니다.' };
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return { success: false, message: '등록된 사용자가 없습니다.' };
  }

  const headers = data[0];
  const idCol = headers.indexOf('ID');
  const pwCol = headers.indexOf('비밀번호');
  const nameCol = headers.indexOf('이름');
  const activeCol = headers.indexOf('활성');
  const joinDateCol = headers.indexOf('가입일');
  const lastLoginCol = headers.indexOf('마지막 로그인');
  const countCol = headers.indexOf('로그인 횟수');
  const deviceCol = headers.indexOf('마지막 로그인 기기');

  if (idCol === -1 || pwCol === -1) {
    return { success: false, message: '시트 헤더에 "ID" 또는 "비밀번호" 컬럼이 없습니다.' };
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[idCol]).trim() === String(id).trim()
        && String(row[pwCol]) === String(password)) {

      if (activeCol !== -1) {
        const isActive = row[activeCol] === true
                      || String(row[activeCol]).toUpperCase() === 'TRUE'
                      || row[activeCol] === 1;
        if (!isActive) {
          return { success: false, message: '비활성화된 계정입니다. 본사에 문의하세요.' };
        }
      }

      const rowIdx = i + 1;
      if (lastLoginCol !== -1) sheet.getRange(rowIdx, lastLoginCol + 1).setValue(new Date());
      if (countCol !== -1) sheet.getRange(rowIdx, countCol + 1).setValue((Number(row[countCol]) || 0) + 1);
      if (deviceCol !== -1) {
        const deviceVal = (device === 'mo' || device === 'pc') ? device : 'unknown';
        sheet.getRange(rowIdx, deviceCol + 1).setValue(deviceVal);
      }

      return {
        success: true,
        id: String(row[idCol]),
        name: nameCol !== -1 ? String(row[nameCol] || '') : '',
        joinDate: joinDateCol !== -1 && row[joinDateCol] ? formatDate(row[joinDateCol]) : ''
      };
    }
  }

  return { success: false, message: 'ID 또는 비밀번호가 일치하지 않습니다.' };
}

// ============================================================
// 이미지 제작 카운트 증가
// ============================================================
function handleCardCreated(id) {
  if (!id) return { success: false, message: 'ID가 필요합니다.' };

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) return { success: false, message: '시트를 찾을 수 없습니다.' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('ID');
  const imgCountCol = headers.indexOf('이미지 제작 횟수');

  if (idCol === -1 || imgCountCol === -1) {
    return { success: false, message: '필수 컬럼이 없습니다.' };
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]).trim() === String(id).trim()) {
      const cur = Number(data[i][imgCountCol]) || 0;
      sheet.getRange(i + 1, imgCountCol + 1).setValue(cur + 1);
      return { success: true, newCount: cur + 1 };
    }
  }
  return { success: false, message: '사용자를 찾을 수 없습니다.' };
}

// ============================================================
// AI 사용(모델컷 생성) 카운트 증가
// ============================================================
function handleAiUsed(id) {
  if (!id) return { success: false, message: 'ID가 필요합니다.' };

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) return { success: false, message: '시트를 찾을 수 없습니다.' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('ID');
  let aiCountCol = headers.indexOf('AI 사용 횟수');

  if (idCol === -1) {
    return { success: false, message: 'ID 컬럼이 없습니다.' };
  }
  // "AI 사용 횟수" 헤더가 없으면 마지막 컬럼 다음에 자동 생성
  if (aiCountCol === -1) {
    aiCountCol = headers.length;
    sheet.getRange(1, aiCountCol + 1).setValue('AI 사용 횟수');
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]).trim() === String(id).trim()) {
      const cur = Number(data[i][aiCountCol]) || 0;
      sheet.getRange(i + 1, aiCountCol + 1).setValue(cur + 1);
      return { success: true, newCount: cur + 1 };
    }
  }
  return { success: false, message: '사용자를 찾을 수 없습니다.' };
}

// ============================================================
// 유틸
// ============================================================
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDate(d) {
  if (!d) return '';
  try {
    const date = new Date(d);
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  } catch (e) {
    return String(d);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput('VIA ELITE Auth Service v2 — 정상 작동 중입니다.')
    .setMimeType(ContentService.MimeType.TEXT);
}

// 테스트용
function testConnection() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    Logger.log('헤더: ' + data[0].join(' | '));
    Logger.log('등록된 사용자 수: ' + (data.length - 1));
  } catch (err) {
    Logger.log('오류: ' + err.message);
  }
}
