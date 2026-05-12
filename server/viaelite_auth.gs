/**
 * VIA ELITE 컨시어지 로그인 인증 시스템
 * Google Apps Script
 *
 * 셋업:
 *  1. 이 코드를 시트의 "확장 프로그램 → Apps Script"에 통째로 붙여넣기
 *  2. 아래 SHEET_ID 값을 본인 시트 ID로 변경
 *  3. 저장 후 "배포 → 새 배포 → 웹 앱" 으로 게시
 *     - 다음 사용자 인증: 나
 *     - 액세스 권한: 모든 사용자
 *  4. 배포 후 표시되는 웹 앱 URL을 HTML의 APPS_SCRIPT_URL 자리에 붙여넣기
 *
 * 시트 컬럼 (1행은 헤더):
 *   ID | 비밀번호 | 이름 | 활성 | 레벨 | 소속 | 가입일 | 마지막 로그인 | 로그인 횟수
 */

// === 본인 시트 ID로 변경하세요 ===
const SHEET_ID = '여기에_본인_시트_ID_붙여넣기';
const SHEET_NAME = '컨시어지명단';

// ============================================================
// 로그인 처리
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'login') {
      return jsonResponse(handleLogin(data.id, data.password));
    }
    return jsonResponse({ success: false, message: '알 수 없는 요청입니다.' });
  } catch (err) {
    return jsonResponse({ success: false, message: '서버 오류: ' + err.message });
  }
}

function handleLogin(id, password) {
  if (!id || !password) {
    return { success: false, message: 'ID와 비밀번호를 모두 입력해주세요.' };
  }

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) {
    return { success: false, message: '시트를 찾을 수 없습니다. 시트 이름을 확인하세요.' };
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
  const levelCol = headers.indexOf('레벨');
  const teamCol = headers.indexOf('소속');
  const lastLoginCol = headers.indexOf('마지막 로그인');
  const countCol = headers.indexOf('로그인 횟수');

  if (idCol === -1 || pwCol === -1) {
    return { success: false, message: '시트 헤더에 "ID" 또는 "비밀번호" 컬럼이 없습니다.' };
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[idCol]).trim() === String(id).trim()
        && String(row[pwCol]) === String(password)) {

      // 활성 여부 체크
      if (activeCol !== -1) {
        const isActive = row[activeCol] === true
                      || String(row[activeCol]).toUpperCase() === 'TRUE'
                      || row[activeCol] === 1;
        if (!isActive) {
          return { success: false, message: '비활성화된 계정입니다. 본사에 문의하세요.' };
        }
      }

      // 마지막 로그인 시각, 로그인 횟수 업데이트
      const rowIdx = i + 1;
      if (lastLoginCol !== -1) {
        sheet.getRange(rowIdx, lastLoginCol + 1).setValue(new Date());
      }
      if (countCol !== -1) {
        const cur = Number(row[countCol]) || 0;
        sheet.getRange(rowIdx, countCol + 1).setValue(cur + 1);
      }

      return {
        success: true,
        id: String(row[idCol]),
        name: nameCol !== -1 ? String(row[nameCol] || '') : '',
        level: levelCol !== -1 ? String(row[levelCol] || '') : '',
        '소속': teamCol !== -1 ? String(row[teamCol] || '') : ''
      };
    }
  }

  return { success: false, message: 'ID 또는 비밀번호가 일치하지 않습니다.' };
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 헬스 체크 (브라우저에서 URL 직접 열면 호출됨)
function doGet(e) {
  return ContentService
    .createTextOutput('VIA ELITE Auth Service — 정상 작동 중입니다.\n\nPOST /login 으로 인증 요청을 보내세요.')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ============================================================
// 테스트용: Apps Script 편집기에서 직접 실행해서 시트 연결 확인
// ============================================================
function testConnection() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    Logger.log('헤더: ' + data[0].join(' | '));
    Logger.log('등록된 사용자 수: ' + (data.length - 1));
    if (data.length > 1) {
      Logger.log('첫 사용자 ID: ' + data[1][0]);
    }
  } catch (err) {
    Logger.log('오류: ' + err.message);
  }
}
