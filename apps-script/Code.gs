/**
 * 서강대학교 채용박람회 성향/직무 테스트 — 구글 시트 연동용 Apps Script
 *
 * 설정 방법
 * 1. 결과를 쌓을 구글 스프레드시트를 새로 만듭니다.
 * 2. 메뉴에서 확장 프로그램 > Apps Script 를 엽니다.
 * 3. 기본 생성된 코드를 지우고 이 파일의 내용을 붙여넣습니다.
 * 4. 저장 후 배포 > 새 배포 > 유형 선택(웹 앱)을 클릭합니다.
 *    - 실행 계정: 나(본인)
 *    - 액세스 권한: 전체 허용(누구나)
 * 5. 배포 후 나오는 웹 앱 URL을 복사해서, index.html 상단의
 *    GOOGLE_SHEET_WEB_APP_URL 상수 값으로 붙여넣습니다.
 * 6. 코드를 수정한 뒤에는 "새 배포"가 아니라 기존 배포를 "관리 > 편집 > 새 버전"으로
 *    업데이트해야 URL이 바뀌지 않습니다.
 */

const SHEET_NAME = '참가자기록';

// 시트에 저장할 컬럼 순서. index.html의 buildLeadPayload / buildResultPayload가
// 보내는 필드명과 1:1로 맞춰져 있습니다.
const COLUMNS = [
  'participant_id',
  'timestamp',
  'stage', // 'lead' = 정보+동의만 제출한 단계, 'result' = 테스트까지 완료한 단계
  'name',
  'grade',
  'major',
  'email',
  'consent_agreed',
  'consent_at',
  'personality_code',
  'trait_scores_json',
  'character_animal',
  'nickname',
  'recommended_job_1',
  'recommended_job_2',
  'job_scores_json',
  'answers_json',
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getOrCreateSheet();
    const data = JSON.parse(e.postData.contents);
    const row = COLUMNS.map(col => (data[col] !== undefined ? data[col] : ''));

    const existingRow = findRowByParticipantId(sheet, data.participant_id);
    if (existingRow > 0) {
      // 같은 참가자가 이미 lead(정보+동의) 단계로 한 번 기록되어 있으면
      // 테스트 완료 시점에 같은 행을 결과로 덮어씁니다 (한 사람 = 한 행 유지).
      sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findRowByParticipantId(sheet, participantId) {
  if (!participantId) return -1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const idColIndex = COLUMNS.indexOf('participant_id') + 1;
  const ids = sheet.getRange(2, idColIndex, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === participantId) return i + 2;
  }
  return -1;
}
