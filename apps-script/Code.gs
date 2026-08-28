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
 * 5. 처음 배포/실행 시 Google Drive·Gmail 접근 권한 승인 팝업이 뜹니다.
 *    (사원증 카드 이미지를 드라이브에 저장하고, 필요하면 메일로 보내기 위한 권한입니다)
 *    본인 계정으로 승인해주세요.
 * 6. 배포 후 나오는 웹 앱 URL을 복사해서, index.html 상단의
 *    GOOGLE_SHEET_WEB_APP_URL 상수 값으로 붙여넣습니다.
 * 7. 코드를 수정한 뒤에는 "새 배포"가 아니라 기존 배포를 "관리 > 편집 > 새 버전"으로
 *    업데이트해야 URL이 바뀌지 않습니다. (Drive/Gmail 권한이 새로 추가된 경우
 *    새 버전 배포 시 권한 재승인 팝업이 다시 뜰 수 있어요 — 승인해주세요)
 */

const SHEET_NAME = '참가자기록';

// 참가자 결과 카드(PNG) 이미지를 저장할 구글 드라이브 폴더 이름.
// 없으면 첫 업로드 때 자동으로 만들어집니다.
const DRIVE_FOLDER_NAME = '채용박람회 사원증 카드';

// 카드 이미지가 생성될 때마다 이 주소로 메일을 보낼지 여부.
// 비워두면('') 구글 드라이브 저장만 하고 메일은 보내지 않습니다.
// 채워두면 카드가 완성되는 즉시 해당 주소로 PNG가 첨부된 메일이 발송됩니다.
// (참가자가 아주 많으면 Gmail 발송 일일 한도(개인 계정 100통/워크스페이스 계정 1500통)에
//  걸릴 수 있으니, 참가자 수가 많을 땐 비워두고 드라이브 폴더만 확인하는 걸 추천해요)
const NOTIFY_EMAIL = 'ahjung@spigen.com';

// 시트에 저장할 컬럼 순서. index.html의 buildLeadPayload / buildResultPayload /
// autoUploadCardPNG가 보내는 필드명과 1:1로 맞춰져 있습니다.
const COLUMNS = [
  'participant_id',
  'timestamp',
  'stage', // 'lead' = 정보+동의, 'result' = 테스트 완료, 'card' = 결과 카드 이미지 업로드
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
  'card_png_url',
  'interest_job',
];
// ⚠️ 컬럼을 더 추가할 일이 생기면 항상 이 배열의 "맨 끝"에 추가하세요.
// 중간에 끼워 넣으면 이미 만들어진 시트의 기존 컬럼과 순서가 어긋나서
// 예전 데이터 위의 헤더가 틀어질 수 있어요. (아래 getOrCreateSheet가
// 새로 추가된 컬럼의 헤더는 자동으로 뒤에 이어붙여줍니다)

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getOrCreateSheet();
    const data = JSON.parse(e.postData.contents);

    // 사원증 카드 이미지가 왔으면 드라이브에 저장하고, base64 원본 대신 파일 URL만 시트에 남깁니다.
    // (base64 원본은 구글 시트 셀 용량 제한(약 5만자)을 넘기기 쉬워서 시트에 직접 넣지 않아요)
    if (data.card_png_base64) {
      data.card_png_url = saveCardToDrive(data.participant_id, data.card_png_base64);
      delete data.card_png_base64;
    }

    const existingRow = findRowByParticipantId(sheet, data.participant_id);
    if (existingRow > 0) {
      // lead → result → card 순서로 여러 번 요청이 들어올 수 있으므로,
      // 이번 요청에 실려온 필드만 덮어쓰고 나머지 열은 기존 값을 그대로 둡니다.
      const current = sheet.getRange(existingRow, 1, 1, COLUMNS.length).getValues()[0];
      const merged = COLUMNS.map((col, i) => (data[col] !== undefined ? data[col] : current[i]));
      sheet.getRange(existingRow, 1, 1, merged.length).setValues([merged]);
    } else {
      const row = COLUMNS.map(col => (data[col] !== undefined ? data[col] : ''));
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

// base64로 인코딩된 카드 PNG를 구글 드라이브에 저장하고, 설정되어 있으면 메일로도 보냅니다.
// 참가자 기기·브라우저의 다운로드 지원 여부와 완전히 무관하게 동작하는 "무조건 저장" 경로입니다.
function saveCardToDrive(participantId, dataUrl) {
  const folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
  const base64 = dataUrl.indexOf(',') >= 0 ? dataUrl.split(',')[1] : dataUrl;
  const fileName = `${participantId || 'unknown'}.png`;
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/png', fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  if (NOTIFY_EMAIL) {
    try {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: `[채용박람회] 사원증 카드 - ${participantId || ''}`,
        body: '참가자의 결과 카드 이미지가 첨부되어 있습니다. 사원증 제작에 사용해주세요.\n\n드라이브 링크: ' + file.getUrl(),
        attachments: [blob],
      });
    } catch (mailErr) {
      // 메일 발송에 실패해도 드라이브 저장은 이미 끝났으므로 계속 진행합니다.
    }
  }

  return file.getUrl();
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
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
  } else {
    // 예전 버전 스크립트로 이미 만들어진 시트라 헤더가 COLUMNS보다 짧을 수 있어요.
    // 이 경우 새로 추가된 컬럼의 헤더만 뒤에 이어붙입니다 (기존 컬럼 위치는 그대로 유지).
    const currentHeaderLen = sheet.getLastColumn();
    if (currentHeaderLen < COLUMNS.length) {
      const missingHeaders = COLUMNS.slice(currentHeaderLen);
      sheet.getRange(1, currentHeaderLen + 1, 1, missingHeaders.length).setValues([missingHeaders]);
    }
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
