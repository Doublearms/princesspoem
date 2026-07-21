/**
 * 사용법
 * 1) 구글시트를 새로 만든다 (예: "마음시집 방명록")
 * 2) 메뉴에서 확장 프로그램 > Apps Script 를 연다
 * 3) 기존 코드를 지우고 이 파일 내용 전체를 붙여넣는다
 * 4) 아래 ADMIN_KEY 값을 원하는 비밀번호로 바꾼다 (index.html에서 입력할 값과 똑같아야 함)
 * 5) 저장 후 오른쪽 위 "배포하기" > "새 배포" 클릭
 *    - 유형: 웹 앱
 *    - 실행 계정: 나
 *    - 액세스 권한이 있는 사용자: 모든 사용자   (반드시 이렇게 설정해야 방문자가 접속 가능)
 * 6) 배포하면 나오는 웹 앱 URL(.../exec 로 끝남)을 복사해서
 *    index.html 의 SHEET_WEBAPP_URL 자리에 붙여넣는다
 * 7) 구글시트 자체 링크(공유 > 링크 보기, "링크가 있는 모든 사용자: 뷰어")도 복사해서
 *    index.html 의 SHEET_VIEW_URL 자리에 붙여넣는다
 *
 * 보안 메모: 이름을 "쓰는" 것(doPost)은 방문자 누구나 할 수 있어야 하므로 그대로 열려 있습니다.
 * 목록을 "읽는" 것(doGet)만 아래 ADMIN_KEY와 일치할 때만 데이터를 내려주도록 잠갔습니다.
 *
 * 이메일 알림 관련 안내 (중요, 기존에 이미 배포해두신 분들만 해당):
 * 이번에 MailApp(이메일 발송) 기능을 추가하면서 새로운 권한이 필요해졌습니다.
 * 코드를 붙여넣고 저장한 뒤, 아래 순서로 한 번 권한을 승인해줘야 정상 작동합니다.
 * 1) 위쪽 함수 선택 드롭다운에서 doGet 선택 후 "실행" 버튼 클릭
 * 2) "권한 검토" 팝업이 뜨면 본인 구글 계정 선택 → "고급" → "이동(안전하지 않음)" → "허용"
 * 3) 그 다음 "배포하기 > 배포 관리" 에서 연필(수정) 아이콘 클릭 → 버전을 "새 버전"으로 선택 → "배포"
 *    (이 단계를 꼭 해야 실제 웹 앱 URL에도 새 코드가 반영됩니다)
 */

const ADMIN_KEY = 'star2026'; // ← 원하는 비밀번호로 바꾸세요. index.html의 값과 동일해야 합니다.
const NOTIFY_EMAIL = 'bynatura@naver.com'; // ← 방문자 알림을 받을 이메일 주소

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  let data = {};
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    data = {};
  }
  const name = (data.name || '').toString().slice(0, 50);
  if (name) {
    const now = new Date();
    sheet.appendRow([now, name]);

    // 방문자 등록 알림 이메일 발송 (실패해도 방명록 저장 자체는 문제 없이 진행됨)
    try {
      const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: '[시집 방명록] ' + name + '님이 방문했어요',
        body: name + '님이 ' + timeStr + '에 방문해서 이름을 남겼어요.\n\n구글시트에서 전체 목록 확인하기:\n' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
      });
    } catch (mailErr) {
      // 메일 발송 권한이 아직 승인 안 됐거나 일일 발송 한도를 초과한 경우 등
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const key = (e.parameter && e.parameter.key) || '';
  if (key !== ADMIN_KEY) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'unauthorized' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const rows = sheet.getDataRange().getValues();
  const visitors = rows
    .filter(r => r[1])
    .map(r => ({ ts: r[0], name: r[1] }))
    .reverse();
  return ContentService
    .createTextOutput(JSON.stringify({ visitors }))
    .setMimeType(ContentService.MimeType.JSON);
}
