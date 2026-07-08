/**
 * 용디쌤 퍼스널 페이지 — '오늘 할 한 가지' 구독자 수집 백엔드 (Google Apps Script)
 *
 * 동작: 페이지 구독폼 제출 → 이 스크립트가 받아서
 *   1) 스프레드시트 '구독자' 시트에 한 줄 추가 (= 구독자 명단)
 *   2) NOTIFY_EMAIL로 즉시 메일 알림 발송
 *
 * 설치 방법은 같은 폴더의 SETUP.md 참고 (5분 소요)
 */

// ===== 설정 (여기만 수정하면 됩니다) =====
const SHEET_NAME = '구독자';
const NOTIFY_EMAIL = 'skdyddns@gmail.com'; // 알림 받을 메일. 끄려면 '' 로.

// ===== 이하 수정 불필요 =====
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const email = String(data.email || '').trim().toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) {
      return json({ ok: false, error: 'bad_email' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['신청일시', '이메일', '상태', '메모']);
      sheet.setFrozenRows(1);
      sheet.getRange('A1:D1').setFontWeight('bold').setBackground('#E8F4EC');
      sheet.setColumnWidths(1, 4, 180);
    }

    // 중복 구독 방지
    const emails = sheet.getRange(2, 2, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat();
    if (emails.indexOf(email) !== -1) {
      return json({ ok: true, dup: true }); // 이미 구독 중이어도 사용자에겐 성공으로
    }

    const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([now, email, '구독중', '']);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: '[용디쌤 구독 신청] ' + email,
        body:
          '새 구독 신청이 접수되었습니다.\n\n' +
          '신청일시: ' + now + '\n' +
          '이메일: ' + email + '\n\n' +
          '전체 명단(스프레드시트): ' + ss.getUrl()
      });
    }

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
