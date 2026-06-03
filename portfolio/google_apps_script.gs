/**
 * 포트폴리오 대시보드 → 구글 시트 백업용 Apps Script
 *
 * [붙여넣는 곳] 구글 시트 → 확장 프로그램(Extensions) → Apps Script
 * 아래 코드를 통째로 붙여넣고 저장 → "배포(Deploy) → 새 배포 → 웹 앱"
 *   - 실행 주체(Execute as): 나(Me)
 *   - 액세스 권한(Who has access): 모든 사용자(Anyone)   ← 대시보드가 익명으로 보내기 때문
 * 그러면 나오는 웹앱 URL(…/exec)을 대시보드의 '시트연결' 버튼에 붙여넣으세요.
 *
 * ⚠️ 이 URL을 아는 사람은 이 시트에 쓸 수 있어요(읽기는 불가). 남에게 공유하지 마세요.
 */
function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  function writeSheet(name, rows) {
    if (!rows || !rows.length) return;
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.clearContents();
    var width = Math.max.apply(null, rows.map(function (r) { return r.length; }));
    var norm = rows.map(function (r) {
      var a = r.slice();
      while (a.length < width) a.push('');
      return a;
    });
    sh.getRange(1, 1, norm.length, width).setValues(norm);
    sh.setFrozenRows(1);
  }

  writeSheet('보유종목', data.holdings);
  writeSheet('거래내역', data.txlog);
  writeSheet('자산추이', data.history);
  writeSheet('입출금', data.cashlog);
  writeSheet('월별요약', data.monthly);
  writeSheet('정보', data.meta);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// (선택) 브라우저로 URL을 직접 열었을 때 동작 확인용
function doGet() {
  return ContentService.createTextOutput('OK - 포트폴리오 백업 웹앱이 동작 중입니다.');
}
