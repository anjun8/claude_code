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

/**
 * doGet: 대시보드 '시트에서 불러오기'가 호출.
 * 이 스프레드시트의 '모든 탭'을 읽어 JSON으로 돌려줍니다. (탭 이름 = 계좌 이름)
 * 백업으로 만들어진 탭(보유종목/거래내역/…)은 제외하므로, 같은 시트에 백업이 있어도 안전합니다.
 * → 계좌가 늘면 탭만 추가하면 됩니다. 대시보드 설정은 URL 하나로 끝.
 */
function doGet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reserved = { '보유종목': 1, '거래내역': 1, '자산추이': 1, '입출금': 1, '월별요약': 1, '정보': 1 };
  var tabs = [];
  ss.getSheets().forEach(function (sh) {
    var name = sh.getName();
    if (reserved[name]) return;
    var rng = sh.getDataRange();
    if (rng.getNumRows() < 2) return;
    tabs.push({ name: name, rows: rng.getDisplayValues() });   // 표시값(날짜 등은 보이는 문자열로)
  });
  return ContentService
    .createTextOutput(JSON.stringify({ tabs: tabs }))
    .setMimeType(ContentService.MimeType.JSON);
}
