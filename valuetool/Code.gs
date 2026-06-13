/**
 * 밸류툴 자동 업데이트 (Google Apps Script)
 * ------------------------------------------------------------------
 * 가격     : GOOGLEFINANCE("KRX:<코드>")  → _price 시트
 * 펀더멘털 : Open DART API (다중회사 주요계정) → _dart 시트
 *
 * 기존 모델은 _price / _dart 시트를 VLOOKUP 으로 참조하면 됩니다.
 * (지금 시트 "1"에 가격을 붙여넣고 VLOOKUP 하던 방식과 동일)
 *
 * 사용 전:
 *   1) 메뉴 "확장 프로그램 > Apps Script" 에 이 파일을 붙여넣기
 *   2) 아래 CONFIG.DART_KEY 에 OpenDART 인증키 입력
 *      (https://opendart.fss.or.kr → 인증키 신청, 이메일 가입, 무료)
 *   3) 종목 목록 시트 "_codes" 준비 (codes_seed.csv 를 import).
 *      A열=stock_code(6자리), B열=name. 1행은 헤더.
 *   4) 시트 새로고침 후 상단 "📊 밸류툴" 메뉴 사용
 * ------------------------------------------------------------------
 */

var CONFIG = {
  // ── OpenDART 인증키 (필수) ──────────────────────────────
  DART_KEY: 'YOUR_OPENDART_API_KEY_HERE',

  // ── 기준 연도 / 분기순익 분리 대상 ──────────────────────
  BSNS_YEAR: 2025,          // 조회할 사업연도

  // ── 시트 이름 ───────────────────────────────────────────
  CODES_SHEET: '_codes',    // 입력: stock_code(A), name(B)
  PRICE_SHEET: '_price',    // 출력: 코드/종목/현재가(GOOGLEFINANCE)
  DART_SHEET:  '_dart',     // 출력: 펀더멘털
  CORPMAP_SHEET: '_corpmap',// 캐시: stock_code → corp_code

  // ── DART 다중회사 호출 시 한 번에 묶는 종목 수 ──────────
  // 응답이 비거나 오류면 50 → 20 으로 줄여보세요.
  DART_CHUNK: 100,

  // ── 한 번 실행에서 처리할 청크 수 (6분 제한 회피) ───────
  // 다 못 끝내면 자동으로 1분 뒤 이어서 실행됩니다.
  CHUNKS_PER_RUN: 8
};

var REPRT = { Q1: '11013', H1: '11012', Q3: '11014', FY: '11011' };

// ════════════════════════════════════════════════════════════════
// 메뉴
// ════════════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 밸류툴')
    .addItem('① 가격 시트 만들기/갱신 (GOOGLEFINANCE)', 'setupPriceSheet')
    .addSeparator()
    .addItem('② corp_code 매핑 구축 (최초 1회)', 'buildCorpMap')
    .addItem('③ DART 펀더멘털 갱신 (분기마다)', 'updateDartFundamentals')
    .addSeparator()
    .addItem('진행상태 초기화', 'resetProgress')
    .addToUi();
}

// ════════════════════════════════════════════════════════════════
// ① 가격: GOOGLEFINANCE 수식을 _price 시트에 작성
// ════════════════════════════════════════════════════════════════
function setupPriceSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var codes = readCodes_();
  if (!codes.length) { alert_('_codes 시트가 비어 있습니다. codes_seed.csv 를 먼저 import 하세요.'); return; }

  var sh = getOrCreateSheet_(CONFIG.PRICE_SHEET);
  sh.clear();
  sh.getRange(1, 1, 1, 4).setValues([['stock_code', 'name', '현재가', '갱신시각']]);

  var rows = codes.map(function (c) {
    // KRX: 접두사는 코스피/코스닥 구분 불필요. GOOGLEFINANCE 가 가격 자동 조회.
    return [c.code, c.name, '=IFERROR(GOOGLEFINANCE("KRX:' + c.code + '"),"")', ''];
  });
  sh.getRange(2, 1, rows.length, 4).setValues(rows);
  // 갱신시각 한 칸 (전체 공통)
  sh.getRange('D2').setValue(new Date());
  sh.setFrozenRows(1);
  alert_('가격 시트 완료: ' + rows.length + '개 종목. GOOGLEFINANCE 가 약 20분 지연가로 자동 갱신됩니다.');
}

// ════════════════════════════════════════════════════════════════
// ② corp_code 매핑: corpCode.xml(zip) 다운로드 → 파싱 → _corpmap 캐시
// ════════════════════════════════════════════════════════════════
function buildCorpMap() {
  requireKey_();
  var url = 'https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key=' + CONFIG.DART_KEY;
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) { alert_('corpCode 다운로드 실패: HTTP ' + resp.getResponseCode()); return; }

  // 응답은 zip. 압축 해제 후 CORPCODE.xml 파싱.
  var blob = resp.getBlob().setContentType('application/zip');
  var files = Utilities.unzip(blob);
  var xmlBlob = null;
  for (var i = 0; i < files.length; i++) {
    if (/\.xml$/i.test(files[i].getName())) { xmlBlob = files[i]; break; }
  }
  if (!xmlBlob) { alert_('zip 안에서 XML 을 찾지 못했습니다.'); return; }

  var doc = XmlService.parse(xmlBlob.getDataAsString('UTF-8'));
  var list = doc.getRootElement().getChildren('list');
  var map = {}; // stock_code -> corp_code
  for (var j = 0; j < list.length; j++) {
    var stock = (list[j].getChildText('stock_code') || '').trim();
    var corp = (list[j].getChildText('corp_code') || '').trim();
    if (stock && corp) map[stock] = corp; // 상장사만 (stock_code 존재)
  }

  // 우리 종목 목록에 매칭해서 _corpmap 에 저장
  var codes = readCodes_();
  var out = [['stock_code', 'corp_code', 'name']];
  var miss = 0;
  codes.forEach(function (c) {
    var corp = map[c.code];
    if (corp) out.push([c.code, corp, c.name]);
    else { out.push([c.code, '', c.name]); miss++; }
  });
  var sh = getOrCreateSheet_(CONFIG.CORPMAP_SHEET);
  sh.clear();
  sh.getRange(1, 1, out.length, 3).setValues(out);
  sh.setFrozenRows(1);
  alert_('corp_code 매핑 완료: ' + (out.length - 1) + '개 중 미매칭 ' + miss + '개.');
}

// ════════════════════════════════════════════════════════════════
// ③ DART 펀더멘털 갱신 (다중회사 주요계정 + 분기순익 분리)
//    6분 제한 회피: 청크 단위로 처리하고 진행상태를 저장,
//    남으면 1분 뒤 자동으로 이어서 실행.
// ════════════════════════════════════════════════════════════════
function updateDartFundamentals() {
  requireKey_();
  var pairs = readCorpMap_();           // [{code, corp, name}]
  if (!pairs.length) { alert_('먼저 ②corp_code 매핑을 구축하세요.'); return; }
  pairs = pairs.filter(function (p) { return p.corp; });

  var sh = getOrCreateSheet_(CONFIG.DART_SHEET);
  var header = ['stock_code', 'name', '매출액', '영업이익',
                '순익_1Q', '순익_2Q', '순익_3Q', '순익_4Q', '순익_연간', '자본총계', '갱신'];
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.setFrozenRows(1);
  }

  // 진행 커서
  var props = PropertiesService.getDocumentProperties();
  var cursor = parseInt(props.getProperty('DART_CURSOR') || '0', 10);

  var chunkSize = CONFIG.DART_CHUNK;
  var done = 0;
  for (var run = 0; run < CONFIG.CHUNKS_PER_RUN && cursor < pairs.length; run++) {
    var slice = pairs.slice(cursor, cursor + chunkSize);
    processChunk_(slice, sh);
    cursor += slice.length;
    done += slice.length;
  }
  props.setProperty('DART_CURSOR', String(cursor));

  if (cursor < pairs.length) {
    scheduleContinuation_();
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '진행 ' + cursor + '/' + pairs.length + ' … 1분 뒤 자동 이어서 실행', '밸류툴', 8);
  } else {
    props.deleteProperty('DART_CURSOR');
    clearContinuation_();
    alert_('DART 펀더멘털 갱신 완료: ' + pairs.length + '개 종목.');
  }
}

// 한 청크(여러 종목) 처리: 4개 보고서를 묶어 받아 분기순익 분리
function processChunk_(slice, sh) {
  var corps = slice.map(function (p) { return p.corp; });
  // 보고서별로 다중회사 호출 (corp_code 콤마 결합)
  var data = {
    Q1: fetchMulti_(corps, REPRT.Q1),
    H1: fetchMulti_(corps, REPRT.H1),
    Q3: fetchMulti_(corps, REPRT.Q3),
    FY: fetchMulti_(corps, REPRT.FY)
  };

  var rows = slice.map(function (p) {
    var q1 = pick_(data.Q1, p.corp);
    var h1 = pick_(data.H1, p.corp);
    var q3 = pick_(data.Q3, p.corp);
    var fy = pick_(data.FY, p.corp);

    // 손익은 누적 → 분기 분리
    var ni1 = num_(q1.net);
    var ni2 = sub_(h1.net, q1.net);
    var ni3 = sub_(q3.net, h1.net);
    var ni4 = sub_(fy.net, q3.net);
    var niY = num_(fy.net);

    // 매출/영업이익/자본총계: 가장 최근 가용 보고서 기준
    var latest = fy.has ? fy : (q3.has ? q3 : (h1.has ? h1 : q1));
    return [p.code, p.name, num_(latest.sales), num_(latest.op),
            ni1, ni2, ni3, ni4, niY, num_(latest.equity), new Date()];
  });

  // _dart 시트에 코드 기준 upsert
  upsertRows_(sh, rows);
}

// 다중회사 주요계정 호출 → { corp_code: {sales,op,net,equity,has} }
function fetchMulti_(corps, reprt) {
  var url = 'https://opendart.fss.or.kr/api/fnlttMultiAcnt.json'
          + '?crtfc_key=' + CONFIG.DART_KEY
          + '&corp_code=' + corps.join(',')
          + '&bsns_year=' + CONFIG.BSNS_YEAR
          + '&reprt_code=' + reprt;
  var out = {};
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(resp.getContentText());
    if (json.status !== '000' || !json.list) return out; // 013=데이터없음 등
    json.list.forEach(function (it) {
      var cc = it.corp_code;
      if (!out[cc]) out[cc] = { has: true };
      var amt = it.thstrm_amount; // 당기금액(누적)
      switch (it.account_nm) {
        case '매출액': out[cc].sales = amt; break;
        case '영업이익': out[cc].op = amt; break;
        case '당기순이익': out[cc].net = amt; break;
        case '자본총계': out[cc].equity = amt; break;
      }
    });
  } catch (e) { /* 네트워크/파싱 오류는 빈 결과로 처리 */ }
  return out;
}

// ════════════════════════════════════════════════════════════════
// 헬퍼
// ════════════════════════════════════════════════════════════════
function pick_(map, corp) { return map[corp] || { has: false }; }

function num_(v) {
  if (v === undefined || v === null || v === '' || v === '-') return '';
  var n = parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? '' : n / 1e8; // 원 → 억원
}
function sub_(a, b) {
  var x = num_(a), y = num_(b);
  if (x === '' && y === '') return '';
  return (x === '' ? 0 : x) - (y === '' ? 0 : y);
}

function readCodes_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.CODES_SHEET);
  if (!sh) return [];
  var vals = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < vals.length; i++) {
    var code = String(vals[i][0] || '').trim();
    if (!code) continue;
    code = ('000000' + code).slice(-6); // 6자리 0패딩
    out.push({ code: code, name: String(vals[i][1] || '').trim() });
  }
  return out;
}

function readCorpMap_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.CORPMAP_SHEET);
  if (!sh) return [];
  var vals = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < vals.length; i++) {
    out.push({ code: String(vals[i][0] || '').trim(),
               corp: String(vals[i][1] || '').trim(),
               name: String(vals[i][2] || '').trim() });
  }
  return out;
}

// 코드 기준 upsert (기존 행 갱신, 없으면 추가)
function upsertRows_(sh, rows) {
  var last = sh.getLastRow();
  var index = {};
  if (last > 1) {
    var codes = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < codes.length; i++) index[String(codes[i][0])] = i + 2;
  }
  rows.forEach(function (r) {
    var key = String(r[0]);
    var rowNum = index[key];
    if (rowNum) sh.getRange(rowNum, 1, 1, r.length).setValues([r]);
    else { sh.appendRow(r); index[key] = sh.getLastRow(); }
  });
}

function getOrCreateSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function requireKey_() {
  if (!CONFIG.DART_KEY || CONFIG.DART_KEY === 'YOUR_OPENDART_API_KEY_HERE')
    throw new Error('CONFIG.DART_KEY 에 OpenDART 인증키를 입력하세요.');
}

function alert_(msg) { SpreadsheetApp.getUi().alert(msg); }

// ── 6분 제한 회피용 1회성 트리거 ───────────────────────────
function scheduleContinuation_() {
  clearContinuation_();
  ScriptApp.newTrigger('updateDartFundamentals').timeBased().after(60 * 1000).create();
}
function clearContinuation_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'updateDartFundamentals') ScriptApp.deleteTrigger(t);
  });
}
function resetProgress() {
  PropertiesService.getDocumentProperties().deleteProperty('DART_CURSOR');
  clearContinuation_();
  alert_('진행상태를 초기화했습니다.');
}
