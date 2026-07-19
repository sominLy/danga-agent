/**
 * 단가 에이전트 v1 — Apps Script
 * 원천(01_원천) → 정제(02_관측DB) → 집계(03_채널마스터)
 *
 * 규칙 정의: schema.md 참조. 이 파일은 그 계약을 그대로 구현한다.
 * 설치: 확장프로그램 > Apps Script 에 붙여넣고 저장 → 시트 새로고침 → 상단 "단가에이전트" 메뉴.
 */

// 원천 탭 후보(위에서부터 먼저 찾은 것을 사용). '사본 만들기'로 만든 시트는 'History'가 원천.
var SHEET_RAW_CANDIDATES = ['01_원천', 'History'];
var SHEET_OBS = '02_관측DB';
var SHEET_CH  = '03_채널마스터';

/** 원천 탭을 이름 후보 순서대로 찾아 반환 */
function getRawSheet_(ss) {
  for (var i = 0; i < SHEET_RAW_CANDIDATES.length; i++) {
    var sh = ss.getSheetByName(SHEET_RAW_CANDIDATES[i]);
    if (sh) return sh;
  }
  return null;
}

var OBS_HEADERS = [
  'obs_id', '연도', '월', '연월', '제품', '테마', '닉네임',
  '채널링크_정규화', '플랫폼', '업로드링크', '콘텐츠유형',
  '실비_원본', '실비_숫자', '실비_상태', '대행사', '담당자', '협력사', '링크없음'
];
var CH_HEADERS = [
  '채널링크', '플랫폼', '대표닉네임', '관측수',
  '최신실비', '최신집행월', '실비_중앙값', '최소', '최대'
];

/** 시트 열 때 커스텀 메뉴 생성 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('단가에이전트')
    .addItem('① 정제 실행 (원천 → 관측DB)', 'runClean')
    .addItem('② 채널마스터 재생성 (관측DB → 채널마스터)', 'runChannelMaster')
    .addSeparator()
    .addItem('①+② 모두 실행', 'runAll')
    .addToUi();
}

function runAll() {
  runClean();
  runChannelMaster();
}

/* =========================================================
 * ① 정제: 01_원천 → 02_관측DB
 * ======================================================= */
function runClean() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var raw = getRawSheet_(ss);
  if (!raw) { alertMissing_(SHEET_RAW_CANDIDATES.join(' 또는 ')); return; }

  var values = raw.getDataRange().getValues();
  if (values.length < 2) { toast_('원천 데이터가 비어 있습니다.'); return; }

  // 앞에 빈 열/행이 있어도 되도록, 헤더 이름으로 열 위치를 찾는다(위치 가정 안 함).
  var map = locateColumns_(values);
  if (!map) {
    SpreadsheetApp.getUi().alert('원천에서 헤더(예: "채널 링크", "제품")를 찾지 못했습니다. History/01_원천 탭인지 확인하세요.');
    return;
  }

  var out = [OBS_HEADERS];
  for (var r = map.headerRow + 1; r < values.length; r++) {
    var row = values[r];
    if (isBlankRow_(row)) continue;
    // 제품·닉네임·링크·실비가 모두 비면 데이터 행이 아니라고 보고 건너뜀
    if (cell_(row, map.제품) === '' && cell_(row, map.닉네임) === ''
        && cell_(row, map.채널링크) === '' && cell_(row, map.실비) === '') continue;

    var year  = cell_(row, map.연도);
    var month = parseMonth_(cell_(row, map.월));
    var ym    = (year !== '' && month !== '') ? (year + '-' + pad2_(month)) : '';
    var link  = normalizeLink_(cell_(row, map.채널링크));
    var platform = classifyPlatform_(link);
    var uploadLink = cell_(row, map.업로드링크);
    var feeRaw = cell_(row, map.실비);
    var fee = parseFee_(feeRaw);

    out.push([
      r + 1,                       // obs_id = 원천 행번호(1-based)
      year,                        // 연도
      month,                       // 월
      ym,                          // 연월
      cell_(row, map.제품),         // 제품
      cell_(row, map.테마),         // 테마
      cell_(row, map.닉네임),       // 닉네임
      link,                        // 채널링크_정규화
      platform,                    // 플랫폼
      uploadLink,                  // 업로드링크
      classifyContent_(uploadLink),// 콘텐츠유형
      feeRaw,                      // 실비_원본
      fee.amount,                  // 실비_숫자
      fee.status,                  // 실비_상태
      cell_(row, map.대행사),       // 대행사
      cell_(row, map.담당자),       // 담당자
      cell_(row, map.협력사),       // 협력사
      (link === '')                // 링크없음
    ]);
  }

  writeSheet_(ss, SHEET_OBS, out);
  toast_('정제 완료: 관측 ' + (out.length - 1) + '건');
}

/* =========================================================
 * ② 집계: 02_관측DB → 03_채널마스터 (채널링크 = 키)
 * ======================================================= */
function runChannelMaster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var obs = ss.getSheetByName(SHEET_OBS);
  if (!obs) { alertMissing_(SHEET_OBS + ' (먼저 ① 정제 실행)'); return; }

  var values = obs.getDataRange().getValues();
  if (values.length < 2) { toast_('관측DB가 비어 있습니다. 먼저 ① 정제 실행.'); return; }

  var idx = headerIndex_(values[0]);
  var groups = {}; // 채널링크 -> 관측 배열

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (row[idx['링크없음']] === true || row[idx['링크없음']] === 'TRUE') continue;
    var link = row[idx['채널링크_정규화']];
    if (link === '' || link == null) continue;
    if (!groups[link]) groups[link] = [];
    groups[link].push({
      ym: row[idx['연월']],
      nick: row[idx['닉네임']],
      platform: row[idx['플랫폼']],
      amount: row[idx['실비_숫자']],
      status: row[idx['실비_상태']]
    });
  }

  var out = [CH_HEADERS];
  Object.keys(groups).forEach(function (link) {
    var arr = groups[link];
    // 최신순 정렬(연월 내림차순)
    arr.sort(function (a, b) { return String(b.ym).localeCompare(String(a.ym)); });

    var latest = null; // 정상/무상 중 최신
    for (var i = 0; i < arr.length; i++) {
      if (arr[i].status === '정상' || arr[i].status === '무상') { latest = arr[i]; break; }
    }
    var normalAmts = arr
      .filter(function (a) { return a.status === '정상' && a.amount !== '' && a.amount != null; })
      .map(function (a) { return Number(a.amount); });

    out.push([
      link,
      arr[0].platform,
      arr[0].nick,                                   // 최신 연월의 닉네임
      arr.length,
      latest ? latest.amount : '',
      latest ? latest.ym : '',
      normalAmts.length ? median_(normalAmts) : '',
      normalAmts.length ? Math.min.apply(null, normalAmts) : '',
      normalAmts.length ? Math.max.apply(null, normalAmts) : ''
    ]);
  });

  // 관측수 많은 채널 우선 정렬
  var body = out.slice(1).sort(function (a, b) { return b[3] - a[3]; });
  writeSheet_(ss, SHEET_CH, [CH_HEADERS].concat(body));
  toast_('채널마스터 완료: 채널 ' + body.length + '개');
}

/* =========================================================
 * 정제 헬퍼 (규칙은 schema.md와 일치)
 * ======================================================= */

function parseMonth_(v) {
  if (v === '' || v == null) return '';
  var m = String(v).match(/\d+/);
  return m ? Number(m[0]) : '';
}

function normalizeLink_(v) {
  if (v === '' || v == null) return '';
  var s = String(v).trim();
  if (s === '') return '';
  s = s.split('?')[0].split('#')[0];      // 쿼리/프래그먼트 제거
  s = s.replace(/\/+$/, '');               // 끝 슬래시 제거
  s = s.replace(/^http:\/\//i, 'https://');
  if (!/^https?:\/\//i.test(s)) s = 'https://' + s;
  // 호스트만 소문자화, 경로 대소문자 보존
  s = s.replace(/^(https:\/\/)([^\/]+)/i, function (_, p, host) { return p + host.toLowerCase(); });
  return s;
}

function classifyPlatform_(link) {
  if (link === '' || link == null) return '(없음)';
  var l = link.toLowerCase();
  if (l.indexOf('instagram.com') >= 0) return 'instagram';
  if (l.indexOf('youtube.com') >= 0 || l.indexOf('youtu.be') >= 0) return 'youtube';
  if (l.indexOf('tiktok.com') >= 0) return 'tiktok';
  return '기타';
}

function classifyContent_(v) {
  if (v === '' || v == null) return '';
  var s = String(v).toLowerCase();
  if (s.indexOf('/reel/') >= 0) return 'reel';
  if (s.indexOf('/p/') >= 0) return 'post';
  if (s.indexOf('/shorts/') >= 0) return 'shorts';
  if (s.indexOf('/watch') >= 0) return 'watch';
  if (s.indexOf('/video/') >= 0) return 'video';
  return String(v).trim() === '' ? '' : '기타';
}

function parseFee_(v) {
  var raw = (v == null) ? '' : String(v).trim();
  if (raw === '') return { amount: '', status: '진행중' };
  if (/^(-|협의중|tbd|미정)$/i.test(raw)) return { amount: '', status: '미정' };
  var digits = raw.replace(/[₩,\s]/g, '').replace(/원$/,'');
  if (!/^-?\d+(\.\d+)?$/.test(digits)) return { amount: '', status: '미정' };
  var num = Number(digits);
  if (num === 0) return { amount: 0, status: '무상' };
  return { amount: num, status: '정상' };
}

/* =========================================================
 * 공용 유틸
 * ======================================================= */

function trim_(v) { return (v == null) ? '' : String(v).trim(); }
function pad2_(n) { return (n < 10 ? '0' : '') + n; }

/** 값을 가진 셀(열 인덱스가 유효할 때만) */
function cell_(row, colIdx) {
  return (colIdx == null || colIdx < 0) ? '' : trim_(row[colIdx]);
}

/**
 * 원천 2D 배열에서 헤더 행과 각 필드의 열 인덱스를 찾는다.
 * 앞쪽 빈 열/행이 있어도 동작. 실패 시 null.
 */
function locateColumns_(values) {
  var aliases = {
    연도:   ['협업 집행 연도', '집행 연도', '연도', '집행연도'],
    월:     ['집행 월', '월', '집행월'],
    제품:   ['제품'],
    테마:   ['협업 콘텐츠 테마', '콘텐츠 테마', '테마'],
    닉네임: ['채널대표닉네임(or설명)', '채널대표닉네임', '닉네임', '채널명', '대표닉네임'],
    채널링크: ['채널 링크', '채널링크', '채널 url', '채널url'],
    업로드링크: ['업로드 링크', '업로드링크', '콘텐츠 링크', '게시물 링크'],
    실비:   ['광고 집행 실비용', '집행 실비용', '실비용', '실비', '단가', '금액'],
    대행사: ['대표 대행사', '대행사'],
    담당자: ['담당자'],
    협력사: ['협력사', '협업사']
  };
  var norm = function (s) { return String(s == null ? '' : s).replace(/\s+/g, '').toLowerCase(); };

  // 헤더 행 = '채널링크' 별칭을 포함한 첫 행(상단 30행 내 탐색)
  for (var r = 0; r < Math.min(values.length, 30); r++) {
    var row = values[r];
    var map = { headerRow: r };
    var found = {};
    for (var c = 0; c < row.length; c++) {
      var v = norm(row[c]);
      if (v === '') continue;
      for (var key in aliases) {
        if (found[key]) continue;
        var list = aliases[key];
        for (var i = 0; i < list.length; i++) {
          if (v === norm(list[i])) { map[key] = c; found[key] = true; break; }
        }
      }
    }
    // 채널링크 + 실비를 찾았으면 이 행을 헤더로 확정
    if (found['채널링크'] && found['실비']) return map;
  }
  return null;
}

function isBlankRow_(row) {
  for (var i = 0; i < row.length; i++) {
    if (row[i] !== '' && row[i] != null) return false;
  }
  return true;
}

function median_(nums) {
  var a = nums.slice().sort(function (x, y) { return x - y; });
  var mid = Math.floor(a.length / 2);
  return a.length % 2 ? a[mid] : (a[mid - 1] + a[mid]) / 2;
}

/** 헤더 배열 -> {헤더명: 열인덱스} */
function headerIndex_(headerRow) {
  var idx = {};
  for (var i = 0; i < headerRow.length; i++) idx[String(headerRow[i]).trim()] = i;
  return idx;
}

/** 시트를 지우고 2D 배열을 통째로 씀 */
function writeSheet_(ss, name, values) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clearContents();
  if (values.length === 0) return;
  sh.getRange(1, 1, values.length, values[0].length).setValues(values);
  sh.setFrozenRows(1);
}

function alertMissing_(name) {
  SpreadsheetApp.getUi().alert('시트 "' + name + '" 가 없습니다. schema.md의 셋업 순서를 확인하세요.');
}

function toast_(msg) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, '단가에이전트', 5);
}
