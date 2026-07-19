/**
 * 단가 조회 웹앱 — 시트 내장 버전 (보안 강화)
 *
 * 데이터가 구글 밖으로 나가지 않는다:
 * - 웹앱은 이 시트의 Apps Script 안에서 실행
 * - 구글 로그인 + 이 시트를 공유받은 사람만 접속 가능
 *
 * ── 설치 (시트 소유자, 한 번만) ──
 * 1. 단가 관리 시트에서 확장 프로그램 > Apps Script
 * 2. 기본 Code.gs에 이 파일(서버.gs) 내용 붙여넣기
 * 3. 좌측 파일 옆 + 버튼 > HTML → 이름을 "index"로 → index.html 내용 붙여넣기
 * 4. [배포] > [새 배포] > 유형: 웹 앱
 *    - 실행 계정: 나
 *    - 액세스 권한: "Google 계정이 있는 모든 사용자"  ← 로그인 필수가 됨
 * 5. 나온 웹 앱 URL을 팀에 공유 (시트도 팀원에게 공유돼 있어야 함)
 *
 * ── 사용 ──
 * - 시트에 "제품명(월)" 형식 탭을 추가하면 웹이 자동 인식
 * - 시트 수정 후엔 웹 새로고침만 하면 반영
 *
 * ── 접근 통제 ──
 * getTabs()가 실행될 때 접속자의 시트 열람 권한을 함께 확인한다.
 * 시트를 공유받지 못한 사람은 로그인해도 데이터를 볼 수 없다.
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('단가 조회')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function getTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tabs = [];
  ss.getSheets().forEach(function (sh) {
    var name = sh.getName();
    // "제품명(월)" 형식 탭만 (전각 괄호 허용, '템플릿' 포함 탭 제외)
    if (!/[（(][^)）]+[)）]\s*$/.test(name)) return;
    if (name.indexOf('템플릿') >= 0) return;
    var vals = sh.getDataRange().getDisplayValues();
    if (vals.length < 2) return;
    tabs.push({ name: name, rows: vals });
  });
  return { ok: true, updated: new Date().toISOString(), tabs: tabs };
}
