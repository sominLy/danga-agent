/**
 * 단가 시트 → 웹 연결용 API
 *
 * 설치 (시트 소유자가 한 번만):
 * 1. 단가 관리 시트에서 확장 프로그램 > Apps Script → 이 코드 붙여넣고 저장
 * 2. 우상단 [배포] > [새 배포] > 유형: 웹 앱
 *    - 실행 계정: 나
 *    - 액세스 권한: "링크가 있는 모든 사용자"
 * 3. 나온 웹 앱 URL(https://script.google.com/macros/s/.../exec)을 복사해
 *    단가조회 웹의 ⚙(연결 설정)에 붙여넣기
 *
 * 동작:
 * - 이름이 "제품명(월)" 형식인 탭만 자동으로 골라 전부 내려준다.
 *   → 탭을 새로 추가하면 웹이 다음 로드 때 자동 인식. 별도 작업 불필요.
 * - (템플릿), Dashboard 같은 형식 밖 탭은 무시된다.
 * - 열 해석(공유가/실비/컨택 매핑)은 웹 쪽에서 하므로 여기선 값을 그대로 보낸다.
 *
 * 시트 내용을 수정한 뒤에는 웹을 새로고침하면 반영된다.
 */
function doGet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tabs = [];
  ss.getSheets().forEach(function (sh) {
    var name = sh.getName();
    // "제품명(월)" 형식 탭만: 끝이 (…)로 끝나는 이름 (전각 괄호 허용, '템플릿' 포함 탭 제외)
    if (!/[（(][^)）]+[)）]\s*$/.test(name)) return;
    if (name.indexOf('템플릿') >= 0) return;
    var vals = sh.getDataRange().getDisplayValues();
    if (vals.length < 2) return;
    tabs.push({ name: name, rows: vals });
  });
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, updated: new Date().toISOString(), tabs: tabs }))
    .setMimeType(ContentService.MimeType.JSON);
}
