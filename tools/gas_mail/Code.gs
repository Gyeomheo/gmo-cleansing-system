/* GMO 주간 메일 · 리마인드 — Gmail 초안 자동 생성 (보내지는 않는다)
 *
 * 흐름: GUI [④ 메일] → 파이썬이 '내 드라이브/GMO_메일초안/메일_W<nn>.json' 저장
 *       → 이 스크립트(5분마다)가 파일을 읽어 Gmail 초안을 만들고 파일을 '처리됨' 으로 옮긴다
 *   · 주간 메일: 검색어(thread_query, 기본 in:sent subject:"GMO D2C Pacing Report") 로 찾은 가장 최근 스레드에
 *               '전체 회신' 초안 (본문 = Collection status 표 포함 HTML)
 *   · 리마인드: 법인별 새 초안(받는 사람은 비워 둠 — 초안에서 넣고 보낸다)
 *
 * 한 번 할 일
 *   1. script.google.com → 새 프로젝트 → 이 파일 내용 붙여넣기 → 저장
 *   2. dryRun() 실행 → 권한 허용(Gmail · Drive) → 로그에 찾은 스레드 제목이 맞는지 확인
 *   3. installTrigger() 실행 → 5분마다 자동
 *   (회사 계정에서 Apps Script 가 막혀 있으면 2 에서 오류가 난다 — 그때는 GUI 의 클립보드 방식만 쓴다)
 */

var DROP_FOLDER = 'GMO_메일초안';        // 내 드라이브 아래 폴더 이름
var DONE_FOLDER = '처리됨';

function folder_(name, parent) {
  var it = (parent || DriveApp.getRootFolder()).getFoldersByName(name);
  return it.hasNext() ? it.next() : (parent || DriveApp.getRootFolder()).createFolder(name);
}

function thread_(query) {
  var ts = GmailApp.search(query || 'in:sent subject:"GMO D2C Pacing Report"', 0, 1);
  return ts.length ? ts[0] : null;
}

/** 확인만: 찾은 스레드와 대기 중인 파일 */
function dryRun() {
  var t = thread_();
  Logger.log('스레드: ' + (t ? t.getFirstMessageSubject() + ' (' + t.getLastMessageDate() + ')' : '못 찾음'));
  var all = folder_(DROP_FOLDER).getFiles();
  while (all.hasNext()) Logger.log('대기: ' + all.next().getName());
}

/** 대기 파일 → 초안 */
function makeDrafts() {
  var drop = folder_(DROP_FOLDER), done = folder_(DONE_FOLDER, drop);
  var files = drop.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    if (!/\.json$/i.test(f.getName())) continue;
    var d = JSON.parse(f.getBlob().getDataAsString('UTF-8'));
    var t = thread_(d.thread_query);
    if (t) {
      t.createDraftReplyAll('', { htmlBody: d.html });
    } else {
      GmailApp.createDraft('', d.subject, '', { htmlBody: d.html });
    }
    (d.reminders || []).forEach(function (r) {
      GmailApp.createDraft('', r.subject, r.text);
    });
    f.moveTo(done);
    Logger.log(f.getName() + ' → 초안 ' + (1 + (d.reminders || []).length) + '개');
  }
}

function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'makeDrafts') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('makeDrafts').timeBased().everyMinutes(5).create();
  Logger.log('5분마다 makeDrafts 실행');
}
