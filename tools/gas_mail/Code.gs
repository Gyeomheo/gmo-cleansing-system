/* GMO 주간 메일 · 리마인드 — Gmail 초안 자동 생성 · 리마인드 자동 발송(켤 때만)
 *
 * 흐름: GUI [④ 메일] → 파이썬이 '내 드라이브/GMO_메일초안/메일_W<nn>.json' 저장
 *       → 이 스크립트(5분마다)가 파일을 읽어 Gmail 초안을 만들고 파일을 '처리됨' 으로 옮긴다
 *   · 주간 메일: 검색어(thread_query, 기본 in:sent 제목 "GMO Weekly Report" 또는 옛 제목) 로 찾은 가장 최근 스레드에
 *               '전체 회신' 초안 (본문 = Collection status 표 포함 HTML)
 *   · 리마인드: 법인마다 그 법인 스레드(r.query, 기본 제목에 법인 코드) 중 가장 최신 것에 '전체 회신'
 *               — 받는 사람 · 참조는 스레드 그대로. 기본은 초안, 스크립트 속성 AUTO_SEND = true 면 바로 발송
 *               같은 주 · 같은 법인(r.key)은 한 번만. 스레드를 못 찾으면 새 메일을 만들지 않고 건너뛴다(로그)
 *   · 주간 메일(PCG)은 늘 초안 — 보내기는 사람
 *   · 기록: 내 드라이브/GMO_메일초안/리마인드기록.txt 에 한 줄씩
 *
 * 한 번 할 일
 *   1. script.google.com → 새 프로젝트 → 이 파일 내용 붙여넣기 → 저장
 *   2. dryRun() 실행 → 권한 허용(Gmail · Drive) → 로그에 찾은 스레드 제목이 맞는지 확인
 *   3. installTrigger() 실행 → 5분마다 자동
 *   4. (발송 켜기) 프로젝트 설정 → 스크립트 속성 AUTO_SEND = true. 끄기는 false 또는 지우기
 *      켜기 전에 dryRun() 로그의 '리마인드 → 스레드' 가 법인마다 맞는지 볼 것
 *   (회사 계정에서 Apps Script 가 막혀 있으면 2 에서 오류가 난다 — 그때는 GUI 의 클립보드 방식만 쓴다)
 */

var DROP_FOLDER = 'GMO_메일초안';        // 내 드라이브 아래 폴더 이름
var DONE_FOLDER = '처리됨';

function folder_(name, parent) {
  var it = (parent || DriveApp.getRootFolder()).getFoldersByName(name);
  return it.hasNext() ? it.next() : (parent || DriveApp.getRootFolder()).createFolder(name);
}

function thread_(query) {
  var ts = GmailApp.search(query || 'in:sent (subject:"GMO Weekly Report" OR subject:"GMO D2C Weekly Report" OR subject:"GMO D2C Pacing Report")', 0, 1);
  return ts.length ? ts[0] : null;
}

function autoSend_() {
  return PropertiesService.getScriptProperties().getProperty('AUTO_SEND') === 'true';
}

function log_(line) {
  var drop = folder_(DROP_FOLDER), it = drop.getFilesByName('리마인드기록.txt');
  var stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  if (it.hasNext()) {
    var f = it.next();
    f.setContent(f.getBlob().getDataAsString('UTF-8') + stamp + '  ' + line + '\n');
  } else {
    drop.createFile('리마인드기록.txt', stamp + '  ' + line + '\n');
  }
  Logger.log(line);
}

/** 확인만: PCG 스레드, 대기 파일, 리마인드마다 찾은 스레드(가장 최신) */
function dryRun() {
  var t = thread_();
  Logger.log('PCG 스레드: ' + (t ? t.getFirstMessageSubject() + ' (' + t.getLastMessageDate() + ')' : '못 찾음'));
  Logger.log('AUTO_SEND = ' + autoSend_());
  var all = folder_(DROP_FOLDER).getFiles();
  while (all.hasNext()) {
    var f = all.next();
    if (!/\.json$/i.test(f.getName())) continue;
    Logger.log('대기: ' + f.getName());
    var d = JSON.parse(f.getBlob().getDataAsString('UTF-8'));
    (d.reminders || []).forEach(function (r) {
      var rt = r.query ? GmailApp.search(r.query, 0, 1)[0] : null;
      Logger.log('  리마인드 ' + r.market + ' → ' + (rt ? rt.getFirstMessageSubject() + ' (' + rt.getLastMessageDate() + ')' : '스레드 없음 — 건너뜀') + '   [' + r.query + ']');
    });
  }
}

/** 리마인드 한 통: 그 법인 최신 스레드에 전체 회신(초안 또는 발송). 같은 key 는 한 번만 */
function remind_(r) {
  var props = PropertiesService.getScriptProperties(), done = 'done_' + (r.key || r.subject);
  if (props.getProperty(done)) { log_('건너뜀(이미 처리) ' + r.market + ' ' + r.key); return; }
  var t = r.query ? GmailApp.search(r.query, 0, 1)[0] : null;
  if (!t) { log_('스레드 없음 — 안 보냄 ' + r.market + ' [' + r.query + ']'); return; }
  if (autoSend_()) {
    t.replyAll(r.text);
    log_('발송 ' + r.market + ' → ' + t.getFirstMessageSubject());
  } else {
    t.createDraftReplyAll(r.text);
    log_('초안 ' + r.market + ' → ' + t.getFirstMessageSubject());
  }
  props.setProperty(done, new Date().toISOString());
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
    (d.reminders || []).forEach(remind_);
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
