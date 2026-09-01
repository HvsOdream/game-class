/*************************************************************
 * 캡스톤디자인Ⅲ 게임반 — 워크시트 제출 수신 서버 (Google Apps Script)
 *
 * 학생이 워크시트에서 [구글드라이브로 제출]을 누르면
 *   1) 내 드라이브 > 캡스톤3_게임반_제출 > 1주차  폴더에
 *      팀 이름으로 된 구글 문서를 만들고(재제출 시 최신본으로 교체)
 *   2) 같은 폴더의 「제출현황」 스프레드시트에 한 줄로 기록합니다.
 *
 * 배포 방법은 「구글드라이브_제출_설치안내.md」 참고.
 *************************************************************/

var ROOT_FOLDER = '캡스톤3_게임반_제출';   // 내 드라이브에 자동 생성됩니다
var SHEET_NAME  = '제출현황';
var TZ          = 'Asia/Seoul';

/* ---------- 진입점 ---------- */

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.check) {
    // 워크시트가 "내 제출이 도착했는지" 확인하는 요청 (JSONP)
    return reply(p.callback, lookup(p.team, p.week));
  }
  return HtmlService.createHtmlOutput(
    '<meta charset="utf-8"><p style="font-family:sans-serif">캡스톤Ⅲ 게임반 제출 서버가 동작 중입니다.</p>'
  );
}

function reply(callback, obj) {
  var s = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + s + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}

function lookup(team, week) {
  try {
    var sh = openSheet(rootFolder()).getSheets()[0];
    var last = sh.getLastRow();
    if (last < 2) return { ok: true, found: false };
    var vals = sh.getRange(2, 1, last - 1, 9).getValues();
    var wk = String(week || 1) + '주차';
    for (var i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][1]) === wk && String(vals[i][2]) === String(team)) {
        return { ok: true, found: true, at: String(vals[i][0]), url: String(vals[i][8]) };
      }
    }
    return { ok: true, found: false };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return json({ ok: false, error: '서버가 바쁩니다. 잠시 후 다시 시도하세요.' });
  }
  try {
    // 폼 전송(payload 파라미터)과 raw JSON 본문을 모두 받는다
    var body = (e && e.parameter && e.parameter.payload) ||
               (e && e.postData && e.postData.contents) || '';
    if (!body) return json({ ok: false, error: '전달된 내용이 없습니다.' });
    var p = JSON.parse(body);
    var team = String(p.team || '이름없는팀').trim();
    var week = p.week || 1;

    var folder = subFolder(rootFolder(), week + '주차');
    var docName = safeName(team) + '_' + week + '주차';

    // 같은 이름의 이전 제출본은 휴지통으로 (최신본 1개만 유지)
    var old = folder.getFilesByName(docName);
    while (old.hasNext()) { old.next().setTrashed(true); }

    var doc = DocumentApp.create(docName);
    writeDoc(doc, p);
    doc.saveAndClose();

    DriveApp.getFileById(doc.getId()).moveTo(folder);   // 주차 폴더로 이동

    logRow(rootFolder(), p, doc.getUrl());

    return json({ ok: true, url: doc.getUrl(), name: docName });
  } catch (err) {
    return json({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 문서 작성 ---------- */

function writeDoc(doc, p) {
  var b = doc.getBody();
  b.clear();

  var title = b.appendParagraph((p.course || '캡스톤디자인Ⅲ 게임반') + ' · ' + (p.week || 1) + '주차 팀 워크시트');
  title.setHeading(DocumentApp.ParagraphHeading.TITLE);

  var sub = b.appendParagraph(
    '제출 ' + fmt(p.submittedAt) + ' · 작성률 ' + (p.pct != null ? p.pct + '%' : '—')
  );
  sub.setForegroundColor('#666666').setFontSize(9);

  var blocks = p.blocks || [];
  for (var i = 0; i < blocks.length; i++) {
    var blk = blocks[i] || {};
    if (blk.type === 'meta') {
      table(b, null, ['항목', '내용'], (blk.items || []).map(function (kv) {
        return [String(kv[0] || ''), String(kv[1] || '—')];
      }));
    } else if (blk.type === 'h') {
      b.appendParagraph(String(blk.text || '')).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    } else if (blk.type === 'p') {
      b.appendParagraph(String(blk.text || ''));
    } else if (blk.type === 'kv') {
      table(b, blk.title, ['항목', '내용'], (blk.items || []).map(function (kv) {
        return [String(kv[0] || ''), String(kv[1] || '') || '—'];
      }));
    } else if (blk.type === 'list') {
      if (blk.title) b.appendParagraph(String(blk.title)).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      (blk.items || []).forEach(function (t) {
        b.appendListItem(String(t)).setGlyphType(DocumentApp.GlyphType.BULLET);
      });
    } else if (blk.type === 'table') {
      table(b, blk.title, blk.head || [], (blk.rows || []).map(function (r) {
        return (r || []).map(function (c) { return String(c == null || c === '' ? '—' : c); });
      }));
    }
  }

  b.appendParagraph('');
  var foot = b.appendParagraph('제출자 ' + (p.submitter || '—') + ' · 자동 생성 문서');
  foot.setForegroundColor('#888888').setFontSize(8);
}

function table(body, title, head, rows) {
  if (title) body.appendParagraph(String(title)).setHeading(DocumentApp.ParagraphHeading.HEADING3);
  if (!rows || !rows.length) { body.appendParagraph('— 미입력'); return; }
  var cells = [];
  if (head && head.length) cells.push(head.map(String));
  rows.forEach(function (r) { cells.push(r.map(String)); });
  var t = body.appendTable(cells);
  if (head && head.length) {
    var hr = t.getRow(0);
    for (var c = 0; c < hr.getNumCells(); c++) {
      hr.getCell(c).setBackgroundColor('#EEF0F8').editAsText().setBold(true);
    }
  }
  t.setBorderColor('#CCCCCC');
}

/* ---------- 제출현황 시트 ---------- */

function logRow(root, p, url) {
  var ss = openSheet(root);
  var sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  var head = ['제출시각', '주차', '팀명', '게임 타이틀', '제출자', '작성률', '판정', '트랙', '문서 링크'];
  if (sh.getLastRow() === 0) {
    sh.appendRow(head);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#EEF0F8');
    sh.setFrozenRows(1);
  }
  var row = [
    fmt(p.submittedAt), (p.week || 1) + '주차', p.team || '', p.project || '',
    p.submitter || '', (p.pct != null ? p.pct / 100 : ''), p.verdict || '', p.track || '', url
  ];
  // 같은 주차·같은 팀이면 해당 줄을 갱신, 없으면 새 줄
  var last = sh.getLastRow();
  var target = 0;
  if (last > 1) {
    var vals = sh.getRange(2, 2, last - 1, 2).getValues();   // 주차, 팀명
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][0]) === row[1] && String(vals[i][1]) === row[2]) { target = i + 2; break; }
    }
  }
  if (!target) target = last + 1;
  sh.getRange(target, 1, 1, row.length).setValues([row]);
  sh.getRange(target, 6).setNumberFormat('0%');
}

function openSheet(root) {
  var it = root.getFilesByName(SHEET_NAME);
  if (it.hasNext()) return SpreadsheetApp.open(it.next());
  var ss = SpreadsheetApp.create(SHEET_NAME);
  DriveApp.getFileById(ss.getId()).moveTo(root);
  return ss;
}

/* ---------- 유틸 ---------- */

function rootFolder() {
  var it = DriveApp.getFoldersByName(ROOT_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(ROOT_FOLDER);
}

function subFolder(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function safeName(s) {
  return String(s).replace(/[\\\/:*?"<>|]/g, '_').slice(0, 60);
}

function fmt(iso) {
  var d = iso ? new Date(iso) : new Date();
  if (isNaN(d.getTime())) d = new Date();
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd HH:mm');
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 배포 전 점검용 (에디터에서 직접 실행) ---------- */

function 테스트제출() {
  var demo = {
    course: '캡스톤디자인Ⅲ 게임반', week: 1, team: '테스트팀', project: '테스트 게임',
    submitter: '김민철', pct: 42, verdict: '존속', track: '계승',
    submittedAt: new Date().toISOString(),
    blocks: [
      { type: 'meta', items: [['팀명', '테스트팀'], ['게임 타이틀', '테스트 게임']] },
      { type: 'h', text: 'A. 빌드 현황 카드' },
      { type: 'kv', items: [['장르 · 플랫폼', '2D 로그라이크 · PC']] },
      { type: 'list', title: '되는 것', items: ['한 판이 끝난다'] },
      { type: 'table', title: '미구현', head: ['#', '내용'], rows: [['1', '보스전']] }
    ]
  };
  var res = doPost({ postData: { contents: JSON.stringify(demo) } });
  Logger.log(res.getContent());
}
