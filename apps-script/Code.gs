/**
 * KGC 2027 survey backend (Google Apps Script, bound to the "KGC 2027 Survey Responses" sheet).
 *
 * The website reads totals and comments from this script and sends answers to it.
 * Names live only in this spreadsheet; the website only ever receives totals.
 *
 * Setup: Extensions > Apps Script, paste this file, Save, then
 * Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.
 */

const GOLFERS = { tyler: 'Tyler', granger: 'Granger', kyle: 'Kyle', kevin: 'Kevin', james: 'James', colin: 'Colin' };
const WEEKENDS = [
  ['mar18', 'Mar 18–21'], ['mar25', 'Mar 25–28'], ['apr01', 'Apr 1–4'], ['apr15', 'Apr 15–18'],
  ['apr22', 'Apr 22–25'], ['apr29', 'Apr 29–May 2'], ['may13', 'May 13–16']
];
const DESTS = [['stg', 'St. George + Vegas'], ['scot', 'Scottsdale'], ['peb', 'Pebble Beach']];
const AVAIL = { yes: 'Available', maybe: 'Could make it work', no: "Can't" };
const RATING = { love: 'Obsessed', cool: 'Would be down', no: 'Not interested' };

const RESPONSE_HEADERS = ['Golfer', 'Updated']
  .concat(WEEKENDS.map(w => w[1]))
  .concat(DESTS.map(d => d[1]))
  .concat(['#1 pick', 'Notes']);
const COMMENT_HEADERS = ['Posted', 'Golfer', 'Destination', 'Comment'];

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}
function responses_() { return sheet_('Responses', RESPONSE_HEADERS); }
function comments_() { return sheet_('Comments', COMMENT_HEADERS); }

/** Run once from the editor to create the tabs (optional; they are also created on first use). */
function setup() {
  responses_();
  comments_();
  const first = SpreadsheetApp.getActive().getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && SpreadsheetApp.getActive().getSheets().length > 1) {
    SpreadsheetApp.getActive().deleteSheet(first);
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
const invert_ = map => Object.fromEntries(Object.entries(map).map(([k, v]) => [v, k]));

/** GET: totals (no names) and comments. */
function doGet() {
  const availKey = invert_(AVAIL), rateKey = invert_(RATING);
  const destKey = Object.fromEntries(DESTS.map(d => [d[1], d[0]]));
  const rows = responses_().getDataRange().getValues().slice(1).filter(r => r[0]);
  const tallies = rows.map(r => {
    const weekends = {}, ratings = {};
    WEEKENDS.forEach((w, i) => { const v = availKey[r[2 + i]]; if (v) weekends[w[0]] = v; });
    DESTS.forEach((d, i) => { const v = rateKey[r[2 + WEEKENDS.length + i]]; if (v) ratings[d[0]] = v; });
    return { weekends, ratings, top: destKey[r[2 + WEEKENDS.length + DESTS.length]] || null };
  });
  const comments = comments_().getDataRange().getValues().slice(1).filter(r => r[3]).map(r => ({
    at: new Date(r[0]).getTime(),
    golfer: invert_(GOLFERS)[r[1]] || '',
    dest: destKey[r[2]] || '',
    text: String(r[3])
  }));
  return json_({ ok: true, count: tallies.length, tallies, comments });
}

/** POST (text/plain JSON): {action:"submit", ...answers} or {action:"comment", golfer, dest, text}. */
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }
  const golfer = GOLFERS[body.golfer];
  if (!golfer) return json_({ ok: false, error: 'unknown_golfer' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (body.action === 'submit') {
      const w = body.weekends || {}, r = body.ratings || {};
      const destName = (DESTS.find(d => d[0] === body.top) || [])[1] || '';
      const row = [golfer, new Date()]
        .concat(WEEKENDS.map(x => AVAIL[w[x[0]]] || ''))
        .concat(DESTS.map(x => RATING[r[x[0]]] || ''))
        .concat([destName, String(body.notes || '').slice(0, 1000)]);
      const sh = responses_();
      const names = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues().map(v => v[0]);
      const idx = names.indexOf(golfer);
      if (idx > 0) sh.getRange(idx + 1, 1, 1, row.length).setValues([row]);
      else sh.appendRow(row);
      return json_({ ok: true });
    }
    if (body.action === 'comment') {
      const dest = (DESTS.find(d => d[0] === body.dest) || [])[1];
      const text = String(body.text || '').trim().slice(0, 800);
      if (!dest || !text) return json_({ ok: false, error: 'bad_comment' });
      const sh = comments_();
      if (sh.getLastRow() > 1000) return json_({ ok: false, error: 'full' });
      sh.appendRow([new Date(), golfer, dest, text]);
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown_action' });
  } finally {
    lock.releaseLock();
  }
}
