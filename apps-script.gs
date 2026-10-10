// Who's the Problem? signups → Google Sheet
// Paste this into the sheet: Extensions > Apps Script, then Deploy > New deployment > Web app
// (Execute as: Me, Who has access: Anyone). Copy the web app URL into SHEET_POST in index.html.
const SHEET_ID = '1BRNkkg8z4R9hNsV15E5yg167KPSPeeXpF4lFm9p2oqg';

function doPost(e) {
  const d = JSON.parse(e.postData.contents || '{}');
  const clean = v => String(v == null ? '' : v).slice(0, 200).replace(/^[=+\-@]/, "'$&");
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(d.email || ''))) return ContentService.createTextOutput('skip');
  const answers = (Array.isArray(d.answers) ? d.answers : []).slice(0, 10).map(clean);
  while (answers.length < 10) answers.push('');
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  sheet.appendRow([new Date(), clean(d.name), clean(d.email), Number(d.score) || 0, clean(d.rank), clean(d.verdict)]
    .concat(answers, [clean(d.source)]));
  return ContentService.createTextOutput('ok');
}
