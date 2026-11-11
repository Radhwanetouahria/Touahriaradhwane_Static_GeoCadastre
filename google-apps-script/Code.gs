// GeoCadastre Pro - Google Apps Script
// Spreadsheet ID supplied by the owner: 18mqtkghTzuU55izsXWS09v4X1au-Pya8
// 1) Create a new Apps Script project at https://script.google.com
// 2) Paste this file. Set SECRET to a private phrase of your choice.
// 3) Deploy > New deployment > Web app; Execute as: Me; Who has access: Anyone.
// 4) Copy the /exec URL into GeoCadastre > Paramètres.
const SPREADSHEET_ID = '18mqtkghTzuU55izsXWS09v4X1au-Pya8';
const SECRET = 'CHANGE-ME-TO-A-LONG-PRIVATE-SECRET';
const DATA_SHEET = 'GeoCadastre_Data';
function doGet(e) {
  try {
    checkToken_(e.parameter.token);
    const sheet = getSheet_();
    const raw = sheet.getRange('A1').getValue();
    const data = raw ? JSON.parse(raw) : {dossiers: [], interventions: []};
    return json_({ok:true, data:data});
  } catch (err) { return json_({ok:false, error:String(err.message || err)}); }
}
function doPost(e) {
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    checkToken_(body.token);
    if (body.action !== 'save') throw new Error('Action inconnue');
    getSheet_().getRange('A1').setValue(JSON.stringify(body.data || {}));
    return json_({ok:true, savedAt:new Date().toISOString()});
  } catch (err) { return json_({ok:false, error:String(err.message || err)}); }
}
function checkToken_(token) {
  if (!SECRET || SECRET.indexOf('CHANGE-ME') === 0) throw new Error('Configurez SECRET dans Apps Script');
  if (token !== SECRET) throw new Error('Code de protection incorrect');
}
function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(DATA_SHEET);
  if (!sh) sh = ss.insertSheet(DATA_SHEET);
  return sh;
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
