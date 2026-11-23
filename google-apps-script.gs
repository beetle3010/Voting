/** @OnlyCurrentDoc */
// Speichert die Stimmen der Wohnzimmer-Abstimmung in diesem Google Sheet.
// Update einspielen: Code ersetzen -> Speichern -> Bereitstellen -> Bereitstellungen verwalten ->
// Stift-Symbol -> Version: "Neue Version" -> Bereitstellen (die URL bleibt gleich).

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Stimmen");
  if (!sh) {
    sh = ss.insertSheet("Stimmen");
    sh.appendRow(["Zeit", "Voter-ID", "Name", "Variante", "Sterne"]);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  if (!e || !e.postData) return json_({ ok: false, error: "no data" });
  const d = JSON.parse(e.postData.contents);
  const stars = Math.max(0, Math.min(5, Math.round(Number(d.stars))));
  const name = String(d.name || "").slice(0, 60);
  const design = String(d.design || "").slice(0, 40);
  const voter = String(d.voter || "").slice(0, 60);
  if (!design || !voter) return json_({ ok: false, error: "missing" });
  // Eine Zeile pro Person (Gerät) und Variante: vorhandene Zeile wird aktualisiert statt neu angehängt.
  const lock = LockService.getScriptLock();
  const locked = lock.tryLock(20000);
  try {
    const sh = sheet_();
    const data = sh.getDataRange().getValues();
    let row = -1;
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][1]) === voter && String(data[i][3]) === design) { row = i + 1; break; }
    }
    if (row > 0) sh.getRange(row, 1, 1, 5).setValues([[new Date(), voter, name, design, stars]]);
    else sh.appendRow([new Date(), voter, name, design, stars]);
  } finally { if (locked) lock.releaseLock(); }
  return json_({ ok: true });
}

function doGet(e) {
  const me = e && e.parameter && e.parameter.voter ? String(e.parameter.voter) : "";
  const rows = sheet_().getDataRange().getValues().slice(1);
  const latest = {}; // letzte Stimme je Person und Variante zählt
  rows.forEach(r => { latest[r[1] + "|" + r[3]] = { voter: String(r[1]), design: String(r[3]), stars: Number(r[4]) }; });
  const agg = {}, mine = {};
  Object.values(latest).forEach(v => {
    agg[v.design] = agg[v.design] || { sum: 0, n: 0 };
    agg[v.design].sum += v.stars; agg[v.design].n++;
    if (me && v.voter === me) mine[v.design] = v.stars;
  });
  const results = {};
  Object.keys(agg).forEach(k => results[k] = { avg: agg[k].sum / agg[k].n, n: agg[k].n });
  return json_({ results: results, mine: mine });
}
