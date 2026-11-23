/** @OnlyCurrentDoc */
// Speichert die Stimmen der Wohnzimmer-Abstimmung in diesem Google Sheet.
// Einrichtung: Google Sheet anlegen -> Erweiterungen -> Apps Script -> diesen Code einfügen ->
// Bereitstellen -> Neue Bereitstellung -> Typ "Web-App" -> Ausführen als: Ich, Zugriff: Jeder -> URL kopieren.

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Stimmen");
  if (!sh) {
    sh = ss.insertSheet("Stimmen");
    sh.appendRow(["Zeit", "Voter-ID", "Name", "Variante", "Sterne"]);
  }
  return sh;
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const stars = Math.max(0, Math.min(5, Math.round(Number(d.stars))));
  const name = String(d.name || "").slice(0, 60);
  const design = String(d.design || "").slice(0, 40);
  const voter = String(d.voter || "").slice(0, 60);
  if (!design || !voter) return ContentService.createTextOutput("missing");
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try { sheet_().appendRow([new Date(), voter, name, design, stars]); }
  finally { lock.releaseLock(); }
  return ContentService.createTextOutput("ok");
}

function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const latest = {}; // letzte Stimme je Person und Variante zählt
  rows.forEach(r => { latest[r[1] + "|" + r[3]] = { design: r[3], stars: Number(r[4]) }; });
  const agg = {};
  Object.values(latest).forEach(v => {
    agg[v.design] = agg[v.design] || { sum: 0, n: 0 };
    agg[v.design].sum += v.stars; agg[v.design].n++;
  });
  const out = {};
  Object.keys(agg).forEach(k => out[k] = { avg: agg[k].sum / agg[k].n, n: agg[k].n });
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
