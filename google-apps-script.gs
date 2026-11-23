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
  // Früher: waitLock(5000) – bei vielen gleichzeitigen Stimmen lief das in einen Fehler
  // und die Stimme ging verloren. appendRow braucht keine Sperre; wir warten nur kurz, wenn möglich.
  const lock = LockService.getScriptLock();
  const locked = lock.tryLock(20000);
  try { sheet_().appendRow([new Date(), voter, name, design, stars]); }
  finally { if (locked) lock.releaseLock(); }
  return json_({ ok: true });
}

function doGet(e) {
  const me = e && e.parameter && e.parameter.voter ? String(e.parameter.voter) : "";
  const rows = sheet_().getDataRange().getValues().slice(1);
  // Pro Person (Name, Groß-/Kleinschreibung egal) und Variante zählt nur die letzte Stimme –
  // auch wenn jemand vom Handy und vom Laptop abstimmt.
  const latest = {}, mine = {};
  rows.forEach(r => {
    const voter = String(r[1]), name = String(r[2]).trim().toLowerCase(), design = String(r[3]), stars = Number(r[4]);
    latest[(name || voter) + "|" + design] = { design: design, stars: stars };
    if (me && voter === me) mine[design] = stars;
  });
  const agg = {};
  Object.values(latest).forEach(v => {
    agg[v.design] = agg[v.design] || { sum: 0, n: 0 };
    agg[v.design].sum += v.stars; agg[v.design].n++;
  });
  const results = {};
  Object.keys(agg).forEach(k => results[k] = { avg: agg[k].sum / agg[k].n, n: agg[k].n });
  return json_({ results: results, mine: mine });
}
