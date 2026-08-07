/**
 * IBI Coconut Sales Tracker — Google Drive backend (v2)
 * ⚠ LOCAL MIRROR — this file only takes effect when pasted into Apps Script and deployed.
 *
 * ONE-TIME SETUP:
 *  1. In Google Drive, create a Google Sheet named e.g. "Coconut Sales Tracker".
 *  2. Extensions → Apps Script → delete any code there → paste this whole file → Save.
 *  3. Deploy → New deployment → type: Web app →
 *       Execute as: Me   |   Who has access: Anyone
 *     → Deploy, authorize, and copy the URL ending in /exec.
 *  4. In the app: Report tab → Google Drive Backend → paste the URL → Save URL.
 *  After that, every entry auto-syncs to this Sheet (Sales / Batches / Labour /
 *  Rates tabs + a RawBackup used by "Restore from Drive").
 */

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    if (req.action !== 'sync') return json_({ ok: false, error: 'bad action' });
    var d = req.data || {};
    var ss = SpreadsheetApp.getActive();

    var salesRows = [];
    (d.sales || []).slice().reverse().forEach(function (s) {
      (s.items || []).forEach(function (i, ix) {
        salesRows.push([
          s.date || '', s.time || '', s.buyer || '', i.label || '',
          i.unit === 'kg' ? i.qty : '',
          (i.wts && i.wts.length) ? i.wts.join(' + ') : '',
          i.count || '', i.rate || 0, i.labRate || '',
          i.amount || 0, ix === 0 ? s.total : ''
        ]);
      });
    });
    writeSheet_(ss, 'Sales',
      ['Date', 'Time', 'Buyer', 'Item', 'Weight (kg)', 'Weighments', 'Count (pcs)', 'Rate', 'Labour/pc', 'Amount', 'Bill Total'],
      salesRows);

    writeSheet_(ss, 'Batches', ['Date', 'Coconut Count', 'Note'],
      (d.batches || []).slice().reverse().map(function (b) { return [b.date || '', b.count || 0, b.note || '']; }));

    writeSheet_(ss, 'Labour', ['Date', 'Persons', 'Coconuts De-husked', 'Rate/Coconut', 'Amount'],
      (d.labour || []).slice().reverse().map(function (l) { return [l.date || '', l.persons || '', l.coconuts || 0, l.rate || 0, l.amount || 0]; }));

    writeSheet_(ss, 'Rates', ['Date', 'Good/kg', 'Lower/kg', 'Husk-Kadambai/pc', 'Labour/coconut'],
      (d.rateHistory || []).slice().reverse().map(function (r) { return [r.date || '', r.good || 0, r.low || 0, r.kad || 0, r.lab || 0]; }));

    var raw = ensureSheet_(ss, 'RawBackup');
    raw.clear();
    raw.getRange(1, 1).setNumberFormat('@').setValue(JSON.stringify(d));
    raw.getRange(1, 2).setNumberFormat('@').setValue('synced ' + new Date());

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'load') {
    var raw = SpreadsheetApp.getActive().getSheetByName('RawBackup');
    var txt = raw ? String(raw.getRange(1, 1).getValue()) : '';
    var data = null;
    try { data = txt ? JSON.parse(txt) : null; } catch (err) {}
    return json_({ ok: true, data: data });
  }
  return json_({ ok: true, ping: 'coconut-sales-tracker' });
}

// Dates arrive as yyyy-mm-dd strings; the whole range is forced to plain text so
// Sheets can never locale-swap day/month (see IBI sheets-date rule).
function writeSheet_(ss, name, header, rows) {
  var sh = ensureSheet_(ss, name);
  sh.clear();
  var all = [header].concat(rows);
  sh.getRange(1, 1, all.length, header.length).setNumberFormat('@').setValues(all);
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
}

function ensureSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
