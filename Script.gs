function doGet() {
  var template = HtmlService.createTemplateFromFile("Form");
  return template.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function submitData(siapterbang, ppjk, alamat, consignee, noSurat, tglSurat, alasan, detilAlasan, bc11, pos, subpos, subsubpos, tglbc11, mawb, hawb, tglmawb, tglhawb, jumlah, satuan, berat, uraian, kodehs, fungsi, shipper, nshipper, ntujuan, spbl, tglspbl, dasarHukum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1"); // change "Sheet1" to your sheet name
  
  // get last row number
  const lastRow = sheet.getLastRow();
  
  // insert data into sheet
  sheet.getRange(lastRow + 1, 1, 1, 29).setValues([
    [siapterbang, ppjk, alamat, consignee, noSurat, tglSurat, alasan, detilAlasan, bc11, pos, subpos, subsubpos, tglbc11, mawb, hawb, tglmawb, tglhawb, jumlah, satuan, berat, uraian, kodehs, fungsi, shipper, nshipper, ntujuan, spbl, tglspbl, dasarHukum]
  ]);
}
