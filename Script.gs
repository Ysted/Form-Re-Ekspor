function doGet() {
  var template = HtmlService.createTemplateFromFile("Form");
  return template.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function submitData(siapterbang, ppjk, alamat, consignee, noSurat, tglSurat, category, detilAlasan, bc11, pos, subpos, subsubpos, tglbc11, mawb, hawb, tglmawb, tglhawb, jumlah, satuan, berat, uraian, kodehs, shipper, nshipper, ntujuan, spbl, tglspbl, ketentuan) {
  var username = Session.getActiveUser();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("INPUT"); // change "Sheet1" to your sheet name
  // get last row number
  const lastRow = sheet.getLastRow();

  // Remove dot separators from Kode HS
  kodehs = kodehs ? kodehs.replace(/\./g, '') : '00000000';
  
  // insert data into sheet
  sheet.getRange(lastRow + 1, 1, 1, 29).setValues([
    [username, siapterbang, ppjk, alamat, consignee, noSurat, tglSurat, category, detilAlasan, bc11, pos, subpos, subsubpos, tglbc11, mawb, hawb, tglmawb, tglhawb, jumlah, satuan, berat, uraian, kodehs, shipper, nshipper, ntujuan, spbl, tglspbl, ketentuan]
  ]);

}

// Function to check if siapterbang exists in the "INPUT" sheet
function checkSiapterbangInSpreadsheet(siapterbang) {
  const spreadsheetId = '1HJ_yQkr-vws403N1rWKnBl0cU8dRd8p9T2fDTtlCECY'; // Replace with your spreadsheet ID
  const sheetName = 'INPUT'; // Replace with the name of your sheet
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName(sheetName);
  const data = sheet.getRange("B:B").getValues(); // Assuming Siapterbang values are in column B

  // Find the row with the matching Siapterbang value
  const rowIndex = data.findIndex(row => row[0] === siapterbang);

  return { found: rowIndex !== -1, rowIndex: rowIndex + 1 }; // Adding 1 to convert from 0-based index to 1-based row number
}

// Function to update the spreadsheet with Nomor Surat and Tanggal Surat
function updateSpreadsheet(siapterbang, nomorSurat, tanggalSurat, barcode) {
    const spreadsheetId = '1HJ_yQkr-vws403N1rWKnBl0cU8dRd8p9T2fDTtlCECY';
    const sheetInput = SpreadsheetApp.openById(spreadsheetId).getSheetByName('INPUT');
    const data = sheetInput.getRange("B:B").getValues();
    const matchingRows = [];

    // Find all matching rows with the same siapterbang
    data.forEach((row, index) => {
        if (row[0] === siapterbang) {
            matchingRows.push(index + 1); // Adding 1 to convert from 0-based index to 1-based row number
        }
    });

    // Check if multiple rows with the same siapterbang are found
    if (matchingRows.length > 1) {
        // Iterate through each matching row and copy data to 'Database' sheet
        const sheetDatabase = SpreadsheetApp.openById(spreadsheetId).getSheetByName('Database');
        matchingRows.forEach(rowIndex => {
            // Update 'Nomor Surat' and 'Tanggal Surat' columns in 'INPUT' sheet
            sheetInput.getRange(rowIndex, 30).setValue(nomorSurat);
            sheetInput.getRange(rowIndex, 31).setValue(tanggalSurat);

            // Set hyperlink in column 30 with the URL from extractedinfo.barcode
            sheetInput.getRange(rowIndex, 30).setFormula(`=HYPERLINK("${barcode}"; "${nomorSurat}")`);
            
            // Copy value from 'INPUT' sheet to 'Database' sheet
            const hawb = sheetInput.getRange(rowIndex, 16).getValue();
            const mawb = sheetInput.getRange(rowIndex, 15).getValue();
            const tglHawb = sheetInput.getRange(rowIndex, 18).getValue();
            const consignee = sheetInput.getRange(rowIndex, 5).getValue();
            const ppjk = sheetInput.getRange(rowIndex, 3).getValue();
            const kategori = sheetInput.getRange(rowIndex, 8).getValue();
            const jumlah = sheetInput.getRange(rowIndex, 19).getValue();
            const satuan = sheetInput.getRange(rowIndex, 20).getValue();
            const berat = sheetInput.getRange(rowIndex, 21).getValue();
            const uraian = sheetInput.getRange(rowIndex, 22).getValue();
            const nTujuan = sheetInput.getRange(rowIndex, 26).getValue();
            const analis = sheetInput.getRange(rowIndex, 1).getValue();

            // Get the last empty row in 'Database' sheet
            const lastRow = sheetDatabase.getLastRow() + 1;

            // Set hyperlink in column 30 with the URL from extractedinfo.barcode
            sheetInput.getRange(rowIndex, 30).setFormula(`=HYPERLINK("${barcode}"; "${nomorSurat}")`);

            // Set values in 'Database' sheet
            sheetDatabase.getRange(lastRow, 2).setValue(hawb);
            sheetDatabase.getRange(lastRow, 3).setValue(mawb);
            sheetDatabase.getRange(lastRow, 4).setValue(tglHawb);
            sheetDatabase.getRange(lastRow, 5).setValue(consignee);
            sheetDatabase.getRange(lastRow, 6).setValue(ppjk);
            sheetDatabase.getRange(lastRow, 7).setValue(kategori);
            sheetDatabase.getRange(lastRow, 8).setFormula(`=HYPERLINK("${barcode}"; "${nomorSurat.split('/')[0]}")`);
            sheetDatabase.getRange(lastRow, 9).setValue(tanggalSurat);
            sheetDatabase.getRange(lastRow, 10).setValue(jumlah + " " + satuan);
            sheetDatabase.getRange(lastRow, 11).setValue(String(berat).replace(".", ",") + " kg");
            sheetDatabase.getRange(lastRow, 12).setValue(uraian);
            sheetDatabase.getRange(lastRow, 14).setValue(nTujuan);
            sheetDatabase.getRange(lastRow, 15).setValue(siapterbang);
            sheetDatabase.getRange(lastRow, 16).setValue(analis);
        });
      } else if (matchingRows.length === 1) {
        // Only one row with the siapterbang found, update as before
        const rowIndex = matchingRows[0];
        // Update 'Nomor Surat' and 'Tanggal Surat' columns in 'INPUT' sheet
          sheetInput.getRange(rowIndex, 30).setValue(nomorSurat);
          sheetInput.getRange(rowIndex, 31).setValue(tanggalSurat);
        // Set hyperlink in column 30 with the URL from extractedinfo.barcode
          sheetInput.getRange(rowIndex, 30).setFormula(`=HYPERLINK("${barcode}"; "${nomorSurat}")`);
          
        // Check the last empty row in the 'Database' sheet
        const sheetDatabase = SpreadsheetApp.openById(spreadsheetId).getSheetByName('Database');
        const lastRow = sheetDatabase.getLastRow();

        // Copy value from 'INPUT' sheet to 'Database' sheet
        const hawb = sheetInput.getRange(rowIndex, 16).getValue();
        sheetDatabase.getRange(lastRow + 1, 2).setValue(hawb);
        const mawb = sheetInput.getRange(rowIndex, 15).getValue();
        sheetDatabase.getRange(lastRow + 1, 3).setValue(mawb);
        const tglHawb = sheetInput.getRange(rowIndex, 18).getValue();
        sheetDatabase.getRange(lastRow + 1, 4).setValue(tglHawb);
        const consignee = sheetInput.getRange(rowIndex, 5).getValue();
        sheetDatabase.getRange(lastRow + 1, 5).setValue(consignee);
        const ppjk = sheetInput.getRange(rowIndex, 3).getValue();
        sheetDatabase.getRange(lastRow + 1, 6).setValue(ppjk);
        const kategori = sheetInput.getRange(rowIndex, 8).getValue();
        sheetDatabase.getRange(lastRow + 1, 7).setValue(kategori);
        sheetDatabase.getRange(lastRow + 1, 8).setFormula(`=HYPERLINK("${barcode}"; "${nomorSurat.split('/')[0]}")`);
        sheetDatabase.getRange(lastRow + 1, 9).setValue(tanggalSurat);
        const jumlah = sheetInput.getRange(rowIndex, 19).getValue();
        const satuan = sheetInput.getRange(rowIndex, 20).getValue();
        sheetDatabase.getRange(lastRow + 1, 10).setValue(jumlah + " " + satuan);
        const berat = sheetInput.getRange(rowIndex, 21).getValue();
        sheetDatabase.getRange(lastRow + 1, 11).setValue(String(berat).replace(".", ",") + " kg");
        const uraian = sheetInput.getRange(rowIndex, 22).getValue();
        sheetDatabase.getRange(lastRow + 1, 12).setValue(uraian);
        const nTujuan = sheetInput.getRange(rowIndex, 26).getValue();
        sheetDatabase.getRange(lastRow + 1, 14).setValue(nTujuan);
        sheetDatabase.getRange(lastRow + 1, 15).setValue(siapterbang);
        const analis = sheetInput.getRange(rowIndex, 1).getValue();
        sheetDatabase.getRange(lastRow + 1, 16).setValue(analis);
    }
}

function mailMerge(siapterbang) {
  var spreadsheetId = '1HJ_yQkr-vws403N1rWKnBl0cU8dRd8p9T2fDTtlCECY'; // Replace with your spreadsheet ID
  var sheetName = 'INPUT'; // Replace with the name of your sheet
  var templateNotaId = '1Z0_woJZ8K4gNq55J1M6KzXqcFvifrysNT34rrIGohVE'; // Replace with your template ID
  var templateNotaLartasId = '1IH_ETOkKQQVbWf8CH4RSwMIqJwkHrWgqsignMXQbjZ0'; // Replace with your template ID
  var templateSuratId = '15gdtA6-AePMaFL1XQha5BfXUUk32hRwgWnjvHbaza6M';
  var folderId = '1SLLRyCcxBqClDA0pExlidBrIV_bCtKGl'; // Replace with your folder ID
  
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);

  // Get the row data based on siapterbang
  var rowDataResult = checkSiapterbangInSpreadsheet(siapterbang);

  // Check if siapterbang exists in the spreadsheet
  if (!rowDataResult.found) {
    throw new Error('Nomor Siap Terbang tidak ditemukan. Coba lagi Ngab!');
  }

  // Get the row data
  var rowData = sheet.getRange(rowDataResult.rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var templateNota = DriveApp.getFileById(templateNotaId);
  var templateNotaLartas = DriveApp.getFileById(templateNotaLartasId);
  var templateSurat = DriveApp.getFileById(templateSuratId);
  var folder = DriveApp.getFolderById(folderId);
  var newFileNameNota = "[KONSEP] Pendapat Atas Permohonan Ekspor Kembali " + rowData[2] + " Untuk " + rowData[4] + " (" + rowData[1] + ")";
  var newFileNameNotaLartas = "[KONSEP] Pendapat Atas Permohonan Ekspor Kembali " + rowData[2] + " Untuk " + rowData[4] + " (" + rowData[1] + ")";
  var newFileNameSurat = "[KONSEP] Persetujuan Permohonan Ekspor Kembali " + rowData[2] + " Untuk " + rowData[4] + " (" + rowData[1] + ")";
  var copyNota = templateNota.makeCopy(newFileNameNota, folder);
  var copyNotaLartas = templateNotaLartas.makeCopy(newFileNameNotaLartas, folder);
  var copySurat = templateSurat.makeCopy(newFileNameSurat, folder);
  var docNota = DocumentApp.openById(copyNota.getId());
  var docNotaLartas = DocumentApp.openById(copyNotaLartas.getId());
  var docSurat = DocumentApp.openById(copySurat.getId());
  var bodyNota = docNota.getBody();
  var bodyNotaLartas = docNotaLartas.getBody();
  var bodySurat = docSurat.getBody();
  
  var placeholders = {
    '{{PPJK}}': rowData[2],
    '{{Alamat}}': rowData[3],
    '{{Consignee}}': rowData[4],
    '{{Siapterbang}}': rowData[1],
    '{{No Surat Permohonan}}': rowData[5],
    '{{Tgl Surat}}': rowData[6] ? rowData[6].toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' }) : '',
    '{{BC 1.1}}': rowData[9] ? rowData[9].toString().padStart(6, '0') : '000000',
    '{{Tgl BC 1.1}}': rowData[13] ? rowData[13].toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' }) : '',
    '{{Pos}}': rowData[10] ? rowData[10].toString().padStart(4, '0') : '0000',
    '{{Sub Pos}}': rowData[11] ? rowData[11].toString().padStart(4, '0') : '0000',
    '{{Sub-sub Pos}}': rowData[12] ? rowData[12].toString().padStart(4, '0') : '0000',
    '{{MAWB}}': rowData[14],
    '{{Tgl MAWB}}': rowData[16] ? rowData[16].toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' }) : '',
    '{{HAWB}}': rowData[15],
    '{{Tgl HAWB}}': rowData[17] ? rowData[17].toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' }) : '',
    '{{Jumlah}}': rowData[18],
    '{{Satuan}}': rowData[19],
    '{{Berat}}': rowData[20] ? rowData[20].toString().replace('.', ',') : '',
    '{{Uraian}}': rowData[21],
    '{{Shipper}}': rowData[23],
    '{{Negara Shipper}}': rowData[24],
    '{{Negara Tujuan}}': rowData[25],
    '{{Alasan Surat}}': rowData[31],
    '{{Detil Alasan}}': rowData[8],
    '{{Uraian}}': rowData[21],
    '{{Kode HS}}': rowData[22] ? rowData[22].toString().padStart(8, '0').replace(/(\d{4})(\d{2})(\d{2})/, "$1.$2.$3") : '0000.00.00',
    '{{Nomor SPBL}}': rowData[26] ? rowData[26].toString().padStart(6, '0') : '000000',
    '{{Tgl SPBL}}': rowData[27] ? rowData[27].toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' }) : '',
    '{{Peraturan Terkait}}': rowData[28]
  };
  
  for (var placeholder in placeholders) {
    var value = placeholders[placeholder];
    bodyNota.replaceText(placeholder, value);
    bodyNotaLartas.replaceText(placeholder, value);
    bodySurat.replaceText(placeholder, value);
  }
  
  docNota.saveAndClose();
  docNotaLartas.saveAndClose();
  docSurat.saveAndClose();
  
  // Get the ID
  var docNotaId = docNota.getId();
  var docNotaLartasId = docNotaLartas.getId();
  var docSuratId = docSurat.getId();

  // Create download links
  var downloadLinkNota = 'https://docs.google.com/document/d/' + docNotaId + '/export?format=docx';
  var downloadLinkNotaLartas = 'https://docs.google.com/document/d/' + docNotaLartasId + '/export?format=docx';
  var downloadLinkSurat = 'https://docs.google.com/document/d/' + docSuratId + '/export?format=docx';

  // Return the download links
  return {
    downloadLinkNota: downloadLinkNota,
    downloadLinkNotaLartas: downloadLinkNotaLartas,
    downloadLinkSurat: downloadLinkSurat
  };
}
