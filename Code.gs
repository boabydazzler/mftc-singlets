/**
 * Mulligans Flat Track Club singlet numbers
 * Collects name and singlet size only.
 */
const SHEET_NAME = "Singlet Numbers";

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(["Timestamp", "Name", "Size"]);
      sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    const data = JSON.parse(e.postData.contents || "{}");

    if (!data.name || !data.size) {
      throw new Error("Missing required information.");
    }

    const validSizes = ["XXS", "XS", "S", "M", "L", "XL", "XXL"];
    if (!validSizes.includes(data.size)) {
      throw new Error("Invalid size.");
    }

    sheet.appendRow([
      new Date(),
      String(data.name).trim(),
      data.size
    ]);

    return jsonResponse({ok:true});
  } catch (err) {
    return jsonResponse({ok:false, error:String(err.message || err)});
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return jsonResponse({ok:true, service:"MFTC Singlet Numbers"});
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
