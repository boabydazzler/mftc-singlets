const SHEET_NAME = "Singlet Numbers";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || "{}");

    const name = String(data.name || "").trim();
    const cut = String(data.cut || "").trim();
    const size = String(data.size || "").trim();

    const allowedCuts = ["Mens", "Womens"];
    const allowedSizes = ["XXS", "XS", "S", "M", "L", "XL", "XXL"];

    if (!name || !allowedCuts.includes(cut) || !allowedSizes.includes(size)) {
      return jsonResponse({ok:false, error:"Invalid submission"});
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Ensure the four current headers are correct.
    sheet.getRange(1, 1, 1, 4).setValues([["Timestamp", "Name", "Cut", "Size"]]);

    sheet.appendRow([new Date(), name, cut, size]);

    return jsonResponse({ok:true});
  } catch (err) {
    return jsonResponse({ok:false, error:String(err)});
  }
}

function doGet() {
  return jsonResponse({ok:true, service:"MFTC singlet form"});
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
