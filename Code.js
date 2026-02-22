// ================= CONFIG =================
const TEMPLATE_ID = "1Djt2F79MKPIkzpfplyLV6-FRhjbOL3eDhbdHj3VtUFM/"; // Google Doc template ID
const FOLDER_ID = "1DRC1MdcayN7Rn1KlOOsYezeN5gn4tXM8"; // Folder for PDFs
const SHEET_FILE_ID = "15Mn11rlLY-OItZuhUOXZ-MhvfRABRIh6c1mfHqm4Z-c";
const SHEET_NAME = "MASTERLIST";

// ==========================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("TRANSMITTAL")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheetData() {
  const ss = SpreadsheetApp.openById(SHEET_FILE_ID);
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  data.shift(); // remove headers
  return data;
}

/**
 * Return a data URL for a Drive file (image) so the client can embed it reliably.
 * Returns null on error (e.g. permissions).
 */
function getDriveImageBase64(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var contentType = blob.getContentType() || "image/png";
    var base64 = Utilities.base64Encode(blob.getBytes());
    return "data:" + contentType + ";base64," + base64;
  } catch (e) {
    return null;
  }
}

/**
 * Accepts an array of refs: {type: 'drive', id: '...'} or {type: 'url', url: '...'}
 * Returns an array of data URLs (or null) in the same order.
 */
function getImagesBase64(refs) {
  var out = [];
  refs.forEach(function (ref) {
    try {
      var blob;
      if (ref.type === "drive") {
        var file = DriveApp.getFileById(ref.id);
        blob = file.getBlob();
      } else if (ref.type === "url") {
        var res = UrlFetchApp.fetch(ref.url, { muteHttpExceptions: true });
        if (res.getResponseCode() !== 200) {
          out.push(null);
          return;
        }
        blob = res.getBlob();
      } else {
        out.push(null);
        return;
      }

      var contentType = blob.getContentType() || "image/png";
      var base64 = Utilities.base64Encode(blob.getBytes());
      out.push("data:" + contentType + ";base64," + base64);
    } catch (e) {
      out.push(null);
    }
  });
  return out;
}
