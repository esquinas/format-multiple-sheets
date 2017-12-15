function isSameSheet (sheetA, sheetB) {
  return (sheetA.getSheetId() === sheetB.getSheetId())
}

function isValidA1Notation (a1ref) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getActiveSheet()
  var result = false
  var range

  try {
    range = sheet.getRange(a1ref)
    result = true
  } catch (error) {
    ss.toast(error.message, error.name)
  }
  return result
}

/*
 *       Polyfills
 */
if (!Array.isArray) {
  Array.isArray = function (arg) {
    return Object.prototype.toString.call(arg) === '[object Array]'
  }
}

