/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

'use strict'

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen (e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi()
}

/**
 * @returns {array} list of the names of the sheets in the Spreadsheet
 */
function getSheetNames () {
  var sheetNames = []
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()

  for (var i = 0; i < sheets.length; i++) {
     sheetNames.push(sheets[i].getName());
  }
 
  return sheetNames;
}

/**
 * propagateFormat takes a template sheet and copies its format
 * to every other non-hidden sheet. Including width of cols, height of rows,
 * frozen cols/rows and even tab color. 
 * TODO: Copy also charts, protections and images?
 * TODO2: Checkboxes to select if copy chart, protections, tab colors, ...
 * TODO3: Input text to set a placeholder text/number that we want ignored, otherwise content gets copied.
 * TODO4: Option to make a non-destructive propagation.
 * @param   {string}       Name of the template sheet.
 * @returns {string|Error} Message or error.
 */
function propagateFormat (template) {
  var ss = SpreadsheetApp.getActiveSpreadsheet() 
  var source = ss.getSheetByName(template)

  if (source === null) { 
    throw new ReferenceError('Sheet with name: "' + template +  '" no longer exists!', 'format_multiple_sheets.gs', 54)
    return
  }

  var destinations = ss.getSheets()
  var destination = destinations[0]
  var data = prepareData(source)
  
  var tasks = [unhide, freeze, copyTabColor, copyOnlyFormats, copyHeights, copyWidths]
  
  for (var i = 0; i < destinations.length; i++) {
    destination = destinations[i]

    if (destination.isSheetHidden() || isSameSheet(destination, source)) continue
    
    for (var j = 0; j < tasks.length; j++) tasks[j](destination, data)
  }
  
  return 'Propagation succeeded!'
}

function prepareData (source, placeholders) {
  var data = { 'range': source.getDataRange() }
  data.placeholders = placeholders || [/text/i, /1234/i, /11\/11\/11/]

  data.values   = data.range.getValues()
  data.rangeInA1Notation = data.range.getA1Notation()
  data.numCols  = data.range.getNumColumns()
  data.numRows  = data.range.getNumRows()

  data.firstCol = data.range.getColumn()
  data.firstRow = data.range.getRow()

  data.frozenCols = source.getFrozenColumns()
  data.frozenRows = source.getFrozenRows()

  data.tabColor   = source.getTabColor()
  data.source = source
  return data
}

function freeze (target, data) {
    target.setFrozenColumns(data.frozenCols)
    target.setFrozenRows(   data.frozenRows)
}

function unhide (target, data) {
    target.showColumns(data.firstCol, data.numCols)
    target.showRows(   data.firstRow, data.numRows)
}

function copyTabColor (target, data) {
  target.setTabColor(data.tabColor)
}

function copyOnlyFormats (target, data) {
  target.clearFormats()
  data.range.copyFormatToRange(target, 
                               data.firstCol, 
                               data.numCols, 
                               data.firstRow, 
                               data.numRows)
}

function copyWidths (target, data) {
  var colPosition = 0
  var desiredWidth = 0
  var currentWidth = 0
    
  for (var i = 0; i <= data.numCols; i++) {
      colPosition = data.firstCol + i
      desiredWidth  = data.source.getColumnWidth(colPosition)
      currentWidth = target.getColumnWidth(colPosition)
      if (currentWidth !== desiredWidth) target.setColumnWidth(colPosition, desiredWidth)
    }
}

function copyHeights (target, data) {
  var rowPosition = 0
  var desiredHeight = 0
  var currentHeight = 0
    
  for (var i = 0; i <= data.numRows; i++) {
    rowPosition = data.firstRow + i
    desiredHeight = data.source.getRowHeight(rowPosition)
    currentHeight = target.getRowHeight(rowPosition)
    if (currentHeight !== desiredHeight) target.setRowHeight(rowPosition, desiredHeight)
  }
}

function isSameSheet (sheetA, sheetB) {
  return (sheetA.getSheetId() === sheetB.getSheetId())
}


/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall (e) {
  onOpen(e)
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function showSidebar () {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Format Multiple Sheets')
  SpreadsheetApp.getUi().showSidebar(ui)
}
