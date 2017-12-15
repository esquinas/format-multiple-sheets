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
 * @returns {array.<string>} list of names of the sheets in active Spreadsheet.
 */
function getSheetNames () {
  var sheetNames = []
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()

  for (var i = 0; i < sheets.length; i++) {
     sheetNames.push(sheets[i].getName())
  }
 
  return sheetNames;
}

/**
 * propagateFormat takes a template sheet and copies its format
 * to every other non-hidden sheet. Including col widths, row heights,
 * frozen cols/rows and even tab color. 
 *
 * TODO1: Copy also charts, protections and images, how?
 * TODO2: Checkboxes to select whether to copy chart, protections, tab colors, etc.
 * TODO3: Option to make the script non-destructive 
 *          -> YAGNY if user duplicates spreadsheet before running script.
 * Todo3B: Change "tip" to a "tips carrousel"?
 *
 * FIX: At input for contentOnly range, HTML5 RegExp for input validation FAILs (b/c no submit button?).
 * @param   {string}       Name of the template sheet.
 * @returns {string|Error} Message or error.
 */
function propagateFormat (template, options) {
  var opts = options || {}
  var ss = SpreadsheetApp.getActiveSpreadsheet() 
  var source = ss.getSheetByName(template)

  if (source === null)
    throw new ReferenceError('Sheet with name: "' + template +  '" no longer exists!', 'format_multiple_sheets.gs')

  var destinations = ss.getSheets()
  var destination = destinations[0]
  var data = prepareData(source, opts)
  
  var tasks = [{ 'execute': unhideCells,        'skip': false },
               { 'execute': freezeCells,        'skip': false },
               { 'execute': copyTabColor,       'skip': false },
               { 'execute': cleanUp,            'skip': true  },
               { 'execute': copyFormatsOnly,    'skip': false },
               { 'execute': copyHeights,        'skip': false },
               { 'execute': copyWidths,         'skip': false },
               { 'execute': copyContentsOnly,   'skip': false }]
  
  for (var i = 0; i < destinations.length; i++) {
    destination = destinations[i]

    if (destination.isSheetHidden() || isSameSheet(destination, data.source)) continue
    
    for (var j = 0; j < tasks.length; j++) {
      if (!tasks[j].skip) tasks[j].execute(destination, data)
    }
  }
  
  return 'Done!'
}

function prepareData (source, opts) {
  var data = { 'range': source.getDataRange() }

  if (opts.contentsOnlyRange && isValidA1Notation(opts.contentsOnlyRange))
    data.contentsOnlyRange = source.getRange(opts.contentsOnlyRange)

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

function freezeCells (target, data) {
    target.setFrozenColumns(data.frozenCols)
    target.setFrozenRows(   data.frozenRows)
}

function unhideCells (target, data) {
    target.showColumns(data.firstCol, data.numCols)
    target.showRows(   data.firstRow, data.numRows)
}

function copyTabColor (target, data) {
  target.setTabColor(data.tabColor)
}

function cleanUp (target, _) {
  target.clearFormats()
}

function copyFormatsOnly (target, data) {
  // TODO: Refactor using range.copyTo(destRange, { formatOnly: true })
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

function copyContentsOnly (target, data) {
  var destination
  if (!data.contentsOnlyRange) return
  
  destination = target.getRange(data.contentsOnlyRange.getA1Notation())
  data.contentsOnlyRange.copyTo(destination, { contentsOnly: true })
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
