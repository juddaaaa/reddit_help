/**
 * @author u/juddaaaa <https://reddit.com/user/juddaaaa>
 * @file https://www.reddit.com/r/googlesheets/comments/1c45785/apps_script_to_autoinsertdelete_cells_based_on/
 * @description Answer to Reddit Question. A function that is triggered on sheet edit to add or delete cells in a range
 * @license MIT
 * @version 1.0
 */

/**
 * This function will be triggered on sheet edit and will add or delete cells from a specific range
 * @param {Object<GoogleAppsScript.Events>} event
 * @returns void
 */
function exerciseArrayRows({range, value}) {
  const targetSheets = [1765965136] // Function will ony run if sheets with these Ids are edited
  const targetSheetId = range.getSheet().getSheetId() // Get the edited sheet's Id

  if (!targetSheets.includes(targetSheetId) || range.rowStart !== range.rowEnd) return // Exit function if edited sheet's Id is not in array or the range size is greater than 1

  const targetSheet = range.getSheet() // Get the edited sheet
  const row = range.getRow() // Get the edited row
  const column = range.getColumn() // Get the edited column
  const targetColumns = [14, 15, 16, 18, 19, 20, 22, 23, 24, 26, 27, 28, 29, 30, 31, 32, 33] // Target columns

  if (column === 11 && row >= 3) {
    // If column K was edited
    if (value) {
      targetSheet.getRange(row, 4, 1, 6).insertCells(SpreadsheetApp.Dimension.ROWS) // Insert cells if edited cell is not blank
    } else {
      // If one of the target columns was edited
      targetSheet.getRange(row, 4, 1, 6).deleteCells(SpreadsheetApp.Dimension.ROWS) // Delete cells if edited cell is blank
    }
  } else if (targetColumns.includes(column) && row >= 3) {
    // If any of the target columns was edited
    const lastRow = targetSheet.getLastRow() // Get last row of sheet
    const editedCategory = targetSheet.getRange(2, column).getValue() // Get the category where the edit was made
    const nextCategory = targetSheet.getRange(2, column + 1).getValue() // Get the adjacent category
    const rowAbove = targetSheet.getRange(row - 1, column).getValue() // Get the value of the row above the edited cell
    const lookupStart = targetSheet
      .getRange(2, 3, lastRow - 1, 1)
      .createTextFinder(editedCategory)
      .findNext()
      ?.getRow() // Get the start row of lookup in column C

    const lookupEnd = targetSheet
      .getRange(2, 3, lastRow - 1, 1)
      .createTextFinder(nextCategory)
      .findNext()
      ?.getRow() // Get the end row of lookup in column C

    if (lookupStart && lookupEnd) {
      const numRows = lookupEnd - lookupStart // Number of rows in lookup
      const targetRow = targetSheet.getRange(lookupStart, 3, numRows, 1).createTextFinder(rowAbove).findNext()?.getRow() // Get row of lookup value (if found)
      if (targetRow) {
        if (value) {
          targetSheet.getRange(targetRow + 1, 4, 1, 6).insertCells(SpreadsheetApp.Dimension.ROWS) // Insert cells if edited cell is not blank
        } else {
          targetSheet.getRange(targetRow + 1, 4, 1, 6).deleteCells(SpreadsheetApp.Dimension.ROWS) // Delete cells if edited cell is blank
        }
      }
    }
  }
}
