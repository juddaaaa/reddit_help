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

  if (!targetSheets.includes(targetSheetId)) return // Exit function if edited sheet's Id is not in array

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
    const nextCategory = targetSheet.getRange(2, column + 1).getValue() // Get the category from row 2 of the column next to the edited column
    const nextCategoryRow = targetSheet
      .getRange(2, 3, lastRow - 1, 1)
      .createTextFinder(nextCategory || "foobar")
      .findNext()
      ?.getRow() // Get the row in column C that matches the next category value (if found)

    const targetRow = nextCategoryRow ? nextCategoryRow - 1 : null // Set the target row to insert or delete cells

    if (value) {
      targetSheet.getRange(targetRow ? targetRow : lastRow, 4, 1, 6).insertCells(SpreadsheetApp.Dimension.ROWS) // Insert cells if edited cell is not blank
    } else {
      targetSheet.getRange(targetRow ? targetRow + 1 : lastRow, 4, 1, 6).deleteCells(SpreadsheetApp.Dimension.ROWS) // Delete cells if edited cell is blank
    }
  }
}
