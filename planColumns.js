/**
 * @author u/juddaaaa <https://reddit.com/user/juddaaaa>
 * @file https://www.reddit.com/r/googlesheets/comments/1c45785/apps_script_to_autoinsertdelete_cells_based_on/
 * @description Answer to Reddit Question. A function to hide specific columns based on what is entered in two cells
 * @license MIT
 * @version 1.0
 */

/**
 * This function will hide specific columns depending on what is entered in two cells
 * @param {Object<GoogleAppsScript.Events>} event
 * @returns void
 */
function planColumns({range, source}) {
  const targetSheets = [786484840, 402743240, 1704182648, 1412478509, 1096188587] // Function will ony run if sheets with these Ids are edited
  const targetSheetId = range.getSheet().getSheetId() // Get the edited sheet's Id

  if (!targetSheets.includes(targetSheetId) && !["C4", "C6"].includes(range.getA1Notation())) return // Exit function if edited sheet's Id is not in array and edited cell address is not in array

  const spreadsheet = source // Get the spreadsheet
  const targetSheet = range.getSheet() // Get the edited sheet
  const sheetTimelines = getSheetById(spreadsheet, 1053704024) // Get the Timelines sheet
  const sport = targetSheet.getRange(4, 3).getValue() // Get the value of the sport cell
  const season = targetSheet.getRange(6, 3).getValue() // Get the value of the season cell
  const sportRow = sheetTimelines
    .getRange(5, 1, sheetTimelines.getLastRow() - 4, 1)
    .createTextFinder(sport)
    .findNext()
    ?.getRow() // Get the row in the Timeline sheet that matched the value of the sport cell (if found)

  if (!sportRow) return // Exit function if the sport is not found in the Timelines sheet
  const seasonCount = sheetTimelines
    .getRange(sportRow, 2, 1, sheetTimelines.getLastColumn() - 1)
    .createTextFinder(season)
    .findAll()?.length // How many weeks in the season?

  // Hide columns based on how many weeks in season
  if (seasonCount) {
    targetSheet.showColumns(4, 18)
    targetSheet.hideColumns(4 + 3 * seasonCount, 18 - 3 * seasonCount)
  }
}
