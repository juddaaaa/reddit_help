/**
 * @author u/juddaaaa <https://reddit.com/user/juddaaaa>
 * @file https://www.reddit.com/r/googlesheets/comments/1c45785/apps_script_to_autoinsertdelete_cells_based_on/
 * @description Answer to Reddit Question. A function to get a sheet by it's unique Id
 * @license MIT
 * @version 1.0
 */

/**
 * This function will retur a sheet based on the containg spreadsheet and the unique id of the sheet
 * @param {Object<GoogleAppsScript.Spreadsheet>} spreadsheet
 * @param {Number} id
 * @returns {Object<GoogleAppsScript.Spreadsheet.Sheet>}
 */
function getSheetById(spreadsheet, id) {
  const sheets = spreadsheet.getSheets() // Get all the sheets in the spreadsheet

  for (let sheet of sheets) {
    // Loop through the sheets. If the sheet with the passed in Id is found return the sheet
    if (sheet.getSheetId() === id) return sheet
  }
}
