/**
 * This function will be triggered on sheet edit and call other functions
 * @param {Object<GoogleAppsScript.Events} event
 */
function onEdit(event) {
  exerciseArrayRows(event)
  planColumns(event)
}
