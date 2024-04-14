/**
 * @author u/juddaaaa <https://www.reddit.com/user/juddaaaaa/>
 * @file   https://www.reddit.com/r/GoogleAppsScript/comments/1c34i1g/help_consistent_error_code_for_my_workflow/
 * @description Answer to Reddit Question. A function to update monthly sheets with event data from Google Calendar
 * @license MIT
 * @version 1.0
 */

/**
 * This function update monthly sheets with event data from 60 days in the past to 90 days in the future
 */
function updateEventsFromCalendar() {
  const calendarId = "johnhaliburton@atlantaschoolofmassage.edu"
  const calendar = CalendarApp.getCalendarById(calendarId)
  const spreadsheet = SpreadsheetApp.getActive()

  // Get events from 60 days in the past to 90 days in the future
  const today = new Date() // Today's date
  const startDate = new Date(today.getTime() - 60 * 24 * 60 * 60 * 1000) // 60 days ago
  const endDate = new Date(today.getTime() + 90 * 24 * 60 * 60 * 1000) // 90 days from now
  const events = calendar.getEvents(startDate, endDate)

  // Reduce events down to an object of months and weeks in months
  const eventsByMonthWeek = events.reduce((events, event) => {
    const month = Utilities.formatDate(event.getStartTime(), Session.getScriptTimeZone(), "MMM") // Format for month
    const week = Utilities.formatDate(event.getStartTime(), Session.getScriptTimeZone(), "W") // Format for week in month
    const eventName = event.getTitle()

    if (!events[month]) events[month] = {} // Object for the month
    if (!events[month][week]) events[month][week] = [] // Array for the week within the month
    if (eventName.includes("Tour")) events[month][week].push(getEventData(event)) // Push qualifying events into this weeks array

    return events
  }, {})

  // Loop through eventsByMonthWeek and set up monthly sheets
  for (const [month, weeks] of Object.entries(eventsByMonthWeek)) {
    const sheetName = month.toUpperCase()
    let sheet = spreadsheet.getSheetByName(sheetName)

    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName)
      sheet.getRange(1, 1, 1, 6).setValues([["Date", "Start Time", "Event Name", "Status", "Program", "APP"]])
    } else {
      const lastRow = sheet.getLastRow() // Last row of sheet
      const lastColumn = sheet.getLastColumn() // Last column of sheet
      if (lastRow >= 2) sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent() // If sheet contains data (excluding headers), clear contents
    }

    // Loop through weeks and write to sheet, leaving 1 row separation between each week
    if (weeks.keys.length) {
      for (const week of Object.values(weeks)) {
        const lastRow = sheet.getLastRow()
        const seperation = lastRow === 1 ? 1 : 2
        sheet.getRange(lastRow + seperation, 1, week.length, week[0].length).setValues(week)
      }
    }
  }
}

/**
 * This function formats a calendar event in preperation for insertion into sheet
 * @param {Object<GoogleAppsScript.Calendar.CalendarEvent>} event
 * @returns {Array}
 */
function getEventData(event) {
  const startTime = event.getStartTime()
  const date = Utilities.formatDate(startTime, Session.getScriptTimeZone(), "MM/dd/yyyy")
  const startTimeFormatted = Utilities.formatDate(startTime, Session.getScriptTimeZone(), "HH:mm")
  const eventName = event.getTitle()
  const description = event.getDescription()

  if (description) {
    const descriptionLines = description.split("\n")
    var [status, program, app] = descriptionLines
  } else {
    var [status, program, app] = ["", "", ""]
  }

  return [date, startTimeFormatted, eventName, status, program, app]
}
