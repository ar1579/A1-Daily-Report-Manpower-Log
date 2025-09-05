/**
 * A1 Daily Report & Manpower Log - Google Apps Script Backend
 * 
 * This script provides the backend functionality for the Daily Report & Manpower Log system,
 * including form submission handling, data processing, and email notifications.
 * 
 * @version 1.0
 * @date 2025-09-04
 */

// Global variables
const WEATHER_API_KEY = "YOUR_WEATHER_API_KEY"; // Replace with actual API key
const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID"; // Replace with actual spreadsheet ID

/**
 * Creates a custom menu when the spreadsheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Daily Report Tools')
    .addItem('Generate Tomorrow\'s Report', 'generateNextDayReport')
    .addItem('Update Weather Data', 'updateWeatherData')
    .addItem('Email Daily Report', 'emailDailyReport')
    .addItem('Export as PDF', 'exportAsPDF')
    .addSeparator()
    .addSubMenu(ui.createMenu('Analytics')
      .addItem('Generate Manpower Histogram', 'generateManpowerHistogram')
      .addItem('Analyze Productivity Trends', 'analyzeProductivityTrends')
      .addItem('Track Delay Impact', 'trackDelayImpact'))
    .addToUi();
}

/**
 * Handles GET requests to the web app
 * Serves the HTML form to users
 * 
 * @param {Object} e - Event object from web app
 * @returns {HtmlOutput} - HTML content to display
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('improved_A1_Daily_Report_Manpower_Log_Form')
    .setTitle('Daily Report & Manpower Log')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Handles POST requests to the web app (form submissions)
 * 
 * @param {Object} e - Event object containing form data
 * @returns {Object} - JSON response with success/error information
 */
function doPost(e) {
  try {
    // Log the received data for debugging
    Logger.log("Received form submission: " + JSON.stringify(e.parameter));
    
    // Validate required fields
    if (!validateRequiredFields(e.parameter)) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: "Missing required fields"
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Process and save the form data
    const result = processFormData(e.parameter);
    
    // Send confirmation email
    sendConfirmationEmail(e.parameter);
    
    // Return success response
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      reportId: result.reportId,
      message: "Report submitted successfully"
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    // Log the error
    Logger.log("Error processing form submission: " + error.toString());
    
    // Return error response
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Validates that all required fields are present in the form data
 * 
 * @param {Object} formData - The form data to validate
 * @returns {boolean} - True if all required fields are present, false otherwise
 */
function validateRequiredFields(formData) {
  const requiredFields = [
    'project-name',
    'project-number',
    'project-location',
    'report-date',
    'superintendent'
  ];
  
  for (const field of requiredFields) {
    if (!formData[field] || formData[field].trim() === '') {
      Logger.log(`Missing required field: ${field}`);
      return false;
    }
  }
  
  return true;
}

/**
 * Processes form data and saves it to the spreadsheet
 * 
 * @param {Object} formData - The form data to process
 * @returns {Object} - Object containing processing results
 */
function processFormData(formData) {
  // Get the active spreadsheet
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID) || SpreadsheetApp.getActiveSpreadsheet();
  
  // Get or create the daily report sheet
  const reportDate = new Date(formData['report-date']);
  const sheetName = Utilities.formatDate(reportDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    // Create a new sheet based on the template
    const templateSheet = ss.getSheetByName("Template");
    if (templateSheet) {
      sheet = templateSheet.copyTo(ss);
      sheet.setName(sheetName);
    } else {
      // If no template exists, create a new blank sheet
      sheet = ss.insertSheet(sheetName);
      setupNewSheet(sheet); // Set up headers and formatting
    }
  }
  
  // Map form data to spreadsheet
  mapFormDataToSheet(sheet, formData);
  
  // Update the historical data tab
  updateHistoricalData(ss, formData);
  
  // Return the result
  return {
    reportId: sheetName,
    sheetUrl: ss.getUrl() + "#gid=" + sheet.getSheetId()
  };
}

/**
 * Sets up a new sheet with headers and formatting
 * 
 * @param {Sheet} sheet - The sheet to set up
 */
function setupNewSheet(sheet) {
  // Set up header
  sheet.getRange("A1:F1").merge().setValue("DAILY REPORT & MANPOWER LOG")
    .setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  
  // Set up project info section
  sheet.getRange("A3").setValue("Project Name:");
  sheet.getRange("A4").setValue("Project Number:");
  sheet.getRange("A5").setValue("Project Location:");
  sheet.getRange("A6").setValue("Weather Conditions:");
  sheet.getRange("A7").setValue("Superintendent:");
  
  sheet.getRange("E4").setValue("Report #:");
  sheet.getRange("E7").setValue("Date:");
  
  // Set up section headers
  const sections = [
    { row: 9, title: "MANPOWER" },
    { row: 33, title: "EQUIPMENT" },
    { row: 47, title: "MATERIALS RECEIVED" },
    { row: 61, title: "WORK COMPLETED" },
    { row: 72, title: "ISSUES/DELAYS" },
    { row: 82, title: "SAFETY OBSERVATIONS & INCIDENTS" },
    { row: 92, title: "QUALITY CONTROL INSPECTIONS" },
    { row: 102, title: "VISITOR LOG" },
    { row: 112, title: "DAILY PHOTOS" },
    { row: 122, title: "APPROVAL" }
  ];
  
  sections.forEach(section => {
    sheet.getRange(`A${section.row}:F${section.row}`).merge()
      .setValue(section.title)
      .setBackground("#f0f0f0")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
  });
  
  // Set column widths
  sheet.setColumnWidth(1, 100); // Column A
  sheet.setColumnWidth(2, 200); // Column B
  sheet.setColumnWidth(3, 150); // Column C
  sheet.setColumnWidth(4, 150); // Column D
  sheet.setColumnWidth(5, 150); // Column E
  sheet.setColumnWidth(6, 150); // Column F
}

/**
 * Maps form data to the appropriate cells in the spreadsheet
 * 
 * @param {Sheet} sheet - The sheet to update
 * @param {Object} formData - The form data to map
 */
function mapFormDataToSheet(sheet, formData) {
  // Project information
  sheet.getRange("C3").setValue(formData['project-name']);
  sheet.getRange("C4").setValue(formData['project-number']);
  sheet.getRange("F4").setValue(formData['report-number']);
  sheet.getRange("C5").setValue(formData['project-location']);
  sheet.getRange("C6").setValue(formData['weather-conditions']);
  sheet.getRange("E6").setValue(formData['temperature']);
  sheet.getRange("C7").setValue(formData['superintendent']);
  sheet.getRange("F7").setValue(new Date(formData['report-date']));
  
  // Manpower section
  for (let i = 1; i <= 8; i++) {
    const rowIndex = 10 + i;
    sheet.getRange(`B${rowIndex}`).setValue(formData[`contractor-${i}`] || "");
    sheet.getRange(`C${rowIndex}`).setValue(formData[`trade-${i}`] || "");
    sheet.getRange(`D${rowIndex}`).setValue(formData[`workers-${i}`] ? Number(formData[`workers-${i}`]) : "");
    sheet.getRange(`E${rowIndex}`).setValue(formData[`hours-${i}`] ? Number(formData[`hours-${i}`]) : "");
    sheet.getRange(`F${rowIndex}`).setValue(formData[`areas-${i}`] || "");
    sheet.getRange(`G${rowIndex}`).setValue(formData[`productivity-${i}`] || "");
  }
  
  // Manpower totals
  sheet.getRange("D31").setValue(formData['total-workers'] ? Number(formData['total-workers']) : 0);
  sheet.getRange("E31").setValue(formData['total-hours'] ? Number(formData['total-hours']) : 0);
  
  // Equipment section
  for (let i = 1; i <= 5; i++) {
    const rowIndex = 34 + i;
    sheet.getRange(`B${rowIndex}`).setValue(formData[`equipment-${i}`] || "");
    sheet.getRange(`C${rowIndex}`).setValue(formData[`equipment-qty-${i}`] ? Number(formData[`equipment-qty-${i}`]) : "");
    sheet.getRange(`D${rowIndex}`).setValue(formData[`equipment-hours-${i}`] ? Number(formData[`equipment-hours-${i}`]) : "");
    sheet.getRange(`E${rowIndex}`).setValue(formData[`equipment-activity-${i}`] || "");
  }
  
  // Equipment totals
  sheet.getRange("C45").setValue(formData['total-equipment'] ? Number(formData['total-equipment']) : 0);
  sheet.getRange("D45").setValue(formData['total-equipment-hours'] ? Number(formData['total-equipment-hours']) : 0);
  
  // Materials section
  for (let i = 1; i <= 5; i++) {
    const rowIndex = 48 + i;
    sheet.getRange(`B${rowIndex}`).setValue(formData[`material-${i}`] || "");
    sheet.getRange(`C${rowIndex}`).setValue(formData[`material-qty-${i}`] ? Number(formData[`material-qty-${i}`]) : "");
    sheet.getRange(`D${rowIndex}`).setValue(formData[`material-unit-${i}`] || "");
    sheet.getRange(`E${rowIndex}`).setValue(formData[`material-supplier-${i}`] || "");
    sheet.getRange(`F${rowIndex}`).setValue(formData[`material-location-${i}`] || "");
  }
  
  // Work completed section
  for (let i = 1; i <= 5; i++) {
    const rowIndex = 62 + i;
    sheet.getRange(`B${rowIndex}`).setValue(formData[`work-${i}`] || "");
    sheet.getRange(`E${rowIndex}`).setValue(formData[`work-area-${i}`] || "");
    sheet.getRange(`F${rowIndex}`).setValue(formData[`work-complete-${i}`] ? Number(formData[`work-complete-${i}`]) : "");
  }
  
  // Overall progress
  sheet.getRange("F70").setValue(formData['overall-progress'] ? Number(formData['overall-progress']) : "");
  
  // Issues/delays section
  for (let i = 1; i <= 3; i++) {
    const rowIndex = 73 + i;
    sheet.getRange(`B${rowIndex}`).setValue(formData[`issue-${i}`] || "");
    sheet.getRange(`D${rowIndex}`).setValue(formData[`issue-impact-${i}`] || "");
    sheet.getRange(`E${rowIndex}`).setValue(formData[`issue-resolution-${i}`] || "");
  }
  
  // Safety observations section
  for (let i = 1; i <= 3; i++) {
    const rowIndex = 83 + i;
    sheet.getRange(`B${rowIndex}`).setValue(formData[`safety-${i}`] || "");
    sheet.getRange(`E${rowIndex}`).setValue(formData[`safety-action-${i}`] || "");
    sheet.getRange(`F${rowIndex}`).setValue(formData[`safety-reportable-${i}`] || "");
  }
  
  // Quality control section
  for (let i = 1; i <= 3; i++) {
    const rowIndex = 93 + i;
    sheet.getRange(`B${rowIndex}`).setValue(formData[`inspection-${i}`] || "");
    sheet.getRange(`D${rowIndex}`).setValue(formData[`inspection-area-${i}`] || "");
    sheet.getRange(`E${rowIndex}`).setValue(formData[`inspector-${i}`] || "");
    sheet.getRange(`F${rowIndex}`).setValue(formData[`inspection-result-${i}`] || "");
  }
  
  // Visitor log section
  for (let i = 1; i <= 3; i++) {
    const rowIndex = 103 + i;
    sheet.getRange(`B${rowIndex}`).setValue(formData[`visitor-${i}`] || "");
    sheet.getRange(`C${rowIndex}`).setValue(formData[`visitor-company-${i}`] || "");
    sheet.getRange(`D${rowIndex}`).setValue(formData[`visitor-purpose-${i}`] || "");
    sheet.getRange(`E${rowIndex}`).setValue(formData[`visitor-in-${i}`] || "");
    sheet.getRange(`F${rowIndex}`).setValue(formData[`visitor-out-${i}`] || "");
  }
  
  // Daily photos
  sheet.getRange("A114:F114").merge().setValue(formData['photo-references'] || "");
  
  // Approval section
  sheet.getRange("C124").setValue(formData['prepared-by'] || "");
  sheet.getRange("E124").setValue(formData['prepared-date'] ? new Date(formData['prepared-date']) : "");
  sheet.getRange("C126").setValue(formData['superintendent-signature'] || formData['superintendent'] || "");
  sheet.getRange("E126").setValue(formData['superintendent-date'] ? new Date(formData['superintendent-date']) : "");
  sheet.getRange("C128").setValue(formData['pm-signature'] || "");
  sheet.getRange("E128").setValue(formData['pm-date'] ? new Date(formData['pm-date']) : "");
  
  // Apply conditional formatting
  applyConditionalFormatting(sheet);
}

/**
 * Updates the historical data tab with information from the current report
 * 
 * @param {Spreadsheet} ss - The spreadsheet
 * @param {Object} formData - The form data
 */
function updateHistoricalData(ss, formData) {
  // Get or create the historical data sheet
  let historySheet = ss.getSheetByName("Historical Data");
  if (!historySheet) {
    historySheet = ss.insertSheet("Historical Data");
    
    // Set up headers
    const headers = [
      "Date", "Project Name", "Weather", "Temperature", "Total Workers", 
      "Total Hours", "Productivity Average", "Issues Count", "Safety Incidents", 
      "Quality Failures", "Visitors Count"
    ];
    
    historySheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold").setBackground("#f0f0f0");
  }
  
  // Calculate derived values
  const issuesCount = countIssues(formData);
  const safetyIncidents = countSafetyIncidents(formData);
  const qualityFailures = countQualityFailures(formData);
  const visitorsCount = countVisitors(formData);
  const productivityAvg = calculateProductivityAverage(formData);
  
  // Prepare row data
  const rowData = [
    new Date(formData['report-date']),
    formData['project-name'],
    formData['weather-conditions'],
    formData['temperature'],
    Number(formData['total-workers'] || 0),
    Number(formData['total-hours'] || 0),
    productivityAvg,
    issuesCount,
    safetyIncidents,
    qualityFailures,
    visitorsCount
  ];
  
  // Append to the sheet
  historySheet.appendRow(rowData);
}

/**
 * Counts the number of issues/delays reported
 * 
 * @param {Object} formData - The form data
 * @returns {number} - The count of issues
 */
function countIssues(formData) {
  let count = 0;
  for (let i = 1; i <= 3; i++) {
    if (formData[`issue-${i}`] && formData[`issue-${i}`].trim() !== '') {
      count++;
    }
  }
  return count;
}

/**
 * Counts the number of reportable safety incidents
 * 
 * @param {Object} formData - The form data
 * @returns {number} - The count of reportable safety incidents
 */
function countSafetyIncidents(formData) {
  let count = 0;
  for (let i = 1; i <= 3; i++) {
    if (formData[`safety-reportable-${i}`] === 'Y') {
      count++;
    }
  }
  return count;
}

/**
 * Counts the number of quality control failures
 * 
 * @param {Object} formData - The form data
 * @returns {number} - The count of quality failures
 */
function countQualityFailures(formData) {
  let count = 0;
  for (let i = 1; i <= 3; i++) {
    if (formData[`inspection-result-${i}`] === 'Fail') {
      count++;
    }
  }
  return count;
}

/**
 * Counts the number of visitors
 * 
 * @param {Object} formData - The form data
 * @returns {number} - The count of visitors
 */
function countVisitors(formData) {
  let count = 0;
  for (let i = 1; i <= 3; i++) {
    if (formData[`visitor-${i}`] && formData[`visitor-${i}`].trim() !== '') {
      count++;
    }
  }
  return count;
}

/**
 * Calculates the average productivity rating
 * 
 * @param {Object} formData - The form data
 * @returns {number} - The average productivity rating
 */
function calculateProductivityAverage(formData) {
  let total = 0;
  let count = 0;
  
  for (let i = 1; i <= 8; i++) {
    if (formData[`productivity-${i}`] && formData[`productivity-${i}`] !== '') {
      total += Number(formData[`productivity-${i}`]);
      count++;
    }
  }
  
  return count > 0 ? Math.round((total / count) * 10) / 10 : 0; // Round to 1 decimal place
}

/**
 * Applies conditional formatting to the sheet
 * 
 * @param {Sheet} sheet - The sheet to format
 */
function applyConditionalFormatting(sheet) {
  // Productivity rating formatting
  const productivityRange = sheet.getRange("G11:G18");
  
  // Clear existing rules
  const rules = sheet.getConditionalFormatRules();
  sheet.setConditionalFormatRules([]);
  
  // Low productivity (1-2) - Red
  let rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThanOrEqualTo(2)
    .setBackground("#ffcccc")
    .setRanges([productivityRange])
    .build();
  rules.push(rule);
  
  // High productivity (4-5) - Green
  rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(4)
    .setBackground("#ccffcc")
    .setRanges([productivityRange])
    .build();
  rules.push(rule);
  
  // Major impact issues - Red
  const impactRange = sheet.getRange("D74:D76");
  rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Major")
    .setBackground("#ffcccc")
    .setRanges([impactRange])
    .build();
  rules.push(rule);
  
  // Reportable safety incidents - Red
  const reportableRange = sheet.getRange("F84:F86");
  rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Y")
    .setBackground("#ffcccc")
    .setRanges([reportableRange])
    .build();
  rules.push(rule);
  
  // Quality failures - Red
  const qualityRange = sheet.getRange("F94:F96");
  rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Fail")
    .setBackground("#ffcccc")
    .setRanges([qualityRange])
    .build();
  rules.push(rule);
  
  // Overall progress formatting
  const progressRange = sheet.getRange("F70");
  
  // Low progress (<50%) - Yellow
  rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(50)
    .setBackground("#fff2cc")
    .setRanges([progressRange])
    .build();
  rules.push(rule);
  
  // High progress (>=80%) - Green
  rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(80)
    .setBackground("#ccffcc")
    .setRanges([progressRange])
    .build();
  rules.push(rule);
  
  // Apply all rules
  sheet.setConditionalFormatRules(rules);
}

/**
 * Sends a confirmation email to the submitter
 * 
 * @param {Object} formData - The form data
 */
function sendConfirmationEmail(formData) {
  // Check if email is provided
  const emailAddress = formData['submitter-email'];
  if (!emailAddress) {
    Logger.log("No email address provided for confirmation");
    return;
  }
  
  try {
    // Get project and report information
    const projectName = formData['project-name'];
    const reportDate = new Date(formData['report-date']);
    const formattedDate = Utilities.formatDate(reportDate, Session.getScriptTimeZone(), "MMMM dd, yyyy");
    
    // Create email body
    const emailBody = `
      <html>
        <body>
          <h2>Daily Report Submission Confirmation</h2>
          <p>Your daily report for <strong>${projectName}</strong> on <strong>${formattedDate}</strong> has been successfully submitted.</p>
          
          <h3>Report Summary:</h3>
          <ul>
            <li><strong>Project:</strong> ${projectName}</li>
            <li><strong>Date:</strong> ${formattedDate}</li>
            <li><strong>Superintendent:</strong> ${formData['superintendent']}</li>
            <li><strong>Total Workers:</strong> ${formData['total-workers'] || '0'}</li>
            <li><strong>Total Hours:</strong> ${formData['total-hours'] || '0'}</li>
            <li><strong>Overall Progress:</strong> ${formData['overall-progress'] || '0'}%</li>
          </ul>
          
          <p>You can view the full report in the project spreadsheet.</p>
          
          <p>Thank you for your submission.</p>
          
          <p><em>This is an automated message. Please do not reply to this email.</em></p>
        </body>
      </html>
    `;
    
    // Send the email
    MailApp.sendEmail({
      to: emailAddress,
      subject: `Daily Report Confirmation - ${projectName} - ${formattedDate}`,
      htmlBody: emailBody
    });
    
    Logger.log("Confirmation email sent to: " + emailAddress);
    
  } catch (error) {
    Logger.log("Error sending confirmation email: " + error.toString());
  }
}

/**
 * Updates weather data for the current report based on project location
 */
function updateWeatherData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const projectLocation = sheet.getRange("C5").getValue(); // Project location
  
  if (!projectLocation) {
    SpreadsheetApp.getUi().alert("Please enter a project location first.");
    return;
  }
  
  try {
    // Call weather API (this is a placeholder - actual implementation would use specific weather API)
    const weatherData = fetchWeatherData(projectLocation);
    
    // Update weather fields
    sheet.getRange("C6").setValue(weatherData.conditions);
    sheet.getRange("E6").setValue(weatherData.tempLow + "°F - " + weatherData.tempHigh + "°F");
    
    SpreadsheetApp.getUi().alert("Weather data updated successfully.");
  } catch (error) {
    SpreadsheetApp.getUi().alert("Error updating weather data: " + error.toString());
  }
}

/**
 * Fetches weather data from API (placeholder function)
 * 
 * @param {string} location - The location to get weather for
 * @returns {Object} - Weather data object
 */
function fetchWeatherData(location) {
  // This is a placeholder. In a real implementation, you would:
  // 1. Make an HTTP request to a weather API using UrlFetchApp
  // 2. Parse the JSON response
  // 3. Return the relevant weather data
  
  try {
    // Example API call (replace with actual API)
    const url = `https://api.weatherapi.com/v1/current.json?key=${WEATHER_API_KEY}&q=${encodeURIComponent(location)}`;
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    
    // Extract relevant weather information
    return {
      conditions: data.current.condition.text,
      tempLow: Math.round(data.current.temp_f - 5), // Approximate low temp
      tempHigh: Math.round(data.current.temp_f + 5), // Approximate high temp
      precipitation: data.current.precip_in,
      windSpeed: data.current.wind_mph
    };
  } catch (error) {
    // For demonstration, returning mock data
    Logger.log("Using mock weather data due to error: " + error.toString());
    return {
      conditions: "Partly Cloudy",
      tempLow: 65,
      tempHigh: 78,
      precipitation: 20,
      windSpeed: 8
    };
  }
}

/**
 * Generates the next day's report based on the current report
 */
function generateNextDayReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getActiveSheet();
  const currentDate = currentSheet.getRange("F7").getValue();
  
  if (!currentDate) {
    SpreadsheetApp.getUi().alert("Current sheet has no date set. Please set a date first.");
    return;
  }
  
  // Calculate next day's date
  const nextDate = new Date(currentDate);
  nextDate.setDate(nextDate.getDate() + 1);
  
  // Format date for sheet name
  const nextDateFormatted = Utilities.formatDate(nextDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  // Check if sheet for next day already exists
  let nextSheet;
  try {
    nextSheet = ss.getSheetByName(nextDateFormatted);
    if (nextSheet) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        'Sheet already exists',
        'A sheet for ' + nextDateFormatted + ' already exists. Do you want to replace it?',
        ui.ButtonSet.YES_NO
      );
      if (response === ui.Button.YES) {
        ss.deleteSheet(nextSheet);
      } else {
        return;
      }
    }
  } catch (e) {
    // Sheet doesn't exist, continue
  }
  
  // Create new sheet for next day
  nextSheet = currentSheet.copyTo(ss);
  nextSheet.setName(nextDateFormatted);
  
  // Update date
  nextSheet.getRange("F7").setValue(nextDate);
  
  // Clear daily data but keep project information
  const rangesToClear = [
    "B11:G18", // Manpower table
    "B35:E39", // Equipment section
    "B49:F53", // Materials section
    "B63:F67", // Work completed
    "B74:E76", // Issues/delays
    "B84:F86", // Safety observations
    "B94:F96", // Quality control
    "B104:F106", // Visitor log
    "A114:F114", // Photos
    "C124:E128" // Signatures
  ];
  
  rangesToClear.forEach(range => {
    nextSheet.getRange(range).clearContent();
  });
  
  // Reset calculated fields
  nextSheet.getRange("D31").setValue(0); // Total workers
  nextSheet.getRange("E31").setValue(0); // Total hours
  nextSheet.getRange("C45").setValue(0); // Total equipment
  nextSheet.getRange("D45").setValue(0); // Total equipment hours
  nextSheet.getRange("F70").setValue(0); // Overall progress
  
  // Activate the new sheet
  ss.setActiveSheet(nextSheet);
  
  SpreadsheetApp.getUi().alert("New daily report created for " + nextDateFormatted);
}

/**
 * Emails the current daily report to specified recipients
 */
function emailDailyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const projectName = sheet.getRange("C3").getValue();
  const reportDate = sheet.getRange("F7").getValue();
  
  if (!projectName || !reportDate) {
    SpreadsheetApp.getUi().alert("Missing project name or date. Please fill in these fields first.");
    return;
  }
  
  const formattedDate = Utilities.formatDate(reportDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  // Get email recipients from settings sheet or prompt user
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Email Daily Report',
    'Enter recipient email addresses (separated by commas):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.CANCEL) {
    return;
  }
  
  const recipients = response.getResponseText();
  
  try {
    // Export current sheet as PDF
    const pdfBlob = exportSheetAsPDF();
    
    // Send email with PDF attachment
    MailApp.sendEmail({
      to: recipients,
      subject: projectName + " - Daily Report " + formattedDate,
      body: "Please find attached the daily report for " + projectName + " on " + formattedDate + ".\n\n" +
            "This report was automatically generated from the Daily Report & Manpower Log system.",
      attachments: [pdfBlob]
    });
    
    ui.alert("Daily report sent successfully to: " + recipients);
  } catch (error) {
    ui.alert("Error sending email: " + error.toString());
  }
}

/**
 * Exports the current sheet as PDF
 * 
 * @returns {Blob} - PDF blob
 */
function exportAsPDF() {
  try {
    const pdfBlob = exportSheetAsPDF();
    const ui = SpreadsheetApp.getUi();
    
    // Save to Google Drive
    const projectName = SpreadsheetApp.getActiveSheet().getRange("C3").getValue();
    const reportDate = SpreadsheetApp.getActiveSheet().getRange("F7").getValue();
    const formattedDate = Utilities.formatDate(reportDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const fileName = projectName.replace(/[^a-z0-9]/gi, '_') + "_DailyReport_" + formattedDate + ".pdf";
    
    const file = DriveApp.createFile(pdfBlob.setName(fileName));
    
    ui.alert("PDF exported successfully. Saved to Google Drive as: " + fileName);
    
    // Open the PDF in a new tab
    const fileUrl = file.getUrl();
    const html = HtmlService.createHtmlOutput('<script>window.open("' + fileUrl + '", "_blank");</script>')
      .setWidth(10)
      .setHeight(10);
    ui.showModalDialog(html, "Opening PDF...");
    
    return file;
    
  } catch (error) {
    SpreadsheetApp.getUi().alert("Error exporting PDF: " + error.toString());
    return null;
  }
}

/**
 * Helper function to export the current sheet as PDF
 * 
 * @returns {Blob} - PDF blob
 */
function exportSheetAsPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // Get sheet ID and sheet name for the URL
  const sheetId = ss.getId();
  const sheetGid = sheet.getSheetId();
  
  // Construct the URL for exporting as PDF
  const url = 'https://docs.google.com/spreadsheets/d/' + sheetId + '/export?' +
    'format=pdf' +
    '&size=letter' +
    '&portrait=true' +
    '&fitw=true' +
    '&sheetnames=false' +
    '&printtitle=false' +
    '&pagenumbers=true' +
    '&gridlines=false' +
    '&fzr=false' +
    '&gid=' + sheetGid;
  
  // Fetch the PDF content
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  });
  
  // Get the project name and date for the filename
  const projectName = sheet.getRange("C3").getValue();
  const reportDate = sheet.getRange("F7").getValue();
  const formattedDate = Utilities.formatDate(reportDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const fileName = projectName.replace(/[^a-z0-9]/gi, '_') + "_DailyReport_" + formattedDate + ".pdf";
  
  // Create a blob from the response
  return response.getBlob().setName(fileName);
}

/**
 * Generates a manpower histogram for analysis
 */
function generateManpowerHistogram() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("This feature would generate a manpower histogram based on historical data. Implementation would require creating a chart in a separate analytics sheet.");
  
  // In a full implementation, this would:
  // 1. Collect manpower data across multiple daily reports
  // 2. Organize by trade and date
  // 3. Create a histogram chart
  // 4. Display in a dedicated analytics dashboard
}

/**
 * Analyzes productivity trends over time
 */
function analyzeProductivityTrends() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("This feature would analyze productivity trends based on historical data. Implementation would require collecting and processing data across multiple reports.");
  
  // In a full implementation, this would:
  // 1. Collect productivity ratings across multiple daily reports
  // 2. Calculate trends by trade and overall
  // 3. Create trend line charts
  // 4. Identify patterns and outliers
}

/**
 * Tracks cumulative impact of delays
 */
function trackDelayImpact() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("This feature would track the cumulative impact of delays. Implementation would require analyzing delay data across multiple reports.");
  
  // In a full implementation, this would:
  // 1. Collect delay information across multiple daily reports
  // 2. Categorize by reason and impact level
  // 3. Calculate cumulative impact on schedule and cost
  // 4. Generate impact analysis report
}

/**
 * Calculates total manpower hours and costs when the sheet is edited
 * 
 * @param {Object} e - The edit event object
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  // Check if edit was in the manpower table
  if (range.getRow() >= 11 && range.getRow() <= 18 && 
      range.getColumn() >= 2 && range.getColumn() <= 5) {
    calculateManpowerTotals(sheet);
  }
  
  // Check if edit was in the equipment table
  if (range.getRow() >= 35 && range.getRow() <= 39 && 
      range.getColumn() >= 2 && range.getColumn() <= 4) {
    calculateEquipmentTotals(sheet);
  }
  
  // Check if edit was in the work completed table
  if (range.getRow() >= 63 && range.getRow() <= 67 && 
      range.getColumn() === 6) {
    calculateOverallProgress(sheet);
  }
}

/**
 * Calculates manpower totals for the sheet
 * 
 * @param {Sheet} sheet - The sheet to calculate totals for
 */
function calculateManpowerTotals(sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSheet();
  
  // Define the range of the manpower table
  const manpowerRange = sheet.getRange("D11:E18");
  const manpowerData = manpowerRange.getValues();
  
  let totalWorkers = 0;
  let totalHours = 0;
  
  // Calculate totals
  for (let i = 0; i < manpowerData.length; i++) {
    const workers = manpowerData[i][0]; // Workers column
    const hours = manpowerData[i][1]; // Hours column
    
    if (!isNaN(workers) && workers !== "") {
      totalWorkers += Number(workers);
    }
    
    if (!isNaN(hours) && hours !== "") {
      totalHours += Number(hours);
    }
  }
  
  // Update totals in the sheet
  sheet.getRange("D31").setValue(totalWorkers);
  sheet.getRange("E31").setValue(totalHours);
}

/**
 * Calculates equipment totals for the sheet
 * 
 * @param {Sheet} sheet - The sheet to calculate totals for
 */
function calculateEquipmentTotals(sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSheet();
  
  // Define the range of the equipment table
  const equipmentRange = sheet.getRange("C35:D39");
  const equipmentData = equipmentRange.getValues();
  
  let totalEquipment = 0;
  let totalHours = 0;
  
  // Calculate totals
  for (let i = 0; i < equipmentData.length; i++) {
    const qty = equipmentData[i][0]; // Quantity column
    const hours = equipmentData[i][1]; // Hours column
    
    if (!isNaN(qty) && qty !== "") {
      totalEquipment += Number(qty);
    }
    
    if (!isNaN(hours) && hours !== "") {
      totalHours += Number(hours);
    }
  }
  
  // Update totals in the sheet
  sheet.getRange("C45").setValue(totalEquipment);
  sheet.getRange("D45").setValue(totalHours);
}

/**
 * Calculates overall progress for the sheet
 * 
 * @param {Sheet} sheet - The sheet to calculate progress for
 */
function calculateOverallProgress(sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSheet();
  
  // Define the range of the work completed table
  const progressRange = sheet.getRange("F63:F67");
  const progressData = progressRange.getValues();
  
  let totalProgress = 0;
  let count = 0;
  
  // Calculate average progress
  for (let i = 0; i < progressData.length; i++) {
    const progress = progressData[i][0];
    
    if (!isNaN(progress) && progress !== "") {
      totalProgress += Number(progress);
      count++;
    }
  }
  
  // Update overall progress in the sheet
  if (count > 0) {
    sheet.getRange("F70").setValue(Math.round(totalProgress / count));
  } else {
    sheet.getRange("F70").setValue(0);
  }
}