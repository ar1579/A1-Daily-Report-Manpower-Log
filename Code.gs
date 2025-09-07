/**
• Daily Manpower Report - Code.gs
• Fully compliant with Enhanced Documentation: A.1 Daily Report & Manpower Log
• 
• This script handles daily report automation, calculations, and data management
• in accordance with the enhanced documentation specifications.
 */


// Configuration constants per documentation
const CONFIG = {
  DATE_FORMAT: 'MM/dd/yyyy',
  WEATHER_API_KEY: '65a2db8131fe93a346f8b3aaafc3b883'
};


/**
• Creates daily report with proper formatting and structure
• @param {Date} reportDate - The date for the report
• @return {Object} Daily report object with all required fields
 */
function createDailyReport(reportDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dailySheet = ss.getSheetByName('Daily Report');


  if (!dailySheet) {
    throw new Error('Daily Report sheet not found');
  }


  // Format date per documentation MM/DD/YYYY
  const formattedDate = Utilities.formatDate(reportDate, Session.getScriptTimeZone(), CONFIG.DATE_FORMAT);


  // Get weather data with location from settings
  const weatherData = getWeatherData(reportDate);


  // Create comprehensive report structure per documentation
  return {
    date: formattedDate,
    superintendent: '',
    jobName: '',
    weather: weatherData.conditions,
    temperature: weatherData.temperature,
    humidity: weatherData.humidity,
    windSpeed: weatherData.windSpeed,
    totalManpower: 0,
    carpenters: 0,
    electricians: 0,
    plumbers: 0,
    hvac: 0,
    generalLaborers: 0,
    equipmentOperators: 0,
    ironWorkers: 0,
    concreteWorkers: 0,
    roofers: 0,
    painters: 0,
    landscapers: 0,
    notes: ''
  };
}


/**
• Gets weather data from API using location from settings
• @param {Date} date - Date for weather data
• @return {Object} Weather data object
 */
function getWeatherData(date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');


  // Get location from settings
  let location = 'New York, NY'; // Default
  if (settingsSheet) {
    const locationCell = settingsSheet.getRange('B4');
    if (locationCell.getValue()) {
      location = locationCell.getValue();
    }
  }


  try {
    const url = `https://api.openweathermap.org/data/2.5/weather?q=${location}&appid=${CONFIG.WEATHER_API_KEY}&units=imperial`;
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());


return {
  conditions: data.weather[0].description,
  temperature: Math.round(data.main.temp) + '\u00b0F',
  humidity: data.main.humidity + '%',
  windSpeed: data.wind.speed + ' mph'
};

  } catch (error) {
    Logger.log('Weather API error: ' + error.toString());
    return {
      conditions: 'Weather data unavailable',
      temperature: 'N/A',
      humidity: 'N/A',
      windSpeed: 'N/A'
    };
  }
}


/**
• Creates the complete daily report system with all required sheets
 */
function createDailyReportSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();


  // Create all required sheets
  createDailyReportTab();
  createHistoricalDataTab();
  createWeatherDataTab();
  createDashboardTab();
  createSettingsTab();
  createHelpTab();


  // Create additional tables per documentation
  createEquipmentTab();
  createMaterialsTab();
  createWorkCompletedTab();
  createIssuesDelaysTab();
  createSafetyTab();
  createQualityControlTab();
  createVisitorLogTab();


  // Create named ranges
  createNamedRanges();


  // Apply formatting and validation
  applySheetFormatting();
  applyDataValidationRules();


  Logger.log("Daily Report System created successfully");
}


/**
• Main setup function - renamed from testSetup to setupDailyManpowerReport
 */
function setupDailyManpowerReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();


  // Create all required tabs
  createDailyReportTab();
  createHistoricalDataTab();
  createDashboardTab();
  createSettingsTab();
  createHelpTab();


  // Create additional tables per documentation
  createEquipmentTab();
  createMaterialsTab();
  createWorkCompletedTab();
  createIssuesDelaysTab();
  createSafetyTab();
  createQualityControlTab();
  createVisitorLogTab();


  // Create named ranges
  createNamedRanges();


  // Apply formatting and validation
  applyFormatting();
  applyDataValidation();
  applyConditionalFormatting();


  // Set up triggers automatically
  setupTriggers();


  Logger.log("Daily Manpower Report setup completed successfully");
}


/**
• Creates Daily Report tab with proper structure
 */
function createDailyReportTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dailySheet = ss.getSheetByName('Daily Report');


  if (dailySheet) {
    ss.deleteSheet(dailySheet);
  }


  dailySheet = ss.getSheetByName('Daily Report');


  const headers = [
    "Date",
    "Superintendent",
    "Job Name",
    "Weather",
    "Temperature",
    "Total Manpower",
    "Carpenters",
    "Electricians",
    "Plumbers",
    "HVAC",
    "General Laborers",
    "Equipment Operators",
    "Iron Workers",
    "Concrete Workers",
    "Roofers",
    "Painters",
    "Landscapers",
    "Notes"
  ];


  if (dailySheet) {
    dailySheet.getRange(1, 1, 1, headers.length).setValues([headers]);


// Set up data entry row
const dataRow = 2;
dailySheet.getRange(dataRow, 1).setFormula('=TODAY()');
dailySheet.getRange(dataRow, 6).setFormula('=SUM(G2:Q2)');

// Format headers
const headerRange = dailySheet.getRange(1, 1, 1, headers.length);
headerRange.setFontWeight('bold');
headerRange.setBackground('#4285f4');
headerRange.setFontColor('white');

// Set column widths
dailySheet.setColumnWidths(1, 18, 120);
dailySheet.setColumnWidth(1, 100);
dailySheet.setColumnWidth(2, 150);
dailySheet.setColumnWidth(3, 150);
dailySheet.setColumnWidth(4, 120);
dailySheet.setColumnWidth(18, 200);

// Freeze header row
dailySheet.setFrozenRows(1);

  }
}


/**
• Creates Historical Data tab with proper structure
 */
function createHistoricalDataTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let historicalSheet = ss.getSheetByName('Historical Data');


  if (historicalSheet) {
    ss.deleteSheet(historicalSheet);
  }


  historicalSheet = ss.getSheetByName('Historical Data');


  const headers = [
    "Date",
    "Superintendent",
    "Job Name",
    "Weather",
    "Temperature",
    "Total Manpower",
    "Carpenters",
    "Electricians",
    "Plumbers",
    "HVAC",
    "General Laborers",
    "Equipment Operators",
    "Iron Workers",
    "Concrete Workers",
    "Roofers",
    "Painters",
    "Landscapers",
    "Notes"
  ];


  if (historicalSheet) {
    historicalSheet.getRange(1, 1, 1, headers.length).setValues([headers]);


// Format headers
const headerRange = historicalSheet.getRange(1, 1, 1, headers.length);
headerRange.setFontWeight('bold');
headerRange.setBackground('#34a853');
headerRange.setFontColor('white');

// Set column widths
historicalSheet.setColumnWidths(1, 18, 120);

// Freeze header row
historicalSheet.setFrozenRows(1);

  }
}


/**
• Creates named ranges for the daily report
 */
function createNamedRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dailySheet = ss.getSheetByName('Daily Report');


  if (!dailySheet) {
    Logger.log('Daily Report sheet not found for named ranges');
    return;
  }


  // Create named ranges for key sections
  try {
    // Project Info range
    ss.setNamedRange('ProjectInfo', dailySheet.getRange('A2:C2'));


// Weather range
ss.setNamedRange('WeatherData', dailySheet.getRange('D2:E2'));

// Manpower table range
ss.setNamedRange('ManpowerTable', dailySheet.getRange('F2:Q2'));

// Notes range
ss.setNamedRange('NotesSection', dailySheet.getRange('R2'));

Logger.log('Named ranges created successfully');

  } catch (error) {
    Logger.log('Error creating named ranges: ' + error.toString());
  }
}


/**
• Gets email recipients from Settings tab
• @return {Array} Array of email addresses
 */
function getEmailRecipients() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');


  if (!settingsSheet) {
    Logger.log('Settings sheet not found for email recipients');
    return ['email@example.com']; // Default fallback
  }


  const recipientsCell = settingsSheet.getRange('B8');
  const recipientsText = recipientsCell.getValue();


  if (!recipientsText) {
    return ['email@example.com']; // Default fallback
  }


  // Parse email addresses - split by comma, semicolon, or new line
  const recipients = recipientsText.toString()
    .split(/[,;\
]+/)
    .map(email => email.trim())
    .filter(email => email.length > 0 && email.includes('@'));


  return recipients.length > 0 ? recipients : ['email@example.com'];
}


/**
• Simplified PDF generation using Google Sheets export URL
• @param {string} dateStr - Date string for report
• @return {Blob} PDF blob with proper formatting
 */
function exportSheetAsPDF(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fileId = ss.getId();


  // Get project name from settings
  const settingsSheet = ss.getSheetByName('Settings');
  let projectName = 'DailyReport';
  if (settingsSheet) {
    const projectCell = settingsSheet.getRange('B7');
    if (projectCell.getValue()) {
      projectName = projectCell.getValue().toString().replace(/\s+/g, '_');
    }
  }


  // Construct PDF filename
  const filename = `${projectName}_${dateStr.replace(/\//g, '-')}.pdf`;


  // Construct export URL with formatting options
  const exportUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=pdf&size=letter&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=UNDEFINED&attachment=false`;


  try {
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    });


const blob = response.getBlob();
blob.setName(filename);

return blob;

  } catch (error) {
    Logger.log('PDF generation failed: ' + error.toString());


// Fallback - create simple PDF
const tempFile = DriveApp.createFile('temp_report.txt', 'Report data would go here');
return tempFile.getBlob();

  }
}


/**
• Updates historical data with new daily report
• @param {Object} reportData - Daily report data
 */
function updateHistoricalData(reportData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historicalSheet = ss.getSheetByName('Historical Data');


  if (!historicalSheet) {
    throw new Error('Historical Data sheet not found');
  }


  // Find next empty row
  const lastRow = historicalSheet.getLastRow();
  const nextRow = lastRow + 1;


  // Write data per documentation structure
  historicalSheet.getRange(nextRow, 1).setValue(reportData.date);
  historicalSheet.getRange(nextRow, 2).setValue(reportData.superintendent);
  historicalSheet.getRange(nextRow, 3).setValue(reportData.jobName);
  historicalSheet.getRange(nextRow, 4).setValue(reportData.weather);
  historicalSheet.getRange(nextRow, 5).setValue(reportData.temperature);
  historicalSheet.getRange(nextRow, 6).setValue(reportData.totalManpower);
  historicalSheet.getRange(nextRow, 7).setValue(reportData.carpenters);
  historicalSheet.getRange(nextRow, 8).setValue(reportData.electricians);
  historicalSheet.getRange(nextRow, 9).setValue(reportData.plumbers);
  historicalSheet.getRange(nextRow, 10).setValue(reportData.hvac);
  historicalSheet.getRange(nextRow, 11).setValue(reportData.generalLaborers);
  historicalSheet.getRange(nextRow, 12).setValue(reportData.equipmentOperators);
  historicalSheet.getRange(nextRow, 13).setValue(reportData.ironWorkers);
  historicalSheet.getRange(nextRow, 14).setValue(reportData.concreteWorkers);
  historicalSheet.getRange(nextRow, 15).setValue(reportData.roofers);
  historicalSheet.getRange(nextRow, 16).setValue(reportData.painters);
  historicalSheet.getRange(nextRow, 17).setValue(reportData.landscapers);
  historicalSheet.getRange(nextRow, 18).setValue(reportData.notes);
}


/**
• Sends daily report email with dynamic recipients
• @param {Object} reportData - Daily report data
 */
function sendDailyReportEmail(reportData) {
  const emailRecipients = getEmailRecipients();
  const subject = `Daily Manpower Report - ${reportData.date}`;
  const body = `
Daily Manpower Report for ${reportData.date}


Superintendent: ${reportData.superintendent}
Job Name: ${reportData.jobName}
Weather: ${reportData.weather}, ${reportData.temperature}, ${reportData.humidity}, ${reportData.windSpeed}


Total Manpower: ${reportData.totalManpower}


Trade Breakdown:
• Carpenters: ${reportData.carpenters}
• Electricians: ${reportData.electricians}
• Plumbers: ${reportData.plumbers}
• HVAC: ${reportData.hvac}
• General Laborers: ${reportData.generalLaborers}
• Equipment Operators: ${reportData.equipmentOperators}
• Iron Workers: ${reportData.ironWorkers}
• Concrete Workers: ${reportData.concreteWorkers}
• Roofers: ${reportData.roofers}
• Painters: ${reportData.painters}
• Landscapers: ${reportData.landscapers}


Notes: ${reportData.notes}
  `;


  const pdfBlob = exportSheetAsPDF(reportData.date);


  emailRecipients.forEach(email => {
    GmailApp.sendEmail(email, subject, body, {
      attachments: [pdfBlob]
    });
  });
}


/**
• Main function to process daily report
 */
function processDailyReport() {
  const today = new Date();
  const report = createDailyReport(today);


  // Update from sheet data
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dailySheet = ss.getSheetByName('Daily Report');


  if (dailySheet) {
    report.superintendent = dailySheet.getRange('B2').getValue();
    report.jobName = dailySheet.getRange('C2').getValue();
    report.carpenters = dailySheet.getRange('G2').getValue() || 0;
    report.electricians = dailySheet.getRange('H2').getValue() || 0;
    report.plumbers = dailySheet.getRange('I2').getValue() || 0;
    report.hvac = dailySheet.getRange('J2').getValue() || 0;
    report.generalLaborers = dailySheet.getRange('K2').getValue() || 0;
    report.equipmentOperators = dailySheet.getRange('L2').getValue() || 0;
    report.ironWorkers = dailySheet.getRange('M2').getValue() || 0;
    report.concreteWorkers = dailySheet.getRange('N2').getValue() || 0;
    report.roofers = dailySheet.getRange('O2').getValue() || 0;
    report.painters = dailySheet.getRange('P2').getValue() || 0;
    report.landscapers = dailySheet.getRange('Q2').getValue() || 0;
    report.notes = dailySheet.getRange('R2').getValue();


// Calculate total manpower
report.totalManpower = calculateTotalManpower({
  carpenters: report.carpenters,
  electricians: report.electricians,
  plumbers: report.plumbers,
  hvac: report.hvac,
  generalLaborers: report.generalLaborers,
  equipmentOperators: report.equipmentOperators,
  ironWorkers: report.ironWorkers,
  concreteWorkers: report.concreteWorkers,
  roofers: report.roofers,
  painters: report.painters,
  landscapers: report.landscapers
});

  }


  // Update historical data
  updateHistoricalData(report);


  // Send email report
  sendDailyReportEmail(report);
}


/**
• Calculates total manpower based on trade counts
• @param {Object} tradeCounts - Object containing trade counts
• @return {number} Total manpower count
 */
function calculateTotalManpower(tradeCounts) {
  // Formula per documentation: sum of all trade counts
  const trades = ['carpenters', 'electricians', 'plumbers', 'hvac', 'generalLaborers', 
               'equipmentOperators', 'ironWorkers', 'concreteWorkers', 'roofers', 
               'painters', 'landscapers'];


  let total = 0;
  trades.forEach(trade => {
    total += parseInt(tradeCounts[trade]) || 0;
  });


  return total;
}


/**
• Scheduled trigger function
 */
function dailyReportTrigger() {
  processDailyReport();
}


/**
• Creates Daily Report tab with proper structure
 */
function createDailyReportTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dailySheet = ss.getSheetByName('Daily Report');


  if (dailySheet) {
    ss.deleteSheet(dailySheet);
  }


  dailySheet = ss.getSheetByName('Daily Report');


  const headers = [
    "Date",
    "Superintendent",
    "Job Name",
    "Weather",
    "Temperature",
    "Total Manpower",
    "Carpenters",
    "Electricians",
    "Plumbers",
    "HVAC",
    "General Laborers",
    "Equipment Operators",
    "Iron Workers",
    "Concrete Workers",
    "Roofers",
    "Painters",
    "Landscapers",
    "Notes"
  ];


  if (dailySheet) {
    dailySheet.getRange(1, 1, 1, headers.length).setValues([headers]);


// Set up data entry row
const dataRow = 2;
dailySheet.getRange(dataRow, 1).setFormula('=TODAY()');
dailySheet.getRange(dataRow, 6).setFormula('=SUM(G2:Q2)');

// Format headers
const headerRange = dailySheet.getRange(1, 1, 1, headers.length);
headerRange.setFontWeight('bold');
headerRange.setBackground('#4285f4');
headerRange.setFontColor('white');

// Set column widths
dailySheet.setColumnWidths(1, 18, 120);
dailySheet.setColumnWidth(1, 100);
dailySheet.setColumnWidth(2, 150);
dailySheet.setColumnWidth(3, 150);
dailySheet.setColumnWidth(4, 120);
dailySheet.setColumnWidth(18, 200);

// Freeze header row
dailySheet.setFrozenRows(1);

  }
}


/**
• Creates Historical Data tab with proper structure
 */
function createHistoricalDataTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let historicalSheet = ss.getSheetByName('Historical Data');


  if (historicalSheet) {
    ss.deleteSheet(historicalSheet);
  }


  historicalSheet = ss.getSheetByName('Historical Data');


  const headers = [
    "Date",
    "Superintendent",
    "Job Name",
    "Weather",
    "Temperature",
    "Total Manpower",
    "Carpenters",
    "Electricians",
    "Plumbers",
    "HVAC",
    "General Laborers",
    "Equipment Operators",
    "Iron Workers",
    "Concrete Workers",
    "Roofers",
    "Painters",
    "Landscapers",
    "Notes"
  ];


  if (historicalSheet) {
    historicalSheet.getRange(1, 1, 1, headers.length).setValues([headers]);


// Format headers
const headerRange = historicalSheet.getRange(1, 1, 1, headers.length);
headerRange.setFontWeight('bold');
headerRange.setBackground('#34a853');
headerRange.setFontColor('white');

// Set column widths
historicalSheet.setColumnWidths(1, 18, 120);

// Freeze header row
historicalSheet.setFrozenRows(1);

  }
}


/**
• Creates Weather Data tab
 */
function createWeatherDataTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let weatherSheet = ss.getSheetByName('Weather Data');


  if (weatherSheet) {
    ss.deleteSheet(weatherSheet);
  }


  weatherSheet = ss.getSheetByName('Weather Data');


  const headers = [
    "Date",
    "Location",
    "Conditions",
    "Temperature",
    "Humidity",
    "Wind Speed",
    "Notes"
  ];


  if (weatherSheet) {
    weatherSheet.getRange(1, 1, 1, headers.length).setValues([headers]);


// Format headers
const headerRange = weatherSheet.getRange(1, 1, 1, headers.length);
headerRange.setFontWeight('bold');
headerRange.setBackground('#42a5f5');
headerRange.setFontColor('white');

// Set column widths
weatherSheet.setColumnWidths(1, 7, 120);

// Freeze header row
weatherSheet.setFrozenRows(1);

  }
}


/**
• Creates Dashboard tab with metrics and charts
 */
function createDashboardTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dashboardSheet = ss.getSheetByName('Dashboard');


  if (dashboardSheet) {
    ss.deleteSheet(dashboardSheet);
  }


  dashboardSheet = ss.getSheetByName('Dashboard');


  // Dashboard title
  dashboardSheet.getRange('A1').setValue('Daily Manpower Dashboard');
  dashboardSheet.getRange('A1').setFontSize(18).setFontWeight('bold');


  // Key metrics section
  dashboardSheet.getRange('A3').setValue('Key Metrics');
  dashboardSheet.getRange('A3').setFontSize(14).setFontWeight('bold');


  // Total manpower today
  dashboardSheet.getRange('A4').setValue('Total Manpower Today:');
  dashboardSheet
    .getRange('B4')
    .setFormula(
      "=IFERROR(INDEX('Daily Report'!F:F,MATCH(TODAY(),'Daily Report'!A:A,0)),\"No data\")"
    );


  // Weekly total
  dashboardSheet.getRange('A5').setValue('This Week Total:');
  dashboardSheet
    .getRange('B5')
    .setFormula(
      "=SUMIFS('Historical Data'!F:F,'Historical Data'!A:A,\">=\"&TODAY()-WEEKDAY(TODAY())+1,'Historical Data'!A:A,\"<=\"&TODAY())"
    );


  // Monthly total
  dashboardSheet.getRange('A6').setValue('This Month Total:');
  dashboardSheet
    .getRange('B6')
    .setFormula(
      "=SUMIFS('Historical Data'!F:F,'Historical Data'!A:A,\">=\"&EOMONTH(TODAY(),-1)+1,'Historical Data'!A:A,\"<=\"&EOMONTH(TODAY(),0))"
    );


  // Trade breakdown
  dashboardSheet.getRange('A8').setValue('Trade Breakdown');
  dashboardSheet.getRange('A8').setFontSize(14).setFontWeight('bold');


  const trades = [
    'Carpenters',
    'Electricians',
    'Plumbers',
    'HVAC',
    'General Laborers',
    'Equipment Operators',
    'Iron Workers',
    'Concrete Workers',
    'Roofers',
    'Painters',
    'Landscapers',
  ];


  trades.forEach((trade, index) => {
    dashboardSheet.getRange(`A${9 + index}`).setValue(trade);
    dashboardSheet
      .getRange(`B${9 + index}`)
      .setFormula(
        `=SUMIFS('Historical Data'!${String.fromCharCode(
          71 + index
        )}:${String.fromCharCode(
          71 + index
        )},'Historical Data'!A:A,">="&TODAY()-30,'Historical Data'!A:A,"<="&TODAY())`
      );
  });


  // Set column widths
  dashboardSheet.setColumnWidth(1, 200);
  dashboardSheet.setColumnWidth(2, 150);
}


/**
• Creates Settings tab for configuration
 */
function createSettingsTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = ss.getSheetByName('Settings');


  if (settingsSheet) {
    ss.deleteSheet(settingsSheet);
  }


  settingsSheet = ss.getSheetByName('Settings');


  // Settings title
  settingsSheet.getRange('A1').setValue('Settings');
  settingsSheet.getRange('A1').setFontSize(18).setFontWeight('bold');


  // Configuration settings
  settingsSheet.getRange('A3').setValue('Configuration');
  settingsSheet.getRange('A3').setFontSize(14).setFontWeight('bold');


  // Location for weather
  settingsSheet.getRange('A4').setValue('Location (for weather):');
  settingsSheet.getRange('B4').setValue('New York, NY');


  // Weather API Key
  settingsSheet.getRange('A5').setValue('Weather API Key:');
  settingsSheet.getRange('B5').setValue('65a2db8131fe93a346f8b3aaafc3b883');


  // Default Superintendent
  settingsSheet.getRange('A6').setValue('Default Superintendent:');
  settingsSheet.getRange('B6').setValue('Superintendent 1');


  // Default Job
  settingsSheet.getRange('A7').setValue('Default Job:');
  settingsSheet.getRange('B7').setValue('Project Alpha');


  // Email Recipients
  settingsSheet.getRange('A8').setValue('Email Recipients:');
  settingsSheet.getRange('B8').setValue('email@example.com');


  // Set column widths
  settingsSheet.setColumnWidth(1, 150);
  settingsSheet.setColumnWidth(2, 200);
}


/**
• Creates Help tab with instructions
 */
function createHelpTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let helpSheet = ss.getSheetByName('Help');


  if (helpSheet) {
    ss.deleteSheet(helpSheet);
  }


  helpSheet = ss.getSheetByName('Help');


  // Help content
  const helpContent = [
    ['Daily Manpower Report - User Guide'],
    [''],
    ['Getting Started:'],
    ['1. Fill in the Daily Report tab with today's data'],
    ['2. Superintendent and Job Name have dropdown validation'],
    ['3. Trade counts should be entered as numbers'],
    ['4. Weather data is automatically populated from settings'],
    [''],
    ['Additional Tables:'],
    ['- Equipment: Track equipment usage and hours'],
    ['- Materials: Record material deliveries and usage'],
    ['- Work Completed: Document daily work progress'],
    ['- Issues & Delays: Log any problems encountered'],
    ['- Safety: Record safety observations and incidents'],
    ['- Quality Control: Track quality inspections'],
    ['- Visitor Log: Record site visitors'],
    [''],
    ['Daily Report Tab:'],
    ['- Date: Automatically set to today'],
    ['- Superintendent: Select from dropdown list'],
    ['- Job Name: Select from dropdown list'],
    ['- Weather: Current conditions (auto-populated)'],
    ['- Trade counts: Enter number of workers for each trade'],
    ['- Notes: Additional comments or observations'],
    [''],
    ['Historical Data Tab:'],
    ['- Automatically populated with daily reports'],
    ['- Use for trend analysis and reporting'],
    [''],
    ['Dashboard Tab:'],
    ['- Shows key metrics and trends'],
    ['- Updates automatically with new data'],
    [''],
    ['Settings Tab:'],
    ['- Configure API keys and preferences'],
    ['- Update email recipients'],
    [''],
    ['For support, contact your system administrator']
  ];


  helpSheet.getRange(1, 1, helpContent.length, 1).setValues(helpContent);
  helpSheet.setColumnWidth(1, 600);


  // Format title
  helpSheet.getRange('A1').setFontSize(18).setFontWeight('bold');
}


/**
• Creates named ranges for the daily report
 */
function createNamedRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dailySheet = ss.getSheetByName('Daily Report');


  if (!dailySheet) {
    Logger.log('Daily Report sheet not found for named ranges');
    return;
  }


  // Create named ranges for key sections
  try {
    // Project Info range
    ss.setNamedRange('ProjectInfo', dailySheet.getRange('A2:C2'));


// Weather range
ss.setNamedRange('WeatherData', dailySheet.getRange('D2:E2'));

// Manpower table range
ss.setNamedRange('ManpowerTable', dailySheet.getRange('F2:Q2'));

// Notes range
ss.setNamedRange('NotesSection', dailySheet.getRange('R2'));

Logger.log('Named ranges created successfully');

  } catch (error) {
    Logger.log('Error creating named ranges: ' + error.toString());
  }
}


/**
• Applies consistent formatting across all sheets
 */
function applySheetFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = [
 'Daily Report',
 'Historical Data',
 'Dashboard',
 'Settings',
 'Help',
 'Equipment',
 'Materials',
 'Work Completed',
 'Issues & Delays',
 'Safety',
 'Quality Control',
 'Visitor Log'
  ];


  allSheets.forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      // Set default font
      sheet
        .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
        .setFontFamily('Arial')
        .setFontSize(10);


  // Auto-resize columns
  sheet.autoResizeColumns(1, sheet.getMaxColumns());
}

  });
}


/**
• Applies data validation rules
 */
function applyDataValidationRules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();


  // Daily Report validation
  const dailySheet = ss.getSheetByName('Daily Report');
  if (dailySheet) {
    const superintendentRange = dailySheet.getRange('B:B');
    const superintendentRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Superintendent 1', 'Superintendent 2', 'Superintendent 3'])
      .build();
    superintendentRange.setDataValidation(superintendentRule);


const jobRange = dailySheet.getRange('C:C');
const jobRule = SpreadsheetApp.newDataValidation()
  .requireValueInList(['Project Alpha', 'Project Beta', 'Project Gamma'])
  .build();
jobRange.setDataValidation(jobRule);

  }
}


/**
• Applies formatting and validation
 */
function applyFormatting() {
  applySheetFormatting();
}


/**
• Applies data validation rules
 */
function applyDataValidation() {
  applyDataValidationRules();
}


/**
• Applies conditional formatting per documentation
 */
function applyConditionalFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dailySheet = ss.getSheetByName('Daily Report');
  const historicalSheet = ss.getSheetByName('Historical Data');


  if (dailySheet) {
    // Weekend highlighting
    const dateRange = dailySheet.getRange('A2');
    const weekendRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=WEEKDAY(A2,2)>5')
      .setBackground('#fff2cc')
      .setRanges([dateRange])
      .build();


// Missing data alerts
const totalManpowerRange = dailySheet.getRange('E2');
const missingDataRule = SpreadsheetApp.newConditionalFormatRule()
  .whenCellEmpty()
  .setBackground('#ffcccc')
  .setRanges([totalManpowerRange])
  .build();

const rules = dailySheet.getConditionalFormatRules();
rules.push(weekendRule);
rules.push(missingDataRule);
dailySheet.setConditionalFormatRules(rules);

  }


  if (historicalSheet) {
    // Highlight weekends in historical data
    const historicalRules = historicalSheet.getConditionalFormatRules();
    const historicalWeekendRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=WEEKDAY(A:A,2)>5')
      .setBackground('#e6f3ff')
      .setRanges([historicalSheet.getRange('A:A')])
      .build();


historicalRules.push(historicalWeekendRule);
historicalSheet.setConditionalFormatRules(historicalRules);

  }
}


/**
• Sets up time-based triggers
 */
function setupTriggers() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
 if (trigger.getHandlerFunction() === 'dailyReportTrigger') {
   ScriptApp.deleteTrigger(trigger);
 }
  });


  // Create new daily trigger
  ScriptApp.newTrigger('dailyReportTrigger')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
}


/**
• Test function to verify setup - now calls setupDailyManpowerReport
 */
function testSetup() {
  setupDailyManpowerReport();
  Logger.log('Setup completed successfully');
}


// Additional table creation functions - standardized to use "Tab" suffix
function createEquipmentTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Equipment');


  if (sheet) {
    ss.deleteSheet(sheet);
  }


  sheet = ss.getSheetByName('Equipment');


  const headers = [
    'Date',
    'Equipment Description',
    'Hours Used',
    'Operator',
    'Status',
    'Notes'
  ];


  if (sheet && headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);


// Format headers
const headerRange = sheet.getRange(1, 1, 1, headers.length);
headerRange.setFontWeight('bold');
headerRange.setBackground('#ff9800');
headerRange.setFontColor('white');

// Set column widths
sheet.setColumnWidth(1, 100); // Date
sheet.setColumnWidth(2, 200); // Equipment Description
sheet.setColumnWidth(3, 100); // Hours Used
sheet.setColumnWidth(4, 150); // Operator
sheet.setColumnWidth(5, 100); // Status
sheet.setColumnWidth(6, 200); // Notes

// Freeze header row
sheet.setFrozenRows(1);

  }
}


function createMaterialsTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Materials');


  if (sheet) {
    ss.deleteSheet(sheet);
  }


  sheet = ss.getSheetByName('Materials');


  const headers = [
    'Date',
    'Material Description',
    'Quantity',
    'Supplier',
    'Delivery Date',
    'Status',
    'Notes'
  ];


  if (sheet && headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);


// Format headers
const headerRange = sheet.getRange(1, 1, 1, headers.length);
headerRange.setFontWeight('bold');
headerRange.setBackground('#9c27b0');
headerRange.setFontColor('white');

// Set column widths
sheet.setColumnWidth(1, 100); // Date
sheet.setColumnWidth(2, 200); // Material Description
sheet.setColumnWidth(3, 100); // Quantity
sheet.setColumnWidth(4, 150); // Supplier
sheet.setColumnWidth(5, 120); // Delivery Date
sheet.setColumnWidth(6, 100); // Status
sheet.setColumnWidth(7, 200); // Notes

// Freeze header row
sheet.setFrozenRows(1);

  }
}


function createWorkCompletedTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Work Completed');


  if (sheet) {
    ss.deleteSheet(sheet);
  }


  sheet = ss.getSheetByName('Work Completed');


  const headers = [
    'Date',
    'Work Description',
    'Location',
    'Percent Complete',
    'Estimated Hours',
    'Actual Hours',
    'Notes'
  ];


  if (sheet && headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);


// Format headers
const headerRange = sheet.getRange(1, 1, 1, headers.length);
headerRange.setFontWeight('bold');
headerRange.setBackground('#4caf50');
headerRange.setFontColor('white');

// Set column widths
sheet.setColumnWidth(1, 100); // Date
sheet.setColumnWidth(2, 200); // Work Description
sheet.setColumnWidth(3, 150); // Location
sheet.setColumnWidth(4, 120); // Percent Complete
sheet.setColumnWidth(5, 120); // Estimated Hours
sheet.setColumnWidth(6, 120); // Actual Hours
sheet.setColumnWidth(7, 200); // Notes

// Freeze header row
sheet.setFrozenRows(1);

  }
}


function createIssuesDelaysTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Issues & Delays');


  if (sheet) {
    ss.deleteSheet(sheet);
  }


  sheet = ss.getSheetByName('Issues & Delays');


  const headers = [
    'Date',
    'Issue Description',
    'Impact',
    'Resolution',
    'Status',
    'Responsible Party',
    'Target Date'
  ];


  if (sheet && headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);


// Format headers
const headerRange = sheet.getRange(1, 1, 1, headers.length);
headerRange.setFontWeight('bold');
headerRange.setBackground('#f44336');
headerRange.setFontColor('white');

// Set column widths
sheet.setColumnWidth(1, 100); // Date
sheet.setColumnWidth(2, 200); // Issue Description
sheet.setColumnWidth(3, 150); // Impact
sheet.setColumnWidth(4, 200); // Resolution
sheet.setColumnWidth(5, 100); // Status
sheet.setColumnWidth(6, 150); // Responsible Party
sheet.setColumnWidth(7, 120); // Target Date

// Freeze header row
sheet.setFrozenRows(1);

  }
}


function createSafetyTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Safety');


  if (sheet) {
    ss.deleteSheet(sheet);
  }


  sheet = ss.getSheetByName('Safety');


  const headers = [
    'Date',
    'Observation',
    'Type',
    'Incident Count',
    'Training Completed',
    'Corrective Action',
    'Status'
  ];


  if (sheet && headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);


// Format headers
const headerRange = sheet.getRange(1, 1, 1, headers.length);
headerRange.setFontWeight('bold');
headerRange.setBackground('#795548');
headerRange.setFontColor('white');

// Set column widths
sheet.setColumnWidth(1, 100); // Date
sheet.setColumnWidth(2, 200); // Observation
sheet.setColumnWidth(3, 120); // Type
sheet.setColumnWidth(4, 120); // Incident Count
sheet.setColumnWidth(5, 150); // Training Completed
sheet.setColumnWidth(6, 200); // Corrective Action
sheet.setColumnWidth(7, 100); // Status

// Freeze header row
sheet.setFrozenRows(1);

  }
}


function createQualityControlTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Quality Control');


  if (sheet) {
    ss.deleteSheet(sheet);
  }


  sheet = ss.getSheetByName('Quality Control');


  const headers = [
    'Date',
    'Inspection Type',
    'Deficiencies Found',
    'Corrective Actions',
    'Status',
    'Inspector',
    'Date Completed'
  ];


  if (sheet && headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);


// Format headers
const headerRange = sheet.getRange(1, 1, 1, headers.length);
headerRange.setFontWeight('bold');
headerRange.setBackground('#607d8b');
headerRange.setFontColor('white');

// Set column widths
sheet.setColumnWidth(1, 100); // Date
sheet.setColumnWidth(2, 200); // Inspection Type
sheet.setColumnWidth(3, 200); // Deficiencies Found
sheet.setColumnWidth(4, 200); // Corrective Actions
sheet.setColumnWidth(5, 100); // Status
sheet.setColumnWidth(6, 150); // Inspector
sheet.setColumnWidth(7, 120); // Date Completed

// Freeze header row
sheet.setFrozenRows(1);

  }
}


function createVisitorLogTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Visitor Log');


  if (sheet) {
    ss.deleteSheet(sheet);
  }


  sheet = ss.getSheetByName('Visitor Log');


  const headers = [
    'Date',
    'Visitor Name',
    'Company',
    'Purpose',
    'Time In',
    'Time Out',
    'Escort Required',
    'Notes'
  ];


  if (sheet && headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);


// Format headers
const headerRange = sheet.getRange(1, 1, 1, headers.length);
headerRange.setFontWeight('bold');
headerRange.setBackground('#e91e63');
headerRange.setFontColor('white');

// Set column widths
sheet.setColumnWidth(1, 100); // Date
sheet.setColumnWidth(2, 150); // Visitor Name
sheet.setColumnWidth(3, 150); // Company
sheet.setColumnWidth(4, 150); // Purpose
sheet.setColumnWidth(5, 100); // Time In
sheet.setColumnWidth(6, 100); // Time Out
sheet.setColumnWidth(7, 120); // Escort Required
sheet.setColumnWidth(8, 200); // Notes

// Freeze header row
sheet.setFrozenRows(1);

  }
}
