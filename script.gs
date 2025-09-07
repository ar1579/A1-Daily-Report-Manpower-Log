/**
• Daily Manpower Report - Setup.gs
• Fully compliant with Enhanced Documentation: A.1 Daily Report & Manpower Log
• 
• This script handles initial setup, tab creation, formatting, and configuration
• in accordance with the enhanced documentation specifications.
 */

// Configuration per documentation
const SETUP_CONFIG = {
  TAB_NAMES: {
    DAILY_REPORT: "Daily Report",
    HISTORICAL_DATA: "Historical Data",
    DASHBOARD: "Dashboard",
    SETTINGS: "Settings",
    HELP: "Help",
    EQUIPMENT: "Equipment",
    MATERIALS: "Materials",
    WORK_COMPLETED: "Work Completed",
    ISSUES_DELAYS: "Issues & Delays",
    SAFETY: "Safety",
    QUALITY_CONTROL: "Quality Control",
    VISITOR_LOG: "Visitor Log",
  },
  COLUMN_HEADERS: {
    DAILY_REPORT: [
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
      "Notes",
    ],
    HISTORICAL_DATA: [
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
      "Notes",
    ],
    EQUIPMENT: [
      "Date",
      "Equipment Description",
      "Hours Used",
      "Operator",
      "Status",
      "Notes",
    ],
    MATERIALS: [
      "Date",
      "Material Description",
      "Quantity",
      "Supplier",
      "Delivery Date",
      "Status",
      "Notes",
    ],
    WORK_COMPLETED: [
      "Date",
      "Work Description",
      "Location",
      "Percent Complete",
      "Estimated Hours",
      "Actual Hours",
      "Notes",
    ],
    ISSUES_DELAYS: [
      "Date",
      "Issue Description",
      "Impact",
      "Resolution",
      "Status",
      "Responsible Party",
      "Target Date",
    ],
    SAFETY: [
      "Date",
      "Observation",
      "Type",
      "Incident Count",
      "Training Completed",
      "Corrective Action",
      "Status",
    ],
    QUALITY_CONTROL: [
      "Date",
      "Inspection Type",
      "Deficiencies Found",
      "Corrective Actions",
      "Status",
      "Inspector",
      "Date Completed",
    ],
    VISITOR_LOG: [
      "Date",
      "Visitor Name",
      "Company",
      "Purpose",
      "Time In",
      "Time Out",
      "Escort Required",
      "Notes",
    ],
  },
  VALIDATION_RANGES: {
    SUPERINTENDENT_LIST: [
      "Superintendent 1",
      "Superintendent 2",
      "Superintendent 3",
      "Superintendent 4",
    ],
    JOB_LIST: [
      "Project Alpha",
      "Project Beta",
      "Project Gamma",
      "Project Delta",
    ],
    EQUIPMENT_LIST: [
      "Excavator",
      "Bulldozer",
      "Crane",
      "Concrete Pump",
      "Loader",
      "Truck",
      "Forklift",
      "Generator",
    ],
    MATERIAL_STATUS: ["Ordered", "Delivered", "In Use", "Complete", "Rejected"],
    WORK_STATUS: [
      "Not Started",
      "In Progress",
      "Complete",
      "On Hold",
      "Cancelled",
    ],
    ISSUE_STATUS: ["Open", "In Progress", "Resolved", "Closed"],
    SAFETY_TYPES: [
      "Observation",
      "Near Miss",
      "Incident",
      "Training",
      "Inspection",
    ],
    QUALITY_STATUS: ["Pass", "Fail", "Pending", "Re-inspection Required"],
    ESCORT_REQUIRED: ["Yes", "No"],
  },
};

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

  // Set up triggers
  setupTriggers();

  Logger.log("Daily Manpower Report setup completed successfully");
}

/**
• Creates Daily Report tab with proper structure
 */
function createDailyReportTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dailySheet = ss.getSheetByName(SETUP_CONFIG.TAB_NAMES.DAILY_REPORT);

  if (dailySheet) {
    ss.deleteSheet(dailySheet);
  }

  dailySheet = ss.insertSheet(SETUP_CONFIG.TAB_NAMES.DAILY_REPORT);

  // Set up column headers per documentation
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
    "Notes",
  ];

  dailySheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Set up data entry row
  const dataRow = 2;
  dailySheet.getRange(dataRow, 1).setFormula("=TODAY()");
  dailySheet.getRange(dataRow, 6).setFormula("=SUM(G2:Q2)");

  // Format headers
  const headerRange = dailySheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#4285f4");
  headerRange.setFontColor("white");

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

/**
• Creates Historical Data tab with proper structure
 */
function createHistoricalDataTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let historicalSheet = ss.getSheetByName(
    SETUP_CONFIG.TAB_NAMES.HISTORICAL_DATA
  );

  if (historicalSheet) {
    ss.deleteSheet(historicalSheet);
  }

  historicalSheet = ss.insertSheet(SETUP_CONFIG.TAB_NAMES.HISTORICAL_DATA);

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
    "Notes",
  ];

  dailySheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format headers
  const headerRange = historicalSheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#34a853");
  headerRange.setFontColor("white");

  // Set column widths
  historicalSheet.setColumnWidths(1, 18, 120);

  // Freeze header row
  historicalSheet.setFrozenRows(1);
}

/**
• Creates Weather Data tab
 */
function createWeatherDataTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let weatherSheet = ss.getSheetByName("Weather Data");

  if (weatherSheet) {
    ss.deleteSheet(weatherSheet);
  }

  weatherSheet = ss.insertSheet("Weather Data");

  const headers = [
    "Date",
    "Location",
    "Conditions",
    "Temperature",
    "Humidity",
    "Wind Speed",
    "Notes",
  ];

  weatherSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format headers
  const headerRange = weatherSheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#42a5f5");
  headerRange.setFontColor("white");

  // Set column widths
  weatherSheet.setColumnWidths(1, 7, 120);

  // Freeze header row
  weatherSheet.setFrozenRows(1);
}

/**
• Creates Dashboard tab with metrics and charts
 */
function createDashboardTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dashboardSheet = ss.getSheetByName(SETUP_CONFIG.TAB_NAMES.DASHBOARD);

  if (dashboardSheet) {
    ss.deleteSheet(dashboardSheet);
  }

  dashboardSheet = ss.insertSheet(SETUP_CONFIG.TAB_NAMES.DASHBOARD);

  // Dashboard title
  dashboardSheet.getRange("A1").setValue("Daily Manpower Dashboard");
  dashboardSheet.getRange("A1").setFontSize(18).setFontWeight("bold");

  // Key metrics section
  dashboardSheet.getRange("A3").setValue("Key Metrics");
  dashboardSheet.getRange("A3").setFontSize(14).setFontWeight("bold");

  // Total manpower today
  dashboardSheet.getRange("A4").setValue("Total Manpower Today:");
  dashboardSheet
    .getRange("B4")
    .setFormula(
      "=IFERROR(INDEX('Daily Report'!F:F,MATCH(TODAY(),'Daily Report'!A:A,0)),\"No data\")"
    );

  // Weekly total
  dashboardSheet.getRange("A5").setValue("This Week Total:");
  dashboardSheet
    .getRange("B5")
    .setFormula(
      "=SUMIFS('Historical Data'!F:F,'Historical Data'!A:A,\">=\"&TODAY()-WEEKDAY(TODAY())+1,'Historical Data'!A:A,\"<=\"&TODAY())"
    );

  // Monthly total
  dashboardSheet.getRange("A6").setValue("This Month Total:");
  dashboardSheet
    .getRange("B6")
    .setFormula(
      "=SUMIFS('Historical Data'!F:F,'Historical Data'!A:A,\">=\"&EOMONTH(TODAY(),-1)+1,'Historical Data'!A:A,\"<=\"&EOMONTH(TODAY(),0))"
    );

  // Trade breakdown
  dashboardSheet.getRange("A8").setValue("Trade Breakdown");
  dashboardSheet.getRange("A8").setFontSize(14).setFontWeight("bold");

  const trades = [
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
function createSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = ss.getSheetByName(SETUP_CONFIG.TAB_NAMES.SETTINGS);

  if (settingsSheet) {
    ss.deleteSheet(settingsSheet);
  }

  settingsSheet = ss.insertSheet(SETUP_CONFIG.TAB_NAMES.SETTINGS);

  // Settings title
  settingsSheet.getRange("A1").setValue("Settings");
  settingsSheet.getRange("A1").setFontSize(18).setFontWeight("bold");

  // Configuration settings
  settingsSheet.getRange("A3").setValue("Configuration");
  settingsSheet.getRange("A3").setFontSize(14).setFontWeight("bold");

  // Location for weather
  settingsSheet.getRange("A4").setValue("Location (for weather):");
  settingsSheet.getRange("B4").setValue("New York, NY");

  // Weather API Key
  settingsSheet.getRange("A5").setValue("Weather API Key:");
  settingsSheet.getRange("B5").setValue("your-weather-api-key");

  // Default Superintendent
  settingsSheet.getRange("A6").setValue("Default Superintendent:");
  settingsSheet.getRange("B6").setValue("Superintendent 1");

  // Default Job
  settingsSheet.getRange("A7").setValue("Default Job:");
  settingsSheet.getRange("B7").setValue("Project Alpha");

  // Email Recipients
  settingsSheet.getRange("A8").setValue("Email Recipients:");
  settingsSheet.getRange("B8").setValue("email@example.com");

  // Set column widths
  settingsSheet.setColumnWidth(1, 150);
  settingsSheet.setColumnWidth(2, 200);
}

/**
• Creates Help tab with instructions
 */
function createHelpSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let helpSheet = ss.getSheetByName(SETUP_CONFIG.TAB_NAMES.HELP);

  if (helpSheet) {
    ss.deleteSheet(helpSheet);
  }

  helpSheet = ss.insertSheet(SETUP_CONFIG.TAB_NAMES.HELP);

  // Help content
  const helpContent = [
    ["Daily Manpower Report - User Guide"],
    [""],
    ["Getting Started:"],
    ["1. Fill in the Daily Report tab with today's data"],
    ["2. Superintendent and Job Name have dropdown validation"],
    ["3. Trade counts should be entered as numbers"],
    ["4. Weather data is automatically populated from settings"],
    [""],
    ["Additional Tables:"],
    ["- Equipment: Track equipment usage and hours"],
    ["- Materials: Record material deliveries and usage"],
    ["- Work Completed: Document daily work progress"],
    ["- Issues & Delays: Log any problems encountered"],
    ["- Safety: Record safety observations and incidents"],
    ["- Quality Control: Track quality inspections"],
    ["- Visitor Log: Record site visitors"],
    [""],
    ["Daily Report Tab:"],
    ["- Date: Automatically set to today"],
    ["- Superintendent: Select from dropdown list"],
    ["- Job Name: Select from dropdown list"],
    ["- Weather: Current conditions (auto-populated)"],
    ["- Trade counts: Enter number of workers for each trade"],
    ["- Notes: Additional comments or observations"],
    [""],
    ["Historical Data Tab:"],
    ["- Automatically populated with daily reports"],
    ["- Use for trend analysis and reporting"],
    [""],
    ["Dashboard Tab:"],
    ["- Shows key metrics and trends"],
    ["- Updates automatically with new data"],
    [""],
    ["Settings Tab:"],
    ["- Configure API keys and preferences"],
    ["- Update email recipients"],
    [""],
    ["For support, contact your system administrator"],
  ];

  helpSheet.getRange(1, 1, helpContent.length, 1).setValues(helpContent);
  helpSheet.setColumnWidth(1, 600);

  // Format title
  helpSheet.getRange("A1").setFontSize(18).setFontWeight("bold");
}

/**
• Creates named ranges for the daily report
 */
function createNamedRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dailySheet = ss.getSheetByName("Daily Report");

  if (!dailySheet) {
    Logger.log("Daily Report sheet not found for named ranges");
    return;
  }

  // Create named ranges for key sections
  try {
    // Project Info range
    ss.setNamedRange("ProjectInfo", dailySheet.getRange("A2:C2"));

    // Weather range
    ss.setNamedRange("WeatherData", dailySheet.getRange("D2:E2"));

    // Manpower table range
    ss.setNamedRange("ManpowerTable", dailySheet.getRange("F2:Q2"));

    // Notes range
    ss.setNamedRange("NotesSection", dailySheet.getRange("R2"));

    Logger.log("Named ranges created successfully");
  } catch (error) {
    Logger.log("Error creating named ranges: " + error.toString());
  }
}

/**
• Applies consistent formatting across all sheets
 */
function applySheetFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = [
    SETUP_CONFIG.TAB_NAMES.DAILY_REPORT,
    SETUP_CONFIG.TAB_NAMES.HISTORICAL_DATA,
    SETUP_CONFIG.TAB_NAMES.DASHBOARD,
    SETUP_CONFIG.TAB_NAMES.SETTINGS,
    SETUP_CONFIG.TAB_NAMES.HELP,
    SETUP_CONFIG.TAB_NAMES.EQUIPMENT,
    SETUP_CONFIG.TAB_NAMES.MATERIALS,
    SETUP_CONFIG.TAB_NAMES.WORK_COMPLETED,
    SETUP_CONFIG.TAB_NAMES.ISSUES_DELAYS,
    SETUP_CONFIG.TAB_NAMES.SAFETY,
    SETUP_CONFIG.TAB_NAMES.QUALITY_CONTROL,
    SETUP_CONFIG.TAB_NAMES.VISITOR_LOG,
  ];

  allSheets.forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      // Set default font
      sheet
        .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
        .setFontFamily("Arial")
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
  const dailySheet = ss.getSheetByName(SETUP_CONFIG.TAB_NAMES.DAILY_REPORT);
  if (dailySheet) {
    const superintendentRange = dailySheet.getRange("B:B");
    const superintendentRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        "Superintendent 1",
        "Superintendent 2",
        "Superintendent 3",
      ])
      .build();
    superintendentRange.setDataValidation(superintendentRule);

    const jobRange = dailySheet.getRange("C:C");
    const jobRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Project Alpha", "Project Beta", "Project Gamma"])
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
  const dailySheet = ss.getSheetByName(SETUP_CONFIG.TAB_NAMES.DAILY_REPORT);
  const historicalSheet = ss.getSheetByName(
    SETUP_CONFIG.TAB_NAMES.HISTORICAL_DATA
  );

  if (dailySheet) {
    // Weekend highlighting
    const dateRange = dailySheet.getRange("A2");
    const weekendRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=WEEKDAY(A2,2)>5")
      .setBackground("#fff2cc")
      .setRanges([dateRange])
      .build();

    // Missing data alerts
    const totalManpowerRange = dailySheet.getRange("E2");
    const missingDataRule = SpreadsheetApp.newConditionalFormatRule()
      .whenCellEmpty()
      .setBackground("#ffcccc")
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
      .whenFormulaSatisfied("=WEEKDAY(A:A,2)>5")
      .setBackground("#e6f3ff")
      .setRanges([historicalSheet.getRange("A:A")])
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
    if (trigger.getHandlerFunction() === "dailyReportTrigger") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new daily trigger
  ScriptApp.newTrigger("dailyReportTrigger")
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
  Logger.log("Setup completed successfully");
}

// Additional table creation functions - standardized to use "Tab" suffix
function createEquipmentTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SETUP_CONFIG.TAB_NAMES.EQUIPMENT);

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet(SETUP_CONFIG.TAB_NAMES.EQUIPMENT);

  const headers = [
    "Date",
    "Equipment Description",
    "Hours Used",
    "Operator",
    "Status",
    "Notes",
  ];

  if (headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#ff9800");
    headerRange.setFontColor("white");

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
  let sheet = ss.getSheetByName(SETUP_CONFIG.TAB_NAMES.MATERIALS);

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet(SETUP_CONFIG.TAB_NAMES.MATERIALS);

  const headers = [
    "Date",
    "Material Description",
    "Quantity",
    "Supplier",
    "Delivery Date",
    "Status",
    "Notes",
  ];

  if (headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#9c27b0");
    headerRange.setFontColor("white");

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
  let sheet = ss.getSheetByName(SETUP_CONFIG.TAB_NAMES.WORK_COMPLETED);

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet(SETUP_CONFIG.TAB_NAMES.WORK_COMPLETED);

  const headers = [
    "Date",
    "Work Description",
    "Location",
    "Percent Complete",
    "Estimated Hours",
    "Actual Hours",
    "Notes",
  ];

  if (headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#4caf50");
    headerRange.setFontColor("white");

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
  let sheet = ss.getSheetByName(SETUP_CONFIG.TAB_NAMES.ISSUES_DELAYS);

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet(SETUP_CONFIG.TAB_NAMES.ISSUES_DELAYS);

  const headers = [
    "Date",
    "Issue Description",
    "Impact",
    "Resolution",
    "Status",
    "Responsible Party",
    "Target Date",
  ];

  if (headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#f44336");
    headerRange.setFontColor("white");

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
  let sheet = ss.getSheetByName(SETUP_CONFIG.TAB_NAMES.SAFETY);

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet(SETUP_CONFIG.TAB_NAMES.SAFETY);

  const headers = [
    "Date",
    "Observation",
    "Type",
    "Incident Count",
    "Training Completed",
    "Corrective Action",
    "Status",
  ];

  if (headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#795548");
    headerRange.setFontColor("white");

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
  let sheet = ss.getSheetByName(SETUP_CONFIG.TAB_NAMES.QUALITY_CONTROL);

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet(SETUP_CONFIG.TAB_NAMES.QUALITY_CONTROL);

  const headers = [
    "Date",
    "Inspection Type",
    "Deficiencies Found",
    "Corrective Actions",
    "Status",
    "Inspector",
    "Date Completed",
  ];

  if (headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#607d8b");
    headerRange.setFontColor("white");

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
  let sheet = ss.getSheetByName(SETUP_CONFIG.TAB_NAMES.VISITOR_LOG);

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet(SETUP_CONFIG.TAB_NAMES.VISITOR_LOG);

  const headers = [
    "Date",
    "Visitor Name",
    "Company",
    "Purpose",
    "Time In",
    "Time Out",
    "Escort Required",
    "Notes",
  ];

  if (headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#e91e63");
    headerRange.setFontColor("white");

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
