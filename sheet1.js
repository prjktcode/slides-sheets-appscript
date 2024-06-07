// Function to append data to Google Sheets
function appendDataToSheets(data) {
  Logger.log('Appending data to sheets: ' + JSON.stringify(data));
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  if (!sheet) {
    Logger.log('Sheet1 not found. Creating new sheet.');
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Sheet1');
  }

  if (!data.seoAnalysisData || !data.seoAnalysisData.data) {
    Logger.log('SEO Analysis Data is missing or malformed: ' + JSON.stringify(data.seoAnalysisData));
    return;
  }

  var seoAnalysisData = data.seoAnalysisData.data;
  var domainAuthority = seoAnalysisData.domain_authority;
  var pageAuthority = seoAnalysisData.page_authority;
  var seoDifficulty = seoAnalysisData.seo_difficulty;
  var onPageDifficulty = seoAnalysisData.on_page_difficulty;
  var offPageDifficulty = seoAnalysisData.off_page_difficulty;

  Logger.log('Domain Authority: ' + domainAuthority);
  Logger.log('Page Authority' + pageAuthority);
  Logger.log('SEO Difficulty' +seoDifficulty);
  Logger.log('On Page Difficulty' + onPageDifficulty);
  Logger.log('Off Page Difficulty' + offPageDifficulty);

  // Clear any existing content to avoid conflicts
  sheet.clear();

  // Append the domain authority data to the sheet
  sheet.appendRow(['Metric', 'Value']);
  sheet.appendRow(['Domain Authority', domainAuthority]);
  sheet.appendRow(['Page Authority', pageAuthority]);
  sheet.appendRow(['SEO Difficulty', seoDifficulty]);
  sheet.appendRow(['On Page Difficulty', onPageDifficulty]);
  sheet.appendRow(['On Page Difficulty', offPageDifficulty]);

    // Create chart after appending data
  createChart();
}

function createChart() {
  var sheetName = 'Sheet1';
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);

  var lastRow = sheet.getLastRow();
  var range = sheet.getRange('A2:B' + lastRow);

  // Clear any existing charts first
  var existingCharts = sheet.getCharts();
  for (var i = 0; i < existingCharts.length; i++) {
    sheet.removeChart(existingCharts[i]);
  }


  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(range)
    .setPosition(5, 1, 0, 0)
    .setOption('title', 'SEO Analysis')
    .setOption('vAxis', {minValue: 0, maxValue: 100})
    .build();
  
  sheet.insertChart(chart);
  Logger.log('Chart created and inserted.');
}

// Main function to fetch data and append to sheets
function main() {
  Logger.log('Starting main function');
  var data = fetchAllData();
  Logger.log('Data fetched: ' + JSON.stringify(data));
  appendDataToSheets(data);
  Logger.log('Data appended to sheets');
}
