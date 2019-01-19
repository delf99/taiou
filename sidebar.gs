function showSidebar() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('index');
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}