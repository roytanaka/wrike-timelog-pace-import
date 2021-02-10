const ss = SpreadsheetApp.getActiveSpreadsheet();
const activeSheet = ss.getActiveSheet();

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .createMenu('Wrike Import')
    .addItem('Show sidebar', 'showSidebar')
    .addItem('Missing Timelog Categories', 'getTimelogCategories')
    .addToUi();
  showSidebar();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle(
    'Wrike Timelog Import'
  );
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showSidebar(html);
  return;
}

function getSheetData(sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (sheet !== null) {
    const sheetValues = sheet.getDataRange().getValues();
    return sheetValues;
  }
}

function getCurrentSheet() {
  return activeSheet.getName();
}

function gotToSheet(sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  ss.setActiveSheet(sheet);
}

function getTimelogCategories() {
  const timelogCatReq = wrikeGetReq('/timelog_categories');
  const res = UrlFetchApp.fetch(timelogCatReq.url, timelogCatReq);
  const { data } = JSON.parse(res);
  const categories = data.filter((category) => !category.hidden);
  const activitySheet = ss.getSheetByName('Activity Codes');
  SpreadsheetApp.setActiveSheet(activitySheet);
  const categoryIds = activitySheet
    .getDataRange()
    .getValues()
    .map((row) => row[0]);

  const missingCategories = categories.filter(
    (category) => !categoryIds.includes(category.id)
  );

  if (missingCategories.length > 0) {
    const values = missingCategories.map((category) => {
      return [category.id, category.name];
    });
    values.forEach((row) => {
      activitySheet.appendRow(row);
    });
  } else {
    SpreadsheetApp.getUi().alert('No missing categories found on Wrike');
  }

  // const range = activitySheet.getRange(`A2:B${values.length + 1}`);
  // range.setValues(values);
}
