function RefreshGeoDS() {
  SPREADSHEET.getSheetByName(SHEET.META)
    .getRange('B2')
    .setValue('Geo DS');
  setupEvent();
}
function RefreshHothDS() {
  SPREADSHEET.getSheetByName(SHEET.META)
    .getRange('B2')
    .setValue('Hoth DS');
  setupEvent();
}
function RefreshHothLS() {
  SPREADSHEET.getSheetByName(SHEET.META)
    .getRange('B2')
    .setValue('Hoth LS');
  setupEvent();
}
