function RefreshGeoDS() {
  utils
    .getSheetByNameOrDie(SHEET.META)
    .getRange('B2')
    .setValue('Geo DS');
  setupEvent();
}
function RefreshGeoLS() {
  utils
    .getSheetByNameOrDie(SHEET.META)
    .getRange('B2')
    .setValue('Geo LS');
  setupEvent();
}
function RefreshHothDS() {
  utils
    .getSheetByNameOrDie(SHEET.META)
    .getRange('B2')
    .setValue('Hoth DS');
  setupEvent();
}
function RefreshHothLS() {
  utils
    .getSheetByNameOrDie(SHEET.META)
    .getRange('B2')
    .setValue('Hoth LS');
  setupEvent();
}
