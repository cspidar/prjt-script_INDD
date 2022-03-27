var doc = app.activeDocument;

var mm_pt = 2.834645669291; //단위 mm -> pt 환산
var _Font = "Autonics_TC [2020]"; //폰트 설정

var pg_H = doc.documentPreferences.pageHeight; //페이지 높이 변수 설정

var active_spread = app.activeWindow.activeSpread;
var new_spreads = doc.spreads.add(LocationOptions.AFTER, active_spread);

doc.spreads.everyItem().allowPageShuffle = false;

if (Math.round(new_spreads.pages[0].bounds[3]) != 88) {
  new_spreads.pages[0].resize(
    BoundingBoxLimits.GEOMETRIC_PATH_BOUNDS,
    AnchorPoint.TOP_LEFT_ANCHOR,
    ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH,
    [mm_pt * 88, mm_pt * pg_H]
  ); //추가된 페이지 리사이즈, 높이는 기존 페이지 유지, 너비는 88mm로 변경
}
