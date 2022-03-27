var doc = app.activeDocument;

var mm_pt = 2.834645669291; //단위 mm -> pt 환산
//리사이즈 기준이 mm가 아니라 pt로 움직여서 환산이 필요해서 이렇게 햇ㅅ브니다

var pg_H = doc.documentPreferences.pageHeight; //페이지 높이 변수 설정

var active_page = app.activeWindow.activePage;

active_page.resize(
  BoundingBoxLimits.GEOMETRIC_PATH_BOUNDS,
  AnchorPoint.TOP_LEFT_ANCHOR,
  ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH,
  [mm_pt * 88, mm_pt * pg_H]
); //추가된 페이지 리사이즈, 높이는 기존 페이지 유지, 너비는 88mm로 변경
