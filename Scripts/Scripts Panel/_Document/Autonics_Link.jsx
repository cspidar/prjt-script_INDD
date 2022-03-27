var doc = app.activeDocument;

//단락스타일 변수 지정
var link_drw = doc.links.itemByName("DRW_no");
var link_depth = doc.links.itemByName("Depth_4");
var link_ser = doc.links.itemByName("Series_name");

var par_drw = doc.paragraphStyles.itemByName("Cover) DRW_No / Black");
var par_drw_t = doc.paragraphStyles.itemByName("Cover) DRW_No / White");

var par_dep = doc.paragraphStyles.itemByName("Cover) Depth_4 / Black");
var par_dep_t = doc.paragraphStyles.itemByName("Cover) Depth_4 / White");

var par_ser = doc.paragraphStyles.itemByName("Cover) Series / Black");
var par_ser_t = doc.paragraphStyles.itemByName("Cover) Series / White");

//함수 실행
change_info(par_drw, par_drw_t);
change_info(par_dep, par_dep_t);
change_info(par_ser, par_ser_t);

//함수 실행
remove_line("\r");
remove_line("\n");

//제품 개요서 명칭 카탈로그로 변경됨.. (모든 제품군 제품 개요서 명칭 사라지면 삭제 가능 29~32줄 삭제)
app.findGrepPreferences = app.changeGrepPreferences = null;
app.findGrepPreferences.findWhat = "제품 개요서";
app.changeGrepPreferences.changeTo = "카탈로그";
doc.changeGrep();

//변경 함수
function change_info(from, to) {
  app.findGrepPreferences = app.changeGrepPreferences = null;
  app.findGrepPreferences.appliedParagraphStyle = from; //단락 스타일 찾기 "~~~ / Black"
  var found = doc.findGrep();
  if (found.length > 0) {
    //있으면
    var drw_no = found[0].contents; //찾은 내용 변수 지정 (1)
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.appliedParagraphStyle = to; //단락 스타일 찾기 "~~~ / White"
    app.changeGrepPreferences.changeTo = drw_no; // (1) 내용으로 변경
    doc.changeGrep();
  } else {
    //없으면 그냥 넘어감
    try {
      app.findGrepPreferences = app.changeGrepPreferences = null;
      app.findGrepPreferences.appliedParagraphStyle = to;
      var found = doc.findGrep();
      found[0].contents = "";
    } catch (e) {}
  }
}

//줄바꿈 삭제
function remove_line(line_break) {
  app.findGrepPreferences = app.changeGrepPreferences = null;
  app.findGrepPreferences.appliedParagraphStyle = par_dep_t;
  app.findGrepPreferences.findWhat = line_break;
  app.changeGrepPreferences.changeTo = " ";
  doc.changeGrep();
  app.findGrepPreferences = app.changeGrepPreferences = null;
  app.findGrepPreferences.appliedParagraphStyle = par_dep_t;
  app.findGrepPreferences.findWhat = "  ";
  app.changeGrepPreferences.changeTo = " ";
  doc.changeGrep();
}

app.findGrepPreferences = app.changeGrepPreferences = null;

alert("완료되었습니다.", "Autonics TC");
