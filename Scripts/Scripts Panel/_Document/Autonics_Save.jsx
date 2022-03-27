var doc = app.activeDocument;
var name = String(doc.name);

//스타일 지정 안되서 오류 생기는것도 있고 해서 그냥 아예 무시하고 동작하게 함
try {
  save_file();
} catch (e) {}

function save_file() {
  //날짜 생성
  var _date = new Date();
  var dd = _date.getDate() > 9 ? _date.getDate() : "0" + _date.getDate();
  var mm =
    _date.getMonth() + 1 > 9
      ? _date.getMonth() + 1
      : "0" + (_date.getMonth() + 1);
  var yy = _date.getFullYear();

  var today = String(yy) + String(mm) + String(dd);

  //언어 확인
  var language = "";
  app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
  app.findGrepPreferences.appliedParagraphStyle =
    "Cover) Document_type / Black";
  app.findGrepPreferences.findWhat = "[가-힣]";
  var language_found = doc.findGrep();

  language = language_found == 0 ? "_EN" : "_KO";

  //품번 확인
  var number = "";
  app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
  app.findGrepPreferences.appliedParagraphStyle = "Cover) DRW_No / Black";
  var number_found = doc.findGrep();
  if (number_found.length != 0) {
    number = String("_" + number_found[0].contents);
  }

  //시리즈명 확인
  var _seriesName = "";
  app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
  app.findGrepPreferences.appliedParagraphStyle = "Cover) Series / Black";
  var series_found = doc.findGrep();
  if (series_found.length != 0) {
    var series = series_found[0].contents;
    var _series = String(series.substr(0, series.length));

    _seriesName = _series;

    // "/"는 파일명으로 지정이 안되기 때문에 유사한 글리프로 변경
    for (var i = 0; i < _series.length; i++) {
      _seriesName = _seriesName.replace("/", "／");
    }

    for (var i = 0; i < _series.length; i++) {
      _seriesName = _seriesName.replace("Series", "");
    }

    for (var i = 0; i < _seriesName.length; i++) {
      _seriesName = _seriesName.replace(" ", "");
    }
  }

  //품번 없는경우 무시
  if (_seriesName != "") {
    name = _seriesName + language + number + "_" + today + ".indd";
  }

  //신규 파일일 경우
  if (doc.saved == false) {
    var save_file = new File(name);
    var save_loc = save_file.saveDlg(
      "저장하려는 위치를 지정해 주세요.",
      "*.indd"
    );
    doc.save(save_loc);
  }

  //이미 존재하는 파일 수정일 경우
  else {
    var save_loc = Folder(doc.filePath);
    doc.save(File(save_loc + "/" + name));
  }
  app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
}
