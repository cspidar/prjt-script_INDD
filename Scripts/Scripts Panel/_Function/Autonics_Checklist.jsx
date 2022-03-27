var doc = app.activeDocument;
var notice = [];
var count = 0;

var dependant = 0;

var alert_language = "해당 문서에 한글이 포함되어있습니다.";
var alert_overflow = "넘치는 텍스트 프레임이 존재합니다.";
var alert_font = "누락된 글꼴이 있습니다.";
var alert_link = "표지 이미지의 링크가 포함되있지 않습니다.";
var alert_drw = "품번 동기화를 진행해주세요.";
var alert_bookmark = "책갈피를 확인해주세요.";
var alert_symbol = "전압기호를 추가해주세요.";

//한글 확인
app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
app.findGrepPreferences.findWhat = "오토닉스";
var language_found = doc.findGrep();
var language = language_found == 0 ? "EN" : "KO";
if (language == "EN") {
  app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
  app.findGrepPreferences.findWhat = "[가-힣]";
  var found = doc.findGrep();
  if (found.length > 0) {
    alert(alert_language, "Autonics TC");
    count++;
    notice.push(alert_language);
  }
}

//텍스트 넘침 확인
var pg_item = doc.allPageItems;
var overflow_count = 0;
for (i = 0; i < pg_item.length; i++) {
  if (
    (pg_item[i] instanceof TextFrame || pg_item[i] instanceof TextPath) &&
    pg_item[i].overflows == true &&
    pg_item[i].parentPage != null
  ) {
    overflow_count++;
  }
}
if (overflow_count > 0) {
  alert(alert_overflow, "Autonics TC");
  count++;
  notice.push(alert_overflow);
}

//폰트 확인
var font_count = 0;
for (i = 0; i < doc.fonts.length; i++) {
  if (doc.fonts[i].status != FontStatus.INSTALLED) {
    font_count++;
  }
}
if (font_count > 0) {
  alert(alert_font, "Autonics TC");
  count++;
  notice.push(alert_font);
}

//링크 포함 확인
try {
  var link_count = 0;
  app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
  app.findGrepPreferences.appliedParagraphStyle = "Cover) Photo";
  var found = doc.findGrep();
  if (found.length != 0) {
    var image = found[0].pageItems[0];
    for (i = 0; i < image.pageItems.length; i++) {
      if (
        image.pageItems[i].graphics[0].itemLink.status ==
        LinkStatus.LINK_MISSING
      ) {
        link_count++;
      }
    }
  }
  if (link_count > 0) {
    alert(alert_link, "Autonics TC");
    count++;
    notice.push(alert_link);
  }
} catch (e) {}

//품번 확인
try {
  var drw = false;
  app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
  app.findGrepPreferences.appliedParagraphStyle = "Cover) DRW_No / Black";
  var _found = doc.findGrep();
  app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
  app.findGrepPreferences.appliedParagraphStyle = "Cover) DRW_No / White";
  var found = doc.findGrep();
  if (_found.length != 0 && found.length != 0) {
    found[0].contents == _found[0].contents ? (drw = false) : (drw = true);
  }
  if (drw == true) {
    alert(alert_drw, "Autonics TC");
    count++;
    notice.push(alert_drw);
  }
} catch (e) {}

//책갈피 확인
var bookmark_valid = true;
var _bookmark = doc.bookmarks.length;
var _hyperlink = doc.hyperlinkTextDestinations.length;
_bookmark == _hyperlink ? (bookmark_valid = true) : (bookmark_valid = false);
if (bookmark_valid == false) {
  alert(alert_bookmark, "Autonics TC");
  count++;
  notice.push(alert_bookmark);
}

//전압기호 확인
var symbol = true;
app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
app.findGrepPreferences.findWhat = "VDC\\>";
var found = doc.findGrep();
if (found.length != 0) {
  symbol = false;
}
if (symbol == false) {
  alert(alert_symbol, "Autonics TC");
  count++;
  notice.push(alert_symbol);
}

//이상 없을 경우
if (count == 0) {
  if (dependant == 0) {
    alert("발견된 이상이 없습니다.", "Autonics TC");
  }
} else {
  var list = "[    수    정    사    항    ]" + "\n\n 1. " + notice[0];
  for (i = 1; i < notice.length; i++) {
    list += "\n " + (i + 1) + ". " + notice[i];
  }
  alert(list, "Autonics TC");
}
app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
