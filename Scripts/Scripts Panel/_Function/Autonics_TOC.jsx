//기술 해설 목차 작성용으로 만들었습니다.
var doc = app.activeDocument;

var toc_style = doc.tocStyles.itemByName("Autonics_TOC");

if (toc_style.isValid == false) {
  toc_style = doc.tocStyles.add();
  with (toc_style) {
    name = "Autonics_TOC";

    title = "CONTENTS";
    titleStyle = doc.paragraphStyles[0];
    var toc_entry = tocStyleEntries.add("Heading) Title");
    with (toc_entry) {
      formatStyle = "TOC) Contents";
      level = 1;
      pageNumberPosition = PageNumberPosition.AFTER_ENTRY;
      separator = "^n^y ";
    }
  }
  with (doc.tocStyles.itemByName("Autonics_TOC")) {
    createBookmarks = false;
    makeAnchor = true;
    removeForcedLineBreak = true;
  }
}

//목차 생성
var toc_frame = doc.createTOC(
  toc_style,
  true,
  undefined,
  [doc.pages[0].marginPreferences.left, doc.pages[0].marginPreferences.top],
  true
);

toc_frame[0].textContainers[0].move(doc.pages[1]);
toc_frame[0].textContainers[0].move(undefined, [14, 12]);

toc_frame[0].textContainers[0].textFramePreferences.textColumnCount = 2;
toc_frame[0].textContainers[0].textFramePreferences.textColumnGutter = 5;
toc_frame[0].textContainers[0].textFramePreferences.textColumnFixedWidth = 88;
toc_frame[0].textContainers[0].textFramePreferences.insetSpacing = [4, 0, 0, 0];

toc_frame[0].textContainers[0].textFramePreferences.autoSizingReferencePoint =
  AutoSizingReferenceEnum.TOP_LEFT_POINT;
toc_frame[0].textContainers[0].textFramePreferences.useNoLineBreaksForAutoSizing = true;
toc_frame[0].textContainers[0].textFramePreferences.autoSizingType =
  AutoSizingTypeEnum.HEIGHT_ONLY;

app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
app.findGrepPreferences.findWhat = "CONTENTS\r";
app.changeGrepPreferences.appliedFont = "Autonics_TC [2020]";
app.changeGrepPreferences.fontStyle = "Bold";
app.changeGrepPreferences.pointSize = "20";
app.changeGrepPreferences.tracking = 0;
app.changeGrepPreferences.fillTint = 80;
app.changeGrepPreferences.changeTo = "CONTENTS~M";
toc_frame[0].textContainers[0].changeGrep();

app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
app.findGrepPreferences.findWhat = "~y.+";
app.changeGrepPreferences.fontStyle = "Regular";
app.changeGrepPreferences.leading = 30;
toc_frame[0].textContainers[0].changeGrep();
