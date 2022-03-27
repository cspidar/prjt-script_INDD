var doc = app.activeDocument;

//Autonics_Manual 이름의 목차 스타일 변수 지정
var toc_style = doc.tocStyles.itemByName("Autonics_Manual");

//Autonics_Manual 이름의 목차 스타일 없을 경우 목차 스타일 생성
if (toc_style.isValid == false) {
  toc_style = doc.tocStyles.add();
  with (toc_style) {
    name = "Autonics_Manual"; //스타일 명

    title = "Table of Contents"; //타이틀
    titleStyle = doc.paragraphStyles.itemByName("Preface"); //타이틀 단락 스타일
    var toc_entry_P = tocStyleEntries.add("Preface"); //목차 세부 항목

    // 1뎁스 목차 생성
    with (toc_entry_P) {
      formatStyle = "TOC) Preface";
      level = 1;
      pageNumberPosition = PageNumberPosition.AFTER_ENTRY;
      separator = "^y ";
    }
    var toc_entry_1 = tocStyleEntries.add("Heading_1");
    with (toc_entry_1) {
      formatStyle = "TOC) Heading_1";
      level = 1;
      pageNumberPosition = PageNumberPosition.AFTER_ENTRY;
      separator = "^y ";
    }

    //2뎁스 목차 생성
    var toc_entry_2 = tocStyleEntries.add("Heading_2");
    with (toc_entry_2) {
      formatStyle = "TOC) Heading_2";
      level = 2;
      pageNumberPosition = PageNumberPosition.AFTER_ENTRY;
      separator = "^y ";
    }

    //3뎁스 목차 생성
    var toc_entry_3 = tocStyleEntries.add("Heading_3");
    with (toc_entry_3) {
      formatStyle = "TOC) Heading_3";
      level = 3;
      pageNumberPosition = PageNumberPosition.AFTER_ENTRY;
      separator = "^y ";
    }
  }
  //목차 세부 설정
  with (doc.tocStyles.itemByName("Autonics_Manual")) {
    createBookmarks = true;
    includeBookDocuments = true;
    makeAnchor = true;
    removeForcedLineBreak = true;
    numberedParagraphs = NumberedParagraphsOptions.INCLUDE_FULL_PARAGRAPH;
  }
}

//목차 생성
var toc_frame = doc.createTOC(
  toc_style,
  true,
  undefined,
  [
    doc.pages.lastItem().marginPreferences.left,
    doc.pages.lastItem().marginPreferences.top,
  ],
  true
); //목차 만들기

toc_frame[0].textContainers[0].move(doc.pages.lastItem()); //생성된 목차 마지막 페이지로 이동
toc_frame[0].textContainers[0].move(undefined, [17, 30]); //페이지 이동된 목차 위치 이동

//뭐였는지 기억 안나는데 필요한건 맞음 인디자인 기능을 아예 id로 찾아서 지정해놨었는데..
toc_frame[0].texts[0].select();
app.menuActions.itemByID(71442).invoke();

//목차 텍스트 프레임 넘칠 경우 자동으로 다음페이지 넘어가게 함
var overflow_frame = doc.pages.lastItem().textFrames[0];
if (overflow_frame.overflows == true) {
  var new_page = doc.pages.add(LocationOptions.AT_END);
  overflow_frame.nextTextFrame = new_page.textFrames[0];
}

//목차도 책갈피로 추가되게 하는 기능인데 인디자인 내에서 목차 업데이트 하면 다 사라져버려서 쓸모없는 코드
app.findGrepPreferences = app.changeGrepPreferences = null; //GREP 찾기 설정 초기화
app.findGrepPreferences.findWhat = "Table of Contents";
app.findGrepPreferences.appliedParagraphStyle =
  doc.paragraphStyles.itemByName("Preface");

var found = doc.findGrep(); //GREP 찾기
var Bookmark_name = "Table of Contents"; //책갈피 이름 설정
var Bookmark_dest = doc.hyperlinkTextDestinations.add(found[0].texts[0], {
  name: Bookmark_name,
}); //책갈피 목적지 설정
var Bookmark_added = doc.bookmarks.add(Bookmark_dest, { name: Bookmark_name }); //책갈피 추가

doc.bookmarks
  .itemByName("Table of Contents")
  .move(LocationOptions.AFTER, doc.bookmarks.itemByName("취급 시 주의 사항"));
