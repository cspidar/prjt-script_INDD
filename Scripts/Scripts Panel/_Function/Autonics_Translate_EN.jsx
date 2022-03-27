var doc = app.activeDocument;

var save_loc = Folder(
  app.activeScript.parent.parent.parent.parent + "/_Translate"
);
var load_files = save_loc.getFiles("*.indd");

var file_length = load_files.length;

var folder_selected = "폴더 선택 안함";
try {
  folder_selected = dropdown_folder.selection.text;
} catch (e) {}

if (folder_selected != "폴더 선택 안함") {
  var list_folder = Folder(
    app.activeScript.parent.parent.parent.parent +
      "/_Translate/_Series/" +
      folder_selected
  );
  var _files = list_folder.getFiles("*.indd");
  file_length += _files.length;
}

var loading_menu = new Window("palette");
var loading_text = loading_menu.add("statictext", undefined, "변환 시작"); //로딩창 생성
loading_text.orientation = "column";
loading_text.alignChildren = ["left", "top"];
loading_text.bounds = [0, 0, 340, 20];

var progress_bar = loading_menu.add(
  "progressbar",
  [12, 12, 350, 24],
  0,
  file_length
); //진행막대 추가
var progress_text = loading_menu.add("statictext", undefined, "출력 시작..."); //문자 추가
progress_text.bounds = [0, 0, 340, 20]; //문자 크기 설정
progress_text.justify = "right"; //우측 정렬
progress_text.alignment = ["right", "top"]; //우측 정렬

loading_menu.add("statictext", undefined, "Esc를 누르면 취소됩니다.");

loading_menu.show();
loading_menu.active = true;

var key_stats = ScriptUI.environment.keyboardState; //키보드 설정
if (key_stats.keyName == "Escape") {
  //ESC를 누를 경우
  doc_word.close(SaveOptions.NO); //열려있는 파일 닫기 (저장 안함)
  loading_menu.close(); //로딩 창 닫기
}

loading_menu.onClose = function () {
  doc_word.close(SaveOptions.NO); //열려있는 파일 닫기 (저장 안함)
};

//==========문서 변환==========
var kor = ["카탈로그", "제품", "매뉴얼"];
var eng = ["CATALOG", "PRODUCT", "MANUAL"];

loading_text.text = String("한-영 변환 중... "); //로딩창 띄우기
progress_bar.value = 1; //진행막대 채우기
progress_text.text = String(
  "문서 양식 변경 " + progress_bar.value + " / " + file_length + " 완료"
); //진행상황 텍스트 표시
loading_menu.active = true;

for (var i = 0; i < kor.length; i++) {
  app.findTextPreferences = app.changeTextPreferences = null;
  app.findTextPreferences.appliedParagraphStyle =
    "Cover) Document_type / Black";
  translate(kor[i], eng[i]);
}

app.findTextPreferences = app.changeTextPreferences = null;
app.findTextPreferences.appliedParagraphStyle = "Cover) Document_type / White";
translate("취급설명서", "INSTRUCTION MANUAL");

//==========Heading) Title 변환==========
if (folder_selected != "--폴더 선택--") {
  var doc_word = app.open(_files[0], false);
  var word_list = doc_word.textFrames[0].tables[0];

  progress_bar.value++; //진행막대 채우기
  progress_text.text = String(
    "제품군 변경 (" +
      doc_word.name +
      ") " +
      progress_bar.value +
      " / " +
      file_length +
      " 완료"
  ); //진행상황 텍스트 표시
  loading_menu.active = true;

  for (var k = 0; k < word_list.rows.length; k++) {
    var kor = word_list.rows[k].cells[0].contents;
    var eng = word_list.rows[k].cells[1].contents;
    app.findTextPreferences = app.changeTextPreferences = null;
    app.findTextPreferences.appliedParagraphStyle = "Heading) Title";
    translate(kor, eng);
  }

  doc_word.close(SaveOptions.NO);
}

var doc_word = app.open(load_files[0], false);
var word_list = doc_word.textFrames[0].tables[0];

progress_bar.value = 2; //진행막대 채우기
progress_text.text = String(
  "Heading)Title 항목 변경 (" +
    doc_word.name +
    ") " +
    progress_bar.value +
    " / " +
    file_length +
    " 완료"
); //진행상황 텍스트 표시
loading_menu.active = true;

for (var i = 0; i < word_list.rows.length; i++) {
  var kor = word_list.rows[i].cells[0].contents;
  var eng = word_list.rows[i].cells[1].contents;
  app.findTextPreferences = app.changeTextPreferences = null;
  app.findTextPreferences.appliedParagraphStyle = "Heading) Title";
  translate(kor, eng);
}

doc_word.close(SaveOptions.NO);

//==========PL 변환==========
var doc_word = app.open(load_files[1], false);
var word_list = doc_word.textFrames[0].tables[0];

progress_bar.value = 3; //진행막대 채우기
progress_text.text = String(
  "PL 문구 변경 (" +
    doc_word.name +
    ") " +
    progress_bar.value +
    " / " +
    file_length +
    " 완료"
); //진행상황 텍스트 표시
loading_menu.active = true;

for (var i = 0; i < word_list.rows.length; i++) {
  var kor = word_list.rows[i].cells[0].contents;
  var eng = word_list.rows[i].cells[1].contents;
  app.findTextPreferences = app.changeTextPreferences = null;
  translate(kor, eng);
}

doc_word.close(SaveOptions.NO);

//==========문장 변환==========
var doc_word = app.open(load_files[2], false);
var word_list = doc_word.textFrames[0].tables[0];

progress_bar.value++; //진행막대 채우기
progress_text.text = String(
  "문장 변경 (" +
    doc_word.name +
    ") " +
    progress_bar.value +
    " / " +
    file_length +
    " 완료"
); //진행상황 텍스트 표시
loading_menu.active = true;

for (var i = 0; i < word_list.rows.length; i++) {
  var kor = word_list.rows[i].cells[0].contents;
  var eng = word_list.rows[i].cells[1].contents;
  app.findTextPreferences = app.changeTextPreferences = null;
  translate(kor, eng);
}

doc_word.close(SaveOptions.NO);

//==========폴더 변환==========
if (folder_selected != "--폴더 선택--") {
  for (var j = 1; j < _files.length; j++) {
    var doc_word = app.open(_files[j], false);
    var word_list = doc_word.textFrames[0].tables[0];

    progress_bar.value++; //진행막대 채우기
    progress_text.text = String(
      "제품군 변경 (" +
        doc_word.name +
        ") " +
        progress_bar.value +
        " / " +
        file_length +
        " 완료"
    ); //진행상황 텍스트 표시
    loading_menu.active = true;

    for (var k = 0; k < word_list.rows.length; k++) {
      var kor = word_list.rows[k].cells[0].contents;
      var eng = word_list.rows[k].cells[1].contents;
      app.findTextPreferences = app.changeTextPreferences = null;
      translate(kor, eng);
    }

    doc_word.close(SaveOptions.NO);
  }
}

//==========나머지 파일 변환==========
for (var i = 3; i < load_files.length; i++) {
  var doc_word = app.open(load_files[i], false);
  var word_list = doc_word.textFrames[0].tables[0];

  progress_bar.value++; //진행막대 채우기
  progress_text.text = String(
    "기타 변경 (" +
      doc_word.name +
      ") " +
      progress_bar.value +
      " / " +
      file_length +
      " 완료"
  ); //진행상황 텍스트 표시
  loading_menu.active = true;

  for (var j = 0; j < word_list.rows.length; j++) {
    var kor = word_list.rows[j].cells[0].contents;
    var eng = word_list.rows[j].cells[1].contents;
    app.findTextPreferences = app.changeTextPreferences = null;
    translate(kor, eng);
  }

  doc_word.close(SaveOptions.NO);
}

//==========언어 변환==========
app.findTextPreferences = app.changeTextPreferences = null;
app.findTextPreferences.appliedLanguage = "한국어";
app.changeTextPreferences.appliedLanguage = "영어: 미국";
doc.changeText();

//==========매뉴얼 표지 레이아웃 이동==========
try {
  app.findTextPreferences = app.changeTextPreferences = null;
  app.findTextPreferences.appliedParagraphStyle = "Cover) Depth_4 / Black";
  var temp1 = doc.findText()[0].parentTextFrames[0];

  temp1.textFramePreferences.autoSizingReferencePoint =
    AutoSizingReferenceEnum.TOP_LEFT_POINT;
  temp1.textFramePreferences.useNoLineBreaksForAutoSizing = true;
  temp1.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;

  app.findTextPreferences = app.changeTextPreferences = null;
  app.findTextPreferences.findWhat = "Features";
  app.findTextPreferences.appliedParagraphStyle = "Heading) Title";
  var temp2 = doc.findText()[0].parentTextFrames[0];

  doc.distribute(
    [temp1, temp2],
    DistributeOptions.VERTICAL_SPACE,
    AlignDistributeBounds.ITEM_BOUNDS,
    true,
    0
  );
} catch (e) {}

//==========취설 표지 레이아웃 이동==========
try {
  app.findTextPreferences = app.changeTextPreferences = null;
  app.findTextPreferences.findWhat =
    "Thank you for choosing our Autonics product.";
  var temp3 = doc.findText()[0].parentTextFrames[0];

  temp3.textFramePreferences.autoSizingReferencePoint =
    AutoSizingReferenceEnum.TOP_LEFT_POINT;
  temp3.textFramePreferences.useNoLineBreaksForAutoSizing = true;
  temp3.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;

  doc.distribute(
    temp3.parentPage.pageItems.everyItem(),
    DistributeOptions.VERTICAL_SPACE,
    AlignDistributeBounds.ITEM_BOUNDS,
    true,
    5
  );
} catch (e) {}

loading_menu.close();

app.findTextPreferences = app.changeTextPreferences = null;
alert("완료되었습니다.", "Autonics TC");

function translate(k, e) {
  if (k != "") {
    app.findTextPreferences.findWhat = k;
    app.changeTextPreferences.changeTo = e;
    app.selection.length == 0
      ? doc.changeText()
      : app.selection[0].changeText();
  }
}
