// 제품 개요서 생성 스크립트
// 기능: 해당 스크립트를 실행하여 문서를 생성합니다.

//~ var doc = app.activeDocument;
//==========문서 보기 설정==========
var view = true;
var create_folder = false;
var export_subdir = false;
var folder_dir = "";
try {
  view = doc_preview;
  create_folder = check_folder.value;
  export_subdir = check_subfolder.value;
  _;
} catch (e) {}

//==========변수 선언==========
var par_name = "Cover) Document_type / Black"; //문서 단락스타일 변수 지정
var par_ser_t = doc.paragraphStyles.itemByName("Cover) Document_type / Black");

if (PDF_error(doc) == true) {
  var page_range = 0;
  var inst_range = 0;
  for (i = 0; i < doc.pages.length; i++) {
    if (doc.pages[i].pageColor == UIColors.RED) {
      inst_range = i;
    }
    if (doc.pages[i].pageColor == UIColors.BLUE) {
      page_range = i;
      break;
    }
  }
  name_threaded_frame();

  //==========단락 스타일  설정==========
  app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
  app.findGrepPreferences.appliedParagraphStyle = doc.paragraphStyles[1];
  app.changeGrepPreferences.appliedParagraphStyle = doc.paragraphStyles[0];
  doc.changeGrep();

  //==========언어 설정==========
  app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing; //찾기 설정 초기화

  app.findGrepPreferences.findWhat = "[가-힣]"; //한글 찾기
  app.findGrepPreferences.appliedParagraphStyle =
    doc.paragraphStyles.itemByName(par_name); //단락스타일 설정
  var language_found = doc.findGrep(); //찾은 결과 변수 지정

  if (language_found.length != 0) {
    //0개 이상 찾을 경우
    var language = "K-Manual"; //국문
    var safety = "안전을 위한 주의 사항";
    var caution = "취급 시 주의 사항";
    var greeting = "오토닉스 제품을 구입해 주셔서 감사합니다.";
  } else {
    var language = "E-Manual"; //영문
    var safety = "Safety Considerations";
    var caution = "Cautions during Use";
    var greeting = "Thank you for choosing our Autonics product.";
  }

  //==========문서 생성==========
  var new_doc = app.open(
    File(Folder(app.activeScript).parent + "/Autonics_Blank.indd"),
    view,
    OpenOptions.OPEN_COPY
  ); //Autonics_Blank.indd 파일 열기 (보이지 않게)
  with (new_doc) {
    documentPreferences.facingPages = true; //페이지 마주보기 설정
    documentPreferences.pageSize = "A4"; //페이지 크기 A4로 설정
    pages[0].appliedMaster = new_doc.masterSpreads.itemByName("P-Manual");
  }
}

//==========페이지 아이템 복사==========
doc.pages[0].pageItems.everyItem().duplicate(new_doc.pages[0], [15, 0]); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_1st 좌표에 붙여넣기 함
doc.pages[1].pageItems.everyItem().duplicate(new_doc.pages[0], [108, 0]); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_2nd 좌표에 붙여넣기 함

try {
  app.findGrepPreferences = app.changeGrepPreferences = null;
  app.findGrepPreferences.appliedParagraphStyle = par_ser_t;
  var found = doc.findGrep();
  if (found.length > 0) {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.appliedParagraphStyle = par_ser_t;
    var found = doc.findGrep();
    if (found.length != 0) {
      cover_move();
    }
  }
} catch (e) {}

var range = page_range == 0 ? doc.pages.length : page_range;

var count = 2;
for (var i = 2; i < range; i++) {
  duplicate_data(2);
  count++;
}

try {
  new_doc.conditions.itemByName("취급설명서").visible = false;
} catch (e) {}

try {
  new_doc.conditions.itemByName("제품 개요서").visible = false;
} catch (e) {}

try {
  new_doc.conditions.itemByName("제품 매뉴얼").visible = true;
} catch (e) {}

//==========책갈피 추가==========
app.findGrepPreferences = app.changeGrepPreferences = null; //GREP 찾기 설정 초기화
app.findGrepPreferences.appliedParagraphStyle =
  new_doc.paragraphStyles.itemByName("Heading) Title"); //GREP 찾기 단락스타일 지정
var found = new_doc.findGrep(); //GREP 찾기
var grep_find = [];
for (i = 0; i < found.length; i++) {
  //찾은 GREP 수만큼 반복
  if (
    found[i].contents != "" &&
    found[i].parentTextFrames[0].parentPage != null
  ) {
    grep_find.push(found[i]);
  }
}

var found_sort = grep_find.sort(sort_bookmark);
for (i = 0; i < found_sort.length; i++) {
  //찾은 GREP 수만큼 반복
  var grep_fnd = found_sort[i]; //n번째 찾은 GREP 변수 설정
  try {
    //시도
    var Bookmark_name = String(grep_fnd.contents); //책갈피 이름 설정
    var Bookmark_dest = new_doc.hyperlinkTextDestinations.add(
      grep_fnd.texts[0],
      { name: Bookmark_name }
    ); //책갈피 목적지 설정
    var Bookmark_added = new_doc.bookmarks.add(Bookmark_dest, {
      name: Bookmark_name,
    }); //책갈피 추가
  } catch (e) {} //에러 넘기기
}
try {
  if (
    new_doc.bookmarks.itemByName(safety).index >
    new_doc.bookmarks.itemByName(caution).index
  ) {
    new_doc.bookmarks
      .itemByName(safety)
      .move(LocationOptions.BEFORE, new_doc.bookmarks.itemByName(caution));
  }
} catch (e) {}

function sort_bookmark(a, b) {
  var pg_1 = Number(a.parentTextFrames[0].parentPage.name);
  var pg_2 = Number(b.parentTextFrames[0].parentPage.name);
  var hori_1 = Number(a.parentTextFrames[0].geometricBounds[1]);
  var hori_2 = Number(b.parentTextFrames[0].geometricBounds[1]);
  var verti_1 = Number(a.parentTextFrames[0].geometricBounds[0]);
  var verti_2 = Number(b.parentTextFrames[0].geometricBounds[0]);
  if (pg_1 < pg_2) {
    return -1;
  } else if (pg_1 == pg_2) {
    if (Math.round(hori_1) < Math.round(hori_2)) {
      return -1;
    } else if (Math.round(hori_1) == Math.round(hori_2)) {
      if (Math.round(verti_1) < Math.round(verti_2)) {
        return -1;
      } else if (Math.round(verti_1) == Math.round(verti_2)) {
        return 0;
      } else {
        return 1;
      }
    } else {
      return 1;
    }
  } else {
    return 1;
  }
}

if (doc.pages[range - 1].bounds[3] < 90) {
  new_doc.pages.lastItem().appliedMaster =
    new_doc.masterSpreads.itemByName(language);
} else {
  new_doc.pages.lastItem().appliedMaster = new_doc.masterSpreads.itemByName(
    language + "-2"
  );
}

thread_textframe();

PDF_error(new_doc);

function new_folder() {
  if (create_folder == 1) {
    var lang = language == "K-Manual" ? "국문" : "영문";
    var temp_dir = export_subdir == 0 ? save_loc : doc.filePath;
    new Folder(temp_dir + "/_출력/제품 매뉴얼/" + lang).create();
    var folder_output = String("/_출력/제품 매뉴얼/" + lang + "/");
  } else {
    var folder_output = "";
  }
  return folder_output;
}

if (view == false) {
  app.pdfExportPresets.itemByName("TC_Note").pdfPageLayout =
    PageLayoutOptions.TWO_UP_COVER_PAGE_CONTINUOUS; //연속 두페이지 표지 페이지 설정
  app.pdfExportPresets.itemByName("TC_Web").pdfPageLayout =
    PageLayoutOptions.TWO_UP_COVER_PAGE_CONTINUOUS; //연속 두페이지 표지 페이지 설정
  app.pdfExportPresets.itemByName("TC_Print").pdfPageLayout =
    PageLayoutOptions.TWO_UP_COVER_PAGE_CONTINUOUS; //연속 두페이지 표지 페이지 설정

  if (export_subdir == true) {
    save_loc = doc.filePath;
  }

  if (export_manual == 1) {
    folder_dir = new_folder();
    progress_bar.value = progress_count; //진행막대 채우기
    progress_text.text = String(
      "웹 파일 출력 중...   " + progress_count + " / " + file_count + " 완료"
    ); //진행상황 텍스트 표시
    loading_menu.active = true;

    def_PDFpref(1, 1);

    var name = new File(
      save_loc + "/" + folder_dir + doc_name + "_MANUAL_W" + ".pdf"
    ); //새파일 만들기 (폴더 경로/파일명_MANUAL.pdf)
    if (export_subdir == true) {
      name = new File(
        doc.filePath + "/" + folder_dir + doc_name + "_MANUAL_W" + ".pdf"
      ); //새파일 만들기 (폴더 경로/파일명_MANUAL.pdf)
    }
    new_doc.exportFile(
      ExportFormat.PDF_TYPE,
      name,
      false,
      app.pdfExportPresets.itemByName("TC_Web")
    ); //문서형 PDF로 출력

    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    new_doc.conditions.itemByName("제품 매뉴얼").visible = false;

    def_PDFpref(1, 1);

    var name = new File(
      save_loc + "/" + folder_dir + doc_name + "_MANUAL_TEMP" + ".pdf"
    ); //새파일 만들기 (폴더 경로/파일명_MANUAL.pdf)
    if (export_subdir == true) {
      name = new File(
        doc.filePath + "/" + folder_dir + doc_name + "_MANUAL_TEMP" + ".pdf"
      ); //새파일 만들기 (폴더 경로/파일명_MANUAL.pdf)
    }
    new_doc.exportFile(
      ExportFormat.PDF_TYPE,
      name,
      false,
      app.pdfExportPresets.itemByName("TC_Web")
    ); //문서형 PDF로 출력
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    progress_count++;
  }

  //~     if(export_idml == 1) {
  //~         var name = new File (save_loc + "/" + folder_dir + doc_name + "_MANUAL" + ".idml"); //새파일 만들기 (폴더 경로/파일명_IDML.idml)
  //~         if (export_subdir == true) {
  //~             name  = new File (doc.filePath + "/" + folder_dir + doc_name + "_MANUAL" + ".idml"); //새파일 만들기 (폴더 경로/파일명_INST.pdf)
  //~         }
  //~         new_doc.exportFile (ExportFormat.INDESIGN_MARKUP, name, false); //IDML 출력
  //~     }

  if (export_note == 1) {
    var note_frame = new_doc.pages[0].textFrames.add();
    with (note_frame) {
      geometricBounds = [79, 15, 129, 103];
      contents = "수정 이력\r회색 주석: 의뢰사항\n유색 주석: 통일 등의 사유";
      paragraphs.everyItem().appliedFont = "Autonics_TC [2020]";
      paragraphs[0].fontStyle = "Bold";
      paragraphs[0].pointSize = 60;
      paragraphs[0].justification = Justification.CENTER_ALIGN;
      paragraphs[1].pointSize = 20;
      fillColor = "Black";
      fillTint = 20;
      textFramePreferences.verticalJustification =
        VerticalJustification.CENTER_ALIGN;
      transparencySettings.blendingSettings.opacity = 90; //투명도 90%
    }
    new_doc.textFrames.itemByName("end_frame").visible = true;
    progress_bar.value = progress_count; //진행막대 채우기
    progress_text.text = String(
      "주석 파일 출력 중...   " + progress_count + " / " + file_count + " 완료"
    ); //진행상황 텍스트 표시
    loading_menu.active = true;

    def_PDFpref(1, 0);

    var name = new File(save_loc + "/" + doc_name + "_NOTE" + ".pdf"); //새파일 만들기 (폴더 경로/파일명_OVERVIEW.pdf)
    var counter = 2;

    while (name.exists) {
      name = new File(
        save_loc + "/" + doc_name + "_NOTE" + "_" + counter + ".pdf"
      ); //새파일 만들기 (폴더 경로/파일명_MANUAL.pdf)
      counter++;
    }

    new_doc.exportFile(
      ExportFormat.PDF_TYPE,
      name,
      false,
      app.pdfExportPresets.itemByName("TC_Note")
    ); //문서형 PDF로 출력

    progress_count++;
  }
  new_doc.close(SaveOptions.NO);
}

function cover_move() {
  try {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.appliedParagraphStyle =
      "Cover) Document_type / White";
    var found = new_doc.findGrep();
    var result = found[0].contents;
    found[0].parentTextFrames[0].remove();
  } catch (e) {}

  try {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.findWhat = greeting;
    var found = new_doc.findGrep();
    var result = found[0].contents;
    found[0].parentTextFrames[0].remove();
  } catch (e) {}

  var pg_inst = [];
  for (i = 0; i < new_doc.pages[0].pageItems.length; i++) {
    if (new_doc.pages[0].pageItems[i].geometricBounds[1] > 105) {
      pg_inst.push(new_doc.pages[0].pageItems[i]);
    }
  }

  try {
    var group_inst = new_doc.groups.add(pg_inst);
    group_inst.move([108, 12]);
    group_inst.ungroup();
  } catch (e) {
    pg_inst[0].move([108, 12]);
  }
}

function duplicate_data() {
  if (doc.pages[i].bounds[3] < 90) {
    if (count > 1 && count % 2 == 0) {
      new_doc.pages.add();
      if (new_doc.pages.length % 2 == 0) {
        doc.pages[i].pageItems
          .everyItem()
          .duplicate(new_doc.pages.lastItem(), [14, 0]); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_1st 좌표에 붙여넣기 함
        try {
          if (
            !(doc.pages[i + 1].pageColor == UIColors.BLUE) &&
            doc.pages[i + 1].bounds[3] < 90
          ) {
            doc.pages[i + 1].pageItems
              .everyItem()
              .duplicate(new_doc.pages.lastItem(), [107, 0]); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_1st 좌표에 붙여넣기 함
          }
        } catch (e) {}
      } else {
        doc.pages[i].pageItems
          .everyItem()
          .duplicate(new_doc.pages.lastItem(), [15, 0]); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_1st 좌표에 붙여넣기 함
        try {
          if (
            !(doc.pages[i + 1].pageColor == UIColors.BLUE) &&
            doc.pages[i + 1].bounds[3] < 90
          ) {
            doc.pages[i + 1].pageItems
              .everyItem()
              .duplicate(new_doc.pages.lastItem(), [108, 0]); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_1st 좌표에 붙여넣기 함
          }
        } catch (e) {}
      }
    }
  } else {
    new_doc.pages.add();
    if (new_doc.pages.length % 2 == 0) {
      doc.pages[i].pageItems
        .everyItem()
        .duplicate(new_doc.pages.lastItem(), [14, 0]); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_1st 좌표에 붙여넣기 함
    } else {
      doc.pages[i].pageItems
        .everyItem()
        .duplicate(new_doc.pages.lastItem(), [15, 0]); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_1st 좌표에 붙여넣기 함
    }
  }
  if (i == inst_range) {
    var end_inst = new_doc.pages.lastItem().textFrames.add();
    with (end_inst) {
      name = "end_frame";
      contents = "취급 설명서 종료";
      paragraphs[0].appliedFont = "Autonics_TC [2020]";
      paragraphs[0].fontStyle = "Bold";
      paragraphs[0].pointSize = 15;
      paragraphs[0].justification = Justification.RIGHT_ALIGN;
      paragraphs[0].rightIndent = 5;
      fillColor = "Black";
      fillTint = 20;
      textFramePreferences.verticalJustification =
        VerticalJustification.CENTER_ALIGN;
      transparencySettings.blendingSettings.opacity = 90; //투명도 90%
    }
    switch (inst_range % 4) {
      case 0:
        end_inst.geometricBounds = [279, 225, 290, 313];
        break;
      case 1:
        end_inst.geometricBounds = [279, 318, 290, 406];
        break;
      case 2:
        end_inst.geometricBounds = [279, 14, 290, 102];
        break;
      case 3:
        end_inst.geometricBounds = [279, 107, 290, 195];
        break;
    }
    end_inst.visible = false;
  }
}

function PDF_error(doc_type) {
  //PDF 에러함수 설정
  var txtfrms = doc_type.allPageItems;
  var count = 0;
  var frame_overflow = [];
  for (var i = 0; i < txtfrms.length; i++) {
    //모든 텍스트 프레임 수만큼 반복
    if (txtfrms[i] instanceof TextFrame || txtfrms[i] instanceof TextPath) {
      if (txtfrms[i].overflows == true) {
        //n번째 텍스트 프레임 넘칠 경우
        count++;
        frame_overflow.push(txtfrms[i].id);
      }
    }
  }

  if (count == 0) {
    return true;
  } else if (doc_type == doc && count != 0) {
    var error_win = new Window("dialog", "Autonics TC");
    error_win.add(
      "statictext",
      undefined,
      "넘치는 텍스트가 " + count + "개 존재합니다!"
    );
    error_win.add("statictext", undefined, "텍스트 프레임을 수정하시겠습니까?");
    var group_button = error_win.add("group");
    var ok_button = group_button.add("button", undefined, "확인", {
      name: "ok",
    });
    group_button.add("button", undefined, "취소", { name: "cancel" });
    var disp_win = error_win.show();
    if (disp_win == true) {
      auto_sizing(doc_type, frame_overflow);
      return true;
    }
  } else {
    auto_sizing(doc_type, frame_overflow);
  }
}

function auto_sizing(doc_type, frame_overflow) {
  var txtfrms;
  for (var i = 0; i < frame_overflow.length; i++) {
    txtfrms = doc_type.pageItems.itemByID(frame_overflow[i]);
    txtfrms.textFramePreferences.autoSizingReferencePoint =
      AutoSizingReferenceEnum.TOP_LEFT_POINT;
    txtfrms.textFramePreferences.useNoLineBreaksForAutoSizing = true;

    if (
      Math.round(txtfrms.geometricBounds[3] - txtfrms.geometricBounds[1]) >= 88
    ) {
      txtfrms.textFramePreferences.autoSizingType =
        AutoSizingTypeEnum.HEIGHT_ONLY;
    } else {
      txtfrms.textFramePreferences.autoSizingType =
        AutoSizingTypeEnum.HEIGHT_AND_WIDTH;
    }
  }
}

function name_threaded_frame() {
  var frames = doc.textFrames;
  var count = 0;

  for (i = 0; i < frames.length; i++) {
    if (frames[i].name.substr(0, 6) == "thread") {
      frames[i].name = "";
    }
  }

  for (i = 0; i < frames.length; i++) {
    if (
      frames[i].previousTextFrame == null &&
      frames[i].nextTextFrame != null
    ) {
      var frame_0 = frames[i];
      frame_0.parent.toString() == "[object Group]"
        ? frame_0.parent.ungroup()
        : this;
      frame_0.name = "thread_" + count + "-0";
      if (frame_0.nextTextFrame != null) {
        var frame_1 = frame_0.nextTextFrame;
        frame_1.parent.toString() == "[object Group]"
          ? frame_1.parent.ungroup()
          : this;
        frame_1.name = "thread_" + count + "-1";
      }

      try {
        if (frame_1.nextTextFrame != null) {
          var frame_2 = frame_1.nextTextFrame;
          frame_2.parent.toString() == "[object Group]"
            ? frame_2.parent.ungroup()
            : this;
          frame_2.name = "thread_" + count + "-2";
        }
      } catch (e) {}

      try {
        if (frame_2.nextTextFrame != null) {
          var frame_3 = frame_2.nextTextFrame;
          frame_3.parent.toString() == "[object Group]"
            ? frame_3.parent.ungroup()
            : this;
          frame_3.name = "thread_" + count + "-3";
        }
      } catch (e) {}

      try {
        if (frame_3.nextTextFrame != null) {
          var frame_4 = frame_3.nextTextFrame;
          frame_4.parent.toString() == "[object Group]"
            ? frame_4.parent.ungroup()
            : this;
          frame_4.name = "thread_" + count + "-4";
        }
      } catch (e) {}

      count++;
    }
  }
}

function thread_textframe() {
  try {
    for (i = 0; i < 4; i++) {
      var name_0 = "thread_" + i + "-0";
      var frame_0 = new_doc.textFrames.itemByName(name_0);
      try {
        var name_1 = "thread_" + i + "-1";
        var frame_1 = new_doc.textFrames.itemByName(name_1);
        frame_0.nextTextFrame = frame_1;
      } catch (e) {
        0;
      }
      try {
        var name_2 = "thread_" + i + "-2";
        var frame_2 = new_doc.textFrames.itemByName(name_2);
        frame_1.nextTextFrame = frame_2;
      } catch (e) {
        0;
      }
      try {
        var name_3 = "thread_" + i + "-3";
        var frame_3 = new_doc.textFrames.itemByName(name_3);
        frame_2.nextTextFrame = frame_3;
      } catch (e) {
        0;
      }
      try {
        var name_4 = "thread_" + i + "-4";
        var frame_4 = new_doc.textFrames.itemByName(name_4);
        frame_3.nextTextFrame = frame_4;
      } catch (e) {
        0;
      }
    }
  } catch (e) {}
}

app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
