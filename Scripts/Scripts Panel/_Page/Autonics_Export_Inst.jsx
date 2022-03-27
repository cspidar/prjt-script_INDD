// 취급 설명서 생성 스크립트
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
} catch (e) {}

//==========단 좌표 설정==========
var section_cover = [8.5, 0]; //취설 1번째 쪽 x, y 좌표 값
var section_1st = [8.5, 0]; //1번째 단 x, y 좌표 값
var section_2nd = [113.5, 0]; //2번째 단 x, y 좌표 값
var section_3rd = [218.5, 0]; //3번째 단 x, y 좌표 값
var section_4th = [323.5, 0]; //4번째 단 x, y 좌표 값
var section_5th = [428.5, 0]; //5번째 단 x, y 좌표 값

//==========변수 선언==========
var par_name = "Cover) Document_type / Black"; //문서 단락스타일 변수 지정

if (PDF_error(doc) == true) {
  var page_range = 0;
  for (i = 1; i < doc.pages.length; i++) {
    if (doc.pages[i].pageColor == UIColors.RED) {
      page_range = i;
    }
  }

  if (page_range == 0) {
    try {
      page_range = 4;
      doc.pages[4].pageColor = UIColors.RED;
    } catch (e) {
      page_range = 2;
      doc.pages[2].pageColor = UIColors.RED;
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
    var language = "K-Inst"; //국문
    var safety = "안전을 위한 주의 사항";
    var caution = "취급 시 주의 사항";
    var greeting = "오토닉스 제품을 구입해 주셔서 감사합니다.";
  } else {
    var language = "E-Inst"; //영문
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
    documentPreferences.facingPages = false; //페이지 마주보기 설정
    documentPreferences.pageSize = "A3"; //페이지 크기 A4로 설정
    pages[0].appliedMaster = new_doc.masterSpreads.itemByName("I-Inst");
  }
}

//==========페이지 아이템 복사==========
var pg_num = doc.pages.length - 1; //활성화 되어있는 문서의 페이지 크기(수), 0부터 시작, 첫번째 페이지 제외함으로 1을 뺌

doc.pages[1].pageItems.everyItem().duplicate(new_doc.pages[0], section_cover); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_1st 좌표에 붙여넣기 함
for (var i = 2; i < page_range + 1; i++) {
  if (page_range == 5) {
    var mm_pt = 2.834645669291; //단위 mm -> pt 환산
    var pg_H = new_doc.documentPreferences.pageHeight; //페이지 높이 변수 설정
    new_doc.masterSpreads
      .itemByName("I-Inst")
      .pages[0].resize(
        BoundingBoxLimits.GEOMETRIC_PATH_BOUNDS,
        AnchorPoint.TOP_LEFT_ANCHOR,
        ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH,
        [mm_pt * 525, mm_pt * 297]
      );
    new_doc.masterSpreads
      .itemByName("K-Inst")
      .pages[0].resize(
        BoundingBoxLimits.GEOMETRIC_PATH_BOUNDS,
        AnchorPoint.TOP_LEFT_ANCHOR,
        ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH,
        [mm_pt * 525, mm_pt * 297]
      );
    new_doc.masterSpreads
      .itemByName("E-Inst")
      .pages[0].resize(
        BoundingBoxLimits.GEOMETRIC_PATH_BOUNDS,
        AnchorPoint.TOP_LEFT_ANCHOR,
        ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH,
        [mm_pt * 525, mm_pt * 297]
      );
    duplicate_data(5);
  } else {
    duplicate_data(4);
  }
}

try {
  new_doc.conditions.itemByName("취급설명서").visible = true;
} catch (e) {}
try {
  new_doc.conditions.itemByName("제품 개요서").visible = false;
} catch (e) {}
try {
  new_doc.conditions.itemByName("제품 매뉴얼").visible = false;
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

new_doc.pages.lastItem().appliedMaster =
  new_doc.masterSpreads.itemByName(language);

var mm_pt = 2.834645669291; //단위 mm -> pt 환산
var pg_H = new_doc.documentPreferences.pageHeight; //페이지 높이 변수 설정
if (
  (page_range != 0 && page_range < 3) ||
  (page_range == 0 && doc.pages.length < 4)
) {
  new_doc.pages[0].resize(
    BoundingBoxLimits.GEOMETRIC_PATH_BOUNDS,
    AnchorPoint.TOP_LEFT_ANCHOR,
    ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH,
    [mm_pt * 210, mm_pt * 297]
  );
  new_doc.pages[0].appliedMaster.pages[0].pageItems
    .everyItem()
    .move(new_doc.pages[0].appliedMaster.pages[0], [-210, 0]);
} else if (page_range == 5) {
  new_doc.pages[0].appliedMaster.pages[0].pageItems
    .everyItem()
    .move(new_doc.pages[0].appliedMaster.pages[0], [105, 0]);
}

thread_textframe();

PDF_error(new_doc);

function new_folder(doc_format) {
  if (create_folder == 1) {
    var lang = language == "K-Inst" ? "국문" : "영문";
    var temp_dir = export_subdir == 0 ? save_loc : doc.filePath;
    new Folder(temp_dir + "/_출력/" + doc_format + "/" + lang).create();
    var folder_output = String("/_출력/" + doc_format + "/" + lang + "/");
  } else {
    var folder_output = "";
  }
  return folder_output;
}

if (view == false) {
  app.pdfExportPresets.itemByName("TC_Note").pdfPageLayout =
    PageLayoutOptions.DEFAULT_VALUE; //기본 값 설정
  app.pdfExportPresets.itemByName("TC_Web").pdfPageLayout =
    PageLayoutOptions.DEFAULT_VALUE; //기본 값 설정
  app.pdfExportPresets.itemByName("TC_Print").pdfPageLayout =
    PageLayoutOptions.DEFAULT_VALUE; //기본 값 설정

  if (export_subdir == true) {
    save_loc = doc.filePath;
  }
  folder_dir = new_folder("취급 설명서");
  progress_bar.value = progress_count; //진행막대 채우기
  progress_text.text = String(
    "웹 파일 출력 중...   " + progress_count + " / " + file_count + " 완료"
  ); //진행상황 텍스트 표시
  loading_menu.active = true;

  def_PDFpref(0, 1);

  var name = new File(
    save_loc + "/" + folder_dir + doc_name + "_INST_W" + ".pdf"
  ); //새파일 만들기 (폴더 경로/파일명_INST.pdf)
  if (export_subdir == true) {
    name = new File(
      doc.filePath + "/" + folder_dir + doc_name + "_INST_W" + ".pdf"
    ); //새파일 만들기 (폴더 경로/파일명_INST.pdf)
  }
  new_doc.exportFile(
    ExportFormat.PDF_TYPE,
    name,
    false,
    app.pdfExportPresets.itemByName("TC_Web")
  ); //문서형 PDF로 출력

  progress_count++;

  progress_bar.value = progress_count; //진행막대 채우기
  progress_text.text = String(
    "발주 파일 출력 중...   " + progress_count + " / " + file_count + " 완료"
  ); //진행상황 텍스트 표시
  loading_menu.active = true;

  def_PDFpref(0, 0);

  var name = new File(
    save_loc + "/" + folder_dir + doc_name + "_INST_P" + ".pdf"
  ); //새파일 만들기 (폴더 경로/파일명_INST.pdf)
  if (export_subdir == true) {
    name = new File(
      doc.filePath + "/" + folder_dir + doc_name + "_INST_P" + ".pdf"
    ); //새파일 만들기 (폴더 경로/파일명_INST.pdf)
  }
  new_doc.exportFile(
    ExportFormat.PDF_TYPE,
    name,
    false,
    app.pdfExportPresets.itemByName("TC_Print")
  ); //문서형 PDF로 출력

  progress_count++;

  if (export_idml == 1) {
    var name = new File(
      save_loc + "/" + folder_dir + doc_name + "_INST" + ".idml"
    ); //새파일 만들기 (폴더 경로/파일명_IDML.idml)
    if (export_subdir == true) {
      name = new File(
        doc.filePath + "/" + folder_dir + doc_name + "_INST" + ".idml"
      ); //새파일 만들기 (폴더 경로/파일명_INST.pdf)
    }
    new_doc.exportFile(ExportFormat.INDESIGN_MARKUP, name, false); //IDML 출력
  }
  new_doc.close(SaveOptions.NO);
}

function duplicate_data(pg) {
  if (pg == 4) {
    if (i > pg && i % pg == 1) {
      //페이지 2, 4, 6, 8까지가 한 페이지로 구성 되고 그 이후로 새로운 페이지를 생성해야 하므로  8보다 크며 8로 나뉠때 2가 되면 새로운 문서에 새 페이지 추가
      new_doc.pages.add(); //새 페이지 추가
      doc.pages[i].pageItems
        .everyItem()
        .duplicate(new_doc.pages.lastItem(), section_1st); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_1st 좌표에 붙여넣기 함
    } else if (i % pg == 1) {
      //1번째 페이지일 경우
      doc.pages[i].pageItems
        .everyItem()
        .duplicate(new_doc.pages.lastItem(), section_1st); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_1st 좌표에 붙여넣기 함
    } else if (i % pg == 2) {
      //2번째 페이지일 경우
      doc.pages[i].pageItems
        .everyItem()
        .duplicate(new_doc.pages.lastItem(), section_2nd); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_2nd 좌표에 붙여넣기 함
    } else if (i % pg == 3) {
      //3번째 페이지일 경우
      doc.pages[i].pageItems
        .everyItem()
        .duplicate(new_doc.pages.lastItem(), section_3rd); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_3rd 좌표에 붙여넣기 함
    } else if (i % pg == 0) {
      //4번째 페이지일 경우
      doc.pages[i].pageItems
        .everyItem()
        .duplicate(new_doc.pages.lastItem(), section_4th); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_4th 좌표에 붙여넣기 함
    }
  } else {
    if (i > pg && i % pg == 1) {
      //페이지 2, 4, 6, 8까지가 한 페이지로 구성 되고 그 이후로 새로운 페이지를 생성해야 하므로  8보다 크며 8로 나뉠때 2가 되면 새로운 문서에 새 페이지 추가
      new_doc.pages.add(); //새 페이지 추가
      doc.pages[i].pageItems
        .everyItem()
        .duplicate(new_doc.pages.lastItem(), section_1st); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_1st 좌표에 붙여넣기 함
    } else if (i % pg == 1) {
      //1번째 페이지일 경우
      doc.pages[i].pageItems
        .everyItem()
        .duplicate(new_doc.pages.lastItem(), section_1st); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_1st 좌표에 붙여넣기 함
    } else if (i % pg == 2) {
      //2번째 페이지일 경우
      doc.pages[i].pageItems
        .everyItem()
        .duplicate(new_doc.pages.lastItem(), section_2nd); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_2nd 좌표에 붙여넣기 함
    } else if (i % pg == 3) {
      //3번째 페이지일 경우
      doc.pages[i].pageItems
        .everyItem()
        .duplicate(new_doc.pages.lastItem(), section_3rd); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_3rd 좌표에 붙여넣기 함
    } else if (i % pg == 4) {
      //4번째 페이지일 경우
      doc.pages[i].pageItems
        .everyItem()
        .duplicate(new_doc.pages.lastItem(), section_4th); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_3rd 좌표에 붙여넣기 함
    } else if (i % pg == 0) {
      //5번째 페이지일 경우
      doc.pages[i].pageItems
        .everyItem()
        .duplicate(new_doc.pages.lastItem(), section_5th); //활성화 되어있는 문서의 페이지에 해당하는 모든 개체를 새로운 문서의 section_4th 좌표에 붙여넣기 함
    }
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
