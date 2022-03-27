//이제 취설이나 매뉴얼 생성스크립트에 들어있어서 안쓰이는 스크립트에요 없애도 되요

var doc = app.activeDocument; //활성화 된 문서 설정
var Bookmark_win = new Window("dialog", "Autonics TC팀 북마크 스크립트"); //새로운 창 만들기
var para_style_list = doc.paragraphStyles.everyItem().name; //모든 단락스타일 불러오기
para_style_list.unshift("---------단락 스타일 선택---------");

var lists = []; //변수 설정
for (i = 0; i < para_style_list.length; i++) {
  //단락 스타일 수만큼 반복
  lists[i] = para_style_list[i]; //n번째 단락 스타일 list 변수에 넣기
}

with (Bookmark_win) {
  //창 설정
  preferredSize.width = 350; //폭 300pt
  alignChildren = "left"; //왼쪽 정렬

  var grep_Sq = add("radiobutton", undefined, "▣로 시작"); //▣로 시작 라디오 버튼 생성
  var para_st = add("radiobutton", undefined, "단락 스타일"); //단락스타일 라디오 버튼 생성
  para_st.value = true;
  var group_para = add("group"); //그룹 생성
  with (group_para) {
    //그룹 설정
    add("statictext", undefined, "지정할 스타일: "); //문자 추가
    var para_style = group_para.add("dropdownlist", undefined, undefined, {
      items: lists,
    }); //드롭다운 리스트 추가 (단락스타일 리스트)
    try {
      para_style.selection =
        doc.paragraphStyles.itemByName("Heading) Title").index + 1; //드롭다운 리스트 선택 값
    } catch (e) {}
    margins = [20, 0, 10, 0]; //여백 왼쪽 20pt 나머지 0pt
  }

  if (para_st.value == false) {
    //단락스타일이 선택되지 않을 경우
    group_para.enabled = false; //단락스타일 그룹 비활성화
  }

  para_st.addEventListener("mousedown", function (evt) {
    //단락스타일 클릭 시
    group_para.enabled = true; //활성화
  });

  grep_Sq.addEventListener("mousedown", function (evt) {
    //단락스타일 클릭 시
    group_para.enabled = false; //비활성화
  });

  margins = [20, 10, 0, 0]; //여백 왼쪽 위에 20pt, 오른쪽 아래 0pt

  var button_ctrl = add("group"); //그룹 생성
  with (button_ctrl) {
    //설정
    preferredSize.width = 300; //폭 300pt
    alignChildren = "left"; //왼쪽 정렬

    var help = add("button", undefined, "도움말"); //도움말 버튼 추가
    help.onClick = function () {
      //클릭 시
      alert(
        "------------  사용 방법  ------------ \n\n1. 추가하고자 하는 방식을 선택해 주세요.\n\n2. 선택이 완료되면 확인을 누르세요.\n   (문자 스타일의 경우 미리 지정해주세요.)\n\n 기타 문의는 말씀해 주세요! ",
        "도움말"
      );
    };

    add("statictext", [0, 0, 60, 20], ""); //공백 추가

    var ok_button = add("button", undefined, "확인", { name: "ok" }); //확인 버튼 추가
    var cancel_button = add("button", undefined, "취소", { name: "cancel" }); //취소 버튼 추가
    margins = [0, 20, 20, 20]; //여백 왼쪽 0pt 나머지 20pt
  }
}

ok_button.onClick = function () {
  //확인 버튼 선택 시
  if (para_st.value == true && para_style.selection == 0) {
    //단락 스타일 선택 후 스타일 지정 하지 않을 경우
    alert("단락 스타일을 선택해 주세요.", "경고"); //알림창
  } else {
    Bookmark_win.show(); //창 띄우기
  }
};

var Bookmark_Window = Bookmark_win.show(); //창 띄우기

if (Bookmark_Window == true) {
  //확인 버튼 선택 시
  doc.bookmarks.everyItem().remove();
  doc.hyperlinkTextDestinations.everyItem().remove();

  app.findGrepPreferences = app.changeGrepPreferences = null; //GREP 찾기 설정 초기화
  if (grep_Sq.value == true) {
    //▣로 시작 라디오 버튼 선택 시
    app.findGrepPreferences.findWhat = "▣.+";
  } else if (para_st.value == true) {
    //단락 스타일 라디오 버튼 선택 시
    app.findGrepPreferences.appliedParagraphStyle = para_style.selection.text; //GREP 찾기 단락스타일 지정
  }

  var found = doc.findGrep(); //GREP 찾기
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
      var Bookmark_dest = doc.hyperlinkTextDestinations.add(grep_fnd.texts[0], {
        name: Bookmark_name,
      }); //책갈피 목적지 설정
      var Bookmark_added = doc.bookmarks.add(Bookmark_dest, {
        name: Bookmark_name,
      }); //책갈피 추가
    } catch (e) {} //에러 넘기기
  }
  try {
    if (
      doc.bookmarks.itemByName("안전을 위한 주의 사항").index >
      doc.bookmarks.itemByName("취급 시 주의 사항").index
    ) {
      doc.bookmarks
        .itemByName("안전을 위한 주의 사항")
        .move(
          LocationOptions.BEFORE,
          doc.bookmarks.itemByName("취급 시 주의 사항")
        );
    }
  } catch (e) {}
  try {
    if (
      doc.bookmarks.itemByName("Safety Considerations").index >
      doc.bookmarks.itemByName("Cautions during Use").index
    ) {
      doc.bookmarks
        .itemByName("Safety Considerations")
        .move(
          LocationOptions.BEFORE,
          doc.bookmarks.itemByName("Cautions during Use")
        );
    }
  } catch (e) {}

  alert("북마크 추가를 완료하였습니다.", "Autonics TC"); //완료시 알림창

  var bookmarkPanel = app.panels.itemByName("$ID/Bookmarks"); //책갈피 패널 변수 설정
  bookmarkPanel.visible = true; //책갈피 보이기
}
app.findGrepPreferences = app.changeGrepPreferences = null;

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
