var myWin = new Window("dialog", "Autonics TC Scripts 글리프  복원"); //새창 생성
myWin.orientation = "row"; //세로 정렬
myWin.minimumSize.width = 300; //폭 크기 300pt

var sysFonts = app.fonts.everyItem().name; //폰트 이름 불러오기
sysFonts.unshift("글꼴을 선택해주세요."); //선택되지 않았을 경우 초기 문구

//폰트 이름 불러오기
var displayFonts = [];
var displayStyles = [];
var counter = -1;
for (var i = 0; i < sysFonts.length; i++) {
  sysFonts[i] = sysFonts[i].split("\t");
  if (i == 0 || sysFonts[i][0] != sysFonts[i - 1][0]) {
    displayFonts.push(sysFonts[i][0]);
    counter++;
    displayStyles[counter] = [];
  }
  displayStyles[counter].push(sysFonts[i][1]);
}

with (myWin) {
  //창 설정
  var g1 = myWin.add("group"); //그룹 생성
  with (g1) {
    //그룹 설정
    var sg1 = add("group"); //그룹 생성
    with (sg1) {
      //그룹 설정
      add("statictext", undefined, "복원하려는 문자 입력: "); //문자 추가
      var value = sg1.add("edittext", undefined, "", { multiline: false }); //문자 편집 상자 추가
      value.preferredSize.width = 150; //문자 편집상자 폭 150pt
      //valueTxt = value.text;
    }

    var sg2 = add("group"); //그룹 생성
    with (sg2) {
      //그룹 설정
      var charFt = sg2.add("statictext", undefined, "글꼴 지정: "); //문자 추가
      charFt.preferredSize.width = 80; //폭 80pt 설정
      var charFs = sg2.add("dropdownlist", undefined, undefined, {
        items: displayFonts,
      }); //드롭다운 리스트 추가
      charFs.preferredSize.width = 150; //리스트 폭 150pt
      charFs.selection = 0; //초기 1번째 설정
    }
    var sg3 = add("group"); //그룹 추가
    with (sg3) {
      //그룹 설정
      var charSt = sg3.add("statictext", undefined, "크기 지정: "); //문자 추가
      charSt.preferredSize.width = 80; //폭 80pt 설정
      var charSf = sg3.add("edittext", undefined, "", undefined); //문자 편집상자 추가
      charSf.preferredSize.width = 40; //폭 40pt 설정
      var charMsg = sg3.add("statictext", undefined, "pts"); //문자 추가
    }
    orientation = "column"; //방향 가로 설정
    alignChildren = "left"; //왼쪽 정렬
  }

  var g99 = myWin.add("group"); //그룹 추가
  //myWin.alignChildren = "right";
  g99.orientation = "column"; //방향 가로 설정
  g99.add("button", undefined, "확인", { name: "ok" }); //확인 버튼 추가
  g99.add("button", undefined, "취소", { name: "cancel" }); //취소 버튼 추가
}

var myWindow = myWin.show(); //창 띄우기

if (myWindow == true) {
  //확인 버튼 선택 시
  app.findGrepPreferences = app.changeGrepPreferences = null; //GREP 찾기 설정 초기화
  app.findGrepPreferences.findWhat = "~a"; //깨진 GREP 찾기
  //app.findGrepPreferences.fontStyle="bold";
  app.findGrepPreferences.pointSize = charSf.text; //문자 크기 설정
  app.changeGrepPreferences.changeTo = value.text; //변경하려는 문자 설정
  app.changeGrepPreferences.appliedFont = charFs.selection.text; //적용 글꼴 설정
  app.activeDocument.changeGrep(); //GREP 바꾸기

  var doc = app.activeDocument; //활성화 된 문서 변수 설정
  ResizeOverset(); //함수 실행

  app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
  alert("변환이 완료되었습니다.", "완료"); //완료 시 알림창
}

function ResizeOverset() {
  //함수 선언
  var lastFrame, //변수 설정
    stories = app.activeDocument.stories; //스토리 변수 설정
  for (var i = stories.length - 1; i >= 0; i--) {
    //스토리 수 만큼 반복
    while (stories[i].overflows == true) {
      //스토리 중 넘치는 텍스트가 있을 경우
      lastFrame =
        stories[i].texts[0].parentTextFrames[
          stories[i].texts[0].parentTextFrames.length - 1
        ]; //넘친 텍스트 프레임 변수 설정
      lastFrame.textFramePreferences.autoSizingReferencePoint =
        AutoSizingReferenceEnum.TOP_LEFT_POINT; //자동 크기조정 기준점 좌측 상단
      lastFrame.textFramePreferences.autoSizingType =
        AutoSizingTypeEnum.WIDTH_ONLY; //자동 크기 조정 폭만 설정
      lastFrame.fit(FitOptions.FRAME_TO_CONTENT); //내용에 맞게 수정
      //lastFrame.absoluteHorizontalScale+=1;
    }
  }
}
