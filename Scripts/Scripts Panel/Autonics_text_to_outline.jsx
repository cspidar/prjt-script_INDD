//현재 문제점으로는 아웃라인 진행중 텍스트 프레임이 넘칠 경우 에러가 발생하고 중단됩니다.
//만약 에러 발생 시 문제의 텍스트 프레임의 텍스트 넘치는 문제를 해결하고 다시 진행하시면 됩니다.

var myWin = new Window("dialog", "Autonics TC Scripts 글리프 깨기"); //새로운 창 만들기
myWin.orientation = "row"; //창 방향 세로 설정
myWin.minimumSize.width = 300; //창 폭 300pt

var sysFonts = app.fonts.everyItem().name; //모든 폰트 이름 가져오기
sysFonts.unshift("글꼴을 선택해주세요."); //초기 선택되지 않았을 때 문구 설정

var displayFonts = []; //변수 선언
var displayStyles = []; //변수 선언
var counter = -1; //카운터 선언
for (var i = 0; i < sysFonts.length; i++) {
  //폰트 리스트 작성
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
      add("statictext", undefined, "변경하려는 문자 입력: "); //문자 추가
      var value = sg1.add("edittext", undefined, "", { multiline: false }); //문자 편집상자 추가
      value.preferredSize.width = 150; //편집상자 폭 150pt
      valueTxt = value.text; //편집상자에 입력한 문자 읽기
    }

    var sg2 = add("group"); //그룹 생성
    with (sg2) {
      //그룹 설정
      var charFt = sg2.add("checkbox", undefined, "글꼴 지정: "); //글꼴 지정 체크박스 추가
      charFt.preferredSize.width = 80; //폭 설정 80pt
      var charFs = sg2.add("dropdownlist", undefined, undefined, {
        items: displayFonts,
      }); //드롭다운 리스트 추가 (폰트 목록)
      if (charFt.value == false) {
        //글꼴 지정 체크박스 해제 시
        charFs.enabled = false; //글꼴 지정 비활성화
      }
      charFt.addEventListener("mousedown", function (evt) {
        //글꼴지정 체크박스 선택 시
        charFs.enabled = !charFs.enabled; //활성화 반전
      });
      charFs.preferredSize.width = 150; //폭 설정 150pt
      charFs.selection = 0; //초기 글꼴 0번째 선택
    }
    var sg3 = add("group"); //그룹 생성
    with (sg3) {
      //그룹 설정
      var charSt = sg3.add("checkbox", undefined, "크기 지정: "); //크기 지정 체크박스 추가
      charSt.preferredSize.width = 80; //폭 설정 80pt
      var charSf = sg3.add("edittext", undefined, "", undefined); //문자 편집상자 추가
      if (charSt.value == false) {
        //크기 지정 체크박스 해제 시
        charSf.enabled = false; //크기 지정 비활성화
      }
      charSt.addEventListener("mousedown", function (evt) {
        //크기지정 체크박스 선택 시
        charSf.enabled = !charSf.enabled; //활성화 반전
      });
      charSf.preferredSize.width = 40; //폭 40pt
      var charMsg = sg3.add("statictext"); //문자 추가
      charMsg.text = "pts"; //문자 추가
    }
    orientation = "column"; //방향 가로 설정
    alignChildren = "left"; //왼쪽 정렬
  }

  var g99 = myWin.add("group"); //그룹 생성
  g99.orientation = "column"; //방향 가로 설정
  g99.add("button", undefined, "확인", { name: "ok" }); //확인 버튼 생성
  g99.add("button", undefined, "취소", { name: "cancel" }); //취소 버튼 생성
}

var myWindow = myWin.show(); //창 띄우기

if (myWindow == true) {
  //확인 버튼 선택 시

  if (value.text != "") {
    //편집상자가 공백이 아닐 경우
    app.findGrepPreferences = app.changeGrepPreferences = null; //GREP 찾기 설정 초기화
    app.findGrepPreferences.findWhat = value.text; //편집상자에 입력된 텍스트 찾기

    if (charFt.value == true) {
      //글꼴 지정 체크박스 선택 시
      app.findGrepPreferences.appliedFont = charFs.selection.text; //적용 폰트 지정
    }

    if (charSt.value == true) {
      //크기 지정 체크박스 선택 시
      app.findGrepPreferences.pointSize = charSf.text; //적용 크기 지정
    }

    var found = app.activeDocument.findGrep(); //GREP 찾기
    app.findGrepPreferences = app.changeGrepPreferences = null; //GREP 찾기 설정 초기화

    for (var i = 0; i < found.length; i++) {
      //찾은 GREP 수만큼 반복
      var myObj = found[i].createOutlines(); //n번째 찾은 GREP outline으로 깨기
    }
  }
  app.findGrepPreferences = app.changeGrepPreferences = null;
}
