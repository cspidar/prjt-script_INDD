//var Note_layer = app.documents.everyItem().layers.itemByName ("Note");

var _Font = "Autonics_TC [2020]"; //폰트 설정

var win_stats; //윈도우 상태 변수 선언
if (
  !app.documents.length ||
  (app.documents.length && !app.documents[0].visible)
) {
  //실행중인 창이 없을 경우
  win_stats = "batch"; //윈도우 상태: 일괄로 설정
  var title = "Autonics TC팀 PDF  일괄 출력 스크립트"; //창 제목 설정
} else {
  //실행중인 창이 있을 경우
  win_stats = "single"; //윈도우 상태: 개별로 설정
  var doc = app.activeDocument; //활성화 중인 문서 변수 설정
  var title = "Autonics TC팀 PDF  개별 출력 스크립트"; //창 제목 설정
}
var usePrefs = true; //설정 변수 선언 및  사용 설정

var tab_sel = 0; //탭 메뉴 초기 값 설정
var inst_note = 1; //취급설명서 주석 초기 값 설정
var inst_name = 1; //취급설명서 이름 초기 값 설정
var inst_web = 1; //취급설명서 웹 초기 값 설정
var inst_pw = 0; //취급설명서 웹 비밀번호 초기 값 설정
var inst_prnt = 1; //취급설명서 발주 초기 값 설정
var inst_idml = 0; //취급설명서 IDML 초기 값 설정
var cat_note = 1; //카탈로그 주석 초기 값 설정
var cat_name = 1; //카탈로그 이름 초기 값 설정
var cat_web = 1; //카탈로그 웹 초기 값 설정
var cat_pw = 0; //카탈로그 웹 비밀번호 초기 값 설정
var cat_prnt = 1; //카탈로그 발주 초기 값 설정
var cat_slug = 1; //카탈로그 슬러그 초기 값 설정
var cat_idml = 0; //카탈로그 IDML 초기 값 설정
var open_pdf = 1; //파일 열기 초기 값 설정

var prefsFile = File(
  Folder(app.activeScript).parent + "/Autonics_PDF_Export.ini"
); //설정 파일 저장 경로 지정
if (prefsFile.exists) {
  //설정 파일 이 있는 경우
  readPrefs(); //설정파일 읽기 함수 실행
}

function readPrefs() {
  //설정 파일 읽기 함수 설정
  prefsFile.open("r"); //설정 파일 열기 (읽기만 가능하게)
  tab_sel = Number(prefsFile.readln()); //탭 메뉴 값 읽기
  inst_note = Number(prefsFile.readln()); //취급설명서 주석 값 읽기
  inst_name = Number(prefsFile.readln()); //취급설명서 이름 값 읽기
  inst_web = Number(prefsFile.readln()); //취급설명서 웹 값 읽기
  inst_pw = Number(prefsFile.readln()); //취급설명서 웹 비밀번호 값 읽기
  inst_prnt = Number(prefsFile.readln()); //취급설명서 발주  값 읽기
  inst_idml = Number(prefsFile.readln()); //취급설명서 IDML 값 읽기
  cat_note = Number(prefsFile.readln()); //카탈로그 주석 값 읽기
  cat_name = Number(prefsFile.readln()); //카탈로그이름 값 읽기
  cat_web = Number(prefsFile.readln()); //카탈로그 웹 값 읽기
  cat_pw = Number(prefsFile.readln()); //카탈로그 웹 비밀번호 값 읽기
  cat_prnt = Number(prefsFile.readln()); //카탈로그 발주 값 읽기
  cat_slug = Number(prefsFile.readln()); //카탈로그 슬러그 값 읽기
  cat_idml = Number(prefsFile.readln()); //카탈로그 IDML 값 읽기
  open_pdf = Number(prefsFile.readln()); //파일열기 값 읽기
  prefsFile.close(); //설정 파일 파일 닫기
}

function savePrefs() {
  //설정파일 저장 함수 설정
  var newPrefs = //설정 값 변수 선언
    menu_1 +
    "\n" + //탭 메뉴 설정
    (note_i.value ? 1 : 0) +
    "\n" + //취급설명서 주석 값 설정
    (note_kor_1.value ? 1 : 0) +
    "\n" + //취급설명서 이름 값 설정
    (web_i.value ? 1 : 0) +
    "\n" + //취급설명서 웹 값 설정
    (web_pw_1.value ? 1 : 0) +
    "\n" + //취급설명서 비밀번호 값 설정
    (prnt_i.value ? 1 : 0) +
    "\n" + //취급설명서 발주 값 설정
    (idml_i.value ? 1 : 0) +
    "\n" + //취급설명서 IDML 값 설정
    (note.value ? 1 : 0) +
    "\n" + //카탈로그 주석 값 설정
    (note_kor_2.value ? 1 : 0) +
    "\n" + //카탈로그 이름 값 설정
    (web.value ? 1 : 0) +
    "\n" + //카탈로그 웹 값 설정
    (web_pw_2.value ? 1 : 0) +
    "\n" + //카탈로그 비밀번호 값 설정
    (prnt.value ? 1 : 0) +
    "\n" + //카탈로그 발주 값 설정
    menu_2 +
    "\n" + //카탈로그 슬러그 값 설정
    (idml.selection ? 1 : 0) +
    "\n" + //카탈로그 IDML 값 설정
    (prvw.value ? 1 : 0); //파일 열기 값 설정
  prefsFile.open("w"); //설정 파일 열기 (쓰기 가능)
  prefsFile.write(newPrefs); //설정파일에 설정 값 쓰기
  prefsFile.close(); //설정 파일 닫기
}

var PDF_Win = new Window("dialog", title); //새로운 창 생성

with (PDF_Win) {
  //창 설정
  preferredSize.width = 300; //너비 300pt
  alignChildren = "left"; //왼쪽 정렬
  var tabmenu = add("tabbedpanel"); //탭 메뉴 추가
  with (tabmenu) {
    //탭 메뉴 설정
    preferredSize = [350, 200]; //크기 350*200
    var inst = add("tab", undefined, "취급설명서"); //취급설명서 탭 추가
    with (inst) {
      //취급설명서 탭 설정
      alignChildren = "fill"; //정렬 채우기
      var inst_format = add("panel", undefined, "출력 문서"); //패널 추가
      with (inst_format) {
        //패널 설정
        alignChildren = "left"; //왼쪽 정렬
        preferredSize = [280, 100]; //크기 280*100
        var note_i = add("checkbox", undefined, "주석 파일"); //주석 체크박스 추가
        note_i.value = inst_note; //체크박스 설정파일 값 지정
        var name_group_i = add("group"); //그룹 추가
        var note_kor_1 = name_group_i.add(
          "radiobutton",
          undefined,
          "파일명_주석.pdf"
        ); //_주석 라디오 버튼 추가
        var note_eng_1 = name_group_i.add(
          "radiobutton",
          undefined,
          "파일명_NOTE.pdf"
        ); //_NOTE 라디오 버튼 추가
        note_kor_1.value = inst_name; //라디오버튼 설정파일 값 지정
        note_eng_1.value = !note_kor_1.value; //라디오버튼 서로 선택 반전
        name_group_i.margins = [20, 0, 0, 0]; //그룹 여백 설정 (왼쪽: 20pt, 나머지 0pt)
        if (note_i.value == false) {
          //주석 체크박스 해제 시
          name_group_i.enabled = false; //라디오 버튼 비활성화
        }
        note_i.addEventListener("mousedown", function (evt) {
          //주석 체크박스 클릭 시
          name_group_i.enabled = !name_group_i.enabled; //활성화 반전
        });
        var web_i = add("checkbox", undefined, "웹 파일"); //웹파일 체크박스 추가
        web_i.value = inst_web; //체크박스 설정파일 값 지정
        var web_group = add("group"); //그룹 추가
        var web_pw_1 = web_group.add("checkbox", undefined, "PDF 암호화"); //PDF 암호화 체크박스 추가
        web_pw_1.value = inst_pw; //체크박스 설정파일 값 지정
        web_group.margins = [20, 0, 0, 0]; //그룹 여백 설정 (왼쪽: 20pt, 나머지 0pt)
        if (web_i.value == false) {
          //웹 체크박스 해제 시
          web_pw_1.enabled = false; //비밀번호 비활성화
        }
        web_i.addEventListener("mousedown", function (evt) {
          //웹 체크박스 클릭 시
          web_pw_1.enabled = !web_pw_1.enabled; //활성화 반전
        });
        var prnt_i = add("checkbox", undefined, "발주 파일"); //발주파일 체크박스 추가
        prnt_i.value = inst_prnt; //체크박스 설정파일 값 지정
        var idml_i = add("checkbox", undefined, "IDML 파일"); //IDML 체크박스 추가
        idml_i.value = inst_idml; //체크박스 설정파일 값 지정
        margins = [15, 15, 0, 10]; //여백 설정 (왼쪽: 15pt, 위쪽: 15pt, 오른쪽: 0pt, 아래쪽: 10pt)
      }
    }
    var catal = add("tab", undefined, "카탈로그"); //카탈로그 탭 추가 (취급설명서 탭과 유사 주석 생략)
    with (catal) {
      alignChildren = "fill";
      var catal_format = add("panel", undefined, "출력 문서");
      with (catal_format) {
        alignChildren = "left";
        preferredSize = [280, 100];
        var note = add("checkbox", undefined, "주석 파일");
        note.value = cat_note;
        var name_group = add("group");
        var note_kor_2 = name_group.add(
          "radiobutton",
          undefined,
          "파일명_주석.pdf"
        );
        var note_eng_2 = name_group.add(
          "radiobutton",
          undefined,
          "파일명_NOTE.pdf"
        );
        note_kor_2.value = cat_name;
        note_eng_2.value = !note_kor_2.value;
        name_group.margins = [20, 0, 0, 0];
        if (note.value == false) {
          name_group.enabled = false;
        }
        note.addEventListener("mousedown", function (evt) {
          name_group.enabled = !name_group.enabled;
        });
        var web = add("checkbox", undefined, "웹 파일");
        web.value = cat_web;
        var web_group = add("group");
        var web_pw_2 = web_group.add("checkbox", undefined, "PDF 암호화");
        web_pw_2.value = cat_pw;
        web_group.margins = [20, 0, 0, 0];
        if (web.value == false) {
          web_pw_2.enabled = false;
        }
        web.addEventListener("mousedown", function (evt) {
          web_pw_2.enabled = !web_pw_2.enabled;
        });
        var prnt = add("checkbox", undefined, "발주 파일");
        prnt.value = cat_prnt;
        var prnt_view = add("group");
        prnt_view.add("statictext", undefined, "발주 파일 설정:");
        var prnt_sub = prnt_view.add("dropdownlist", undefined, [
          "도련 및 슬러그 없음",
          "도련 및 슬러그 포함",
        ]);
        prnt_sub.selection = cat_slug;
        prnt_view.margins = [20, 0, 0, 0];
        if (prnt.value == false) {
          prnt_sub.enabled = false;
        }
        prnt.addEventListener("mousedown", function (evt) {
          prnt_sub.enabled = !prnt_sub.enabled;
        });
        var idml = add("checkbox", undefined, "IDML 파일");
        idml.selection = cat_idml;
        margins = [15, 15, 0, 10];
      }
    }
  }

  var g3 = add("group"); //그룹 추가
  g3.add("statictext", undefined, "저장 경로:"); //문자 추가
  var outfolder = g3.add("edittext", [0, 0, 240, 20], undefined, {
    readonly: true,
  }); //문자 입력 상자 추가 (크기 240*20, 읽기만 가능)
  var setdir_button = g3.add("iconbutton", undefined, folder_icon(), {
    style: "toolbutton",
  }); //아이콘 버튼 추가 327번째 줄 참고

  if (win_stats == "batch") {
    //일괄 출력 설정 시
    outfolder.text = "경로를 설정해 주세요."; //문자 입력 상자 초기 텍스트 설정
    var save_loc = Folder(""); //폴더 설정
    var load_files; //변수 설정
    setdir_button.onClick = function () {
      //아이콘 클릭 시
      save_loc = Folder("").selectDlg("저장하려는 위치를 지정해 주세요."); //폴더 선택 창 실행
      outfolder.text = save_loc.fullName; //선택된 폴더 경로 문자입력 상자에 표시
      load_files = save_loc.getFiles("*.indd"); //선택 한 폴더의 indd파일 전체 추가
    };
  }

  if (win_stats == "single") {
    //개별 출력 설정 시
    try {
      if (doc.label != "") {
        var file_loc = app.documents.itemByName(doc.label).filePath;
      } else {
        var file_loc = doc.filePath;
      }
    } catch (e) {}
    if (doc.saved == true || doc.label != "") {
      //파일이 저장되어있을 경우
      outfolder.text = folder_dir().fullName; //폴더 함수를 이용하여 현재 파일의 경로 표시
      var save_loc = Folder(file_loc); //저장 폴더 설정
      setdir_button.onClick = function () {
        //아이콘 클릭 시
        save_loc =
          Folder(file_loc).selectDlg("저장하려는 위치를 지정해 주세요."); //폴더 선택 창 실행
        outfolder.text = save_loc.fullName; //선택한 폴더로 경로 문자입력 상자에 표시
      };
    } else {
      //파일이 저장되어 있지 않을 경우
      outfolder.text = "경로를 설정해 주세요.";
      var save_loc = Folder("");
      setdir_button.onClick = function () {
        save_loc = Folder("").selectDlg("저장하려는 위치를 지정해 주세요.");
        outfolder.text = save_loc.fullName;
      };
    }
    function folder_dir() {
      //폴더 함수 설정
      var path = Folder(file_loc); //현재 활성화 중인 파일의 경로
      if (path == "") {
        //경로가 공백일 경우
        path = Folder(save_loc.fullName); //현재 활성화 중인 파일의 경로로 설정
      }
      return path; //경로 값 반환
    }
  }

  var g4 = add("group"); //그룹 생성
  with (g4) {
    //그룹 설정
    var ovrw = add("checkbox", undefined, "덮어 쓰기 (&V)");
    ovrw.shortcutKey = "v"; //덮어 쓰기 체크박스 생성, 단축키 지정
    var prvw = add("checkbox", undefined, "완료 후 파일 열기 (&O)");
    prvw.shortcutKey = "o"; //파일 열기 체크 박스 생성, 단축키 지정
    prvw.value = open_pdf; //체크박스 설정파일 값 지정
    orientation = "row"; //세로방향
    alignChildren = "left"; //왼쪽 정렬
  }

  var g99 = add("group{alignment: 'left'}"); //그룹 생성
  with (g99) {
    //그룹 설정
    var help = add("button", undefined, "도움말"); //도움말 버튼 생성
    help.onClick = function () {
      //버튼 클릭시
      alert(
        "------------      사    용     방    법      ------------ \n\n1. 취급설명서 또는 카탈로그 탭을 선택하여 주세요. \n\n2. 출력하려는 형식을 선택해 주세요.\n   (카탈로그 경우 웹파일의 보기 설정과 발주파일의 \n    도련 및 슬러그 설정이 추가됩니다.)\n\n3. 기타 설정을 선택해 주세요.\n   (덮어쓰기 설정을 하지 않으면 파일 명 뒤에 숫자가 \n    추가됩니다.) \n\n 기타 문의는 말씀해 주세요! ",
        "도움말"
      );
    };
    add("statictext", [0, 0, 80, 20], ""); //공란 추가
    var ok_button = add("button", undefined, "확인", { name: "ok" }); //확인 버튼 생성
    var cancel_button = add("button", undefined, "취소", { name: "cancel" }); //취소 버튼 생성
  }
  var menu_1 = tab_sel ? 1 : 0; //변수에 설정파일 값 지정
  tabmenu.selection = menu_1; //탭 메뉴 변수 값 지정
  var menu_2 = cat_slug ? 1 : 0; //변수에 설정파일 값 지정
  prnt_sub.selection = menu_2; //슬러그 설정에 변수 값 지정
}

ok_button.onClick = function () {
  //확인 버튼 클릭 시
  if (outfolder.text == "경로를 설정해 주세요.") {
    //경로 폴더 지정되있지 않을 경우
    alert("경로를 설정해 주세요.", "경고"); //알림
  } else if (
    (tabmenu.selection == inst &&
      note_i.value == false &&
      web_i.value == false &&
      prnt_i.value == false &&
      idml_i.value == false) ||
    (tabmenu.selection == catal &&
      note.value == false &&
      web.value == false &&
      prnt.value == false &&
      idml.value == false)
  ) {
    //아무것도 선택되지 않을 경우
    alert("출력 형식을 선택해 주세요.", "경고"); //알림
  } else {
    //그 외
    PDF_Win.show(); //창 띄우기
  }
};

var PDF_Window = PDF_Win.show(); //창 띄우기

if (PDF_Window == true) {
  //확인 버튼 누를 경우
  menu_1 = tabmenu.selection == inst ? 0 : 1; //탭 메뉴 값 변수에 지정
  menu_2 = prnt_sub.selection == 0 ? 0 : 1; //슬러그 값 변수에 지정
  savePrefs(); //설정 값 저장 함수 실행
  var loadingmenu = new Window("palette"); //로딩창 생성
  if (win_stats == "single") {
    //개별 실행일 경우
    var loadingtext = loadingmenu.add(
      "statictext",
      undefined,
      doc.name + " 출력 중..."
    ); //출력 텍스트 띄우기
  }
  loadingmenu.preferredSize = [350, 50]; //크기:350*50

  if (win_stats == "batch") {
    //일괄일 경우
    var count_files = 1; //초기 값 1 설정
    var loadingtext = loadingmenu.add("statictext", undefined, "출력 시작"); //로딩창 생성
    loadingtext.alignment = "left"; //왼쪽 정렬
    loadingtext.bounds = [0, 0, 340, 20]; //창 크기 340*20
    var progressbar = loadingmenu.add(
      "progressbar",
      [12, 12, 350, 24],
      0,
      load_files.length
    ); //진행막대 추가
    var progresstext = loadingmenu.add("statictext", undefined, "출력 시작..."); //문자 추가
    progresstext.bounds = [0, 0, 75, 20]; //문자 크기 설정
    progresstext.alignment = "right"; //우측 정렬
  }
  loadingmenu.add("statictext", undefined, "esc를 누르면 취소됩니다."); //로딩창 문자 추가

  loadingmenu.show(); //로딩창 띄우기

  var key_stats = ScriptUI.environment.keyboardState; //키보드 설정
  if (key_stats.keyName == "Escape") {
    //ESC를 누를 경우
    doc.close(SaveOptions.no); //열려있는 파일 닫기 (저장 안함)
    loadingmenu.close(); //로딩 창 닫기
  }

  var error_exist; //에러 변수 설정

  if (win_stats == "batch") {
    //일괄일 경우
    for (var i = 0; i < load_files.length; i++) {
      //지정 한 폴더 안에 indd파일 수 만큼 반복
      var doc = app.open(load_files[i], false); //파일 열기 (보이기 해제)
      var Note_layer = doc.layers.itemByName("Note"); //노트 레이어 변수 설정
      PDF_error(); //PDF 에러 함수 실행
      transGuide(); //투명상자 함수 실행
      if (error_exist != true) {
        //에러가 없을 경우
        var myFileName = doc.name.substr(0, doc.name.lastIndexOf(".")); //파일명 1번째부터 .위치 전까지 설정
        var myFolder = Folder(doc.filePath); //폴더 경로 설정
        loadingtext.text = String(doc.name + " 변환 중... "); //로딩창 띄우기
        progressbar.value = count_files; //진행막대 채우기
        progresstext.text = String(
          count_files + " / " + load_files.length + " 완료"
        ); //진행상황 텍스트 표시

        PDF_out(); //PDF 출력 함수 실행
      }
      doc.close(SaveOptions.yes); //파일 종료 (저장 후)
      count_files++; //진행 막대 숫자 증가
    }
    loadingmenu.close(); //로딩 창 끄기
    alert("PDF 출력을 완료하였습니다.", "Autonics TC"); //알림
    doc = null; //doc 변수 초기화
  } else {
    //개별일 경우
    var Note_layer = doc.layers.itemByName("Note"); //노트 레이어 변수 설정
    PDF_error(); //PDF 에러 함수 실행
    transGuide(); //투명상자 함수 실행
    if (error_exist != true) {
      //에러가 없을 경우
      if (doc.saved == true) {
        var myFileName = doc.name.substr(0, doc.name.lastIndexOf(".")); //파일명 1번째에서 . 위치 전까지 설정
      } else {
        var myFileName = doc.name; //파일명 1번째에서 . 위치 전까지 설정
      }
      var myFolder = save_loc; //폴더 경로 설정

      PDF_out(); //PDF 출력함수 실행
      loadingmenu.close(); //로딩창 끄기
      alert("PDF 출력을 완료하였습니다.", "Autonics TC"); //알림
    }
  }
}

function folder_icon() {
  //폴더 아이콘 설정
  return "\u0089PNG\r\n\x1A\n\x00\x00\x00\rIHDR\x00\x00\x00\x16\x00\x00\x00\x12\b\x06\x00\x00\x00_%.-\x00\x00\x00\tpHYs\x00\x00\x0B\x13\x00\x00\x0B\x13\x01\x00\u009A\u009C\x18\x00\x00\x00\x04gAMA\x00\x00\u00B1\u008E|\u00FBQ\u0093\x00\x00\x00 cHRM\x00\x00z%\x00\x00\u0080\u0083\x00\x00\u00F9\u00FF\x00\x00\u0080\u00E9\x00\x00u0\x00\x00\u00EA`\x00\x00:\u0098\x00\x00\x17o\u0092_\u00C5F\x00\x00\x02\u00DEIDATx\u00DAb\u00FC\u00FF\u00FF?\x03-\x00@\x0011\u00D0\b\x00\x04\x10\u00CD\f\x06\b \x16\x18CFR\x12L\u00CF*\u0092e\u00FE\u00F7\u009F!\u008C\u0097\u008By\x19\u0088\u00FF\u00F7\u00EF\x7F\u0086\u00CF\u00DF\u00FE\u00C6dOz\u00B2\x1C\u00C8\u00FD\x0F\u00C5\x04\x01@\x00\u00A1\u00B8\x18f(##C\u00AD\u009Ak9\u0083\u008E_\x17\u0083i\u00D4<\x06\x16f\u00C6\u009A\t\u00D9\u00D21@%\u00CC@\u00CCH\u008C\u00C1\x00\x01\u00C4\b\u008B<\u0090\u008Bg\x14\u00CAF212,\u00D3q\u00CDb\u00E0\x16Rf`\u00E3\x14f`\u00E5\x14d\u00F8\u00FF\u00E7'\u00C3\u00FE\u00D9a\x18\u009A\u00FF\u00FE\u00FB\u009Fq\u00F3\u00F1\u00CF%\x13\u00D6\u00BE\u00FE\u0086\u00EE\x13\u0080\x00bA\u00B6\x04d\u00A8\u00A1_\x15\u00D8@\u0098\u00A1\u00AC\u00EC\u00FC\f\u00CC<\\\f^\u00A5\u00A7P\f\u00FD\u00F6\u00EE.\u00C3\u00DD\x03\x1D3\u00BE\u00FF<\u00FF\f\u00C8\u00DD\x01\u00C4\x7F\u0090\r\x07\b \x14\u0083A\x04\u00CCP6\x0E!\u0086\u00A3s\x03\x18XY\x19\x19\u00FE\x01\u00C3\x07\x14\u00D6\x7F\u00A1\u00F4\u009F\u00BF\f`\fb\x03}\u00BC\u00A9+U\u0092\u00E1\u00F9\u009B\u00BF\u00BA\u00FD\u00EB_]\u0083\u00C5\x03@\x00\u00B1\u00A0\u00877\u00CC\u00A5\u00F7\x0F\u00F72\u00C8\x1B\x052p\n(\u0080\u00A5\u00FE\u00FD\u00F9\u00C5\u00F0\u00F7\u00F7o\u0086?\u00BF\x7F1\u00FC\u00F9\x05\u00A1\u00FF\u00FE\u00F9\r\u00C6\u009F\u009E_\x00\u00C6\u00C3\u00FDI@\u0085^@\u00FC\x1B\x14J\x00\x01\u00C4\u0084\u00EEb\u0090\u00A1\u00BF>\u00BFd\u00F8\u00FC\u00EA:\x03\u00A7\u00A0\"\u00C3\u00BF\u00BF\u00BF\x19\u00FE\u00FF\u00FD\x034\u00F8\x0F\u00D8\u0090\x7F\u00BFAl \u00FD\u00EF/P\u00EE\x0FX\u00FE\u00C0\u00B1+\f\u008F^\u00FD<\b\u00D4\u00CE\x01\u008B`\u0080\x00\u00C2\b\n\x0E\x1EI\u0086\u009B\u00DB\u00CA\x19\u0084\u0094\u00EC\u0081\u0081\u00CE\u00CA\u00C0\u00C4\x04\u00F4\u00FE\u00AF_`\u0083A\u0086\u0082]\u00F9\x17j8\u0090\u00FE\u00F1\u00E9)\u00C3\u00D6\x13/\x19\u00EE\u00BFa\u00D8\u00C2\u00CE\u00C6\u00CE\n5\u00F8\x0F@\x00ad\u0090W7\u00B60\u00FC\u00FB\u00FF\u0087\u0081KX\x05\u00E8\u00D2\u00DF`\x03\u00FE\u0082]\x0Bq\u00DD\u00BF\u00BF0\u0097\u00FE\x05\u0086\u00EF_\u0086\u00C3G\u008E1\u00DCy\u00FE}9\u00D0\u00D0O\u00C8I\x11 \u00800\f~xr\x06\u0083\u00A0\u00825\u00C3\u00FF\x7FPW\x01\r\x04Y\x00q\u00E9_ \u0086\x1A\x0E\u0094\u00FF\t\f\u00B2\u0095\u00FB\u009F20\u00B3p\u00CC\u0082\u00A6\n\x10\u00FE\x07\u008A<\u0080\x00\u00C20\u0098\u009DO\u0082\u0081\u009DG\x02\x12\u00AE@\u00CD \u0083\u00C0^\x07bP\u00E4\u00FD\u0083\x1A\u00FE\x1F\u00E8\u00ABS'\u008F2\u00DC{\u00FE}\x1D;;\u00C7\x0B\u00A0\u00D6\u009F@\u00FC\x0B\x14q \u0083\x01\x02\u0088\x05\u00C5P6&\u0086\u00F6i\u00DB\x18^\u00BE[\x0FNJ\u00BF\u00FF\u00FCc\x00&\x00\u0086\u00DF\u00BF!l`\x10\x03\u0093\u00D9\x7F0\u00FE\x0B\u00CCX\u00DF\x7F\u00FEe`e\u00E3\u009C\t5\u00F0'\u0092\u008B\x19\x00\x02\b9\u00E7\u0081\x02\u009E\x0B\u0088\u00F9\u00A14+\x119\u00F7\x1F\u00D4\u00D0/P\u00FC\x1Dj8\x03@\x00!\u00BB\u00F8?T\u00F0'\u0096\u00CCC\u00C8\u00E0\u00EFP\u00FA\x1FL\x02 \u0080X\u00D0\x14\u00FD\u0086\u00DA\u00FC\u0083\u00C8\"\x15\u00E6\u0098\u00DF\u00C8\u00C1\x00\x02\x00\x01\x06\x000\u00B2{\u009A\u00B3\x1C#o\x00\x00\x00\x00IEND\u00AEB`\u0082";
}

function def_PDFpref() {
  //PDF 초기값 함수 설정
  with (app.pdfExportPreferences) {
    //PDF 출력 설정
    if (
      tabmenu.selection == catal &&
      doc.pages[0].side == PageSideOptions.RIGHT_HAND
    ) {
      //카탈로그 탭이고 첫 스프레드 페이지가 홀수 일 경우
      pdfPageLayout = PageLayoutOptions.TWO_UP_COVER_PAGE_CONTINUOUS; //연속 두페이지 표지 페이지 설정
    }
    if (
      tabmenu.selection == catal &&
      doc.pages[0].side == PageSideOptions.LEFT_HAND
    ) {
      //카탈로그 탭이고 첫 스프레드 페이지가 짝수 일 경우
      pdfPageLayout = PageLayoutOptions.TWO_UP_FACING_CONTINUOUS; //연속 두페이지 설정
    }
    if (tabmenu.selection == inst) {
      //취설 탭일 경우
      pdfPageLayout = PageLayoutOptions.DEFAULT_VALUE; //기본값 설정
    }
    pdfMagnification = PdfMagnificationOptions.FIT_PAGE; //PDF확대: 페이지 맞추기
    acrobatCompatibility = AcrobatCompatibility.acrobat5; //ACROBAT 5에 호환
    standardsCompliance = PDFXStandards.NONE; //PDFX규격 없음

    pageRange = PageRange.allPages; //출력범위 모든 페이지
    if (Note_layer.isValid == true) {
      //노트 레이어가 있을 경우
      var page_allow = doc.pages[0].name; //페이지 범위 변수에 페이지 1 지정
      for (var i = 1; i < doc.pages.length; i++) {
        //페이지 수만큼 반복
        if (doc.pages[i].bounds[3] - doc.pages[i].bounds[1] > 90) {
          //페이지 폭이 90mm보다 클 경우
          page_allow += ", " + doc.pages[i].name; //페이지 범위 변수에 90mm보다 큰 페이지 추가
        }
      }
      pageRange = page_allow; //페이지 범위: 페이지 범위 변수에 지정한 값
    }

    viewPDF = false; //PDF 보기 설정 해제
    generateThumbnails = false; //썸네일 보기 해제
    optimizePDF = true; //PDF최적화 설정
    exportLayers = false; //레이어 출력 해제
    exportWhichLayers = ExportLayerOptions.EXPORT_VISIBLE_PRINTABLE_LAYERS; //보이는 레이어만 출력 설정
    exportNonPrintingObjects = false; //인쇄되지 않는 개체 출력 해제
    exportReaderSpreads = false; //개별 스프레드 별로 출력 해제
    exportGuidesAndGrids = false; //안내선 및 격자 출력 해제
    includeBookmarks = false; //책갈피 추가 해제
    includeHyperlinks = false; //하이퍼링크 추가 해제
    includeICCProfiles = ICCProfiles.INCLUDE_TAGGED; //ICC 프로필 태그만 포함
    interactiveElementsOption = InteractiveElementsOptions.DO_NOT_INCLUDE; //대화형 개체 미포함

    colorBitmapCompression = BitmapCompression.AUTO_COMPRESSION; //컬러 이미지 압축 자동
    colorBitmapSampling = Sampling.BICUBIC_DOWNSAMPLE; //바이큐빅 다운샘플링 설정
    colorBitmapSamplingDPI = 300; //300인치당 픽셀 수
    colorBitmapQuality = CompressionQuality.MAXIMUM; //이미지 품질 최대

    grayscaleBitmapCompression = BitmapCompression.AUTO_COMPRESSION; //회색 음영 이미지 압축 자동
    grayscaleBitmapQuality = CompressionQuality.maximum; //이미지 품질 최대
    grayscaleBitmapSampling = Sampling.BICUBIC_DOWNSAMPLE; //바이큐백 다운샘플링
    grayscaleBitmapSamplingDPI = 300; //300인치당 픽셀 수

    monochromeBitmapCompression = MonoBitmapCompression.CCIT4; //단색 이미지 압축 CCITT그룹 4
    monochromeBitmapSampling = Sampling.BICUBIC_DOWNSAMPLE; //바이큐빅 다운샘플링
    monochromeBitmapSamplingDPI = 1200; //1200인치당 픽셀 수

    compressTextAndLineArt = true; //텍스트와 라인 아트 압축
    cropImagesToFrames = true; //프레임에 맞게 이미지 데이터 자르기

    cropMarks = false; //재단선 표시 해제
    includeStructure = false; //맞춰찍기 표시 해제
    colorBars = false; //색상막대 표시 해제
    pageInformationMarks = false; //페이지 정보 표시 해제
    bleedMarks = false; //도련 표시 해제
    includeSlugWithPDF = false; //슬러그 표시 해제
    registrationMarks = false; //인증마크 해제
    subsetFontsBelow = 100; //사용된 문자의 비율이 100%보다 작을 경우 하위 세트 글꼴

    bleedInside = "0mm"; //도련 안쪽 0mm
    bleedOutside = "0mm"; //도련 바깥쪽 0mm

    useDocumentBleedWithPDF = false; //문서 도련 설정 사용 해제
    if (prvw.value == 1) {
      //파일열기 체크박스 설정 시
      viewPDF = true; //PDF 열기
    }
  }
}

function notePDF() {
  //주석 함수 설정
  def_PDFpref(); //PDF 초기값 함수 실행
  if (Note_layer.isValid == true) {
    //노트 레이어가 있을 경우
    with (app.interactivePDFExportPreferences) {
      //대화형 PDF 출력 설정
      exportReaderSpreads = true; //스프레드 별로 출력 설정
      pdfPageLayout = PageLayoutOptions.DEFAULT_VALUE; //페이지 레이아웃 기본 값 설정
      pdfMagnification = PdfMagnificationOptions.FIT_PAGE; //페이지 맞추기 설정
    }
  }
  var out_name = "_주석"; //이름 변수 _주석 설정
  if (
    (tabmenu.selection == inst && note_eng_1.value == true) ||
    (tabmenu.selection == catal && note_eng_2.value == true)
  ) {
    //영문 라디오 버튼이 선택되있을 경우
    out_name = "_NOTE"; //이름 변수 _NOTE 설정
  }
  name = new File(myFolder + "/" + myFileName + out_name + ".pdf"); //새파일 만들기 (폴더 경로/파일명_주석(NOTE).pdf)
  var counter = 2; //카운터 변수 설정
  if (ovrw.value !== true) {
    //덮어쓰기가 해제되있을 경우
    while (name.exists) {
      //파일 이 존재하면
      name = new File(
        myFolder + "/" + myFileName + out_name + "_" + counter + ".pdf"
      ); //파일명 뒤에 카운터 추가
      counter++; //카운터 증가
    }
  }
  if (Note_layer.isValid == true) {
    //노트 레이어 있을 경우
    doc.exportFile(ExportFormat.INTERACTIVE_PDF, name, false); //대화형 PDF로 출력
  } else {
    //노트 레이어 없을 경우
    doc.exportFile(ExportFormat.PDF_TYPE, name, false); //문서형 PDF로 출력
  }
}

function webPDF() {
  //웹 함수 설정
  def_PDFpref(); //PDF 초기값 함수 실행
  with (app.pdfExportPreferences) {
    //PDF 출력 설정
    acrobatCompatibility = AcrobatCompatibility.acrobat6; //Acrobat 6 호환성
    includeBookmarks = true; //책갈피 추가
    includeHyperlinks = true; //하이퍼링크 추가
    if (
      (tabmenu.selection == inst && web_pw_1.value == true) ||
      (tabmenu.selection == catal && web_pw_2.value == true)
    ) {
      //비밀번호 체크박스 체크시
      useSecurity = true; //암호 설정
      changeSecurityPassword = "auto1234"; //비밀번호
      disallowChanging = true; //수정 금지 설정
      disallowCopying = true; //복사 금지 설정
      disallowExtractionForAccessibility = true; //추출 금지 설정
    }
  }
  name = new File(myFolder + "/" + myFileName + "_W" + ".pdf"); //파일명_W.pdf 설정
  doc.exportFile(ExportFormat.PDF_TYPE, name, false); //PDF 출력
}

function prntPDF() {
  //발주 함수 설정
  def_PDFpref(); //PDF 초기값 함수 실행
  with (app.pdfExportPreferences) {
    //PDF 출력 설정
    acrobatCompatibility = AcrobatCompatibility.acrobat4; //Acrobat 4 호환성

    colorBitmapSampling = Sampling.NONE; //컬러 이미지 다운샘플링 안함
    grayscaleBitmapSampling = Sampling.NONE; //회색 음영 이미지 다운샘플링 안함
    monochromeBitmapSampling = Sampling.NONE; //단색 이미지 다운샘플링 안함

    appliedFlattenerPreset.convertAllStrokesToOutlines = true; //모든 외곽선 깨기
    appliedFlattenerPreset.convertAllTextToOutlines = true; //모든 텍스트 깨기
    appliedFlattenerPreset.gradientAndMeshResolution = 300; //그래디언트 및 메시 해상도 300
    appliedFlattenerPreset.lineArtAndTextResolution = 1200; //라인 아트 및 텍스트 해상도 1200
    appliedFlattenerPreset.rasterVectorBalance = 100; //래스터/벡터 밸런스 100

    if (tabmenu.selection == catal && prnt_sub.selection == 1) {
      //카탈로그 탭 설정, 도련 및 슬러그 드롭다운 리스트 체크 시
      cropMarks = true; //재단 선 보이기
      bleedMarks = true; //도련 보이기
      registrationMarks = true; //인증마크 보이기
      colorBars = true; //색상막대 보이기
      pageInformationMarks = true; //페이지 정보 보이기

      pdfMarkType = MarkTypes.J_MARK_WITH_CIRCLE; //J표시, 원모향 표시
      printerMarkWeight = PDFMarkWeight.p125pt; //두께 0.125pt

      useDocumentBleedWithPDF = true; //문서 도련 설정 사용
    }
  }
  name = new File(myFolder + "/" + myFileName + "_P" + ".pdf"); //파일명_P.pdf 설정
  doc.exportFile(ExportFormat.PDF_TYPE, name, false); //PDF 출력
}

function saveIDML() {
  //IDML 함수 설정
  name = new File(myFolder + "/" + myFileName + ".idml"); //파일명.idml 설정
  doc.exportFile(ExportFormat.INDESIGN_MARKUP, name, false); //IDML 출력
}

function transGuide() {
  //투명상자 함수 설정
  doc.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN; //각 페이지 기준 왼쪽 상단 0mm
  var mspage = doc.masterSpreads.item(0); //1번째 마스터 스프레드

  app.findTextPreferences = NothingEnum.nothing; //문자 찾기 설정 초기화
  app.changeTextPreferences = NothingEnum.nothing; //문자 바꾸기 설정 초기화

  app.findTextPreferences.findWhat = "-|Transparent setting guide|-"; //-|Transparent setting guide|- 텍스트 찾기
  app.findChangeTextOptions.includeMasterPages = true; //마스터 스프레드 포함

  if (doc.findText() == false) {
    //찾은 결과 없을 경우
    for (var i = 0; i < mspage.pages.length; i++) {
      //마스터 페이지 수만큼 반복
      var msguide = mspage.pages.item(i).textFrames.add(); //텍스트 프레임 추가
      msguide.geometricBounds = [0, 0, 5, 60]; //크기 설정 60*5
      msguide.contents = "-|Transparent setting guide|-"; //텍스트 프레임 내 텍스트 설정
      msguide.paragraphs[0].appliedFont = _Font; //적용 글꼴
      msguide.paragraphs[0].fillColor = "Paper"; //글꼴 색상 (용지)
      msguide.transparencySettings.blendingSettings.opacity = 00; //투명도 0%
    }
  }
  app.findTextPreferences = app.changeTextPreferences = NothingEnum.nothing;
}

function PDF_out() {
  //PDF 출력함수 설정
  if (
    (tabmenu.selection == inst && note_i.value == true) ||
    (tabmenu.selection == catal && note.value == true)
  ) {
    //주석 체크박스 설정 시
    notePDF(); //주석 함수 실행
  }
  if (
    (tabmenu.selection == inst && web_i.value == true) ||
    (tabmenu.selection == catal && web.value == true)
  ) {
    //웹 체크박스 설정 시
    webPDF(); //웹 함수 실행
  }
  if (
    (tabmenu.selection == inst && prnt_i.value == true) ||
    (tabmenu.selection == catal && prnt.value == true)
  ) {
    //발주 체크박스 설정 시
    prntPDF(); //발주 함수 실행
  }
  if (
    (tabmenu.selection == inst && idml_i.value == true) ||
    (tabmenu.selection == catal && idml.value == true)
  ) {
    //IDML 체크박스 설정 시
    saveIDML(); //IDML 함수 실행
  }
}

function PDF_error() {
  //PDF 에러함수 설정
  var txtfrms = doc.pages.everyItem().textFrames.everyItem().getElements(); //모든 페이지의 모든 텍스트 프레임 변수 지정
  error_exist = false; //에러 변수 false 설정
  for (var i = txtfrms.length - 1; i >= 0; i--) {
    //모든 텍스트 프레임 수만큼 반복
    if (txtfrms[i].overflows == true) {
      //n번째 텍스트 프레임 넘칠 경우
      error_exist = true; //에러 변수 true 설정
    }
  }
  if (error_exist == true) {
    //에러 변수 true 일 경우
    alert("넘치는 텍스트가 존재합니다! \n확인 후 다시 출력해주세요. ", "경고"); //알림
  }
}
