var doc = app.activeDocument; //활성화 된 문서 변수 지정

var sel = app.selection[0]; //현재 선택 개체 변수 지정
if (sel.contents == "") {
  //해당 개체에 텍스트가 없을 경우
  sel = sel.parentTextFrames[0]; //해당 개체를 포함하는 텍스트프레임 지정
}

var _contents = String(sel.contents); //선택 개체 내용 변수 문자열 변수 지정

try {
  for (i = 0; i < sel.tables.everyItem().cells.length; i++) {
    //선택된 개체 안에 표가 있는 경우 셀 수만큼 반복
    _contents += String("%0A" + sel.tables.everyItem().cells[i].contents); //다음 셀 내용 넘어갈 때마다 줄바꿈 추가 (줄바꿈 유니코드 000A -> 웹상 유니코드 %0A)
  }
} catch (e) {}

var translate = ""; //변수 설정
for (i = 0; i < _contents.length; i++) {
  //선택 개체 내용 글자 수 만큼 반복 (각 글자 하나씩 확인)
  if (_contents[i] == ("&" || "<" || ">" || '"' || "%" || "+")) {
    //글자가 &, <, >, \, %, + 인지 확인 (html 주소 상에 해당 문자가 포함되면 구분자로 인식하여 끝까지 번역 안됨)
    translate += "%" + _contents[i].charCodeAt(0).toString(16); //해당 글자는 유니코드로 변환 (16진수)
  } else if (_contents[i] == ("\r" || "\n")) {
    //줄바꿈 문자일 경우
    translate += "%0A"; //줄바꿈 유니코드로 변경
  } else if (_contents[i] == "\t") {
    //탭 문자일경우
    translate += "%09"; //탭 유니코드로 변경
  } else if (_contents[i] == "") {
    //텍스트 프레임 내 테이블 같은 고정 개체 일 경우
    translate += ""; //없앰
  } else {
    // 그 외
    translate += _contents[i]; //유지
  }
}

var _link =
  "https://translate.google.com/#view=home&op=translate&sl=ko&tl=en&text="; //구글 번역기 링크 해당 링크 다음에 문자열 추가하면 문자열 번역함
var URL = File(Folder.temp.absoluteURI + "/link.html"); //temp폴더내 link.html 파일 생성
URL.open("w"); //쓰기 전용으로 파일 열기
var _Body =
  '<html><head><META HTTP-EQUIV=Refresh CONTENT="0; URL=' +
  _link +
  translate +
  '"></head><body> <p></body></html>'; //Html 실행 소스
URL.write(_Body); //실행 소스 파일에 쓰기
URL.close(); //파일 닫기
URL.execute(); //해당 파일 실행
