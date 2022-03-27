var _Font = "Autonics_TC [2020]"; //글꼴 지정

var doc = app.activeDocument; //활성화 된 문서 변수 설정
app.findGrepPreferences = app.changeGrepPreferences = null; //GREP 찾기 설정 초기화
app.findGrepPreferences.findWhat = "VDC"; //찾을 GREP: VDC
app.changeGrepPreferences.changeTo = "VDCᜡ"; //변경할 GREP: VDCᜡ
app.changeGrepPreferences.appliedFont = _Font; //폰트 적용

var found = doc.findGrep(); //GREP 찾기
for (i = 0; i < found.length; i++) {
  //찾은 GREP 수만큼 반복
  //if (found[i].parent.constructor.name == "Cell") { //표 안에 있을 경우 (표안의 항목만 변경시 앞의 //를 지워주세요)
  found[i].changeGrep(); //GREP 바꾸기
  //} //(표안의 항목만 변경시 앞의 //를 지워주세요)
}

app.findGrepPreferences.findWhat = "ᜡᜡ"; //찾을 GREP: ᜡᜡ
app.changeGrepPreferences.changeTo = "ᜡ"; //변경할 GREP: ᜡ
doc.changeGrep(); //GREP 변경

app.findGrepPreferences = app.changeGrepPreferences = null; //GREP 찾기 설정 초기화
app.findGrepPreferences.findWhat = "VAC"; //
app.changeGrepPreferences.changeTo = "VACᜠ"; //
app.changeGrepPreferences.appliedFont = _Font; //

var found = doc.findGrep(); //GREP 찾기
for (i = 0; i < found.length; i++) {
  //찾은 GREP 수만큼 반복
  //if (found[i].parent.constructor.name == "Cell") { //표안에 있을 경우 (표안의 항목만 변경시 앞의 //를 지워주세요)
  found[i].changeGrep(); //GREP 바꾸기
  //} //(표안의 항목만 변경시 앞의 //를 지워주세요)
}

app.findGrepPreferences.findWhat = "ᜠᜠ"; //찾을 GREP: ᜠᜠ
app.changeGrepPreferences.changeTo = "ᜠ"; //변경할 GREP: ᜠ
doc.changeGrep(); //GREP 변경

app.findGrepPreferences = app.changeGrepPreferences = null;
alert("전압기호 추가가 완료되었습니다", "Autonics TC");
