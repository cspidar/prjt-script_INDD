var doc = app.activeDocument;

var char_list = doc.characterStyles.length; //문자 스타일 개수 변수 설정
for (i = 0; i < char_list; i++) {
  //문자 스타일 수만큼 반복
  try {
    doc.characterStyles.lastItem().remove(); //마지막에 위치한 문자 스타일 제거
  } catch (e) {}
}
doc.characterStyleGroups.everyItem().remove(); //모든 문자 스타일 그룹 제거

var cell_list = doc.cellStyles.length; //셀 스타일 개수 변수 설정
for (i = 0; i < cell_list; i++) {
  //셀 스타일 수만큼 반복
  try {
    doc.cellStyles.lastItem().remove(); //마지막에 위치한 셀 스타일 제거
  } catch (e) {}
}
doc.cellStyleGroups.everyItem().remove(); //모든 셀 스타일 그룹 제거

var table_list = doc.tableStyles.length; //표 스타일 개수 변수 설정
for (i = 0; i < table_list; i++) {
  //표 스타일 수만큼 반복
  try {
    doc.tableStyles.lastItem().remove(); //마지막에 위치한 표 스타일 제거
  } catch (e) {}
}
doc.tableStyleGroups.everyItem().remove(); //모든 표 스타일 그룹 제거

alert("스타일 제거가 완료되었습니다.");
