var doc = app.activeDocument; //활성화 된 문서 변수 설정
doc.stories.everyItem().hyphenation = false; //모든 스토리의 하이픈 제거
var paraSt = doc.paragraphStyles.everyItem().getElements(); //모든 단락 스타일 변수 설정
for (var i = 0; i < paraSt.length; i++) {
  //단락 스타일 수만큼 반복
  if (paraSt[i].name.indexOf("[") == -1) {
    //단락스타일 이름이 '[' 로 시작할 경우
    paraSt[i].hyphenation = false; //하이픈 제거
  }
}

app.activeDocument.stories
  .everyItem()
  .tables.everyItem()
  .cells.everyItem()
  .paragraphs.everyItem().hyphenation = false; //모든 테이블의 하이픈 제거
alert("하이픈 제거가 완료되었습니다.", "Autonics TC");
