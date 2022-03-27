var doc = app.activeDocument;

var count = 0;

//모든 페이지 아이템 조사
var pg = doc.allPageItems;
for (i = 0; i < pg.length; i++) {
  if (pg[i] instanceof TextFrame || pg[i] instanceof TextPath) {
    //개체가 텍스트프레임이거나 텍스트 패스일 경우
    if (pg[i].overflows == true) {
      //넘치면
      pg[i].textFramePreferences.autoSizingReferencePoint =
        AutoSizingReferenceEnum.TOP_LEFT_POINT; //좌상단 포인트 기준
      pg[i].textFramePreferences.useNoLineBreaksForAutoSizing = true; //자동 크기 정렬  활성화
      if (
        Math.round(pg[i].geometricBounds[3] - pg[i].geometricBounds[1]) >= 88
      ) {
        //템플릿이 88mm니까 88보다 크면 규격외로
        pg[i].textFramePreferences.autoSizingType =
          AutoSizingTypeEnum.HEIGHT_ONLY; //높이만 조절
      } else {
        pg[i].textFramePreferences.autoSizingType =
          AutoSizingTypeEnum.HEIGHT_AND_WIDTH; //88mm 이하면 너비 높이 둘다 조절
      }
      count++;
    }
  }
}
alert(count + "개 수정 완료");
