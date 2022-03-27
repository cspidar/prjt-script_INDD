var doc = app.activeDocument; //활성화 된 문서 변수 지정
var sel = app.selection[0]; //선택 텍스트 변수 지정
var frame = sel.parentTextFrames[0]; //선택 텍스트 포함한 텍스트프레임 변수 지정
app.copy(); //선택 텍스트 내용 복사

var new_frame = app.activeWindow.activePage.textFrames.add({
  geometricBounds: frame.geometricBounds,
}); //텍스트 프레임 추가
with (new_frame) {
  //설정
  textFramePreferences.autoSizingReferencePoint =
    AutoSizingReferenceEnum.TOP_LEFT_POINT; //텍스트 프레임 기준점 좌측 상단
  textFramePreferences.useNoLineBreaksForAutoSizing = true; //자동 줄바꿈 안함
  textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_AND_WIDTH; //텍스트 프레임 크기 자동조절: 높이만

  texts[0].appliedFont = sel.paragraphs[0].appliedFont; //폰트는 선택 항목과 동일
  texts[0].pointSize = sel.paragraphs[0].pointSize; //폰트 크기는 선택 항목과 동일

  insertionPoints[0].select();
}

app.paste(); //붙여넣기 내용은 분자 + 줄바꿈 + 분모

new_frame.contents = new_frame.contents.replace("/", "\r"); // "/"를 줄바꿈으로 변경
new_frame.texts[0].justification = Justification.CENTER_ALIGN; //중앙 정렬

//분수 가운데 나눔 선 생성
var hor = new_frame.geometricBounds[2] - new_frame.geometricBounds[0];
var ver = new_frame.geometricBounds[3] - new_frame.geometricBounds[1];
var _line = app.activeWindow.activePage.graphicLines.add();
_line.strokeWeight = 0.25;
_line.geometricBounds = [
  new_frame.geometricBounds[0] + hor / 2,
  new_frame.geometricBounds[1] - ver * 0.1,
  new_frame.geometricBounds[0] + hor / 2,
  new_frame.geometricBounds[3] + ver * 0.1,
];

//숫자와 선 그룹화
var group_fraction = doc.groups.add([new_frame, _line]);

//분수 그룹 잘라내기, 처음 선택 영역에 붙여넣기
group_fraction.select();
app.cut();

sel.insertionPoints[0].select();
app.paste();

sel.remove(); //선택 텍스트 지우기

frame.fit(FitOptions.FRAME_TO_CONTENT); //선택 텍스트 프레임 크기 맞추기
frame.select();
