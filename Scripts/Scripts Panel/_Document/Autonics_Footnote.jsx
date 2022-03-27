var _Font = "Autonics_TC [2020]"; //폰트명을 입력해 주세요.
var char_number = "Footnote) Upper_Letter"; //각주 윗첨자 문자 스타일을 입력해 주세요.
var par_name = "Footnote) Text / Regular / 5pt"; //각주 하단 설명의 단락스타일명을 입력해 주세요.

var doc = app.activeDocument; //현재 활성화 된 문서 설정

var sel = app.selection[0];
sel.parentTextFrames[0].textFramePreferences.autoSizingReferencePoint =
  AutoSizingReferenceEnum.TOP_LEFT_POINT;
sel.parentTextFrames[0].textFramePreferences.autoSizingType =
  AutoSizingTypeEnum.HEIGHT_ONLY;

//각주 설정
with (doc.footnoteOptions) {
  footnoteNumberingStyle = FootnoteNumberingStyle.SINGLE_LEADING_ZEROS; //01부터 시작
  restartNumbering = FootnoteRestarting.SECTION_RESTART; //새로운 번호 시작
  showPrefixSuffix = FootnotePrefixSuffix.PREFIX_SUFFIX_BOTH; //머릿글 꼬릿글 둘다 추가
  prefix = " "; //버릿글
  suffix = ")"; //꼬릿글
  markerPositioning = FootnoteMarkerPositioning.SUPERSCRIPT_MARKER; //윗첨자 설정
  footnoteTextStyle = doc.paragraphStyles.itemByName(par_name); //단락 스타일 설정
  separatorText = "^>^i"; //탭으로 숫자와 내용 구분
  spacer = "0.5mm"; // 상하간격
  spaceBetween = "0.5mm"; //사이간격
  footnoteFirstBaselineOffset = FootnoteFirstBaseline.LEADING_OFFSET;
  footnoteMinimumFirstBaselineOffset = "0.5mm";
  ruleOn = false;
}

sel.footnotes.add();
