var new_doc = app.open(
  File(Folder(app.activeScript).parent.parent + "/_Page/Autonics_Blank.indd"),
  true,
  OpenOptions.OPEN_COPY
); //Autonics_Blank.indd 파일 열기 (보이지 않게)
with (new_doc) {
  documentPreferences.facingPages = false; //페이지 마주보기 설정
}

alert("생성되었습니다.", "Autoncis TC");
