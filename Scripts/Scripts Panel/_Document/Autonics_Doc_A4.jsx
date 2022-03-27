var new_doc = app.open(
  File(Folder(app.activeScript).parent.parent + "/_Page/Autonics_Blank.indd"),
  true,
  OpenOptions.OPEN_COPY
); //Autonics_Blank.indd 파일 열기 (보이지 않게)
with (new_doc) {
  documentPreferences.facingPages = true; //페이지 마주보기 설정
  documentPreferences.allowPageShuffle = false; //선택한 스프레드 재편성 허용 안함
  documentPreferences.pageSize = "A4"; //페이지 폭 설정
  pages[0].appliedMaster = new_doc.masterSpreads.itemByName("P-Manual");
}

alert("생성되었습니다.", "Autoncis TC");
