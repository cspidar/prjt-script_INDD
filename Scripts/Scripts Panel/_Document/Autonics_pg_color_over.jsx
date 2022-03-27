var doc = app.activeDocument;

var active_page = app.activeWindow.activePage;

active_page.pageColor = UIColors.BLUE; //페이지 색상 지정

var pagePanel = app.panels.itemByName("$ID/Pages"); //페이지 패널 변수 설정
pagePanel.visible = true; //페이지 보이기
