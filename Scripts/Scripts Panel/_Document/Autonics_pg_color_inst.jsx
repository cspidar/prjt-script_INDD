var doc = app.activeDocument;

var active_page = app.activeWindow.activePage;

active_page.pageColor = UIColors.RED;

var pagePanel = app.panels.itemByName("$ID/Pages"); //책갈피 패널 변수 설정
pagePanel.visible = true; //책갈피 보이기
