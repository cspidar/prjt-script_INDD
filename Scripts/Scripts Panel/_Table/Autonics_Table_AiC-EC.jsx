var reg_st = "Regular";
var bld_st = "Bold";

var doc = app.activeDocument;

var sel = app.selection[0]; //선택

for (var i = 0; i < 5; i++) {
  if (!(sel instanceof Table)) {
    sel = sel.parent;
  }
}

var radio_table_value = [false, false, false, false];
try {
  for (var i = 0; i < panel_shade.children.length; i++) {
    radio_table_value[i] = panel_shade.children[i].value;
  }
} catch (e) {}

sel.cells.everyItem().fillColor = "None";

if ((radio_table_value[1] || radio_table_value[3]) == true) {
  sel.rows[0].cells.everyItem().fillColor = "Black"; // 1번째 행 채우기 색상: 검정
  sel.rows[0].cells.everyItem().fillTint = 20; //투명도 20%
}

if ((radio_table_value[2] || radio_table_value[3]) == true) {
  sel.columns[0].cells.everyItem().fillColor = "Black"; //1번째 열 채우기 색상: 검정
  sel.columns[0].cells.everyItem().fillTint = 20; //투명도 20%
}

app.findGrepPreferences = app.changeGrepPreferences = null;
app.findGrepPreferences.position = Position.SUPERSCRIPT;
var found_0 = sel.findGrep();

app.findGrepPreferences = app.changeGrepPreferences = null;
app.findGrepPreferences.position = Position.SUBSCRIPT;
var found_1 = sel.findGrep();

var par_reg = doc.paragraphStyles.itemByName(reg_st);
var par_bld = doc.paragraphStyles.itemByName(bld_st);
var par_style = 0;

try {
  par_style = check_style.value;
} catch (e) {}

if (par_style == 1) {
  sel.cells.everyItem().texts[0].appliedParagraphStyle = par_reg;
  if ((radio_table_value[2] || radio_table_value[3]) == true) {
    sel.columns[0].cells.everyItem().texts[0].appliedParagraphStyle = par_bld;
  }
  if ((radio_table_value[1] || radio_table_value[3]) == true) {
    sel.rows[0].cells.everyItem().texts[0].appliedParagraphStyle = par_bld;
  }
}

for (i = 0; i < found_0.length; i++) {
  found_0[i].position = Position.SUPERSCRIPT;
}

for (i = 0; i < found_1.length; i++) {
  found_1[i].position = Position.SUBSCRIPT;
}

var cell_empty = 0;

try {
  cell_empty = check_empty.value;
} catch (e) {}

if (cell_empty == 1) {
  for (var i = 0; i < sel.cells.length; i++) {
    if (sel.cells[i].contents == "") {
      sel.cells[i].contents = "-";
    }
  }
}

var tab_size = 0;
var tab_width = 15;
try {
  tab_size = check_width.value;
  tab_width = edit_width.text;
} catch (e) {}

if (tab_size == 1) {
  sel.columns[0].width = tab_width;
  var col_width = [];
  if (sel.columns[0].cells[0].contents == "Index") {
    if (sel.columns.length == 8) {
      sel.columns[1].cells[0].contents = "Sub";
      sel.columns[4].cells[0].contents = "PDO Mapping";
      sel.columns[5].cells[0].contents = "Value (defalut)";
      sel.columns[7].cells[0].contents = "Reflect timing";
      col_width = [12, 9, 12, 13, 25, 52, 10, 25];
      for (var i = 0; i < sel.columns.length; i++) {
        sel.columns[i].width = col_width[i];
      }
    } else if (sel.columns.length == 9) {
      sel.columns[1].cells[0].contents = "Sub";
      sel.columns[5].cells[0].contents = "PDO Mapping";
      sel.columns[6].cells[0].contents = "Value (defalut)";
      sel.columns[8].cells[0].contents = "Reflect timing";
      col_width = [12, 9, 45, 12, 13, 17, 26, 10, 14];
      for (var i = 0; i < sel.columns.length; i++) {
        sel.columns[i].width = col_width[i];
      }
    }
  } else {
    //~         for (var i = 1; i < sel.columns.length; i++) {
    //~             sel.columns[i].width =(158-tab_width)/(sel.columns.length-1);
    //~         }
  }
}

var cell_inset = 0;
var radio_table_inset = [0.5, 0.5, 0.5, 0.5];
try {
  cell_inset = check_inset.value;
  radio_table_inset[0] = edittext_left.text;
  radio_table_inset[1] = edittext_right.text;
  radio_table_inset[2] = edittext_top.text;
  radio_table_inset[3] = edittext_bottom.text;
} catch (e) {}

var every_cell = sel.cells.everyItem(); //모든 셀 변수 지정
with (every_cell) {
  //셀 설정
  verticalJustification = VerticalJustification.CENTER_ALIGN;

  topEdgeStrokeWeight = 0; //상단 외곽선 두께 0pt
  bottomEdgeStrokeWeight = 0; //하단 외곽선 두께 0pt
  leftEdgeStrokeWeight = 0; //좌측 외곽선 두께 0pt
  rightEdgeStrokeWeight = 0; //우측 외곽선 두께 0pt

  minimumHeight = 0.1;
  leftInset = 0.5;
  rightInset = 0.5;
  topInset = 0.5;
  bottomInset = 0.5;

  if (cell_inset == 1) {
    leftInset = radio_table_inset[0];
    rightInset = radio_table_inset[1];
    topInset = radio_table_inset[2];
    bottomInset = radio_table_inset[3];
  }
}

sel.cells[0].insertionPoints[0].parentTextFrames[0].strokeColor = "Black"; //텍스트 프레임 색상 없음
sel.cells[0].insertionPoints[0].parentTextFrames[0].strokeWeight = "0mm"; //텍스트 프레임 두께 없음
sel.rows[0].cells.everyItem().topEdgeStrokeWeight = 0.75; //1번째 행 상단 외곽선 두께 0.75pt
for (var i = 0; i < sel.rows.length; i++) {
  sel.rows[i].cells.everyItem().bottomEdgeStrokeWeight = 0.75; //1번째 행 상단 외곽선 두께 0.75pt
}

sel.rows.everyItem().innerColumnStrokeWeight = 0.25; //단락 내부 선 두께 0.25pt
sel.rows.everyItem().innerColumnStrokeColor = "Black";
sel.rows.everyItem().innerColumnStrokeType = "실선"; //선 스타일입니다.  (실선, 굵은선 - 굵은선, 굵은 선 - 가는 선, etc)

sel.columns.everyItem().innerRowStrokeWeight = 0.25; //단락 내부 선 두께 0.25pt
sel.columns.everyItem().innerRowStrokeColor = "Black";
sel.columns.everyItem().innerColumnStrokeType = "실선"; //선 스타일입니다.  (실선, 굵은선 - 굵은선, 굵은 선 - 가는 선, etc)

app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;

var sel_table = sel.parent.tables.itemByID(sel.id);
//~ try {
//~     sel.parent.tables.nextItem(sel_table).select();
//~ } catch (e) {
//~     doc.textFrames.nextItem(sel.parent).tables[0].select();
//~ }
