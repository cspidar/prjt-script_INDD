var margin = 0;
var distribution = DistributeOptions.VERTICAL_SPACE;
var position = AlignOptions.LEFT_EDGES;

try {
  if (check_horizontal.value == true) {
    distribution = DistributeOptions.HORIZONTAL_SPACE;
  }

  var radio_value = [true, false, false, false, false, false, false, false];
  var radio_stats = 0;
  for (i = 0; i < group_align_vertical.children.length; i++) {
    radio_value[i] = group_align_vertical.children[i].value;
  }
  for (i = 0; i < group_align_horizontal.children.length; i++) {
    radio_value[4 + i] = group_align_horizontal.children[i].value;
  }
  for (i = 0; i < radio_value.length; i++) {
    if (radio_value[i] == true) {
      radio_stats = i;
    }
  }
  switch (radio_stats) {
    case 0:
      position = AlignOptions.LEFT_EDGES;
      break;
    case 1:
      position = AlignOptions.HORIZONTAL_CENTERS;
      break;
    case 2:
      position = AlignOptions.RIGHT_EDGES;
      break;
    case 3:
      position = "undefined";
      break;
    case 4:
      position = AlignOptions.TOP_EDGES;
      break;
    case 5:
      position = AlignOptions.VERTICAL_CENTERS;
      break;
    case 6:
      position = AlignOptions.BOTTOM_EDGES;
      break;
    case 7:
      position = "undefined";
      break;
  }
  margin = edit_size.text;
} catch (e) {}

var doc = app.activeDocument; //활성화 된 문서 변수 지정
//var script_run = app.doScript(distribution_Script, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "개체 정렬");
distribution_Script();

function distribution_Script() {
  var sel = app.selection; //선택된 개체 변수 지정
  if (sel.length < 2) {
    //선택된 개체가 2개 미만일 경우
    alert("2개 이상의 개체를 선택해 주세요", "Autonics TC팀"); //알림
  } else {
    //선택된 개체가 2개 이상일 경우
    if (doc.selectionKeyObject == null) {
      //메인 객체가 선택 되지 않았을 경우
      if (position != "undefined") {
        doc.align(sel, position, AlignDistributeBounds.ITEM_BOUNDS); //왼쪽 정렬
      }
      doc.distribute(
        sel,
        distribution,
        AlignDistributeBounds.ITEM_BOUNDS,
        true,
        margin
      ); //간격을 기준으로 세로 정렬
    } else {
      //메인 객체가 선택되었을 경우
      if (position != "undefined") {
        doc.align(
          sel,
          position,
          AlignDistributeBounds.KEY_OBJECT,
          doc.selectionKeyObject
        ); //메인 객체 기준 왼쪽 정렬
      }
      doc.distribute(
        sel,
        distribution,
        AlignDistributeBounds.KEY_OBJECT,
        true,
        margin,
        doc.selectionKeyObject
      ); //메인 객체 기준 간격 세로 정렬
    }
  }
}
