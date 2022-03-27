//초기 값 설정 (스크립트 패널에서 불러오니까 상관은 없는데 그냥 단독 실행으로 오류 잡을때 필요해서 만들어둠)
var margin = 0;

var distribution = DistributeOptions.VERTICAL_SPACE; //기본은 수직 정렬임
var position = AlignOptions.LEFT_EDGES;
var fit_option = AlignDistributeBounds.ITEM_BOUNDS;

//뭔가 오류가 많이나는데 오류 무시해도 알아서 잘 되길래 그냥  try 씀
try {
  margin = margin_value; //여백에 스크립트 패널의 입력값 지정

  if (check_horizontal.value == true) {
    //수평 정렬 선택일 경우
    distribution = DistributeOptions.HORIZONTAL_SPACE;
  }

  if (check_marginal.value == true) {
    //여백 정렬일 경우
    fit_option = AlignDistributeBounds.MARGIN_BOUNDS;
  }

  //스크립트 패널 버튼 읽어오는거
  var radio_value = [true, false, false, false, false, false, false, false]; //왼쪽, 가운데, 오른쪽, 정렬 안함, 위, 가운데, 아래, 정렬 안함
  var radio_stats = 0; //기본
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
        doc.align(sel, position, fit_option); //왼쪽 정렬
      }
      doc.distribute(sel, distribution, fit_option, true, margin); //간격을 기준으로 세로 정렬
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
