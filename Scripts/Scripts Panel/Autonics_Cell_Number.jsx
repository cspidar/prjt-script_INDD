var doc = app.activeDocument;
var sel  =app.selection[0];

app.doScript(fill_cell, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "셀 채우기");

function fill_cell () {
    // 선택 영역 개체가 셀이 될때까지 상위 개체로 이동
    for ( var i = 0 ; i < 5; i++) {
        if (!(sel instanceof Cell)) {
            sel = sel.parent;
        }
    }
    
    if (sel.cells.length ==1) {
        var cell_text = sel.contents.split("-");
        var incremental = Number(cell_text[1]);
        var row_index = 0;
        var column_index = sel.columns[0].index;
        
        for (var i = 0; i < sel.parentColumn.cells.length; i++) {
            if (sel.parentColumn.cells[i].id == sel.id) {
                break;
            }
            else {
                row_index++;
            }
        }
        for (i = row_index; i < sel.parentColumn.cells.length; i++) {
            sel.parentColumn.cells[i].contents = cell_text[0] + "-" + incremental;
            incremental++;
        }
    }
    
    else {
        var cell_text = sel.cells[0].contents.split("-");
        var incremental = Number(cell_text[1]);
        var row_index = 0;
        var column_index = sel.columns[0].index;
        
        for (var i = 0; i < sel.parentColumn.cells.length; i++) {
            if (sel.parentColumn.cells[i].id == sel.id) {
                break;
            }
            else {
                row_index++;
            }
        }
    
        for (i = row_index; i < row_index+sel.cells.length; i++) {
            sel.parentColumn.cells[i].contents = cell_text[0] + "-" + incremental;
            incremental++;
        }
    }
}