#targetengine 'session'

var date = File(app.activeScript).modified;
var update = String(date.getFullYear() + "." + (date.getMonth()+1) + "." + date.getDate());

var usePrefs = true;

var pref_tab = 0;

var pref_inst = 0;
var pref_overview = 0;
var pref_manual = 0;
var pref_note = 0;

var pref_open = 0;
var pref_folder = 0;

var pref_align = 0;

var pref_none = 1;
var pref_row = 0;
var pref_column = 0;
var pref_cr = 0;

var pref_cell_width = 0;
var pref_width_value = 0;

var pref_x = 0;
var pref_y = 0;

var prefsFile = File((Folder(app.activeScript)).parent + "/Autonics_Script_Panel.ini"); //설정 파일 저장 경로 지정
if(prefsFile.exists) { //설정 파일 이 있는 경우
    readPrefs(); //설정파일 읽기 함수 실행
}

function readPrefs() { //설정 파일 읽기 함수 설정
    prefsFile.open("r"); //설정 파일 열기 (읽기만 가능하게)
    pref_tab = Number(prefsFile.readln());
    
    pref_inst = Number(prefsFile.readln());
    pref_overview = Number(prefsFile.readln());
    pref_manual = Number(prefsFile.readln());
    pref_note = Number(prefsFile.readln());
        
    pref_open = Number(prefsFile.readln());
    pref_folder = Number(prefsFile.readln());
    
    pref_align = Number(prefsFile.readln());
    
    pref_none = Number(prefsFile.readln());
    pref_row = Number(prefsFile.readln());
    pref_column = Number(prefsFile.readln());
    pref_cr = Number(prefsFile.readln());
    
    pref_cell_width = Number(prefsFile.readln());
    pref_width_value = String(prefsFile.readln());
    
    pref_x = Number(prefsFile.readln());
    pref_y = Number(prefsFile.readln());
    
    prefsFile.close(); //설정 파일 파일 닫기
}

function savePrefs() { //설정파일 저장 함수 설정
    var newPrefs = //설정 값 변수 선언
    tpanel1.selection.name + "\n" +
     ((check_inst.value)?1:0) + "\n" +
     ((check_overview.value)?1:0) + "\n" +
     ((check_manual.value)?1:0) + "\n" +
     ((check_note.value)?1:0) + "\n" +
     ((check_open.value)?1:0) + "\n" +
     ((check_folder.value)?1:0) + "\n" +
     edit_size.text + "\n" +
     ((radio_cell_none.value)?1:0) + "\n" +
     ((radio_cell_row.value)?1:0) + "\n" +
     ((radio_cell_column.value)?1:0) + "\n" +
     ((radio_cell_cr.value)?1:0) + "\n" +
     ((check_width.value)?1:0) + "\n" +
     edit_width.text + "\n" +
     dialog.frameLocation[0] + "\n" +
     dialog.frameLocation[1];
     prefsFile.open("w"); //설정 파일 열기 (쓰기 가능)
    prefsFile.write(newPrefs); //설정파일에 설정 값 쓰기
    prefsFile.close(); //설정 파일 닫기
}

// DIALOG
// ======
var dialog = new Window("palette");
    dialog.text = "Autonics TC Script Panel - " + update; 
    dialog.orientation = "column"; 
    dialog.alignChildren = ["center","top"]; 
    dialog.spacing = 10; 
    dialog.margins = 0; 
    dialog.frameLocation = [pref_x, pref_y];
    
// TPANEL1
// =======
var tpanel1 = dialog.add("tabbedpanel", undefined, undefined, {name: "tpanel1"}); 
    tpanel1.alignChildren = "fill"; 
    tpanel1.maximumSize.width = 275; 
    tpanel1.margins = 0; 

// TAB_EXPORT
// ========
var tab_export = tpanel1.add("tab");
    tab_export.name = "0";
    tab_export.text = "출력"; 
    tab_export.orientation = "column"; 
    tab_export.alignChildren = ["fill","top"]; 
    tab_export.spacing = 5; 
    tab_export.margins = 10; 
    
// PANEL_DOCUMENT
// ==============
var panel_document = tab_export.add("panel", undefined, undefined, {name: "panel_document"}); 
    panel_document.text = "출력 문서"; 
    panel_document.orientation = "row"; 
    panel_document.alignChildren = ["left","center"]; 
    panel_document.spacing = 5; 
    panel_document.margins = 15; 

var inst_img = File((Folder(app.activeScript)).parent + "/_Icons/file_inst.png");
var check_inst = panel_document.add("iconbutton", undefined, File.decode(inst_img), {name: "check_inst", style: "toolbutton", toggle: true}); 
    check_inst.preferredSize.width = 45; 
    check_inst.value = pref_inst;
    check_inst.onClick = function () {
        app.activate();
    }

var overview_img = File((Folder(app.activeScript)).parent + "/_Icons/file_catal.png");
var check_overview = panel_document.add("iconbutton", undefined, File.decode(overview_img), {name: "check_overview", style: "toolbutton", toggle: true}); 
    check_overview.preferredSize.width = 45; 
    check_overview.value = pref_overview;
    check_overview.onClick = function () {
        app.activate();
    }
    
var manual_img = File((Folder(app.activeScript)).parent + "/_Icons/file_manual.png");
var check_manual = panel_document.add("iconbutton", undefined, File.decode(manual_img), {name: "check_manual", style: "toolbutton", toggle: true}); 
    check_manual.preferredSize.width = 65; 
    check_manual.value = pref_manual;
    check_manual.onClick = function () {
        app.activate();
    }

var note_img = File((Folder(app.activeScript)).parent + "/_Icons/file_note.png");
var check_note = panel_document.add("iconbutton", undefined, File.decode(note_img), {name: "check_note", style: "toolbutton", toggle: true}); 
    check_note.preferredSize.width = 45; 
    check_note.value = pref_note;
    check_note.onClick = function () {
        app.activate();
    }

// PANEL_PREVIEW
// =============
var panel_preview = tab_export.add("panel", undefined, undefined, {name: "panel_preview"}); 
    panel_preview.text = "설정"; 
    panel_preview.orientation = "row"; 
    panel_preview.alignChildren = ["left","center"]; 
    panel_preview.spacing = 20; 
    panel_preview.margins = 15; 

var check_preview = panel_preview.add("checkbox", undefined, undefined, {name: "check_preview"}); 
    check_preview.text = "미리 보기"; 

var check_idml = panel_preview.add("checkbox", undefined, undefined, {name: "check_idml"}); 
    check_idml.text = "IDML 출력"; 

// PANEL_EXPORT
// =============
var panel_export = tab_export.add("panel", undefined, undefined, {name: "panel_export"}); 
    panel_export.text = "출력 설정"; 
    panel_export.orientation = "column"; 
    panel_export.alignChildren = ["left","center"]; 
    panel_export.spacing = 10; 
    panel_export.margins = 15; 

// GROUP_FOLDER
// ============
var group_folder = panel_export.add("group", undefined, {name: "group_folder"}); 
    group_folder.orientation = "row"; 
    group_folder.alignChildren = ["left","center"]; 
    group_folder.spacing = 10; 
    group_folder.margins = 0; 

var statictext1 = group_folder.add("statictext", undefined, undefined, {name: "statictext1"}); 
    statictext1.text = "경로:"; 

var edittext_folder = group_folder.add('edittext {properties: {name: "edittext_folder", readonly: "true"}}'); 
    edittext_folder.text = "경로를 설정해주세요."; 
    edittext_folder.preferredSize.width = 100; 

var file_loc = "";

var icon_update_imgString = File((Folder(app.activeScript)).parent + "/_Icons/refresh.png");
var icon_update = group_folder.add("iconbutton", undefined, File.decode(icon_update_imgString), {name: "icon_update", style: "toolbutton"}); 
    icon_update.preferredSize.width = 25; 
    icon_update.onClick = function () {
        file_loc = app.activeDocument.filePath;
        edittext_folder.text = Folder(file_loc).fullName;
    }

var icon_folder_imgString = File((Folder(app.activeScript)).parent + "/_Icons/folder.png");
var icon_folder = group_folder.add("iconbutton", undefined, File.decode(icon_folder_imgString), {name: "icon_folder", style: "toolbutton"}); 
    icon_folder.preferredSize.width = 30; 
    icon_folder.onClick = function () {
        var folder_select = Folder(file_loc).selectDlg("저장하려는 위치를 선택해주세요.");
        if(folder_select != null) {
            edittext_folder.text = Folder(folder_select).fullName;
            file_loc = folder_select;
        }
    }

// GROUP_SETTING
// ============
var group_setting = panel_export.add("group", undefined, {name: "group_setting"}); 
    group_setting.orientation = "column"; 
    group_setting.alignChildren = ["left","center"]; 
    group_setting.spacing = 5; 


// GROUP_BATCH
// ============
var group_batch = group_setting.add("group", undefined, {name: "group_batch"}); 
    group_batch.orientation = "row"; 
    group_batch.alignChildren = ["left","center"]; 
    group_batch.spacing = 35; 

var check_batch = group_batch.add("checkbox", undefined, undefined, {name: "check_batch"}); 
    check_batch.text = "일괄 출력"; 

var check_subfolder = group_batch.add("checkbox", undefined, undefined, {name: "check_subfolder"}); 
    check_subfolder.text = "하위 폴더 포함"; 
    
    if(check_batch.value == false) {
        icon_update.enabled = true;
        check_subfolder.enabled = false;
    }

    check_batch.addEventListener('mousedown', function (evt) {
        icon_update.enabled = !icon_update.enabled;
        check_subfolder.enabled = !check_subfolder.enabled;        
    });

// GROUP_EXPORT
// ============
var group_export = group_setting.add("group", undefined, {name: "group_export"}); 
    group_export.orientation = "row"; 
    group_export.alignChildren = ["left","center"]; 
    group_export.spacing = 20; 
    
var check_open = group_export.add("checkbox", undefined, undefined, {name: "check_open"}); 
    check_open.text = "완료 후 열기"; 
    check_open.value = pref_open;

var check_folder = group_export.add("checkbox", undefined, undefined, {name: "check_folder"}); 
    check_folder.text = "폴더 생성";
    check_folder.value = pref_folder;
    
// TAB_EXPORT
// ========
var button_export = tab_export.add("button", undefined, undefined, {name: "button_export"}); 
    button_export.text = "출    력"; 
    button_export.preferredSize.height = 30; 

var file_export = File((Folder(app.activeScript)).parent + "/_Page/Autonics_Export.jsx");

button_export.onClick = function () {
    app.doScript(file_export, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "문서 출력");
    app.activate();
};

if(check_preview.value == false) {
        check_idml.enabled = true;
        group_folder.enabled = true;
        group_setting.enabled = true;
    }

    check_preview.addEventListener('mousedown', function (evt) {
        check_idml.enabled = !check_idml.enabled;
        group_folder.enabled = !group_folder.enabled;
        group_setting.enabled = !group_setting.enabled;
    });

// TAB_DOC
// ========
var tab_doc = tpanel1.add("tab");
    tab_doc.name = "1";
    tab_doc.text = "문서"; 
    tab_doc.orientation = "column"; 
    tab_doc.alignChildren = ["fill","top"]; 
    tab_doc.spacing = 7; 
    tab_doc.margins = 10; 

// PANEL_NEWPAGE
// =============
var panel_newpage = tab_doc.add("panel", undefined, undefined, {name: "panel_newpage"}); 
    panel_newpage.text = "새 문서"; 
    panel_newpage.orientation = "row"; 
    panel_newpage.alignChildren = ["center","fill"]; 
    panel_newpage.spacing = 25; 
    panel_newpage.margins = 10; 

var button_blank_imgString = File((Folder(app.activeScript)).parent + "/_Icons/new_blank.png");
var button_blank = panel_newpage.add("iconbutton", undefined, File.decode(button_blank_imgString), {name: "button_blank", style: "toolbutton"}); 
    button_blank.fillBrush = button_blank.graphics.newBrush( button_blank.graphics.BrushType.SOLID_COLOR, [0, 0, 0, 0] );
    button_blank.preferredSize.width = 40; 
    button_blank.preferredSize.height = 40; 

var file_blank = File((Folder(app.activeScript)).parent + "/_Document/Autonics_Doc_Blank.jsx");
button_blank.onClick = function () {
    app.doScript(file_blank, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "문서 생성");
    app.activate();
};    

var button_a4_imgString = File((Folder(app.activeScript)).parent + "/_Icons/new_a4.png");
var button_a4 = panel_newpage.add("iconbutton", undefined, File.decode(button_a4_imgString), {name: "button_a4", style: "toolbutton"}); 
    button_a4.preferredSize.width = 40; 
    button_a4.preferredSize.height = 40; 

var file_a4 = File((Folder(app.activeScript)).parent + "/_Document/Autonics_Doc_A4.jsx");
button_a4.onClick = function () {
    app.doScript(file_a4, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "문서 생성");
    app.activate();
};

var button_a3_imgString = File((Folder(app.activeScript)).parent + "/_Icons/new_a3.png");
var button_a3 = panel_newpage.add("iconbutton", undefined, File.decode(button_a3_imgString), {name: "button_a3", style: "toolbutton"}); 
    button_a3.preferredSize.width = 40; 
    button_a3.preferredSize.height = 40; 

var file_a3 = File((Folder(app.activeScript)).parent + "/_Document/Autonics_Doc_A3.jsx");
button_a3.onClick = function () {
    app.doScript(file_a3, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "문서 생성");
    app.activate();
};

// GROUP_PG2
// =========
var group_pg = tab_doc.add("group", undefined, {name: "group_pg"}); 
    group_pg.orientation = "row"; 
    group_pg.alignChildren = ["left","fill"]; 
    group_pg.spacing = 15; 
    group_pg.margins = 0; 
    group_pg.alignment = ["fill","center"]; 

// PANEL_PAGE_NEW
// ==========
var panel_page_new = group_pg.add("panel", undefined, undefined, {name: "panel_page_new"}); 
    panel_page_new.text = "페이지"; 
    panel_page_new.orientation = "row"; 
    panel_page_new.alignChildren = ["left","top"]; 
    panel_page_new.spacing = 0; 
    panel_page_new.margins = [15,15, 5, 2]; 
    panel_page_new.alignment = ["fill","center"]; 
    panel_page_new.maximumSize.width = 170;

var group_select = panel_page_new.add("group", undefined, {name: "group_select"}); 
    group_select.orientation = "column"; 
    group_select.alignChildren = ["left","center"]; 
    group_select.spacing = 0; 
    group_select.margins = 0; 
    group_select.alignment = ["fill","top"]; 

var radio_new = group_select.add("radiobutton", undefined, undefined, {name: "radio_new"}); 
    radio_new.text = "Add"; 
    radio_new.value = true;

var radio_resize = group_select.add("radiobutton", undefined, undefined, {name: "radio_resize"}); 
    radio_resize.text = "Resize"; 

var new_half_imgString = File((Folder(app.activeScript)).parent + "/_Icons/page_half.png");
var button_half = panel_page_new.add("iconbutton", undefined, File.decode(new_half_imgString), {name: "button_half", style: "toolbutton"}); 
    button_half.preferredSize.width = 20; 
    button_half.preferredSize.height = 30; 

var file_new_half = File((Folder(app.activeScript)).parent + "/_Document/Autonics_half_new.jsx");
var file_resize_half = File((Folder(app.activeScript)).parent + "/_Document/Autonics_half_resize.jsx");

button_half.onClick = function () {
    var file_type = (radio_resize.value == true) ? file_resize_half : file_new_half;
    app.doScript(file_type, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "문서 생성");
    app.activate();
};

var new_full_imgString = File((Folder(app.activeScript)).parent + "/_Icons/page_full.png");
var button_full = panel_page_new.add("iconbutton", undefined, File.decode(new_full_imgString), {name: "button_full", style: "toolbutton"}); 
    button_full.preferredSize.width = 30; 
    button_full.preferredSize.height = 30; 

var file_new_full = File((Folder(app.activeScript)).parent + "/_Document/Autonics_full_new.jsx");
var file_resize_full = File((Folder(app.activeScript)).parent + "/_Document/Autonics_full_resize.jsx");

button_full.onClick = function () {
    var file_type = (radio_resize.value == true) ? file_resize_full : file_new_full;
    app.doScript(file_type, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "문서 생성");
    app.activate();
};

// PANEL_PAGE_RANGE
// ==========
var panel_page_range = group_pg.add("panel", undefined, undefined, {name: "panel_page_range"}); 
    panel_page_range.text = "취설 범위"; 
    panel_page_range.orientation = "row"; 
    panel_page_range.alignChildren = ["fill","center"]; 
    panel_page_range.spacing = 15; 
    panel_page_range.margins = 15; 
    panel_page_range.alignment = ["fill","top"]; 
    panel_page_range.maximumSize.width = 80;

var button_inst = panel_page_range.add("button", undefined, undefined, {name: "button_inst"}); 
    button_inst.text = "설정"; 
    button_inst.maximumSize = [50, 30];

button_inst.onClick = function () {
    var pagePanel = app.panels.itemByName ("$ID/Pages"); //페이지 패널 변수 설정
    pagePanel.visible =true; //페이지 보이기
    if (app.activeWindow.activePage.pageColor == UIColors.RED) {
        app.activeWindow.activePage.pageColor = PageColorOptions.NOTHING;
    }
    else {
        app.activeWindow.activePage.pageColor = UIColors.RED
    }
    app.activate();
};

// PANEL_EDIT
// ==========
var panel_edit = tab_doc.add("panel", undefined, undefined, {name: "panel_edit"}); 
    panel_edit.text = "편집"; 
    panel_edit.orientation = "column"; 
    panel_edit.alignChildren = ["left","center"]; 
    panel_edit.spacing = 15; 
    panel_edit.margins = 15; 
    panel_edit.alignment = ["fill","top"]; 

// GROUP_EDIT0
// ==========
var group_edit0 = panel_edit.add("group", undefined, {name: "group_edit0"}); 
    group_edit0.orientation = "row"; 
    group_edit0.alignChildren = ["left","fill"]; 
    group_edit0.spacing = 10; 
    group_edit0.margins = 0; 
    group_edit0.alignment = ["fill","center"]; 

var button_footnote = group_edit0.add("button", undefined, undefined, {name: "button_footnote"}); 
    button_footnote.text = "각주  추가"; 
    button_footnote.preferredSize = [105, 30];

var file_footnote = File((Folder(app.activeScript)).parent + "/_Document/Autonics_Footnote.jsx");
button_footnote.onClick = function () {
    app.doScript(file_footnote, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "각주 추가");
    app.activate();
};

var button_fraction = group_edit0.add("button", undefined, undefined, {name: "button_fraction"}); 
    button_fraction.text = "분수 추가"; 
    button_fraction.preferredSize = [105, 30];

var file_fraction = File((Folder(app.activeScript)).parent + "/_Document/Autonics_Fraction.jsx");
button_fraction.onClick = function () {
    app.doScript(file_fraction, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "분수 추가");
    app.activate();
};

// GROUP_EDIT1
// ==========
var group_edit1 = panel_edit.add("group", undefined, {name: "group_edit1"});
    group_edit1.orientation = "row"; 
    group_edit1.alignChildren = ["left","fill"]; 
    group_edit1.spacing = 10; 
    group_edit1.margins = 0; 
    group_edit1.alignment = ["fill","center"]; 

var button_link = group_edit1.add("button", undefined, undefined, {name: "button_link"}); 
    button_link.text = "품번 동기화"; 
    button_link.preferredSize = [105, 30];

var file_link = File((Folder(app.activeScript)).parent + "/_Document/Autonics_Link.jsx");
button_link.onClick = function () {
    app.doScript(file_link, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "품번 동기화");
    app.activate();
};

var button_symbol = group_edit1.add("button", undefined, undefined, {name: "button_symbol"}); 
    button_symbol.text = "전압기호 추가"; 
    button_symbol.preferredSize = [105, 30];

var file_symbol = File((Folder(app.activeScript)).parent + "/_Document/Autonics_Symbol.jsx");
button_symbol.onClick = function () {
    app.doScript(file_symbol, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "전압기호 추가");
    app.activate();
};

var button_save = tab_doc.add("button", undefined, undefined, {name: "button_save"}); 
    button_save.text = "파일 저장"; 
    button_save.preferredSize.height = 30;

var file_save = File((Folder(app.activeScript)).parent + "/_Document/Autonics_Save.jsx");
button_save.onClick = function () {
    app.doScript(file_save, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "파일 저장");
    app.activate();
};

// TAB_FUNCTION
// ============
var tab_function = tpanel1.add("tab");
    tab_function.name = "2";
    tab_function.text = "기능"; 
    tab_function.orientation = "column"; 
    tab_function.alignChildren = ["fill","top"]; 
    tab_function.spacing = 5; 
    tab_function.margins = 10; 

// PANEL_TRANSLATE
// =========
var panel_translate = tab_function.add("panel", undefined, undefined, {name: "panel_translate"}); 
    panel_translate.text = "언어"; 
    panel_translate.orientation = "column"; 
    panel_translate.alignChildren = ["fill","top"]; 
    panel_translate.spacing = 15; 
    panel_translate.margins = 15; 

// GROUP_TRANSLATE_FUNCTION
// ===========
var group_translate_function = panel_translate.add("group", undefined, {name: "group_translate_function"}); 
    group_translate_function.preferredSize.width = 140; 
    group_translate_function.orientation = "row"; 
    group_translate_function.alignChildren = ["left","center"]; 
    group_translate_function.spacing = 15; 
    group_translate_function.margins = 0; 

var translate_folder = Folder((app.activeScript).parent.parent.parent + "/_Translate/_Series").getFiles();
var load_folders = [];
load_folders.unshift("--폴더 선택--");

for (i = 0; i < translate_folder.length; i++) {
    if (translate_folder[i] instanceof Folder) {
        load_folders.push (translate_folder[i].displayName);
    }
}

var dropdown_folder = group_translate_function.add("dropdownlist", undefined, undefined, {name: "dropdown_folder", items: load_folders}); 
    dropdown_folder.selection = 0; 
    dropdown_folder.preferredSize.width = 125; 

var button_translate_EN = group_translate_function.add("button", undefined, undefined, {name: "button_translate_EN"}); 
    button_translate_EN.text = "한-영 변환"; 
    button_translate_EN.preferredSize = [80, 30];

var file_translate_EN = File((Folder(app.activeScript)).parent + "/_Function/Autonics_Translate_EN.jsx");
button_translate_EN.onClick = function () {
    app.doScript(file_translate_EN, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "한-영 변환");
    app.activate();
};

// GROUP_TRANSLATE_FUNCTION_2
// ===========
var group_translate_function_2 = panel_translate.add("group", undefined, {name: "group_translate_function_2"}); 
    group_translate_function_2.preferredSize.width = 140; 
    group_translate_function_2.orientation = "row"; 
    group_translate_function_2.alignChildren = ["left","center"]; 
    group_translate_function_2.spacing = 15; 
    group_translate_function_2.margins = 0; 

var icon_google_imgString = File((Folder(app.activeScript)).parent + "/_Icons/google.png");
var icon_google = group_translate_function_2.add("iconbutton", undefined, File.decode(icon_google_imgString), {name: "icon_google"});//, style: "toolbutton"}); 
    icon_google.preferredSize.width = 55;

var file_google = File((Folder(app.activeScript)).parent + "/_Function/Autonics_Google.jsx");
icon_google.onClick = function () {
    app.doScript(file_google);
    app.activate();
};

var icon_papago_imgString = File((Folder(app.activeScript)).parent + "/_Icons/papago.png");
var icon_papago = group_translate_function_2.add("iconbutton", undefined, File.decode(icon_papago_imgString), {name: "icon_papago"});//, style: "toolbutton"}); 
    icon_papago.preferredSize.width = 55;

var file_papago = File((Folder(app.activeScript)).parent + "/_Function/Autonics_Papago.jsx");
icon_papago.onClick = function () {
    app.doScript(file_papago);
    app.activate();
};

var button_file_senten = group_translate_function_2.add("button", undefined, undefined, {name: "button_file_senten"}); 
    button_file_senten.text = "번역 폴더"; 
    button_file_senten.preferredSize = [80, 30];
var folder_trans = Folder((app.activeScript).parent.parent.parent + "/_Translate");
button_file_senten.onClick = function () {
    folder_trans.execute();
    app.activate();
};

// PANEL_ETC
// ==============
var panel_etc = tab_function.add("panel", undefined, undefined, {name: "panel_etc"}); 
    panel_etc.text = "기타"; 
    panel_etc.orientation = "column"; 
    panel_etc.alignChildren = ["fill","top"]; 
    panel_etc.spacing = 15; 
    panel_etc.margins = 15; 
    panel_etc.alignment = ["fill","top"]; 


// GROUP_ADD
// ===========
var group_add = panel_etc.add("group", undefined, {name: "group_add"}); 
    group_add.orientation = "row"; 
    group_add.alignChildren = ["left","top"]; 
    group_add.spacing = 10; 
    group_add.margins = 0; 
    group_add.alignment = ["fill","center"]; 

var button_check = group_add.add("button", undefined, undefined, {name: "button_check"}); 
    button_check.text = "체크 리스트"; 
    button_check.preferredSize = [105, 30];

var file_check = File((Folder(app.activeScript)).parent + "/_Function/Autonics_Checklist.jsx");
button_check.onClick = function () {
    app.doScript(file_check, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "체크 리스트");
    app.activate();
};

var button_style = group_add.add("button", undefined, undefined, {name: "button_style"}); 
    button_style.text = "스타일 제거"; 
    button_style.preferredSize = [105, 30];

var file_style = File((Folder(app.activeScript)).parent + "/_Function/Autonics_Style_remover.jsx");
button_style.onClick = function () {
    app.doScript(file_style, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "스타일 제거");
    app.activate();
};

// GROUP_REMOVE
// ===========
var group_remove = panel_etc.add("group", undefined, {name: "group_remove"}); 
    group_remove.orientation = "row"; 
    group_remove.alignChildren = ["left","top"]; 
    group_remove.spacing = 10; 
    group_remove.margins = 0; 
    group_remove.alignment = ["fill","center"]; 

var button_word = group_remove.add("button", undefined, undefined, {name: "button_word"}); 
    button_word.text = "용어 변경"; 
    button_word.preferredSize = [105, 30];

var file_word = File((Folder(app.activeScript)).parent + "/_Function/Autonics_Words.jsx");
button_word.onClick = function () {
    app.doScript(file_word, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "용어 변경");
    app.activate();
};

var button_file_word = group_remove.add("button", undefined, undefined, {name: "button_file_word"}); 
    button_file_word.text = "용어 추가"; 
    button_file_word.preferredSize = [105, 30];
var file_word_indd = File((Folder(app.activeScript)).parent + "/_Function/wordlist.indd");
button_file_word.onClick = function () {
    file_word_indd.execute();
    app.activate();
};

// GROUP_ETC1
// ===========
var group_etc1 = panel_etc.add("group", undefined, {name: "group_etc1"}); 
    group_etc1.orientation = "row"; 
    group_etc1.alignChildren = ["left","top"]; 
    group_etc1.spacing = 10; 
    group_etc1.margins = 0; 
    group_etc1.alignment = ["fill","center"]; 

var button_pdf = group_etc1.add("button", undefined, undefined, {name: "button_check"}); 
    button_pdf.text = "PDF 출력 (old)"; 
    button_pdf.preferredSize = [105, 30];
    
var file_pdf = File((Folder(app.activeScript)).parent + "/_Function/Autonics_PDF_Export.jsx");
button_pdf.onClick = function () {
    app.doScript(file_pdf);
    app.activate();
};

var button_toc = group_etc1.add("button", undefined, undefined, {name: "button_toc"}); 
    button_toc.text = "목차 생성"; 
    button_toc.preferredSize = [105, 30];

var file_toc = File((Folder(app.activeScript)).parent + "/_Function/Autonics_TOC.jsx");
button_toc.onClick = function () {
    app.doScript(file_toc, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "목차 생성");
    app.activate();
};

// GROUP_TRANS
// ===========
var group_trans = tab_function.add("group", undefined, {name: "group_trans"}); 
    group_trans.orientation = "row"; 
    group_trans.alignChildren = ["center","top"]; 
    group_trans.spacing = 10; 
    group_trans.margins = 0; 
    group_trans.alignment = ["fill","center"]; 

var statictext_trans = group_trans.add("statictext", undefined, undefined, {name: "statictext_trans"}); 
    statictext_trans.text = "투명도 설정"; 

var slider = group_trans.add("slider", undefined, 1, 0.3, 1);
slider.onChanging = function () {dialog.opacity = slider.value};

// TAB_ALIGNMENT
// =============
var tab_alignment = tpanel1.add("tab");
    tab_alignment.name = "3";
    tab_alignment.text = "정렬"; 
    tab_alignment.orientation = "column"; 
    tab_alignment.alignChildren = ["fill","top"]; 
    tab_alignment.spacing = 10; 
    tab_alignment.margins = [10, 20, 10, 0]; 

// GROUP_ROTATION
// ======================
var group_rotation = tab_alignment.add("group", undefined, {name: "group_rotation"}); 
    group_rotation.orientation = "row"; 
    group_rotation.alignChildren = ["left","center"]; 
    group_rotation.spacing = 10; 
    group_rotation.margins = 0; 

var check_vertical = group_rotation.add("radiobutton", undefined, undefined, {name: "check_vertical"}); 
    check_vertical.text = "수직 정렬"; 
    check_vertical.value = true;
    check_vertical.onClick = function () {
        if(check_vertical.value == true) {
            panel_fit.enabled = true;
            group_align_horizontal.size = [0, 0];    
            group_align_horizontal.visible = false;
            group_align_vertical.size = [210, 35];    
            group_align_vertical.visible = true;
            radio_left.value = true;
            radio_center.value = false;
            radio_right.value = false;
            radio_none.value = false;
            radio_top.value = false;
            radio_middle.value = false;
            radio_bottom.value = false;
            radio_none1.value = false;
        }
    }

var check_horizontal = group_rotation.add("radiobutton", undefined, undefined, {name: "check_horizontal"}); 
    check_horizontal.text = "수평 정렬"; 
    check_horizontal.onClick = function () {
        if(check_horizontal.value == true) {
            panel_fit.enabled = true;
            group_align_horizontal.size = [210, 35];    
            group_align_horizontal.visible = true;
            group_align_vertical.size = [0, 0];    
            group_align_vertical.visible = false;
            radio_left.value = false;
            radio_center.value = false;
            radio_right.value = false;
            radio_none.value = false;
            radio_top.value = true;
            radio_middle.value = false;
            radio_bottom.value = false;
            radio_none1.value = false;
        }
    }

var check_marginal = group_rotation.add("radiobutton", undefined, undefined, {name: "check_marginal"}); 
    check_marginal.text = "여백 정렬"; 
    check_marginal.onClick = function () {
        if(check_marginal.value == true) {
            panel_fit.enabled = false;
            radio_left.value = false;
            radio_center.value = false;
            radio_right.value = false;
            radio_none.value = false;
            radio_top.value = false;
            radio_middle.value = false;
            radio_bottom.value = false;
            radio_none1.value = false;
        }
    }

// PANEL_FIT
// =========
var panel_fit = tab_alignment.add("panel", undefined, undefined, {name: "panel_fit"}); 
    panel_fit.text = "맞춤"; 
    panel_fit.orientation = "column"; 
    panel_fit.alignChildren = ["fill","center"]; 
    panel_fit.spacing = 0; 
    panel_fit.margins = 15; 

// GROUP_ALIGN_HORIZONTAL
// ======================
var group_align_horizontal = panel_fit.add("group", undefined, {name: "group_align_horizontal"}); 
    group_align_horizontal.preferredSize.width = 210; 
    group_align_horizontal.preferredSize.height = 50; 
    group_align_horizontal.orientation = "row"; 
    group_align_horizontal.alignChildren = ["center","center"]; 
    group_align_horizontal.spacing = 10; 
    group_align_horizontal.margins = 0; 
    group_align_horizontal.size = [0, 0];

var radio_top_imgString = File((Folder(app.activeScript)).parent + "/_Icons/align_top.png");
var radio_top = group_align_horizontal.add("iconbutton", undefined, File.decode(radio_top_imgString), {name: "radio_top", style: "toolbutton", toggle: true}); 
    radio_top.preferredSize.width = 40;
    radio_top.onClick = function () {
        radio_left.value = false;
        radio_center.value = false;
        radio_right.value = false;
        radio_none.value = false;
        radio_top.value = true;
        radio_middle.value = false;
        radio_bottom.value = false;
        radio_none1.value = false;
    }
    
var radio_middle_imgString = File((Folder(app.activeScript)).parent + "/_Icons/align_hori.png");
var radio_middle = group_align_horizontal.add("iconbutton", undefined, File.decode(radio_middle_imgString), {name: "radio_middle", style: "toolbutton", toggle: true}); 
    radio_middle.preferredSize.width = 40;
    radio_middle.onClick = function () {
        radio_left.value = false;
        radio_center.value = false;
        radio_right.value = false;
        radio_none.value = false;
        radio_top.value = false;
        radio_middle.value = true;
        radio_bottom.value = false;
        radio_none1.value = false;
    }
    
var radio_bottom_imgString = File((Folder(app.activeScript)).parent + "/_Icons/align_bottom.png");
var radio_bottom = group_align_horizontal.add("iconbutton", undefined, File.decode(radio_bottom_imgString), {name: "radio_bottom", style: "toolbutton", toggle: true}); 
    radio_bottom.preferredSize.width = 40;
    radio_bottom.onClick = function () {
        radio_left.value = false;
        radio_center.value = false;
        radio_right.value = false;
        radio_none.value = false;
        radio_top.value = false;
        radio_middle.value = false;
        radio_bottom.value = true;
        radio_none1.value = false;
    }

var radio_none_imgString = File((Folder(app.activeScript)).parent + "/_Icons/align_none_1.png");
var radio_none1 = group_align_horizontal.add("iconbutton", undefined, File.decode(radio_none_imgString), {name: "radio_none1", style: "toolbutton", toggle: true}); 
    radio_none1.preferredSize.width = 40;
    radio_none1.onClick = function () {
        radio_left.value = false;
        radio_center.value = false;
        radio_right.value = false;
        radio_none.value = false;
        radio_top.value = false;
        radio_middle.value = false;
        radio_bottom.value = false;
        radio_none1.value = true;
    }

// GROUP_ALIGN_VERTICAL
// ====================
var group_align_vertical = panel_fit.add("group", undefined, {name: "group_align_vertical"}); 
    group_align_vertical.orientation = "row"; 
    group_align_vertical.alignChildren = ["center","center"]; 
    group_align_vertical.spacing = 10; 
    group_align_vertical.margins = 0; 
    group_align_vertical.size = [210, 35];
    
var radio_left_imgString = File((Folder(app.activeScript)).parent + "/_Icons/align_left.png");
var radio_left = group_align_vertical.add("iconbutton", undefined, File.decode(radio_left_imgString), {name: "radio_left", style: "toolbutton", toggle: true}); 
    radio_left.preferredSize.width = 40; 
    radio_left.value = true;
    radio_left.onClick = function () {
        radio_left.value = true;
        radio_center.value = false;
        radio_right.value = false;
        radio_none.value = false;
        radio_top.value = false;
        radio_middle.value = false;
        radio_bottom.value = false;
        radio_none1.value = false;
    }
    
var radio_center_imgString = File((Folder(app.activeScript)).parent + "/_Icons/align_vert.png");
var radio_center = group_align_vertical.add("iconbutton", undefined, File.decode(radio_center_imgString), {name: "radio_center", style: "toolbutton", toggle: true}); 
    radio_center.preferredSize.width = 40; 
    radio_center.onClick = function () {
        radio_left.value = false;
        radio_center.value = true;
        radio_right.value = false;
        radio_none.value = false;
        radio_top.value = false;
        radio_middle.value = false;
        radio_bottom.value = false;
        radio_none1.value = false;
    }

var radio_right_imgString = File((Folder(app.activeScript)).parent + "/_Icons/align_right.png");
var radio_right = group_align_vertical.add("iconbutton", undefined, File.decode(radio_right_imgString), {name: "radio_right", style: "toolbutton", toggle: true}); 
    radio_right.preferredSize.width = 40; 
    radio_right.onClick = function () {
        radio_left.value = false;
        radio_center.value = false;
        radio_right.value = true;
        radio_none.value = false;
        radio_top.value = false;
        radio_middle.value = false;
        radio_bottom.value = false;
        radio_none1.value = false;
    }

var radio_none = group_align_vertical.add("iconbutton", undefined, File.decode(radio_none_imgString), {name: "radio_none", style: "toolbutton", toggle: true}); 
    radio_none.preferredSize.width = 40; 
    radio_none.onClick = function () {
        radio_left.value = false;
        radio_center.value = false;
        radio_right.value = false;
        radio_none.value = true;
        radio_top.value = false;
        radio_middle.value = false;
        radio_bottom.value = false;
        radio_none1.value = false;
    }

// GROUP_ALIGN_BUTTON
// ==================
var group_align_button = tab_alignment.add("group", undefined, {name: "group_align_button"}); 
    group_align_button.orientation = "column"; 
    group_align_button.alignChildren = ["fill","top"]; 
    group_align_button.spacing = 15; 
    group_align_button.margins = 0; 

var button_5mm = group_align_button.add("button", undefined, undefined, {name: "button_5mm"}); 
    button_5mm.text = "5 mm 간격"; 
    button_5mm.preferredSize.height = 30; 

var file_5mm = File((Folder(app.activeScript)).parent + "/_Alignment/Autonics_Align.jsx");
button_5mm.onClick = function () {
    var margin_value = 5;
    app.doScript(file_5mm, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "개체 정렬");
    app.activate();
};    

var button_2mm = group_align_button.add("button", undefined, undefined, {name: "button_2mm"}); 
    button_2mm.text = "2 mm 간격"; 
    button_2mm.preferredSize.height = 30; 

var file_2mm = File((Folder(app.activeScript)).parent + "/_Alignment/Autonics_Align.jsx");
button_2mm.onClick = function () {
    var margin_value = 2;
    app.doScript(file_2mm, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "개체 정렬");
    app.activate();
};

var button_0mm = group_align_button.add("button", undefined, undefined, {name: "button_0mm"}); 
    button_0mm.text = "0 mm 간격"; 
    button_0mm.preferredSize.height = 30; 

var file_0mm = File((Folder(app.activeScript)).parent + "/_Alignment/Autonics_Align.jsx");
button_0mm.onClick = function () {
    var margin_value = 0;
    app.doScript(file_0mm, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "개체 정렬");
    app.activate();
};

// GROUP_ALIGN_CUSTOM
// ==================
var group_align_custom = tab_alignment.add("group", undefined, {name: "group_align_custom"}); 
    group_align_custom.orientation = "row"; 
    group_align_custom.alignChildren = ["left","center"]; 
    group_align_custom.spacing = 15; 
    group_align_custom.margins = [20,0,0,0]; 

var statictext3 = group_align_custom.add("statictext", undefined, undefined, {name: "statictext3"}); 
    statictext3.text = "직접 설정:"; 

var edit_size = group_align_custom.add('edittext {justify: "right", properties: {name: "edit_size"}}'); 
    edit_size.text = pref_align; 
    edit_size.preferredSize.width = 30; 
    edit_size.addEventListener("keydown", function(k) {handle_key (k, this, 10);});

var statictext4 = group_align_custom.add("statictext", undefined, undefined, {name: "statictext4"}); 
    statictext4.text = "mm"; 

var button_custom = group_align_custom.add("button", undefined, undefined, {name: "button_custom"}); 
    button_custom.text = "설정"; 
    button_custom.preferredSize.width = 75; 
    button_custom.preferredSize.height = 30; 

var file_custom = File((Folder(app.activeScript)).parent + "/_Alignment/Autonics_Align_manual.jsx");
button_custom.onClick = function () {
    app.doScript(file_custom, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "개체 정렬");
    app.activate();
};

// TAB_TABLE
// =========
var tab_table = tpanel1.add("tab");
    tab_table.name = "4";
    tab_table.text = "표"; 
    tab_table.orientation = "column"; 
    tab_table.alignChildren = ["fill","top"]; 
    tab_table.spacing = 5; 
    tab_table.margins = 10; 

// PANEL_SHADE
// ===========
var panel_shade = tab_table.add("panel", undefined, undefined, {name: "panel_shade"}); 
    panel_shade.text = "음영"; 
    panel_shade.orientation = "row"; 
    panel_shade.alignChildren = ["center","top"]; 
    panel_shade.spacing = 15; 
    panel_shade.margins = [15,15,15,10]; 

var radio_cell_none_imgString = File((Folder(app.activeScript)).parent + "/_Icons/table_none.png");
var radio_cell_none = panel_shade.add("iconbutton", undefined, File.decode(radio_cell_none_imgString), {name: "radio_cell_none", style: "toolbutton", toggle: true}); 
    radio_cell_none.value = pref_none; 
    radio_cell_none.onClick = function () {
        radio_cell_none.value = true; 
        radio_cell_row.value = false; 
        radio_cell_column.value = false; 
        radio_cell_cr.value = false; 
    }

var radio_cell_row_imgString = File((Folder(app.activeScript)).parent + "/_Icons/table_row.png");
var radio_cell_row = panel_shade.add("iconbutton", undefined, File.decode(radio_cell_row_imgString), {name: "radio_cell_row", style: "toolbutton", toggle: true}); 
    radio_cell_row.value = pref_row;
    radio_cell_row.onClick = function () {
        radio_cell_none.value = false; 
        radio_cell_row.value = true; 
        radio_cell_column.value = false; 
        radio_cell_cr.value = false; 
    }

var radio_cell_column_imgString = File((Folder(app.activeScript)).parent + "/_Icons/table_column.png");
var radio_cell_column = panel_shade.add("iconbutton", undefined, File.decode(radio_cell_column_imgString), {name: "radio_cell_column", style: "toolbutton", toggle: true}); 
    radio_cell_column.value = pref_column;
    radio_cell_column.onClick = function () {
        radio_cell_none.value = false; 
        radio_cell_row.value = false; 
        radio_cell_column.value = true; 
        radio_cell_cr.value = false; 
    }

var radio_cell_cr_imgString = File((Folder(app.activeScript)).parent + "/_Icons/table_rc.png");
var radio_cell_cr = panel_shade.add("iconbutton", undefined, File.decode(radio_cell_cr_imgString), {name: "radio_cell_cr", style: "toolbutton", toggle: true}); 
    radio_cell_cr.value = pref_cr;
    radio_cell_cr.onClick = function () {
        radio_cell_none.value = false; 
        radio_cell_row.value = false; 
        radio_cell_column.value = false; 
        radio_cell_cr.value = true; 
    }

// PANEL_TABLE
// ===========
var panel_table = tab_table.add("panel", undefined, undefined, {name: "panel_table"}); 
    panel_table.text = "기타 설정"; 
    panel_table.orientation = "column"; 
    panel_table.alignChildren = ["left","top"]; 
    panel_table.spacing = 5; 
    panel_table.margins = 15; 

// GROUP_TABLE
// ===========
var group_table = panel_table.add("group", undefined, {name: "group_table"}); 
    group_table.orientation = "row"; 
    group_table.alignChildren = ["left","center"]; 
    group_table.spacing = 10; 
    group_table.margins = 0; 
    group_table.alignment = ["fill","center"]; 

var check_width = group_table.add("checkbox", undefined, undefined, {name: "check_width"}); 
    check_width.text = "셀 너비 설정"; 
    check_width.value = pref_cell_width;

var edit_width = group_table.add('edittext {properties: {name: "edit_width"}}'); 
    edit_width.text = pref_width_value;
    edit_width.preferredSize.width = 120; 

if(check_width.value == false) {
    edit_width.enabled = false;
}

check_width.addEventListener('mousedown', function (evt) {
    edit_width.enabled = !edit_width.enabled;
});

// PANEL_INSET
// ===========
var panel_inset = tab_table.add("panel", undefined, undefined, {name: "panel_inset"}); 
    panel_inset.text = "인세트"; 
    panel_inset.orientation = "column"; 
    panel_inset.alignChildren = ["left","center"]; 
    panel_inset.spacing = 5; 
    panel_inset.margins = 15; 

// GROUP_USE_INSET
// ======
var group_use_inset = panel_inset.add("group", undefined, {name: "group_use_inset"}); 
    group_use_inset.orientation = "row"; 
    group_use_inset.alignChildren = ["left","center"]; 
    group_use_inset.spacing = 10; 

var radio_inset_table = group_use_inset.add("radiobutton", undefined, undefined, {name: "radio_inset_table"}); 
    radio_inset_table.text = "Table"; 
    radio_inset_table.value = 1;
    radio_inset_table.onClick = function () {
        panel1.enabled = false;
        edittext_left.text = "1";
        edittext_right.text = "0.5";
        edittext_top.text = "0.7";
        edittext_bottom.text = "0.7";
    }

var radio_inset_custom = group_use_inset.add("radiobutton", undefined, undefined, {name: "radio_inset_custom"}); 
    radio_inset_custom.text = "직접 설정"; 
    radio_inset_custom.onClick = function () {
        panel1.enabled = true;
    }

var radio_inset_disable = group_use_inset.add("radiobutton", undefined, undefined, {name: "radio_inset_disable"}); 
    radio_inset_disable.text = "설정 안함"; 
    radio_inset_disable.onClick = function () {
        panel1.enabled = false;
    }

// PANEL1
// ======
var panel1 = panel_inset.add("group", undefined, {name: "panel1"}); 
    panel1.orientation = "column"; 
    panel1.alignChildren = ["center","center"]; 
    panel1.spacing = 15; 
    panel1.margins =[15, 0, 0, 0]; 
    panel1.enabled = false;

// GROUP_INSET_1
// =============
var group_inset_1 = panel1.add("group", undefined, {name: "group_inset_1"}); 
    group_inset_1.orientation = "row"; 
    group_inset_1.alignChildren = ["left","center"]; 
    group_inset_1.spacing = 5; 
    group_inset_1.margins = 0; 

var inset_top_imgString = File((Folder(app.activeScript)).parent + "/_Icons/inset_top.png");
var inset_top = group_inset_1.add("image", undefined, File.decode(inset_top_imgString), {name: "inset_top"}); 

var edittext_top = group_inset_1.add('edittext {properties: {name: "edittext_top"}}'); 
    edittext_top.text = "0.7"; 
    edittext_top.preferredSize.width = 30; 
    edittext_top.addEventListener("keydown", function(k) {handle_key (k, this, 1);});

var unit_1 = group_inset_1.add("statictext", undefined, undefined, {name: "unit_1"}); 
    unit_1.text = "mm"; 
    unit_1.preferredSize.width = 30; 
    
var inset_left_imgString = File((Folder(app.activeScript)).parent + "/_Icons/inset_left.png");
var inset_left = group_inset_1.add("image", undefined, File.decode(inset_left_imgString), {name: "inset_left"}); 

var edittext_left = group_inset_1.add('edittext {justify: "right", properties: {name: "edittext_left"}}'); 
    edittext_left.text = "1"; 
    edittext_left.preferredSize.width = 30; 
    edittext_left.addEventListener("keydown", function(k) {handle_key (k, this, 1);});

var unit_2 = group_inset_1.add("statictext", undefined, undefined, {name: "unit_2"}); 
    unit_2.text = "mm"; 
    unit_2.preferredSize.width = 30; 
    
// GROUP_INSET_2
// =============
var group_inset_2 = panel1.add("group", undefined, {name: "group_inset_2"}); 
    group_inset_2.orientation = "row"; 
    group_inset_2.alignChildren = ["left","center"]; 
    group_inset_2.spacing = 5; 
    group_inset_2.margins = 0; 

var inset_bottom_imgString = File((Folder(app.activeScript)).parent + "/_Icons/inset_bottom.png");
var inset_bottom = group_inset_2.add("image", undefined, File.decode(inset_bottom_imgString), {name: "inset_bottom"}); 

var edittext_bottom = group_inset_2.add('edittext {properties: {name: "edittext_bottom"}}'); 
    edittext_bottom.text = "0.7"; 
    edittext_bottom.preferredSize.width = 30; 
    edittext_bottom.addEventListener("keydown", function(k) {handle_key (k, this, 1);});

var unit_3 = group_inset_2.add("statictext", undefined, undefined, {name: "unit_3"}); 
    unit_3.text = "mm"; 
    unit_3.preferredSize.width = 30; 
    
var inset_right_imgString = File((Folder(app.activeScript)).parent + "/_Icons/inset_right.png");
var inset_right = group_inset_2.add("image", undefined, File.decode(inset_right_imgString), {name: "inset_right"}); 

var edittext_right = group_inset_2.add('edittext {justify: "right", properties: {name: "edittext_right"}}'); 
    edittext_right.text = "0.5"; 
    edittext_right.preferredSize.width = 30; 
    edittext_right.preferredSize.height = 22; 
    edittext_right.addEventListener("keydown", function(k) {handle_key (k, this, 1);});

var unit_4 = group_inset_2.add("statictext", undefined, undefined, {name: "unit_4"}); 
    unit_4.text = "mm"; 
    unit_4.preferredSize.width = 30; 

// TAB_TABLE
// =========
var button_table_set = tab_table.add("button", undefined, undefined, {name: "button_table_set"}); 
    button_table_set.text = "설정"; 
    button_table_set.preferredSize.height = 30; 

var file_table = File((Folder(app.activeScript)).parent + "/_Table/Autonics_Table.jsx")
button_table_set.onClick = function () {
    app.doScript(file_table, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "표 설정");
    app.activate();
};  

// TPANEL1
// =======
tpanel1.selection = pref_tab; 

dialog.show();

dialog.onActivate = function () {
    if(app.documents.length != 0 && check_batch.value == 0) {
        file_loc = app.activeDocument.filePath;
        edittext_folder.text = Folder(file_loc).fullName;
    }
//~     dialog.opacity = 1;
}

//~ dialog.onDeactivate = function () {
//~     dialog.opacity = 0.5;
//~ }
    
dialog.onClose = function() {
    savePrefs();
}

function handle_key (key, control, mult) {
    var step;
    key.shiftKey ? step = 0.5 : step = 0.1;
    switch (key.keyName) {
        case "Up":
            control.text = String(Number(control.text) + step * mult);
            break;
        case "Down":
            control.text = String(Number(control.text) - step * mult);
            break;
    }
}

//~ var myDoc = app.documents.add();
//~ myDoc.addEventListener("afterSelectionChanged", asdf);

//~ function asdf () {
//~     alert("123124");
//~ }

//~ app.addEventListener ("beforeQuit", asdf);