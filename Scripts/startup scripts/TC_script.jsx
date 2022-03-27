#targetengine "myTestMenu"
 
var myFolder = Folder(app.activeScript.path);
myFolder = myFolder.parent + '/Scripts Panel/';

var panel_Handler = function(/*onInvoke*/){
    app.doScript(File(myFolder + 'Autonics_Script_Panel.jsx'));
};

var checklist_Handler = function(/*onInvoke*/){
    app.doScript(File(myFolder + '/_Function/Autonics_Checklist.jsx'));
};

var checklist_HandlerDisplay = function(/*beforeDisplay*/){
    _checklist.enabled = (app.documents.length>0);
};

var pdfout_Handler = function(/*onInvoke*/){
    app.doScript(File(myFolder + '/_Function/Autonics_PDF_Export.jsx'));
};

var bookmark_Handler = function(/*onInvoke*/){
    app.doScript(File(myFolder + '/_Function/Autonics_Bookmarker.jsx'));
};

var bookmark_HandlerDisplay = function(/*beforeDisplay*/){
    _bookmark.enabled = (app.documents.length>0);
};

var save_Handler = function(/*onInvoke*/){
    app.doScript(File(myFolder + '/_Document/Autonics_Save.jsx'));
};

var save_HandlerDisplay = function(/*beforeDisplay*/){
    _save.enabled = (app.documents.length>0);
};

var close_Handler = function(/*onInvoke*/){
    app.doScript(File(myFolder + 'Autonics_Close_All_DOC.jsx'));
};

var cell_Handler = function(/*onInvoke*/){
    app.doScript(File(myFolder + 'Autonics_Cell_Number.jsx'));
};

var cell_HandlerDisplay = function(/*beforeDisplay*/){
    _cell.enabled = (app.documents.length>0);
};

var autosizing_Handler = function(/*onInvoke*/){
    app.doScript(File(myFolder + 'Autonics_Textframe_Autosizing.jsx'));
};

var autosizing_HandlerDisplay = function(/*beforeDisplay*/){
    _autosizing.enabled = (app.documents.length>0);
};

var hyphen_Handler = function(/*onInvoke*/){
    app.doScript(File(myFolder + '/_Document/Autonics_Hyphen_remover.jsx'));
};

var hyphen_HandlerDisplay = function(/*beforeDisplay*/){
    _hyphen.enabled = (app.documents.length>0);
};

var outline_Handler = function(/*onInvoke*/){
    app.doScript(File(myFolder + 'Autonics_text_to_outline.jsx'));
};

var outline_HandlerDisplay = function(/*beforeDisplay*/){
    _outline.enabled = (app.documents.length>0);
};

var grep_Handler = function(/*onInvoke*/){
    app.doScript(File(myFolder + 'Autonics_outline_to_text.jsx')); 
};

var grep_HandlerDisplay = function(/*beforeDisplay*/){
    _grep.enabled = (app.documents.length>0);
};

var menuInstaller = menuInstaller||(function(){
    var panel_menu = "스크립트 패널(&S)",
    checklist_menu = "검토하기 (&C)",
    pdfout_menu = "PDF 변환(&P)",
    bookmark_menu = "책갈피 추가(&B)",
    save_menu = "저장하기 (&V)",
    close_menu = "모든 문서 닫기 (&Q)",
    cell_menu = "셀 번호 채우기(&F)",
    hyphen_menu = "하이픈 제거(&H)",
    outline_menu = "글리프 깨기(&O)",
    grep_menu = "글리프 복원(&G)",
    autosizing_menu = "텍스트 프레임 크기 조정(&A)",
    menuT = "TC팀(&C)",
    subs = app.menus.item("$ID/Main").submenus, sma, mnu; 
    menuT.shortcutKey = "c";
    checklist_menu.shortcutKey = "c";
    pdfout_menu.shortcutKey = "p";
    bookmark_menu.shortcutKey = "b";
    panel_menu.shortcutKey = "s";
    save_menu.shortcutKey = "v";
    close_menu.shortcutKey = "q";
    cell_menu.shortcutKey = "f";
    outline_menu.shortcutKey = "o";
    grep_menu.shortcutKey = "g";
    autosizing_menu.shortcutKey = "a";

var refItem = app.menus.item("$ID/Main").submenus.item("$ID/&Layout");

    _panel = app.scriptMenuActions.add(panel_menu);
    _panel.eventListeners.add("onInvoke", panel_Handler);
 
    _checklist = app.scriptMenuActions.add(checklist_menu);
    _checklist.eventListeners.add("onInvoke", checklist_Handler);
    _checklist.eventListeners.add("beforeDisplay", checklist_HandlerDisplay);
    
    _pdfout = app.scriptMenuActions.add(pdfout_menu);
    _pdfout.eventListeners.add("onInvoke", pdfout_Handler);
    
    _bookmark = app.scriptMenuActions.add(bookmark_menu);
    _bookmark.eventListeners.add("onInvoke", bookmark_Handler);
    _bookmark.eventListeners.add("beforeDisplay", bookmark_HandlerDisplay);
    
    _save = app.scriptMenuActions.add(save_menu);
    _save.eventListeners.add("onInvoke", save_Handler);
    _save.eventListeners.add("beforeDisplay", save_HandlerDisplay);
    
    _close = app.scriptMenuActions.add(close_menu);
    _close.eventListeners.add("onInvoke", close_Handler);    
        
    _cell = app.scriptMenuActions.add(cell_menu);
    _cell.eventListeners.add("onInvoke", cell_Handler);
    _cell.eventListeners.add("beforeDisplay", cell_HandlerDisplay);
    
    _hyphen = app.scriptMenuActions.add(hyphen_menu);
    _hyphen.eventListeners.add("onInvoke", hyphen_Handler);
    _hyphen.eventListeners.add("beforeDisplay", hyphen_HandlerDisplay);
    
    _outline = app.scriptMenuActions.add(outline_menu);
    _outline.eventListeners.add("onInvoke", outline_Handler);
    _outline.eventListeners.add("beforeDisplay", outline_HandlerDisplay);
  
    _grep = app.scriptMenuActions.add(grep_menu);
    _grep.eventListeners.add("onInvoke", grep_Handler);
    _grep.eventListeners.add("beforeDisplay", grep_HandlerDisplay);
    
    _autosizing = app.scriptMenuActions.add(autosizing_menu);
    _autosizing.eventListeners.add("onInvoke", autosizing_Handler);
    _autosizing.eventListeners.add("beforeDisplay", autosizing_HandlerDisplay);  
    
    mnu = subs.item(menuT);
    if(mnu.isValid) mnu = subs.item(menuT).remove();
    mnu = subs.add(menuT, LocationOptions.after, refItem);
        
    mnu.menuItems.add(_panel);
    
    mnu.menuSeparators.add();
    
    mnu.menuItems.add(_checklist);
    mnu.menuItems.add(_pdfout);
    mnu.menuItems.add(_bookmark);
    
    mnu.menuSeparators.add();
    
    mnu.menuItems.add(_save);
    mnu.menuItems.add(_close);
    
    mnu.menuSeparators.add();
    
    mnu.menuItems.add(_cell);
    mnu.menuItems.add(_hyphen);
    
    mnu.menuSeparators.add();
    
    mnu.menuItems.add(_autosizing);
    
    mnu.menuSeparators.add();
    
    mnu.menuItems.add(_outline);
    mnu.menuItems.add(_grep);
    
    return true;
}) ();

