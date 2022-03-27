//초기 값 설정
var export_inst = 0;
var export_overview = 0;
var export_manual = 0;
var export_note = 0;

var doc_preview = false;
var export_idml = 0;

var save_loc = "";
//~ var save_loc = Folder(app.activeDocument.filePath).fullName;

var export_batch = 0;
var view_output = 0;

//스크립트 패널 UI 설정 읽어들이기
try {
    export_inst = check_inst.value;
    export_overview = check_overview.value;
    export_manual = check_manual.value;
    export_note = check_note.value;
    
    doc_preview = check_preview.value;
    export_idml = check_idml.value;
    
    save_loc = edittext_folder.text;
    
    export_batch = check_batch.value;
    export_subdir = check_subfolder.value;
    
    view_output = check_open.value;
} catch(e) {}

//일괄이 아닐 경우 단일 실행
if (export_batch == 0) {
    var doc = app.activeDocument;
    paragraph_style();
    export_PDF();
    alert("완료되었습니다.", "Autonics TC");
}

//미리보기가 아니고 일괄일 경우 
else if (doc_preview == 0 && export_batch == 1) {
    var load_files = Folder(save_loc).getFiles ("*.indd");
    for(var i = 0; i < load_files.length; i++) {
        var doc = app.open(load_files[i], true);
        paragraph_style();
        export_PDF();
        doc.close(SaveOptions.NO);
    }
    if (export_subdir == 1) {
        var load_folders = Folder(save_loc).getFiles();
        for(var j = 0; j < load_folders.length; j++) {
            if (load_folders[j] instanceof Folder) {
                var load_files = load_folders[j].getFiles ("*.indd");
                for(var k = 0; k < load_files.length; k++) {
                    var doc = app.open(load_files[k], true);
                    paragraph_style();
                    export_PDF();
                    doc.close(SaveOptions.NO);
                }
            }
        }
    }
    alert("완료되었습니다.", "Autonics TC");
}

function export_PDF() {
    var cur_doc = app.activeDocument.name; //현재 활성화 중인 문서 이름 변수 지정
    var doc = app.documents.itemByName(cur_doc); //현재 활성화 중인 문서 변수 지정
    var file_count = ((export_inst  * 2)+ (export_overview) + (export_manual) + (export_note));
    
    var loading_menu = new Window("palette");
    var loading_text = loading_menu.add("statictext", undefined, "출력 시작"); //로딩창 생성
    loading_text.alignment = "left";
    loading_text.bounds = [0, 0, 340, 50];
    
    var progress_bar = loading_menu.add ("progressbar", [12, 12, 350, 24], 0, file_count); //진행막대 추가
    var progress_text = loading_menu.add ("statictext", undefined, "출력 준비 중..."); //문자 추가
    progress_text.bounds = [0, 0, 340, 20]; //문자 크기 설정
    progress_text.justify = "right"; //우측 정렬
    
    loading_menu.add("statictext", undefined, "Esc를 연타하면 취소됩니다.");
    
    if(export_inst == 0 && export_overview == 0 && export_manual == 0 && export_note == 0) {
        alert("출력할 문서를 설정해 주세요.", "Autonics TC");
    }
    else if(save_loc == "" || save_loc ==  "경로를 설정해주세요.") {
        alert("경로를 설정해 주세요.", "Autonics TC");
    }
    
    else {
        var progress_count = 1;
        loading_menu.show();
        loading_menu.active = true;
        
        var key_stats = ScriptUI.environment.keyboardState; //키보드 설정
        if (key_stats.keyName == "Escape") { //ESC를 누를 경우
            new_doc.close(SaveOptions.no); //열려있는 파일 닫기 (저장 안함)
            loading_menu.close(); //로딩 창 닫기
        }
        
        if (doc_preview == false) {
            loading_menu.onClose = function () {
                new_doc.close(SaveOptions.no);
            }
        }
        
        var doc_name = doc.name.substr (0, doc.name.lastIndexOf (".")); //파일명 1번째부터 위치 전까지 설정

//////////////////// 취급설명서 출력 //////////////////// 
        if(export_inst == 1) {
            loading_text.text = String (doc.name + "\n" + "취급설명서 출력 중... "); //로딩창 띄우기
            var file_inst = File((Folder(app.activeScript)).parent + "/Autonics_Export_Inst.jsx");
            app.doScript(file_inst, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "문서 출력");
        }
    
//////////////////// 카탈로그 출력 //////////////////// 
        if(export_overview == 1) {
            loading_text.text = String (doc.name + "\n" + " 카탈로그 출력 중... "); //로딩창 띄우기
            progress_text.text = String ("출력 준비 중...   " + (progress_count-1) + " / " + file_count + " 완료"); //진행상황 텍스트 표시
            var file_overview = File((Folder(app.activeScript)).parent + "/Autonics_Export_P-Overview.jsx");
            app.doScript(file_overview, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "문서 출력");
        }
 
 //////////////////// 제품 매뉴얼 출력 //////////////////// 
        if(export_manual == 1 || export_note == 1) {
            loading_text.text = String (doc.name + "\n" + " 제품 매뉴얼 출력 중... "); //로딩창 띄우기
            progress_text.text = String ("출력 준비 중...   " + (progress_count-1) + " / " + file_count + " 완료"); //진행상황 텍스트 표시
            var file_manual = File((Folder(app.activeScript)).parent + "/Autonics_Export_P-Manual.jsx");
            app.doScript(file_manual, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "문서 출력");
        }
        
        loading_menu.close();
    }
}


function paragraph_style() {
    try {
        var par_1 = doc.paragraphStyles.itemByName("Cover) Documnet_type / Black");
        par_1.name = "Cover) Document_type / Black";
    } catch (e) {}
    //표지, 카탈로그 제품매뉴얼 용어 변경 (모든 문서 적용 완료 시 삭제 가능 145~182줄)
    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
    
    if (doc.conditions.itemByName("제품 개요서").isValid) {
        app.findGrepPreferences.appliedParagraphStyle = "Cover) Document_type / Black";
        app.findGrepPreferences.findWhat = "[가-힣].+";
        app.changeGrepPreferences.changeTo = "카탈로그제품 매뉴얼";
        doc.changeGrep();
        app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
        
        app.findGrepPreferences.appliedParagraphStyle = "Cover) Document_type / Black";
        app.findGrepPreferences.findWhat = "카탈로그";
        app.changeGrepPreferences.appliedConditions = ["제품 개요서"];
        doc.changeGrep();
        app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
        
        app.findGrepPreferences.appliedParagraphStyle = "Cover) Document_type / Black";
        app.findGrepPreferences.findWhat = "제품 매뉴얼";
        app.changeGrepPreferences.appliedConditions = ["제품 매뉴얼"];
        doc.changeGrep();
        app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
        
        app.findGrepPreferences.appliedParagraphStyle = "Cover) Document_type / Black";
        app.findGrepPreferences.findWhat = "[A-z].+";
        app.changeGrepPreferences.changeTo = "CATALOGPRODUCT MANUAL";
        doc.changeGrep();
        app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
        
        app.findGrepPreferences.appliedParagraphStyle = "Cover) Document_type / Black";
        app.findGrepPreferences.findWhat = "CATALOG";
        app.changeGrepPreferences.appliedConditions = ["제품 개요서"];
        doc.changeGrep();
        app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
        
        app.findGrepPreferences.appliedParagraphStyle = "Cover) Document_type / Black";
        app.findGrepPreferences.findWhat = "PRODUCT MANUAL";
        app.changeGrepPreferences.appliedConditions = ["제품 매뉴얼"];
        doc.changeGrep();
        app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
    }
    
    //가끔 헤딩의 스타일이 누락되어 책갈피가 추가안되는 경우가 있어 단락스타일 지정을 위해 짜둠
    app.findGrepPreferences.appliedFont = "Autonics_TC [2020]";
    app.findGrepPreferences.fontStyle = "Bold";
    app.findGrepPreferences.pointSize = 9;
    
    var found = doc.findGrep();
    for (var j = 0; j < found.length; j++) {
        if (found[j].parent instanceof Cell && found[j].parent.cells[0].fillColor.name == "Black" && found[j].parent.cells[0].fillTint == 15) {
            var font_temp = [];
            for ( var k = 0; k < found[j].parent.cells[0].characters.length; k++) {
                font_temp.push(found[j].parent.cells[0].characters[k].appliedFont.name);
            }
            found[j].parent.cells[0].paragraphs[0].appliedParagraphStyle = doc.paragraphStyles.itemByName("Heading) Title");
            for ( var k = 0; k < found[j].parent.cells[0].characters.length; k++) {
                found[j].parent.cells[0].characters[k].appliedFont = font_temp[k];
            }
        }
    }
    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
    
    //링크된 텍스트 프레임 신규 문서에서도 링크 유지
    var _link = [];
    for (var i = 0; i < doc.links.length; i++) {
        if (doc.links[i].linkType == "Internal Linked Object") {
            _link.push(doc.links[i].parent.parentPage.name);
            doc.links[i].update();
        }
        
        else if (doc.links[i].linkType == "Internal Linked Story") {
            _link.push(doc.links[i].parent.textContainers[0].parentPage.name);
            doc.links[i].update();
        }
    }
    
    for (var i = 0; i < _link.length; i++) {
        doc.pages.itemByName(_link[i]).pageColor = UIColors.BLUE;
    }
}

function def_PDFpref (_layout, pw) { //PDF 초기값 함수 설정
    with (app.pdfExportPreferences) { //PDF 출력 설정
        pageRange = PageRange.allPages; //출력범위 모든 페이지
        if (pw == 1) {
            useSecurity = true;
            changeSecurityPassword = "auto1234"; //비밀번호
            disallowChanging = true; //수정 금지 설정
            disallowCopying = true; //복사 금지 설정
            disallowExtractionForAccessibility =true; //추출 금지 설정
        }
    
        viewPDF = false; //PDF 열기
        if (view_output == 1) { //파일열기 체크박스 설정 시
            viewPDF = true; //PDF 열기
        }
    }   
}