#targetengine 'session'

var doc = app.activeDocument;

var list_file = app.open (Folder((app.activeScript).parent + "/wordlist.indd"), false); //wordlist.indd파일 열기 (보이지 않게)

//배열 생성
var found_list = [];
var found_contents = [];

var script_run = app.doScript(convert_SCRIPT, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "용어 일괄 변경");
list_file.close(SaveOptions.NO);

app.findGrepPreferences = app.changeGrepPreferences = null;
alert("완료되었습니다.", "Autonics TC");

function convert_SCRIPT() {
    //wordlist.indd 파일 내 표 읽어들임
    var list = list_file.textFrames[0].tables[0];
    
    for (var i = 0; i < list.rows.length; i++) {
        var from = list.rows[i].cells[0].contents; 
        var to = list.rows[i].cells[1].contents;
        
        app.findGrepPreferences = app.changeGrepPreferences = null;
        convert (from, to); //변환 함수 실행
    }
    
    app.findGrepPreferences = app.changeGrepPreferences = null;
    
    //줄바꿈 뒤에 공백 들어가는거 삭제
    app.findGrepPreferences.findWhat="\\n[[:space:]]"; //
    app.changeGrepPreferences.changeTo="\\n"; //
    doc.changeGrep(); //GREP 변경

    //진행결과 창 생성
    var result_panel = new Window ("palette");
    var result_list = result_panel.add ("listbox", undefined, found_contents);
    result_list.maximumSize.width = 700;
    result_list.maximumSize.height = 500;
    result_list.onChange = function () {
        var id_temp = found_list[result_list.selection.index];
        id_temp.select();
        app.activeWindow.zoomPercentage = 600;
    }

    if (found_contents.length > 0) {
        result_panel.show();
    }
    
    //변환 함수
    //이게 바꾸고 변경사항을 표시하려고 했는데 해당 문구 전후로만 찾고싶은데안되서
    //try함수로 하나씩 다 뽑아놨어요
    //코드가 이상하긴 한데 동작은 하고 아직까진 문제가 없었으니 괜찮을꺼에요
    function convert (f, t) {
        if (f != "") {
            app.findGrepPreferences.findWhat = f;
            app.changeGrepPreferences.changeTo = t;
            var found = doc.findGrep();
            doc.changeGrep();
            if (found.length > 0) {
                for (var i = 0; i < found.length; i++) {
                    found_list.push (found[i]);
                    try {
                        found_contents.push(found[i].lines[0].contents);
                    } 
                    catch (e) {
                        try {
                            found_contents.push(found[i].paragraphs[0].contents);
                        } 
                        catch (e) {
                            try {
                                found_contents.push(found[i].parent.contents);
                            }
                            catch (e) {
                                found_contents.push(found[i].contents);
                            }
                        }
                    }
                }
            }
        }
    }
}