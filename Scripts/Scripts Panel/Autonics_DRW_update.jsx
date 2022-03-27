var doc = app.activeDocument;

app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
app.findGrepPreferences.appliedParagraphStyle = doc.paragraphStyles.itemByName("Cover) DRW_No / Black");

var found = doc.findGrep();

if (found[0].contents.length > 11) {
    var drw_no = found[0].contents.substring(0, 11);
    var update = drw_no[drw_no.length-1];
    var increment = String.fromCharCode(update.charCodeAt() + 1);
    var number = found[0].contents.substring(0, 10) + increment;
    app.changeGrepPreferences.changeTo = number;
    doc.changeGrep();
    
    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
    app.findGrepPreferences.appliedParagraphStyle = doc.paragraphStyles.itemByName("Cover) DRW_No / White");
    app.changeGrepPreferences.changeTo = number;
    doc.changeGrep();
}

app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
