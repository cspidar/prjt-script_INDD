var doc = app.activeDocument;

app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
app.findGrepPreferences.findWhat = "[가-힣]";
app.findGrepPreferences.appliedParagraphStyle = doc.paragraphStyles.itemByName("Cover) Document_type / Black");

var lang = doc.findGrep();

if (lang.length > 0) {
    var list  = ["모델 구성", "제품 구성품", "소프트웨어", "정격/성능", "통신 인터페이스"];
}
else {
    var list  = ["Ordering Information", "Product Components", "Software", "Specifications", "Communication Interface"];
}

var found_list = [];

var new_spread = doc.spreads.add(LocationOptions.AT_END);
new_spread.pages[0].pageColor = UIColors.BLUE; //페이지 색상 지정

for (var i = 0; i < list.length; i++) {
    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
    app.findGrepPreferences.findWhat = list[i];
    app.findGrepPreferences.appliedParagraphStyle = doc.paragraphStyles.itemByName("Heading) Title");
    var found = doc.findGrep();
    if (found.length != 0) {
        for ( var j = 0 ; j < 5; j++) {
            if (!(found[0] instanceof TextFrame)) {
                found[0] = found[0].parent;
            }
        }
        found_list[i] = found[0];
        
        doc.spreads.lastItem().contentPlace(found_list[i], true, true, true, [0, (i*50)]);
    }
}
 if (doc.spreads.lastItem().pageItems.length >1) {
     doc.align(doc.spreads.lastItem().pageItems.everyItem(), AlignOptions.LEFT_EDGES,  AlignDistributeBounds.MARGIN_BOUNDS); //왼쪽 정렬
     doc.distribute(doc.spreads.lastItem().pageItems.everyItem(), DistributeOptions.VERTICAL_SPACE, AlignDistributeBounds.MARGIN_BOUNDS, true, 5); //간격을 기준으로 세로 정렬
}

else {
    doc.align(doc.spreads.lastItem().pageItems.everyItem(), AlignOptions.TOP_EDGES,  AlignDistributeBounds.MARGIN_BOUNDS); //왼쪽 정렬
    doc.align(doc.spreads.lastItem().pageItems.everyItem(), AlignOptions.LEFT_EDGES,  AlignDistributeBounds.MARGIN_BOUNDS); //왼쪽 정렬
}

if (lang.length > 0) {
    var list  = ["외형치수도", "각부의 명칭", "별매품"];
}
else {
    var list  = ["Dimensions", "Unit Descriptions", "Sold Separately"];
}

var found_list = [];

var new_spread = doc.spreads.add(LocationOptions.AT_END);
new_spread.pages[0].pageColor = UIColors.BLUE; //페이지 색상 지정

for (var i = 0; i < list.length; i++) {
    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
    app.findGrepPreferences.findWhat = list[i];
    app.findGrepPreferences.appliedParagraphStyle = doc.paragraphStyles.itemByName("Heading) Title");
    var found = doc.findGrep();
    if (found.length != 0) {
        for ( var j = 0 ; j < 5; j++) {
            if (!(found[0] instanceof TextFrame)) {
                found[0] = found[0].parent;
            }
        }
        found_list[i] = found[0];
        
        doc.spreads.lastItem().contentPlace(found_list[i], true, true, true, [0, (i*50)]);
    }
}    

if (doc.spreads.lastItem().pageItems.length >1) {
    doc.align(doc.spreads.lastItem().pageItems.everyItem(), AlignOptions.LEFT_EDGES,  AlignDistributeBounds.MARGIN_BOUNDS); //왼쪽 정렬
    doc.distribute(doc.spreads.lastItem().pageItems.everyItem(), DistributeOptions.VERTICAL_SPACE, AlignDistributeBounds.MARGIN_BOUNDS, true, 5); //간격을 기준으로 세로 정렬
}

else {
    doc.align(doc.spreads.lastItem().pageItems.everyItem(), AlignOptions.TOP_EDGES,  AlignDistributeBounds.MARGIN_BOUNDS); //왼쪽 정렬
    doc.align(doc.spreads.lastItem().pageItems.everyItem(), AlignOptions.LEFT_EDGES,  AlignDistributeBounds.MARGIN_BOUNDS); //왼쪽 정렬
}