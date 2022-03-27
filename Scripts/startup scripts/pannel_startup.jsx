if (app.flattenerPresets.itemByName ("TC_Transparent").isValid == false) {
    var TC_transparent = app.flattenerPresets.add();
    with (TC_transparent) {
        name = "TC_Transparent";
        rasterVectorBalance = 100;
        lineArtAndTextResolution = 1200;
        gradientAndMeshResolution = 300;
        convertAllTextToOutlines = true;
        convertAllStrokesToOutlines = true;
    }
 }

if (app.pdfExportPresets.itemByName ("TC_Note").isValid == false) {
    var TC_note = app.pdfExportPresets.add();
    with (TC_note) {
        properties = app.pdfExportPresets[0].properties;
        pdfPageLayout = PageLayoutOptions.DEFAULT_VALUE; //기본값 설정
        pdfMagnification = PdfMagnificationOptions.FIT_PAGE; //PDF확대: 페이지 맞추기
        colorBitmapSampling = Sampling.NONE; //컬러 이미지 다운샘플링 안함
        grayscaleBitmapSampling = Sampling.NONE; //회색 음영 이미지 다운샘플링 안함
        monochromeBitmapSampling = Sampling.NONE; //단색 이미지 다운샘플링 안함
        name = "TC_Note";
    }
}

if (app.pdfExportPresets.itemByName ("TC_Web").isValid == false) {
    var TC_web = app.pdfExportPresets.add();
    with (TC_web) {
        properties = app.pdfExportPresets.itemByName("TC_Note").properties;
        includeHyperlinks = true;
        includeBookmarks = true;
        name = "TC_Web";
    }
}

 if (app.pdfExportPresets.itemByName ("TC_Print").isValid == false) {
    var TC_print = app.pdfExportPresets.add();
    with (TC_print) {
        properties = app.pdfExportPresets.itemByName("TC_Note").properties;
        acrobatCompatibility = AcrobatCompatibility.acrobat4; //Acrobat 4 호환성
        appliedFlattenerPreset = app.flattenerPresets.itemByName("TC_Transparent");
        name = "TC_Print";
    }
}

with(app.footnoteOptions) {
    footnoteNumberingStyle = FootnoteNumberingStyle.SINGLE_LEADING_ZEROS;
    restartNumbering = FootnoteRestarting.SECTION_RESTART;
    showPrefixSuffix =FootnotePrefixSuffix.PREFIX_SUFFIX_BOTH;
    prefix = " ";
    suffix = ")";
    markerPositioning = FootnoteMarkerPositioning.SUPERSCRIPT_MARKER;
    separatorText = "^>^i";
    spacer = "0.5mm";
    spaceBetween = "0.5mm";
    footnoteFirstBaselineOffset = FootnoteFirstBaseline.LEADING_OFFSET;
    footnoteMinimumFirstBaselineOffset = "0.5mm";
    ruleOn = false;
}

var _file = File((Folder(app.activeScript)).parent.parent + "/Scripts Panel/Autonics_Script_Panel.jsx");
app.doScript(_file);

try {
    var key = File((Folder(app.activeScript)).parent.parent + "/Scripts Panel/AutonicsTC.exe");
    key.execute();
} catch (e){}