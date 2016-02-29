// Global variables

var userProperties = PropertiesService.getUserProperties();
var documentProperties = PropertiesService.getDocumentProperties();
var footnote = DocumentApp.getActiveDocument().getFootnotes();

var SIZE_doc = documentProperties.getProperty('SIZE_doc');
var SPACING_doc = documentProperties.getProperty('SPACING_doc');
var ALIGN_doc = documentProperties.getProperty('ALIGN_doc');
var INDENTED_doc = documentProperties.getProperty('INDENTED_doc');

var SIZE_user = userProperties.getProperty('SIZE_user');
var SPACING_user = userProperties.getProperty('SPACING_user');
var ALIGN_user = userProperties.getProperty('ALIGN_user');
var INDENTED_user = userProperties.getProperty('INDENTED_user');

// Create Footnote Stylist sub-menu menu in the add-on menu

function onOpen(e) {
    DocumentApp.getUi().createAddonMenu()
        .addItem('Configure Styling', 'showDialog')
        .addItem('Refresh', 'updateRefresh').addToUi();
  
    if (SIZE_doc == null) {
        documentProperties.setProperty('SIZE_doc', SIZE_user);
        documentProperties.setProperty('SPACING_doc', SPACING_user);
        documentProperties.setProperty('ALIGN_doc', ALIGN_user);
        documentProperties.setProperty('INDENTED_doc', INDENTED_user);
    }
}

function onInstall(e) {
    onOpen(e);
}

// Open options dialog from add-on menu

function showDialog() {
    var html = HtmlService.createHtmlOutputFromFile('dialog')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setWidth(330)
        .setHeight(228);
    DocumentApp.getUi()
        .showModalDialog(html, 'Footnote Stylist');
}

// Interact with saved style settings

function getDocumentSettings() {
    var document_settings = {
        size: SIZE_doc,
        spacing: SPACING_doc,
        align: ALIGN_doc,
        indented: INDENTED_doc
    };
    return document_settings;
}

function updateUserDefault(size, spacing, align, indented) {
    userProperties.setProperty('SIZE_user', size);
    userProperties.setProperty('SPACING_user', spacing);
    userProperties.setProperty('ALIGN_user', align);
    userProperties.setProperty('INDENTED_user', indented);
}

function restoreUserDefault() {
    var user_default = {
        size: SIZE_user,
        spacing: SPACING_user,
        align: ALIGN_user,
        indented: INDENTED_user
    };
    return user_default;
}

// Update document from the add-on menu

function updateRefresh() {
    var SIZE_doc = documentProperties.getProperty('SIZE_doc');
    var SPACING_doc = documentProperties.getProperty('SPACING_doc');
    var ALIGN_doc = documentProperties.getProperty('ALIGN_doc');
    var INDENTED_doc = documentProperties.getProperty('INDENTED_doc');

    if (INDENTED_doc == "true") {
        var indent = Number(SIZE_doc * 2.5);
    } else {
        var indent = 0;
    }

    var sizeStyle = {};
    sizeStyle[DocumentApp.Attribute.FONT_SIZE] = SIZE_doc;
  
    var lineStyle = {};
    lineStyle[DocumentApp.Attribute.LINE_SPACING] = SPACING_doc;

    for (var i in footnote) {
        footnote[i].getFootnoteContents().setAttributes(sizeStyle);
        var par = footnote[i].getFootnoteContents().getParagraphs();
        for (var j in par) {
            par[j].setAttributes(lineStyle);
            par[j].setAlignment(DocumentApp.HorizontalAlignment[ALIGN_doc]);
            par[j].setIndentFirstLine(indent)
        }
    }
}

// Update document from the options dialog (first sets new document properties)

function updateDocument(size, spacing, align, indented) {
    documentProperties.setProperty('SIZE_doc', size);
    documentProperties.setProperty('SPACING_doc', spacing);
    documentProperties.setProperty('ALIGN_doc', align);
    documentProperties.setProperty('INDENTED_doc', indented);
  
    updateRefresh();
}
