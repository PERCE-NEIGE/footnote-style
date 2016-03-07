// Create Footnote Stylist sub-menu menu in the add-on menu
function onInstall(e) {
    onOpen(e);
}

function onOpen(e) {
    DocumentApp.getUi().createAddonMenu()
        .addItem('Configure Styling', 'showDialog')
        .addItem('Refresh', 'updateRefresh')
        .addToUi();
}


// Open options dialog from add-on menu

function showDialog() {

    var html = HtmlService.createHtmlOutputFromFile('dialog')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setWidth(330)
        .setHeight(228);
    DocumentApp.getUi()
        .showModalDialog(html, 'Configure footnote styling')
}

// Interact with saved style settings

function getDocumentSettings() {

    var documentProperties = PropertiesService.getDocumentProperties();

    if (documentProperties.getProperty('SIZE_doc') == null) {
        documentProperties.setProperty('SIZE_doc', "10.0");
        documentProperties.setProperty('SPACING_doc', "1.0");
        documentProperties.setProperty('ALIGN_doc', 'LEFT');
        documentProperties.setProperty('INDENTED_doc', 'false')
    }
    var document_settings = {
        size: documentProperties.getProperty('SIZE_doc'),
        spacing: documentProperties.getProperty('SPACING_doc'),
        align: documentProperties.getProperty('ALIGN_doc'),
        indented: documentProperties.getProperty('INDENTED_doc')
    };

    return document_settings;
}

function updateUserDefault(size, spacing, align, indented) {

    var userProperties = PropertiesService.getUserProperties();

    userProperties.setProperty('SIZE_user', size);
    userProperties.setProperty('SPACING_user', spacing);
    userProperties.setProperty('ALIGN_user', align);
    userProperties.setProperty('INDENTED_user', indented);
}

function restoreUserDefault() {

    var userProperties = PropertiesService.getUserProperties();

    var user_default = {
        size: userProperties.getProperty('SIZE_user'),
        spacing: userProperties.getProperty('SPACING_user'),
        align: userProperties.getProperty('ALIGN_user'),
        indented: userProperties.getProperty('INDENTED_user')
    };
    return user_default;
}

function noFootnotes() {
    var ui = DocumentApp.getUi();

    ui.alert(
        "There don't seem to be any footnotes in the document.\nPlease try adding one and run again.",
        ui.ButtonSet.OK);
}

function noConfig() {
    var ui = DocumentApp.getUi();

    ui.alert(
        "You appear not to have configured styling for this document.\nPlease try configuring styling and run again.",
        ui.ButtonSet.OK);
}


// Update document from the add-on menu

function updateRefresh() {

    var documentProperties = PropertiesService.getDocumentProperties();
    var userProperties = PropertiesService.getUserProperties();
    var footnote = DocumentApp.getActiveDocument().getFootnotes();

    var SIZE_doc = documentProperties.getProperty('SIZE_doc');
    var SPACING_doc = documentProperties.getProperty('SPACING_doc');
    var ALIGN_doc = documentProperties.getProperty('ALIGN_doc');
    var INDENTED_doc = documentProperties.getProperty('INDENTED_doc');

    var SIZE_user = userProperties.getProperty('SIZE_user');

    var length = footnote.length

    if (length == "0.0") {
        noFootnotes()
    }

    if (SIZE_doc == null) {
        noConfig()
    } else {

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
}

// Update document from the options dialog (first sets new document properties)

function updateDocument(size, spacing, align, indented) {

    var documentProperties = PropertiesService.getDocumentProperties();

    documentProperties.setProperty('SIZE_doc', size);
    documentProperties.setProperty('SPACING_doc', spacing);
    documentProperties.setProperty('ALIGN_doc', align);
    documentProperties.setProperty('INDENTED_doc', indented);

    updateRefresh();
}

// For testing purposes only

function clearProps() {
    var documentProperties = PropertiesService.getDocumentProperties();
    var userProperties = PropertiesService.getUserProperties();

    documentProperties.deleteAllProperties()
    userProperties.deleteAllProperties()
}

function showProps() {
    var props = PropertiesService.getDocumentProperties().getProperties();
    var propsu = PropertiesService.getUserProperties().getProperties();
    Logger.log(props);
    Logger.log(propsu);
}
