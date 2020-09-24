/**
 * GitHub  https://github.com/tanaikech/DocsServiceApp<br>
 * Library name
 * @type {string}
 * @const {string}
 * @readonly
 */
const appName = "DocsServiceApp";

/**
 * @param {String} id Spreasheet ID.
 * @return {DocsServiceApp}
 */
function openBySpreadsheetId(id) {
    return new SpreadsheetAppp(id);
}

/**
 * @param {Object} blob Blob of Excel file (XLSX file).
 * @return {DocsServiceApp}
 */
function openByExcelFileBlob(blob) {
    return new ExcelApp(blob);
}

/**
 * @param {String} id Document ID.
 * @return {DocsServiceApp}
 */
function openByDocumentId(id) {
    return new DocumentAppp(id);
}

/**
 * @param {Object} blob Blob of Word file (DOCX file).
 * @return {DocsServiceApp}
 */
function openByWordFileBlob(blob) {
    return new WordApp(blob);
}

/**
 * @param {object} object Object including parameter for createing new Google Spreadsheet.
 * @return {string} Presentation ID of cerated Google Slides.
 */
function createNewSpreadsheetWithCustomHeaderFooter(object) {
    return new SpreadsheetAppp("create").createNewSpreadsheetWithCustomHeaderFooter(object);
}

/**
 * @param {object} object Object including parameter for createing new Google Slides.
 * @return {string} Presentation ID of cerated Google Slides.
 */
function createNewSlidesWithPageSize(object) {
    return new SlidesAppp("create").createNewSlidesWithPageSize(object);
}

// DriveApp.createFile()  // This is used for automatically detected the scope of "https://www.googleapis.com/auth/drive"
// SpreadsheetApp.create()  // This is used for automatically detected the scope of "https://www.googleapis.com/auth/spreadsheets"
// SlidesApp.create(name)  // This is used for automatically detected the scope of "https://www.googleapis.com/auth/presentations"
;

