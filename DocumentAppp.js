// --- DocumentAppp (DocumentApp plus) ---
(function(r) {
  var DocumentAppp;
  DocumentAppp = (function() {
    var gToM, putError, putInternalError;

    class DocumentAppp {
      constructor(id_) {
        this.name = "DocumentAppp";
        if (id_ !== "create") {
          if (id_ === "" || DriveApp.getFileById(id_).getMimeType() !== MimeType.GOOGLE_DOCS) {
            putError.call(this, "This file ID is not the file ID of Document.");
          }
          this.obj = {
            documentId: id_
          };
        }
        this.headers = {
          Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
        };
        this.mainObj = {};
      }

      // --- begin methods
      getTableColumnWidth() {
        gToM.call(this);
        return new WordApp(this.mainObj.blob).getTableColumnWidth();
      }

    };

    DocumentAppp.name = "DocumentAppp";

    // --- end methods
    gToM = function() {
      var obj, url;
      url = `https://www.googleapis.com/drive/v3/files/${this.obj.documentId}/export?mimeType=${MimeType.MICROSOFT_WORD}`;
      obj = UrlFetchApp.fetch(url, {
        headers: this.headers
      });
      if (obj.getResponseCode() !== 200) {
        putError.call(this, "Document ID might be not correct. Please check it again.");
      }
      this.mainObj.blob = obj.getBlob();
    };

    putError = function(m) {
      throw new Error(`${m}`);
    };

    putInternalError = function(m) {
      throw new Error(`Internal error: ${m}`);
    };

    return DocumentAppp;

  }).call(this);
  return r.DocumentAppp = DocumentAppp;
})(this);
