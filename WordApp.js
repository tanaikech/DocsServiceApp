// --- WordApp ---
(function(r) {
  var WordApp;
  WordApp = (function() {
    var disassembleWord, getXmlObj, parsDOCX, putError, putInternalError;

    class WordApp {
      constructor(blob_) {
        this.name = "WordApp";
        if (!blob_ || blob_.getContentType() !== MimeType.MICROSOFT_WORD) {
          throw new Error("Please set the blob of data of DOCX format.");
        }
        this.obj = {
          excel: blob_
        };
        this.contentTypes = "[Content_Types].xml";
        this.document = "word/document.xml";
        this.mainObj = {};
        parsDOCX.call(this);
      }

      // --- begin methods
      getTableColumnWidth() {
        var body, n1, obj, root, xmlObj;
        if (this.mainObj.fileObj.hasOwnProperty(this.document)) {
          xmlObj = getXmlObj.call(this, this.document);
          root = xmlObj.getRootElement();
          n1 = root.getNamespace("w");
          body = root.getChild("body", n1).getChildren("tbl", n1);
          obj = body.map((e, i) => {
            var tblGrid, tblPr, tblW, temp, w;
            temp = {
              tableIndex: i,
              unit: "pt"
            };
            tblPr = e.getChild("tblPr", n1);
            if (tblPr) {
              tblW = tblPr.getChild("tblW", n1);
              if (tblW) {
                w = tblW.getAttribute("w", n1);
                if (w) {
                  temp.tableWidth = Number(w.getValue()) / 20;
                }
              }
            }
            tblGrid = e.getChild("tblGrid", n1);
            if (tblGrid) {
              temp.tebleColumnWidth = tblGrid.getChildren("gridCol", n1).map((f) => {
                return Number(f.getAttribute("w", n1).getValue()) / 20;
              });
            }
            return temp;
          });
          return obj;
        }
      }

    };

    WordApp.name = "WordApp";

    // --- end methods
    parsDOCX = function() {
      disassembleWord.call(this);
    };

    disassembleWord = function() {
      var blobs;
      blobs = Utilities.unzip(this.obj.excel.setContentType(MimeType.ZIP));
      this.mainObj.fileObj = blobs.reduce((o, b) => {
        return Object.assign(o, {
          [b.getName()]: b
        });
      }, {});
    };

    getXmlObj = function(k_) {
      return XmlService.parse(this.mainObj.fileObj[k_].getDataAsString());
    };

    putError = function(m) {
      throw new Error(`${m}`);
    };

    putInternalError = function(m) {
      throw new Error(`Internal error: ${m}`);
    };

    return WordApp;

  }).call(this);
  return r.WordApp = WordApp;
})(this);
