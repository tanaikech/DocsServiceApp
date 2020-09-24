// --- ExcelApp ---
(function(r) {
  var ExcelApp;
  ExcelApp = (function() {
    var a1NotationToRowCol, columnToLetter, disassembleExcel, getCommentsAsObject, getImagesAsObject, getSharedStrings, getStyleAsObject, getValuesAsObject, getValuesFromObj, getXmlObj, letterToColumn, parsXLSX, putError, putInternalError, sheetNameToFileName;

    class ExcelApp {
      constructor(blob_) {
        this.name = "ExcelApp";
        if (!blob_ || blob_.getContentType() !== MimeType.MICROSOFT_EXCEL) {
          throw new Error("Please set the blob of data of XLSX format.");
        }
        this.obj = {
          excel: blob_
        };
        this.contentTypes = "[Content_Types].xml";
        this.sharedStrings = "xl/sharedStrings.xml";
        this.workbook = "xl/workbook.xml";
        this.styles = "xl/styles.xml";
        this.mainObj = {};
        parsXLSX.call(this);
      }

      // --- begin methods
      getSheetByName(sheetName_) {
        if (!sheetName_ || sheetName_.toString() === "") {
          putError.call(this, "No sheet name.");
        }
        this.obj.sheetName = sheetName_;
        return this;
      }

      getAll() {
        return this.mainObj.sheetsObj.sheetAr.map(({sheetName}, i) => {
          this.obj.sheetName = sheetName;
          return {
            sheetName: sheetName,
            values: getValuesAsObject.call(this),
            images: getImagesAsObject.call(this),
            comments: getCommentsAsObject.call(this)
          };
        });
      }

      getSheets() {
        return this.mainObj.sheetsObj.sheetAr.map(({sheetName}, i) => {
          return {
            index: i,
            sheetName: sheetName
          };
        });
      }

      getValues() {
        return getValuesFromObj.call(this, "value");
      }

      getFormulas() {
        return getValuesFromObj.call(this, "formula");
      }

      getImages() {
        return getImagesAsObject.call(this);
      }

      getComments() {
        return getCommentsAsObject.call(this);
      }

    };

    ExcelApp.name = "ExcelApp";

    // --- end methods
    parsXLSX = function() {
      disassembleExcel.call(this);
      getSharedStrings.call(this);
      getStyleAsObject.call(this);
      sheetNameToFileName.call(this);
    };

    disassembleExcel = function() {
      var blobs;
      blobs = Utilities.unzip(this.obj.excel.setContentType(MimeType.ZIP));
      this.mainObj.fileObj = blobs.reduce((o, b) => {
        return Object.assign(o, {
          [b.getName()]: b
        });
      }, {});
    };

    getSharedStrings = function() {
      var rootChildren, xmlObj;
      if (this.mainObj.fileObj.hasOwnProperty(this.sharedStrings)) {
        xmlObj = getXmlObj.call(this, this.sharedStrings);
        rootChildren = xmlObj.getRootElement().getChildren();
        this.mainObj.valueSharedStrings = rootChildren.map((e) => {
          return e.getChild("t", e.getNamespace()).getValue();
        });
      } else {
        putInternalError.call(this, "This Excel data cannot be analyzed.");
      }
    };

    sheetNameToFileName = function() {
      var rootChildren, sheets, xmlObj;
      if (this.mainObj.fileObj.hasOwnProperty(this.workbook)) {
        xmlObj = getXmlObj.call(this, this.workbook);
        rootChildren = xmlObj.getRootElement();
        sheets = rootChildren.getChild("sheets", rootChildren.getNamespace()).getChildren();
        return this.mainObj.sheetsObj = sheets.reduce((o, e) => {
          var sheetId, sheetName;
          sheetId = e.getAttribute("sheetId").getValue();
          sheetName = e.getAttribute("name").getValue();
          Object.assign(o.sheetsObj, {
            [sheetName]: {
              sheet: `xl/worksheets/sheet${sheetId}.xml`,
              rels: `xl/worksheets/_rels/sheet${sheetId}.xml.rels`
            }
          });
          o.sheetAr.push({
            sheetIndex: sheetId,
            sheetName: sheetName
          });
          return o;
        }, {
          sheetsObj: {},
          sheetAr: []
        });
      } else {
        return putInternalError.call(this, "This Excel data cannot be analyzed.");
      }
    };

    getStyleAsObject = function() {
      var cellXfs, cellXfsC, fills, fillsC, fillsObj, fonts, fontsC, fontsObj, numFmts, numFmtsC, numFmtsObj, root, styleObj, xmlObj;
      if (this.mainObj.fileObj.hasOwnProperty(this.styles)) {
        xmlObj = getXmlObj.call(this, this.styles);
        root = xmlObj.getRootElement();
        numFmtsC = root.getChild("numFmts", root.getNamespace());
        if (numFmtsC) {
          numFmts = numFmtsC.getChildren();
          numFmtsObj = numFmts.reduce((o, e) => {
            return Object.assign(o, {
              [e.getAttribute("numFmtId").getValue()]: e.getAttribute("formatCode").getValue()
            });
          }, {});
        }
        fontsC = root.getChild("fonts", root.getNamespace());
        if (fontsC) {
          fonts = fontsC.getChildren();
          fontsObj = fonts.map((e) => {
            return e.getChildren().reduce((o, f) => {
              return Object.assign(o, {
                [f.getName()]: f.getAttributes().reduce((o2, g) => {
                  return Object.assign(o2, {
                    [g.getName()]: g.getValue() || true
                  });
                }, {})
              });
            }, {});
          });
        }
        fillsC = root.getChild("fills", root.getNamespace());
        if (fillsC) {
          fills = fillsC.getChildren();
          fillsObj = fills.map((e) => {
            var c, children, temp;
            c = e.getChild("patternFill", e.getNamespace());
            temp = {
              patternType: c.getAttribute("patternType").getValue()
            };
            children = c.getChildren();
            if (children.length > 0) {
              children.forEach((f) => {
                return temp[f.getName()] = f.getAttributes().reduce((o, g) => {
                  return Object.assign(o, {
                    [g.getName()]: g.getValue() || true
                  });
                }, {});
              });
            }
            return temp;
          });
        }
        cellXfsC = root.getChild("cellXfs", root.getNamespace());
        if (cellXfsC) {
          cellXfs = cellXfsC.getChildren();
          styleObj = cellXfs.map((e) => {
            return e.getAttributes().reduce((o, f) => {
              var c, n, v;
              n = f.getName();
              v = f.getValue();
              switch (n) {
                case "numFmtId":
                  o.numFmt = numFmtsObj ? numFmtsObj[v] || null : null;
                  break;
                case "fontId":
                  o.font = fontsObj ? fontsObj[v] || null : null;
                  break;
                case "fillId":
                  o.fill = fillsObj ? fillsObj[v] || null : null;
              }
              c = e.getChild("alignment", e.getNamespace());
              if (c) {
                o.alignment = c.getAttributes().reduce((o2, g) => {
                  return Object.assign(o2, {
                    [g.getName()]: g.getValue() || true
                  });
                }, {});
              }
              return o;
            }, {});
          });
          return this.mainObj.styleObj = styleObj;
        }
      } else {
        return putInternalError.call(this, "Style file cannot be used.");
      }
    };

    getValuesAsObject = function() {
      var checkFilename, root, rows, sheetData, xmlObj;
      if (this.mainObj.sheetsObj.sheetsObj.hasOwnProperty(this.obj.sheetName)) {
        checkFilename = this.mainObj.sheetsObj.sheetsObj[this.obj.sheetName].sheet;
        if (this.mainObj.fileObj.hasOwnProperty(checkFilename)) {
          xmlObj = getXmlObj.call(this, checkFilename);
          root = xmlObj.getRootElement();
          sheetData = root.getChild("sheetData", root.getNamespace()).getChildren();
          rows = [];
          sheetData.forEach((s) => {
            var cols, rowNum;
            cols = [];
            rowNum = 0;
            s.getChildren().forEach((c) => {
              var cObj, fv, fval, ra, sa, ta, tav, temp, tobj, vv, vval;
              vval = c.getChild("v", c.getNamespace());
              fval = c.getChild("f", c.getNamespace());
              ra = c.getAttribute("r");
              ta = c.getAttribute("t");
              sa = c.getAttribute("s");
              vv = "";
              fv = "";
              if (ta) {
                tav = ta.getValue();
                if (tav === "s") {
                  if (vval) {
                    vv = this.mainObj.valueSharedStrings[vval.getValue()];
                  }
                  if (fval) {
                    fv = fval.getValue();
                  }
                } else if (tav === "str") {
                  if (vval) {
                    vv = vval.getValue();
                  }
                  if (fval) {
                    fv = fval.getValue();
                  }
                } else if (tav === "b") {
                  if (vval) {
                    vv = vval.getValue() === "1" ? true : false;
                  }
                  if (fval) {
                    fv = fval.getValue();
                  }
                } else {
                  if (vval) {
                    vv = vval.getValue();
                  }
                  if (fval) {
                    fv = fval.getValue();
                  }
                }
              } else {
                if (vval) {
                  vv = Number(vval.getValue());
                }
                if (fval) {
                  fv = fval.getValue();
                }
              }
              if (ra) {
                temp = ra.getValue();
                cObj = a1NotationToRowCol.call(this, temp);
                rowNum = cObj.row;
                tobj = {};
                tobj.value = {};
                tobj.range = {};
                tobj.value.value = vv.toString() !== "" ? vv : null;
                tobj.value.formula = fv.toString() !== "" ? `=${fv}` : null;
                tobj.range.col = cObj.col;
                tobj.range.row = cObj.row;
                tobj.range.a1Notation = temp;
                if (sa) {
                  tobj.style = this.mainObj.styleObj[sa.getValue()];
                }
                return cols[cObj.col - 1] = tobj;
              }
            });
            return rows[rowNum - 1] = cols;
          });
          return rows;
        } else {
          putInternalError.call(this, "No sheet file.");
        }
      } else {
        putInternalError.call(this, "No sheet name.");
      }
    };

    getImagesAsObject = function() {
      var drawingFilename, imageObj, rootChildren, xmlObj;
      if (!this.mainObj.sheetsObj.sheetsObj.hasOwnProperty(this.obj.sheetName)) {
        putError.call(this, "No sheet name.");
      }
      drawingFilename = this.mainObj.sheetsObj.sheetsObj[this.obj.sheetName].sheet.replace("worksheets", "drawings").replace("sheet", "drawing");
      if (this.mainObj.fileObj.hasOwnProperty(drawingFilename)) {
        xmlObj = getXmlObj.call(this, drawingFilename);
        rootChildren = xmlObj.getRootElement().getChildren();
        imageObj = rootChildren.map((e) => {
          var cNvPr, description, filename, form, formn, n, nvPicPr, nvPicPrn, pic, sp, temp, title;
          temp = {};
          n = e.getNamespace();
          form = e.getChild("from", n);
          formn = form.getNamespace();
          temp.range = {
            col: Number(form.getChild("col", formn).getValue()) + 1,
            colOff: Number(form.getChild("colOff", formn).getValue()),
            row: Number(form.getChild("row", formn).getValue()) + 1,
            rowOff: Number(form.getChild("rowOff", formn).getValue())
          };
          temp.range.a1Notation = (columnToLetter.call(this, temp.range.col)) + temp.range.row;
          pic = e.getChild("pic", n);
          sp = e.getChild("sp", n); // Drawing
          if (pic) {
            nvPicPr = pic.getChild("nvPicPr", pic.getNamespace());
            nvPicPrn = nvPicPr.getNamespace();
            cNvPr = nvPicPr.getChild("cNvPr", nvPicPrn);
            filename = "xl/media/" + cNvPr.getAttribute("name").getValue();
            temp.image = {};
            description = cNvPr.getAttribute("descr");
            if (description) {
              temp.image.description = description.getValue();
            }
            title = cNvPr.getAttribute("title");
            if (title) {
              temp.image.title = title.getValue();
            }
            temp.image.blob = this.mainObj.fileObj[filename].setName(filename.split("/").pop()) || null;
            temp.image.innerCell = temp.range.colOff === 0 && temp.range.rowOff === 0 ? true : false;
          } else if (sp) {
            temp.drawing = {
              message: "In the current stage, the object of drawing cannot be directly retrieved as the blob."
            };
          }
          delete temp.range.colOff;
          delete temp.range.rowOff;
          return temp;
        });
        return imageObj;
      } else {
        return putInternalError.call(this, "Internal error: Image files cannot be retrieved.");
      }
    };

    getCommentsAsObject = function() {
      var checkFilename, commentFilename, commentList, comments, f, i, j, ref1, root, rootChildren, t, target, temp, xmlObj;
      commentFilename = "";
      if (this.mainObj.sheetsObj.sheetsObj.hasOwnProperty(this.obj.sheetName)) {
        checkFilename = this.mainObj.sheetsObj.sheetsObj[this.obj.sheetName].rels;
        if (this.mainObj.fileObj.hasOwnProperty(checkFilename)) {
          xmlObj = getXmlObj.call(this, checkFilename);
          rootChildren = xmlObj.getRootElement().getChildren();
          for (i = j = 0, ref1 = rootChildren.length; (0 <= ref1 ? j < ref1 : j > ref1); i = 0 <= ref1 ? ++j : --j) {
            target = rootChildren[i].getAttribute("Target");
            if (target) {
              t = target.getValue();
              if (t.includes("comments")) {
                commentFilename = t;
                break;
              }
            }
          }
        } else {
          putInternalError.call(this, "No sheet file.");
        }
      } else {
        putError.call(this, "No sheet name.");
      }
      temp = commentFilename.split("\/")[1];
      f = Object.entries(this.mainObj.fileObj).filter(([k, v]) => {
        return k.includes(temp);
      });
      comments = [];
      if (f.length > 0) {
        xmlObj = XmlService.parse(f[0][1].getDataAsString());
        root = xmlObj.getRootElement();
        commentList = root.getChild("commentList", root.getNamespace()).getChildren();
        comments = commentList.reduce((ar, e) => {
          var cObj, commentObj, ref, tempComment, text, tmp;
          temp = {};
          ref = e.getAttribute("ref");
          if (ref) {
            tmp = ref.getValue();
            cObj = a1NotationToRowCol.call(this, tmp);
            temp.range = {};
            temp.range.col = cObj.col;
            temp.range.row = cObj.row;
            temp.range.a1Notation = tmp;
            text = e.getChild("text", e.getNamespace());
            t = text.getChild("t", text.getNamespace());
            if (t) {
              tempComment = t.getValue();
              comments = tempComment.split(/\t\-\w.+/);
              commentObj = [...tempComment.matchAll(/\t\-\w.+/g)].map(([h]) => {
                return h.replace(/\t-/, "");
              }).map((h, l) => {
                return {
                  user: h,
                  comment: comments[l]
                };
              });
              temp.comment = commentObj;
            }
            ar.push(temp);
          }
          return ar;
        }, []);
      }
      return comments;
    };

    getValuesFromObj = function(v_) {
      var maxColumnLength, obj, values;
      obj = getValuesAsObject.call(this);
      maxColumnLength = obj.reduce((c, e) => {
        if (c < e.length) {
          return e.length;
        } else {
          return c;
        }
      }, 0);
      values = obj.reduce((ar, e) => {
        var temp;
        temp = new Array(maxColumnLength).fill("");
        e.forEach(({value, range}) => {
          return temp[range.col - 1] = value[v_] === null ? "" : value[v_];
        });
        ar.push(temp);
        return ar;
      }, []);
      return values;
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

    a1NotationToRowCol = function(a1Notation_) {
      var num, str;
      str = a1Notation_.match(/[a-zA-Z]+/)[0];
      num = Number(a1Notation_.match(/\d+/)[0]);
      return {
        row: num,
        col: letterToColumn.call(this, str)
      };
    };

    columnToLetter = function(column) {
      var letter, temp;
      temp = 0;
      letter = "";
      while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
      }
      return letter;
    };

    letterToColumn = function(letter) {
      var column, i, j, length, ref1;
      column = 0;
      length = letter.length;
      for (i = j = 0, ref1 = length; (0 <= ref1 ? j < ref1 : j > ref1); i = 0 <= ref1 ? ++j : --j) {
        column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
      }
      return column;
    };

    return ExcelApp;

  }).call(this);
  return r.ExcelApp = ExcelApp;
})(this);
