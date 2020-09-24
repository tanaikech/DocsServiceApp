// --- SpreadsheetAppp (SpreadsheetApp plus) ---
(function(r) {
  var SpreadsheetAppp;
  SpreadsheetAppp = (function() {
    var ContentTypesXml_, createDrawing1Xml_, createSheet1Xml_, drawing1XmlRels_, gToM, imagesToObj, newXLSXdata, putError, putInternalError, setheaderFooter, xlsxObjToBlob;

    class SpreadsheetAppp {
      constructor(id_) {
        this.name = "SpreadsheetAppp";
        if (id_ !== "create") {
          if (id_ === "" || DriveApp.getFileById(id_).getMimeType() !== MimeType.GOOGLE_SHEETS) {
            putError.call(this, "This file ID is not the file ID of Spreadsheet.");
          }
          this.obj = {
            spreadsheetId: id_
          };
        }
        this.headers = {
          Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
        };
        this.mainObj = {};
      }

      // --- begin methods
      getSheetByName(sheetName_) {
        if (!sheetName_ || sheetName_.toString() === "") {
          putError.call(this, "No sheet name.");
        }
        this.obj.sheetName = sheetName_;
        return this;
      }

      getImages() {
        gToM.call(this);
        if (this.obj.hasOwnProperty("sheetName")) {
          return new ExcelApp(this.mainObj.blob).getSheetByName(this.obj.sheetName).getImages();
        } else {
          return new ExcelApp(this.mainObj.blob).getAll().map(({sheetName, images}) => {
            return {sheetName, images};
          });
        }
      }

      getComments() {
        gToM.call(this);
        if (this.obj.hasOwnProperty("sheetName")) {
          return new ExcelApp(this.mainObj.blob).getSheetByName(this.obj.sheetName).getComments();
        } else {
          return new ExcelApp(this.mainObj.blob).getAll().map(({sheetName, comments}) => {
            return {sheetName, comments};
          });
        }
      }

      insertImage(objAr_) {
        var ar, blob, dstSS, dstSheet, dstSheetId, e, requests, tmpId, tmpSheet, tmpSheetId, xlsxObj;
        if (!Array.isArray(objAr_) || (objAr_.some(({blob, range}) => {
          var height, identification, width;
          if (blob.toString() !== "Blob" || !range.row || !range.column || isNaN(range.row) || isNaN(range.column)) {
            return true;
          }
          ({width, height, identification} = ImgApp.getSize(blob));
          if (width * height > 1048576) {
            return true;
          }
          if (!(["GIF", "PNG", "JPG"].some((e) => {
            return e === identification;
          }))) {
            return true;
          }
          return false;
        }))) {
          putError.call(this, "Wrong object. Please confirm it again. By the way, the maximum image size is 'width x height < 1048576'. And the mimeTypes are PNG, JPG and GIF.");
        }
        xlsxObj = newXLSXdata.call(this);
        ar = imagesToObj.call(this, xlsxObj, objAr_);
        ContentTypesXml_.call(this, xlsxObj, ar);
        createSheet1Xml_.call(this, xlsxObj, ar);
        createDrawing1Xml_.call(this, xlsxObj, ar);
        drawing1XmlRels_.call(this, xlsxObj, ar);
        blob = xlsxObjToBlob.call(this, xlsxObj);
        try {
          tmpId = Drive.Files.insert({
            title: "SpreadsheetAppp_temp",
            mimeType: MimeType.GOOGLE_SHEETS
          }, blob).id;
        } catch (error) {
          e = error;
          if (e.message === "Drive is not defined") {
            putError.call(this, "Please enable Drive API at Advanced Google services, and try again.");
          } else {
            putError.call(this, e.message);
          }
        }
        dstSS = SpreadsheetApp.openById(this.obj.spreadsheetId);
        dstSheet = dstSS.getSheetByName(this.obj.sheetName);
        dstSheetId = dstSheet.getSheetId();
        tmpSheet = SpreadsheetApp.openById(tmpId).getSheets()[0].setName(`SpreadsheetAppp_${Utilities.getUuid()}`).copyTo(dstSS);
        DriveApp.getFileById(tmpId).setTrashed(true);
        tmpSheetId = tmpSheet.getSheetId();
        requests = ar.map((e) => {
          e.from.sheetId = tmpSheetId;
          e.to.sheetId = dstSheetId;
          return {
            copyPaste: {
              source: e.from,
              destination: e.to,
              pasteType: "PASTE_VALUES"
            }
          };
        });
        try {
          Sheets.Spreadsheets.batchUpdate({
            requests: requests
          }, this.obj.spreadsheetId);
        } catch (error) {
          e = error;
          if (e.message === "Sheets is not defined") {
            putError.call(this, "Please enable Sheets API at Advanced Google services, and try again.");
          } else {
            putError.call(this, e.message);
          }
        }
        dstSS.deleteSheet(tmpSheet);
        return null;
      }

      createNewSpreadsheetWithCustomHeaderFooter(obj_) {
        var blob, createObj, e, tmpId, xlsxObj;
        if (!obj_ || Object.keys(obj_).length === 0) {
          putError.call(this, "Object was not found. Please confirm it again.");
        }
        xlsxObj = newXLSXdata.call(this);
        setheaderFooter.call(this, obj_, xlsxObj);
        blob = xlsxObjToBlob.call(this, xlsxObj);
        createObj = {
          title: "SpreadsheetSample",
          mimeType: MimeType.GOOGLE_SHEETS
        };
        if (obj_.hasOwnProperty("title")) {
          createObj.title = obj_.title;
        }
        if (obj_.hasOwnProperty("parent")) {
          createObj.parents = [
            {
              id: obj_.parent
            }
          ];
        }
        try {
          tmpId = Drive.Files.insert(createObj, blob).id;
        } catch (error) {
          e = error;
          if (e.message === "Drive is not defined") {
            putError.call(this, "Please enable Drive API at Advanced Google services, and try again.");
          } else {
            putError.call(this, e.message);
          }
        }
        return tmpId;
      }

    };

    SpreadsheetAppp.name = "SpreadsheetAppp";

    // --- end methods
    gToM = function() {
      var obj, url;
      url = `https://www.googleapis.com/drive/v3/files/${this.obj.spreadsheetId}/export?mimeType=${MimeType.MICROSOFT_EXCEL}`;
      obj = UrlFetchApp.fetch(url, {
        headers: this.headers
      });
      if (obj.getResponseCode() !== 200) {
        putError.call(this, "Spreadsheet ID might be not correct. Please check it again.");
      }
      this.mainObj.blob = obj.getBlob();
    };

    newXLSXdata = function() {
      var data;
      data = "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";
      return JSON.parse(Utilities.unzip(Utilities.newBlob(Utilities.base64Decode(data), MimeType.ZIP))[0].getDataAsString());
    };

    imagesToObj = function(obj_, ar_) {
      var dupCheckAr;
      dupCheckAr = [];
      return ar_.reduce((ar, e, i) => {
        var dupObj, ext, fileSize, filename, imgFilename, mimeType, orgFilename, tempObj;
        orgFilename = e.blob.getName();
        fileSize = e.blob.getBytes().length;
        mimeType = e.blob.getContentType();
        ext = "";
        switch (mimeType) {
          case MimeType.PNG:
            ext = "png";
            break;
          case MimeType.JPEG:
            ext = "jpg";
            break;
          case MimeType.GIF:
            ext = "gif";
            break;
          case MimeType.BMP:
            e.blob = e.blob.getAs(MimeType.PNG);
            ext = "png";
            break;
          default:
            putError.call(this, "In the current stage, this file type cannot be used.");
        }
        filename = `xl/media/image${i + 1}.${ext}`;
        dupObj = dupCheckAr.filter((e) => {
          return e.orgFilename === orgFilename && e.fileSize === fileSize && e.mimeType === mimeType;
        });
        if (dupObj.length === 0) {
          imgFilename = `image${i + 1}.${ext}`;
          dupCheckAr.push({
            filename: imgFilename,
            orgFilename: orgFilename,
            fileSize: fileSize,
            mimeType: mimeType
          });
          tempObj = {
            range: `A${i + 1}`,
            filename: imgFilename,
            rowIndex: i,
            mimeType: mimeType,
            from: {
              startRowIndex: i,
              endRowIndex: i + 1,
              startColumnIndex: 0,
              endColumnIndex: 1
            },
            to: {
              startRowIndex: e.range.row - 1,
              endRowIndex: e.range.row,
              startColumnIndex: e.range.column - 1,
              endColumnIndex: e.range.column
            }
          };
          ar.push(tempObj);
          obj_[filename] = e.blob.copyBlob().setName(filename);
        } else {
          tempObj = {
            range: `A${i + 1}`,
            filename: dupObj[0].filename,
            rowIndex: i,
            mimeType: mimeType,
            from: {
              startRowIndex: i,
              endRowIndex: i + 1,
              startColumnIndex: 0,
              endColumnIndex: 1
            },
            to: {
              startRowIndex: e.range.row - 1,
              endRowIndex: e.range.row,
              startColumnIndex: e.range.column - 1,
              endColumnIndex: e.range.column
            }
          };
          ar.push(tempObj);
        }
        return ar;
      }, []);
    };

    ContentTypesXml_ = function(obj, ar) {
      var filename, mimeTypeToExtension, mimeTypes, n, root, xmlObj;
      mimeTypeToExtension = {
        [MimeType.PNG]: "png",
        [MimeType.JPEG]: "jpg",
        [MimeType.GIF]: "gif"
      };
      mimeTypes = [
        ...new Set(ar.map(({mimeType}) => {
          return mimeType;
        }))
      ];
      filename = "[Content_Types].xml";
      xmlObj = XmlService.parse(obj[filename]);
      root = xmlObj.getRootElement();
      n = root.getNamespace();
      mimeTypes.forEach((e) => {
        var Default;
        Default = XmlService.createElement("Default", n).setAttribute("ContentType", e).setAttribute("Extension", mimeTypeToExtension[e]);
        return root.addContent(Default);
      });
      obj[filename] = XmlService.getRawFormat().format(root);
    };

    createSheet1Xml_ = function(obj, ar) {
      var filename, n, root, xmlObj;
      filename = "xl/worksheets/sheet1.xml";
      xmlObj = XmlService.parse(obj[filename]);
      root = xmlObj.getRootElement();
      n = root.getNamespace();
      ar.forEach((e) => {
        var c, row;
        c = XmlService.createElement("c", n).setAttribute("r", e.range).setAttribute("s", "1");
        row = XmlService.createElement("row", n).setAttribute("r", e.rowIndex + 1).addContent(c);
        return root.getChild("sheetData", n).addContent(row);
      });
      obj[filename] = XmlService.getRawFormat().format(root);
    };

    createDrawing1Xml_ = function(obj, ar) {
      var filename, n, n2, n3, root, xmlObj;
      filename = "xl/drawings/drawing1.xml";
      xmlObj = XmlService.parse(obj[filename]);
      root = xmlObj.getRootElement();
      n = root.getNamespace("xdr");
      n2 = root.getNamespace("r");
      n3 = root.getNamespace("a");
      ar.forEach((e, i) => {
        var avLst, blip, blipFill, cNvPicPr, cNvPr, clientData, col, colOff, ext, fillRect, form, nvPicPr, oneCellAnchor, pic, prstGeom, row, rowOff, spPr, stretch;
        col = XmlService.createElement("col", n).setText(0);
        colOff = XmlService.createElement("colOff", n).setText(0);
        row = XmlService.createElement("row", n).setText(e.rowIndex);
        rowOff = XmlService.createElement("rowOff", n).setText(0);
        form = XmlService.createElement("from", n).addContent(col).addContent(colOff).addContent(row).addContent(rowOff);
        ext = XmlService.createElement("ext", n).setAttribute("cx", "314325").setAttribute("cy", "200025");
        cNvPr = XmlService.createElement("cNvPr", n).setAttribute("id", "0").setAttribute("name", e.filename);
        cNvPicPr = XmlService.createElement("cNvPicPr", n).setAttribute("preferRelativeResize", "0");
        nvPicPr = XmlService.createElement("nvPicPr", n).addContent(cNvPr).addContent(cNvPicPr);
        blip = XmlService.createElement("blip", n3).setAttribute("cstate", "print").setAttribute("embed", `rId${i + 1}`, n2);
        fillRect = XmlService.createElement("fillRect", n3);
        stretch = XmlService.createElement("stretch", n3).addContent(fillRect);
        blipFill = XmlService.createElement("blipFill", n).addContent(blip).addContent(stretch);
        avLst = XmlService.createElement("avLst", n3);
        prstGeom = XmlService.createElement("prstGeom", n3).setAttribute("prst", "rect").addContent(avLst);
        spPr = XmlService.createElement("spPr", n).addContent(prstGeom);
        pic = XmlService.createElement("pic", n).addContent(nvPicPr).addContent(blipFill).addContent(spPr);
        clientData = XmlService.createElement("clientData", n).setAttribute("fLocksWithSheet", "0");
        oneCellAnchor = XmlService.createElement("oneCellAnchor", n).addContent(form).addContent(ext).addContent(pic).addContent(clientData);
        return root.addContent(oneCellAnchor);
      });
      obj[filename] = XmlService.getRawFormat().format(root);
    };

    drawing1XmlRels_ = function(obj, ar) {
      var Relationships, filename, n, root;
      filename = "xl/drawings/_rels/drawing1.xml.rels";
      root = XmlService.createDocument();
      n = XmlService.getNamespace("http://schemas.openxmlformats.org/package/2006/relationships");
      Relationships = XmlService.createElement("Relationships").setNamespace(n);
      ar.forEach(({rowIndex, filename}) => {
        var Relationship;
        Relationship = XmlService.createElement("Relationship", n).setAttribute("Id", `rId${rowIndex + 1}`).setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image").setAttribute("Target", `../media/${filename}`);
        Relationships.addContent(Relationship);
      });
      root.addContent(Relationships);
      obj[filename] = XmlService.getRawFormat().format(root).replace(/<\?xml version="1.0" encoding="UTF-8"\?>/, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`);
    };

    setheaderFooter = function(obj_, obj) {
      var f, filename, h, headerFooter, n, oddFooter, oddHeader, root, unitX, unitY, xmlObj;
      if (obj_.hasOwnProperty("header") || obj_.hasOwnProperty("footer")) {
        unitX = "pixel";
        unitY = "pixel";
        if ((!obj_.header.hasOwnProperty("l") && !obj_.header.hasOwnProperty("r") && !obj_.header.hasOwnProperty("c")) && (!obj_.footer.hasOwnProperty("l") && !obj_.footer.hasOwnProperty("r") && !obj_.footer.hasOwnProperty("c"))) {
          putError.call(this, "Please set header and/or footer.");
        }
        filename = "xl/worksheets/sheet1.xml";
        xmlObj = XmlService.parse(obj[filename]);
        root = xmlObj.getRootElement();
        n = root.getNamespace();
        h = `&L${obj_.header.hasOwnProperty("l") ? obj_.header.l : ""}&C${obj_.header.hasOwnProperty("c") ? obj_.header.c : ""}&R${obj_.header.hasOwnProperty("r") ? obj_.header.r : ""}`;
        f = `&L${obj_.footer.hasOwnProperty("l") ? obj_.footer.l : ""}&C${obj_.footer.hasOwnProperty("c") ? obj_.footer.c : ""}&R${obj_.footer.hasOwnProperty("r") ? obj_.footer.r : ""}`;
        oddHeader = XmlService.createElement("oddHeader", n).setText(h);
        oddFooter = XmlService.createElement("oddFooter", n).setText(f);
        headerFooter = XmlService.createElement("headerFooter", n).addContent(oddHeader).addContent(oddFooter);
        root.addContent(headerFooter);
        obj[filename] = XmlService.getRawFormat().format(root);
      }
    };

    xlsxObjToBlob = function(xlsxObj) {
      var blobs;
      blobs = Object.entries(xlsxObj).reduce((ar, [k, v]) => {
        ar.push(v.toString() === "Blob" ? v : Utilities.newBlob(v, MimeType.PLAIN_TEXT, k));
        return ar;
      }, []);
      return Utilities.zip(blobs, "temp.xlsx").setContentType(MimeType.MICROSOFT_EXCEL);
    };

    putError = function(m) {
      throw new Error(`${m}`);
    };

    putInternalError = function(m) {
      throw new Error(`Internal error: ${m}`);
    };

    return SpreadsheetAppp;

  }).call(this);
  return r.SpreadsheetAppp = SpreadsheetAppp;
})(this);
