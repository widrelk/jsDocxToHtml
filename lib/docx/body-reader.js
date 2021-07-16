/*
- Доделал readParagraphIndent: теперь есть значения по умолчанию и результат возвращается в pt
- readParagraphProperties и readRunProperties теперь считывают "inline" стили и смешивают его
со стилем из styles.xml. Изменена структура результата и есть значение по умолчанию
- Теперь считывается lastRenderedPageBreak
- readBooleanElement работает корректно

остальное не трогал.
*/

exports.createBodyReader = createBodyReader;
exports._readNumberingProperties = readNumberingProperties;

var _ = require("underscore");

var documents = require("../documents");
var Result = require("../results").Result;
var warning = require("../results").warning;
var uris = require("./uris");
var styles;


var findTagByNameInArray = require("./styles-reader").findTagByNameInArray;
var readChildTagVal = require("./styles-reader").readChildTagVal;
var readChildTagAttr = require("./styles-reader").readChildTagAttr;
var readBooleanElement = require("./styles-reader").readBooleanElement;


function createBodyReader(options) {
    styles = options.styles;
    return {
        readXmlElement: function(element) {
            return new BodyReader(options).readXmlElement(element);
        },
        readXmlElements: function(elements) {
            return new BodyReader(options).readXmlElements(elements);
        }
    };
}


function BodyReader(options) {
    var complexFieldStack = [];
    var currentInstrText = [];
    var relationships = options.relationships;
    var contentTypes = options.contentTypes;
    var docxFile = options.docxFile;
    var files = options.files;
    var numbering = options.numbering;
    var styles = options.styles;

    // По какой-то причине, из run у которых стиль был добавлен внутри обработчика отдельно, все поля кроме styleId равны undefined
    // TODO: Как-то это отладить
    function readXmlElements(elements) {
        var results = elements.map(readXmlElement);
        return combineResults(results);                 // Там очень стрёмный js
    }

    function readXmlElement(element) {
        if (element.type === "element") {
            var handler = xmlElementReaders[element.name];
            if (handler) {
                return handler(element);
            } else if (!Object.prototype.hasOwnProperty.call(ignoreElements, element.name)) {
                var message = warning("An unrecognised element was ignored: " + element.name);
                return emptyResultWithMessages([message]);
            }
        }
        return emptyResult();
    }

    /**
     * Reads w:ind tag and returns it as an object
     */
    function readParagraphIndent(element) {
        if (!element) {
            return({
                left: 0,
                right: 0,
                firstLine: 0,
                hanging: 0
            });
        }
        var l = element.attributes["w:start"] || element.attributes["w:left"];
        var r = element.attributes["w:end"] || element.attributes["w:right"];
        var f = element.attributes["w:firstLine"];
        var h = element.attributes["w:hanging"];
        // in Word distance is measured 1/20 of a pt
        l = /^[0-9]+$/.test(l) ? parseInt(l, 10) / 2 : null;
        r = /^[0-9]+$/.test(r) ? parseInt(r, 10) / 2 : null;
        f = /^[0-9]+$/.test(f) ? parseInt(f, 10) / 2 : null;
        h = /^[0-9]+$/.test(h) ? parseInt(h, 10) / 2 : null;
        return {
            /** Specifies the indentation to be placed at the left.*/
            left: l,
            /** Specifies the indentation to be placed at the right.*/
            right: r,
            /** Specifies indentation to be removed from the first line.
             * This attribute and firstLine are mutually exclusive.
             * This attribute controls when both are specified. */
            firstLine: f,
            /** Specifies additional indentation to be applied to the first line.*/
            hanging: h
        };
    }

    /**
     * Reads w:pPr element and returns it as an object.
     * Inner structure and tag names kept the same.
     * Some properties are dropped due to irrelevance for the task.
     * http://officeopenxml.com/WPparagraphProperties.php
     * @param {*} element pPr element
     */
    function readParagraphProperties(element) {
        try {
        //return readParagraphStyle(element).map(function(style) {
            
            var sId = readChildTagVal(element,"w:pStyle");
            var style = styles.findParagraphStyleById(sId);
            if (!style) {
                style = styles.findParagraphStyleById("a");
            }
            result = style.pPr;
            result["styleId"] = style.styleId;
            result["styleName"] = style["styleName"];
            result["type"] = "paragraphProperties";

            var rPr = findTagByNameInArray(element.children, "w:rPr");              // Looking for rPr in pPr, that can rewrite original r style from pPr.styleId
            if (rPr) {
                var runStyle = readChildTagVal(rPr, "w:rStyle");
                if (runStyle) {
                    result.rPr = styles.findCharacterStyleById(runStyle).rPr;
                } 

                var fontSizeString = readChildTagVal(rPr, "w:sz");
                var fontSize = /^[0-9]+$/.test(fontSizeString) ? parseInt(fontSizeString, 10) / 2 : null;
                if (fontSize) {
                    result.rPr.fontSize = fontSize;
                }

                var font = findTagByNameInArray(rPr.children, "w:rFonts");
                if (font) {
                    font = font.attributes["w:ascii"];
                    if (font) {
                        result.rPr.font = font;
                    }
                }
            }
            // Тут надо бы вмешать pPr.rPr в стиль, полученный по rStyle.
            // Действительно, иногда свойства хранятся тут, а не в run.
            // Что зависит на то, где хранится стиль - не понятно

            var spcng = findTagByNameInArray(element.children,"w:spacing");
            if (spcng) {
                var a = spcng.attributes["w:after"];
                var b = spcng.attributes["w:before"];
                var l = spcng.attributes["w:line"];
                result.spacing = {
                    after:  a ? parseInt(a, 10) / 20 : 0,
                    before: b ? parseInt(b, 10) / 20 : 0,
                    line: l ? parseInt(l, 10) / 20 : 0
                };
            } 

            if (readChildTagVal(element, "w:jc")) {
                result.alignment = readChildTagVal(element, "w:jc");
            }

            if (findTagByNameInArray(element.children,"w:numPr")) {
                result.numbering = readNumberingProperties(element.firstOrEmpty("w:numPr"), numbering)
            }

            if (findTagByNameInArray(element.children,"w:ind")) {
                result.indent = readParagraphIndent(findTagByNameInArray(element.children,"w:ind"));
            }

            if (findTagByNameInArray(element.children, "w:outlineLvl")) {
                /** Specifies the outline level associated with the paragraph.
                 * It is used to build the table of contents and does not affect the appearance of the text. */
                result.outlineLvl = readChildTagVal(element, "w:outlineLvl");
            }

            if (findTagByNameInArray(element.children, "w:keeplines")) {
                /** Specifies that all lines of the paragraph are to be kept on a single page when possible.*/
                result.keepLines = readBooleanElement(findTagByNameInArray(element.children, "w:keepLines"));
            }

            if (findTagByNameInArray(element.children, "w:keepNext")) {
                /** Specifies that the paragraph (or at least part of it) should be rendered on the same page as the next paragraph when possible.*/
                result.keepNext = readBooleanElement(findTagByNameInArray(element.children, "w:keepNext"));
            }

            return elementResult(result);
        //});
        } catch(error) {
            debugger;
        }
    }


    function readRunProperties(element) {
        try {
            if (element){
                var sId = readChildTagVal(element,"w:rStyle");
                var style = styles.findCharacterStyleById(sId);
                if (!style) {
                    style = styles.findCharacterStyleById("a0");
                }
                
    
                // TODO: Скорее всего есть способ сделать это лучше
                result = style.rPr;  
                result["styleId"] = sId;  
                result["type"] = "runProperties";                                                         // Using the obtained style as the basis before applying
                if (readBooleanElement(findTagByNameInArray(element.children, "w:b"))) {                              // 'inline' style
                    result["isBold"] = true;
                }
                if (readBooleanElement(findTagByNameInArray(element.children, "w:i"))) {
                    result["isItalic"] = true;
                }
                if (readBooleanElement(findTagByNameInArray(element.children, "w:strike"))) {
                    result["isStrikethrough"] = true;
                }
                
                if (readChildTagVal(element, "w:u")) {
                    result["underline"] = readChildTagVal(element, "w:u");
                }
                if (readChildTagAttr(element, "w:u", "w:color")) {
                    result["underlineColor"] = readChildTagAttr(element, "w:u", "w:color");
                }
                if (readChildTagAttr(element, "w:rFonts", "w:ascii")) {
                    result["font"] = readChildTagAttr(element, "w:rFonts", "w:ascii");
                }

                var fontSizeString = readChildTagVal(element, "w:sz");
                // w:sz gives the font size in half points, so halve the value to get the size in points
                var fontSize = /^[0-9]+$/.test(fontSizeString) ? parseInt(fontSizeString, 10) / 2 : "11";
                result["fontSize"] = fontSize;

                var font = findTagByNameInArray(element.children, "w:rFonts");
                if (font) {
                    font = font.attributes["w:ascii"];
                }
                result["font"] = font;
                
                if (readChildTagVal(element, "w:color")) {
                    result["color"] = readChildTagVal(element, "w:color");
                }
                if (readChildTagVal(element, "w:highlight")) {
                    result["highlight"] = readChildTagVal(element, "w:highlight");
                }
                if (readChildTagVal(element, "w:vertAlign")) {
                    result["verticalAlignment"] = readChildTagVal(element, "w:vertAlign");
                }
                if (readBooleanElement(findTagByNameInArray(element.children, "w:dstrike"))) {
                    result["isDStrikethrough"] = true;
                }
                if (readBooleanElement(findTagByNameInArray(element.children, "w:caps"))) {
                    result["isCaps"] = true;
                }
                if (readBooleanElement(findTagByNameInArray(element.children, "w:smallcaps"))) {
                    result["isSmallCaps"] = true;
                }
    
                
                return elementResult(result);
            } else {
                return elementResult({
                        type: "runProperties",
                        styleId: "a0",                                       // "a0" is a default run style
                        styleName: "Default Paragraph Font",
                        isBold: false,
                        isItalic: false,
                        isStrikethrough: false,
                        underline: false,
                        underlineColor: false, 
                        font: "Calibri",                                    // TODO: сделать проддержку других семей шрифтов
                        fontSize: 11,
                        color: "000000",
                        highlight: false,
                        verticalAlignment: false,
                        isDStrikethrough: false,
                        isCaps:false,
                        isSmallCaps: false
                    });
            }
        } catch(error){
                debugger;
            }
        }

    /**
     * Returns true if element's value exists and not equals false or 0. Returns false otherwise
     * @param {*} element A tag with w:val
     * @returns 
     */
    // TODO: Проверить работу в других местах, может где-то важно, что 0 => false
    //      Раньше не читало теги w:i, w:b и так далее, они пустые и без w:val
    function readBooleanElement(element) {
        if (element) {
            return(true);
            /*var value = element.attributes["w:val"];
            return value !== "false" && value !== "0";*/
        } else {
            return false;
        }
    }


    function readTableStyle(element) {
        return readStyle(element, "w:tblStyle", "Table", styles.findTableStyleById);
    }


    function readStyle(element, styleTagName, styleType, findStyleById) {
        var messages = [];
        var styleElement = element.first(styleTagName);
        var styleId = null;
        var name = null;
        if (styleElement) {
            styleId = styleElement.attributes["w:val"];
            if (styleId) {
                var style = findStyleById(styleId);
                if (style) {
                    name = style.name;
                } else {
                    messages.push(undefinedStyleWarning(styleType, styleId));
                }
            }
        }
        return elementResultWithMessages({styleId: styleId, name: name}, messages);
    }


    var unknownComplexField = {type: "unknown"};


    /*function readFldChar(element) {
        var type = element.attributes["w:fldCharType"];
        if (type === "begin") {
            complexFieldStack.push(unknownComplexField);
            currentInstrText = [];
        } else if (type === "end") {
            complexFieldStack.pop();
        } else if (type === "separate") {
            var href = parseHyperlinkFieldCode(currentInstrText.join(''));
            var complexField = href === null ? unknownComplexField : {type: "hyperlink", href: href};
            complexFieldStack.pop();
            complexFieldStack.push(complexField);
        }
        return emptyResult();
    }*/


    function currentHyperlinkHref() {
        var topHyperlink = _.last(complexFieldStack.filter(function(complexField) {
            return complexField.type === "hyperlink";
        }));
        return topHyperlink ? topHyperlink.href : null;
    }


    /*function parseHyperlinkFieldCode(code) {
        var result = /\s*HYPERLINK "(.*)"/.exec(code);
        if (result) {
            return result[1];
        } else {
            return null;
        }
    }


    function readInstrText(element) {
        currentInstrText.push(element.text());
        return emptyResult();
    }*/


    function noteReferenceReader(noteType) {
        return function(element) {
            var noteId = element.attributes["w:id"];
            return elementResult(new documents.NoteReference({
                noteType: noteType,
                noteId: noteId
            }));
        };
    }


    function readCommentReference(element) {
        return elementResult(documents.commentReference({
            commentId: element.attributes["w:id"]
        }));
    }


    function readChildElements(element) {
        return readXmlElements(element.children);
    }

    /**
     * Contains readers for major tags
     */
    var xmlElementReaders = {
        "w:p": function(element) {
            try {
            return readXmlElements(element.children)
                .map(function(children) {
                    var properties = _.find(children, isParagraphProperties);   // Отдельно выделяем props

                    if (!properties) {
                        properties = styles.findParagraphStyleById("a");
                    }
                    
                    return new documents.Paragraph(                             // функция из documents.js
                        children.filter(negate(isParagraphProperties)),         // Удаляем props из детей
                        properties                                              // и передаём его отдельно
                    );           
                })
                .insertExtra();
            } catch (error) {
                debugger;
            }
                
        },
        "w:pPr": readParagraphProperties,
        "w:r": function(element) {
            return readXmlElements(element.children)
                .map(function(child) {

                    var properties = _.find(child, isRunProperties);
                    child = child.filter(negate(isRunProperties));

                    if (!properties) {
                        properties = styles.findCharacterStyleById("a0");
                    }

                    var hyperlinkHref = currentHyperlinkHref(); // Для чего оно тут - не понятно. Мб гиперссылка хранится отдельно и вставляется по мере надобности в каждый run
                    if (hyperlinkHref !== null) {
                        child = [new documents.Hyperlink(child, {href: hyperlinkHref})];
                    }
                    return new documents.Run(child, properties);  
                });
        },
        "w:rPr": readRunProperties,
        "w:fldChar": readFldChar,
        "w:instrText": readInstrText,
        "w:t": function(element) {
            return elementResult(new documents.Text(element.text()));
        },
        "w:tab": function(element) {
            return elementResult(new documents.Tab());
        },
        "w:noBreakHyphen": function() {
            return elementResult(new documents.Text("\u2011"));
        },
        "w:hyperlink": function(element) {
            var relationshipId = element.attributes["r:id"];
            var anchor = element.attributes["w:anchor"];
            return readXmlElements(element.children).map(function(children) {
                function create(options) {
                    var targetFrame = element.attributes["w:tgtFrame"] || null;

                    return new documents.Hyperlink(
                        children,
                        _.extend({targetFrame: targetFrame}, options)
                    );
                }

                if (relationshipId) {
                    var href = relationships.findTargetByRelationshipId(relationshipId);
                    if (anchor) {
                        href = uris.replaceFragment(href, anchor);
                    }
                    return create({href: href});
                } else if (anchor) {
                    return create({anchor: anchor});
                } else {
                    return children;
                }
            });
        },
        "w:tbl": readTable,
        "w:tr": readTableRow,
        "w:tc": readTableCell,
        "w:footnoteReference": noteReferenceReader("footnote"),
        "w:endnoteReference": noteReferenceReader("endnote"),
        "w:commentReference": readCommentReference,
        "w:br": function(element) {                                       // Page breaks are irrelevant, cause only lRPB matter
            var breakType = element.attributes["w:type"];
            if (breakType == null || breakType === "textWrapping") {
                return elementResult(documents.lineBreak);
            } else if (breakType === "page") {
                return elementResult(documents.pageBreak);
            } else if (breakType === "column") {
                return elementResult(documents.columnBreak);
            } else {
                return emptyResultWithMessages([warning("Unsupported break type: " + breakType)]);
            }
        },
        "w:bookmarkStart": function(element){                   //TODO: разобраться с закладками, надо оно вообще или нет и что это такое
            var name = element.attributes["w:name"];
            if (name === "_GoBack") {
                return emptyResult();
            } else {
                return elementResult(new documents.BookmarkStart({name: name}));
            }
        },

        "mc:AlternateContent": function(element) {
            return readChildElements(element.first("mc:Fallback"));
        },

        "w:sdt": function(element) {
            return readXmlElements(element.firstOrEmpty("w:sdtContent").children);
        },

        "w:ins": readChildElements,
        "w:object": readChildElements,
        "w:smartTag": readChildElements,
        "w:drawing": readChildElements,
        "w:pict": function(element) {
            return readChildElements(element).toExtra();
        },
        "v:roundrect": readChildElements,
        "v:shape": readChildElements,
        "v:textbox": readChildElements,
        "w:txbxContent": readChildElements,
        "wp:inline": readDrawingElement,
        "wp:anchor": readDrawingElement,
        "v:imagedata": readImageData,
        "v:group": readChildElements,
        "v:rect": readChildElements,
        "w:lastRenderedPageBreak" : function(element) {
            return elementResult(documents.lastRenderedPageBreak())
        }
        // TODO: сделать чтение тега w:sectPr
        //"w:sectPr":
    };
    return {
        readXmlElement: readXmlElement,
        readXmlElements: readXmlElements
    };


    function readTable(element) {
        var propertiesResult = readTableProperties(element.firstOrEmpty("w:tblPr"));
        return readXmlElements(element.children)
            .flatMap(calculateRowSpans)
            .flatMap(function(children) {
                return propertiesResult.map(function(properties) {
                    return documents.Table(children, properties);
                });
            });
    }

    function readTableProperties(element) {
        return readTableStyle(element).map(function(style) {
            return {
                styleId: style.styleId,
                styleName: style.name
            };
        });
    }

    function readTableRow(element) {
        var properties = element.firstOrEmpty("w:trPr");
        var isHeader = !!properties.first("w:tblHeader");
        return readXmlElements(element.children).map(function(children) {
            return documents.TableRow(children, {isHeader: isHeader});
        });
    }

    function readTableCell(element) {
        return readXmlElements(element.children).map(function(children) {
            var properties = element.firstOrEmpty("w:tcPr");

            var gridSpan = properties.firstOrEmpty("w:gridSpan").attributes["w:val"];
            var colSpan = gridSpan ? parseInt(gridSpan, 10) : 1;

            var cell = documents.TableCell(children, {colSpan: colSpan});
            cell._vMerge = readVMerge(properties);
            return cell;
        });
    }

    function readVMerge(properties) {
        var element = properties.first("w:vMerge");
        if (element) {
            var val = element.attributes["w:val"];
            return val === "continue" || !val;
        } else {
            return null;
        }
    }

    function calculateRowSpans(rows) {
        var unexpectedNonRows = _.any(rows, function(row) {
            return row.type !== documents.types.tableRow;
        });
        if (unexpectedNonRows) {
            return elementResultWithMessages(rows, [warning(
                "unexpected non-row element in table, cell merging may be incorrect"
            )]);
        }
        var unexpectedNonCells = _.any(rows, function(row) {
            return _.any(row.children, function(cell) {
                return cell.type !== documents.types.tableCell;
            });
        });
        if (unexpectedNonCells) {
            return elementResultWithMessages(rows, [warning(
                "unexpected non-cell element in table row, cell merging may be incorrect"
            )]);
        }

        var columns = {};

        rows.forEach(function(row) {
            var cellIndex = 0;
            row.children.forEach(function(cell) {
                if (cell._vMerge && columns[cellIndex]) {
                    columns[cellIndex].rowSpan++;
                } else {
                    columns[cellIndex] = cell;
                    cell._vMerge = false;
                }
                cellIndex += cell.colSpan;
            });
        });

        rows.forEach(function(row) {
            row.children = row.children.filter(function(cell) {
                return !cell._vMerge;
            });
            row.children.forEach(function(cell) {
                delete cell._vMerge;
            });
        });

        return elementResult(rows);
    }

    function readDrawingElement(element) {
        var blips = element                                 // По сути, получает rId картинки
            .getElementsByTagName("a:graphic")
            .getElementsByTagName("a:graphicData")
            .getElementsByTagName("pic:pic")
            .getElementsByTagName("pic:blipFill")
            .getElementsByTagName("a:blip");

        return combineResults(blips.map(readBlip.bind(null, element)));
    }

    function readBlip(element, blip) {
        var properties = element.first("wp:docPr").attributes;
        var altText = isBlank(properties.descr) ? properties.title : properties.descr;
        return readImage(findBlipImageFile(blip), altText);
    }

    function isBlank(value) {
        return value == null || /^\s*$/.test(value);
    }

    function findBlipImageFile(blip) {
        var embedRelationshipId = blip.attributes["r:embed"];
        var linkRelationshipId = blip.attributes["r:link"];
        if (embedRelationshipId) {
            return findEmbeddedImageFile(embedRelationshipId);
        } else {
            var imagePath = relationships.findTargetByRelationshipId(linkRelationshipId);
            return {
                path: imagePath,
                read: files.read.bind(files, imagePath)
            };
        }
    }

    function readImageData(element) {
        var relationshipId = element.attributes['r:id'];

        if (relationshipId) {
            return readImage(
                findEmbeddedImageFile(relationshipId),
                element.attributes["o:title"]);
        } else {
            return emptyResultWithMessages([warning("A v:imagedata element without a relationship ID was ignored")]);
        }
    }

    function findEmbeddedImageFile(relationshipId) {
        var path = uris.uriToZipEntryName("word", relationships.findTargetByRelationshipId(relationshipId));
        return {
            path: path,
            read: docxFile.read.bind(docxFile, path)
        };
    }

    function readImage(imageFile, altText) {
        var contentType = contentTypes.findContentType(imageFile.path);

        var image = documents.Image({
            readImage: imageFile.read,
            altText: altText,
            contentType: contentType
        });
        var warnings = supportedImageTypes[contentType] ?
            [] : warning("Image of type " + contentType + " is unlikely to display in web browsers");
        return elementResultWithMessages(image, warnings);
    }

    function undefinedStyleWarning(type, styleId) {
        return warning(
            type + " style with ID " + styleId + " was referenced but not defined in the document");
    }
}


function readNumberingProperties(element, numbering) {
    var level = element.firstOrEmpty("w:ilvl").attributes["w:val"];
    var numId = element.firstOrEmpty("w:numId").attributes["w:val"];
    if (level === undefined || numId === undefined) {
        return null;
    } else {
        return numbering.findLevel(numId, level);
    }
}

var supportedImageTypes = {
    "image/png": true,
    "image/gif": true,
    "image/jpeg": true,
    "image/svg+xml": true,
    "image/tiff": true
};

var ignoreElements = {
    "office-word:wrap": true,
    "v:shadow": true,
    "v:shapetype": true,
    "w:annotationRef": true,
    "w:bookmarkEnd": true,
    "w:sectPr": true,
    "w:proofErr": true,
    "w:commentRangeStart": true,
    "w:commentRangeEnd": true,
    "w:del": true,
    "w:footnoteRef": true,
    "w:endnoteRef": true,
    "w:tblPr": true,
    "w:tblGrid": true,
    "w:trPr": true,
    "w:tcPr": true
};

function isParagraphProperties(element) {
    return element.type === "paragraphProperties";
}

function isRunProperties(element) {
    return element.type === "runProperties";
}

function negate(predicate) {
    return function(value) {
        return !predicate(value);
    };
}


function emptyResultWithMessages(messages) {
    return new ReadResult(null, null, messages);
}

function emptyResult() {
    return new ReadResult(null);
}

function elementResult(element) {
    return new ReadResult(element);
}

function elementResultWithMessages(element, messages) {
    return new ReadResult(element, null, messages);
}

function ReadResult(element, extra, messages) {
    this.value = element || [];
    this.extra = extra;
    this._result = new Result({
        element: this.value,
        extra: extra
    }, messages);
    this.messages = this._result.messages;
}

ReadResult.prototype.toExtra = function() {
    return new ReadResult(null, joinElements(this.extra, this.value), this.messages);
};

ReadResult.prototype.insertExtra = function() {
    var extra = this.extra;
    if (extra && extra.length) {
        return new ReadResult(joinElements(this.value, extra), null, this.messages);
    } else {
        return this;
    }
};

ReadResult.prototype.map = function(func) {
    var result = this._result.map(function(value) {
        return func(value.element);
    });
    return new ReadResult(result.value, this.extra, result.messages);
};

ReadResult.prototype.flatMap = function(func) {
    var result = this._result.flatMap(function(value) {
        return func(value.element)._result;
    });
    return new ReadResult(result.value.element, joinElements(this.extra, result.value.extra), result.messages);
};

function combineResults(results) {
    var result = Result.combine(_.pluck(results, "_result"));
    return new ReadResult(
        _.flatten(_.pluck(result.value, "element")),
        _.filter(_.flatten(_.pluck(result.value, "extra")), identity),
        result.messages
    );
}

function joinElements(first, second) {
    return _.flatten([first, second]);
}

function identity(value) {
    return value;
}


