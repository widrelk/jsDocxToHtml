exports.createBodyReader = createBodyReader;
exports._readNumberingProperties = readNumberingProperties;

var _ = require("underscore");

var documents = require("../documents");
var Result = require("../results").Result;
var warning = require("../results").warning;
var uris = require("./uris");


var findTagByNameInArray = require("./styles-reader").findTagByNameInArray;
var readChildTagVal = require("./styles-reader").readChildTagVal;
var readChildTagAttr = require("./styles-reader").readChildTagAttr;
var readTableStyleProperties = require("./styles-reader").readTableStyleProperties;
var readTableRowProperties = require("./styles-reader").readTableRowProperties;
var readTableCellProperties = require("./styles-reader").readTableCellProperties;


function createBodyReader(options) {
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
    var numbering = options.numbering;
    var styles = options.styles;
    var file = options.docxFile

    function readXmlElements(elements) {
        var results = elements.map(readXmlElement);
        return combineResults(results);
        
    }


    function myReadXmlElements(elements) {
        var result = []
        elements.forEach(function(element) {
            result.push(readXmlElement(element))
        })
        return result
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
     * Contains readers for major tags
     */
    var xmlElementReaders = {
        "w:p": function(element) {
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
        "w:commentRangeStart": function(element) {
            return(elementResult({type: "commentRangeStart", commentId: element.attributes["w:id"]}))
        },
        "w:commentRangeEnd": function(element) {
            return(elementResult({type: "commentRangeEnd", commentId: element.attributes["w:id"]}))
        },
        "w:br": function(element) {
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

        "w:drawing": function(element) {
            var val = readChildElements(element).value[0]
            var picId = readPicId(element)
            var extent = findTagByNameInArray(element.children[0].children, "wp:extent")
            return elementResult({
                type: "image",
                cx: extent.attributes.cx / 12700,       // 1pt = 12700 EMUs
                cy: extent.attributes.cy / 12700,
                id: picId,
                val: val.read().then(function(res) {
                        sessionStorage.setItem(picId, URL.createObjectURL( new Blob([res.buffer], { type: 'image/png'}) ))
                    })
            })

            function readPicId(element) {
                var val = element.children[0]
                val = findTagByNameInArray(val.children, "a:graphic")
                val = val.children[0].children[0]
                val = findTagByNameInArray(val.children, "pic:blipFill")
                val = findTagByNameInArray(val.children, "a:blip")
                return val.attributes["r:embed"]
            }
        },
        "w:pict": function(element) {
            return readChildElements(element).toExtra();
        },
        "v:roundrect": readChildElements,
        "v:shape": readChildElements,
        "v:textbox": readChildElements,
        "w:txbxContent": readChildElements,
        // Эти теги отвечают за то, будет ли картинка в строке, или "плавающая". сейчас читается только в строке
        // TODO: сделать поддержку "якоря"
        "wp:inline": readDrawingElement,
        //"wp:anchor": readDrawingElement,
        "v:imagedata": readImageData,
        "v:group": readChildElements,
        "v:rect": readChildElements,
        "w:lastRenderedPageBreak" : function(element) {
            return elementResult(documents.lastRenderedPageBreak())
        },
        // TODO: вынести в отдельную функцию. Почему-то не получилось сразу
        "w:sectPr": function(sectPr) {
            var headers = _.filter(sectPr.children, function(elem) {
                return(elem.name == "w:headerReference");
            });
            for (var i = 0; i < headers.length; i++) {
                headers[i] = {
                    type: "headerReference",
                    headerType: headers[i].attributes["w:type"],    // Тут может быть even, http://officeopenxml.com/WPSectionHeaderReference.php
                    id: headers[i].attributes["r:id"]
                }
            }
            var footers = _.filter(sectPr.children, function(elem) {
                return(elem.name == "w:footerReference");
            });
            for (var i = 0; i < footers.length; i++) {
                footers[i] = {
                    type: "footerReference",
                    footerType: footers[i].attributes["w:type"],
                    id: footers[i].attributes["r:id"]
                }
            }

            var pgSz = _.find(sectPr.children, function(elem){
                return(elem.name == "w:pgSz");
            }).attributes;
            var pgMar = _.find(sectPr.children, function(elem){
                return(elem.name == "w:pgMar");
            }).attributes;
            var cols = _.find(sectPr.children, function(elem){
                return(elem.name == "w:cols");
            });
            var pgNumType = _.find(sectPr.children, function(elem){
                return(elem.name == "w:pgNumType");
            });
            var titlePg = readBooleanElement(findTagByNameInArray(sectPr.children, "w:titlePg"));

            var sectType = readChildTagVal(sectPr, "w:type");
            var vAlign = readChildTagVal(sectPr, "w:vAlign");
            /* Есть ещё свойства, но не факт, что они нужны
             pgBorders
             pgMar
             formProt
             lnNumType
             paperSrc

            */ 

            return(elementResult({
                type: "sectPr",
                headers: headers,
                footers: footers,
                orientation: pgSz["w:orient"],
                pgSz: {
                    h: pgSz["w:h"] / 20,        // В xml эти размеры в 1/20 pt
                    w: pgSz["w:w"] / 20
                },
                pgMar: {
                    left: pgMar["w:left"] / 20,
                    right: pgMar["w:right"] / 20,
                    top: pgMar["w:top"] / 20,
                    bottom: pgMar["w:bottom"] / 20,
                    // Specifies the distance from the bottom edge of the page to the bottom edge of the footer.
                    footer: pgMar["w:footer"] / 20,
                    // Specifies the distance from the top edge of the page to the top edge of the header.
                    header: pgMar["w:header"] / 20,
                    // Specifies the page gutter (the extra space added to the margin, typically to account for binding).
                    gutter: pgMar["w:gutter"] / 20
                },
                cols: cols, // TODO: реализовать
                pgNumType: pgNumType,   // Странная тема, такой тег нигде не встретился, хоть и нумерация есть
                // Specifies whether the section should have a different header and footer for its first page.
                titlePg: titlePg,
                sectType: sectType,
                vAlign: vAlign
            }));
        },
        "w:sym": (element) => {
            return elementResult({type: "symbol", font: element.attributes["w:font"], char: element.attributes["w:char"]})
        }
    };

    /**
     * Reads w:pPr element and returns it as an object.
     * Inner structure and tag names kept the same.
     * Some properties are dropped due to irrelevance for the task.
     * http://officeopenxml.com/WPparagraphProperties.php
     * @param {*} element pPr element
     */
    function readParagraphProperties(element) {
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
                // 1pt = 96/72px
                var fontSize = /^[0-9]+$/.test(fontSizeString) ? parseInt(fontSizeString, 10) * 96 / 144 : null;
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

            if (findTagByNameInArray(element.children, "w:widowControl")){
                //TODO: разобраться с тегом и сделать https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_widowControl_topic_ID0E5EEO.html
            }

            var sectPr = findTagByNameInArray(element.children, "w:sectPr");
            if (sectPr) {
                result["sectPr"] = readXmlElement(sectPr).value;
                // Если бы работало - все sectPr были бы внутри document, сейчас некоторые из них в pPr встречаются.
                //return combineResults([elementResult(result), readXmlElement(sectPr)]);
            }
            
            return elementResult(result);
    }


    function readRunProperties(element) {
            if (element){
                var result = {};

                var sId = readChildTagVal(element,"w:rStyle");
                var style = styles.findCharacterStyleById(sId);         
                if (style) {
                    result = style.rPr;                                                                   // Using the obtained style as the basis before applying
                } else {                                                                                        //  'inline' style
                    result = {                      // для удобства отладки, чтобы всегда были все поля
                        isBold: false,
                        isItalic: false,
                        isStrikethrough: false,
                        underline: false,
                        underlineColor: false, 
                        font: null,                                        // TODO: сделать проддержку других семей шрифтов
                        fontSize: null,
                        color: null,
                        highlight: null,
                        verticalAlignment: null,
                        isDStrikethrough: false,
                        isCaps:false,
                        isSmallCaps: false
                    }
                }
                // TODO: Скорее всего есть способ сделать это лучше
                result["styleId"] = sId;  
                result["type"] = "runProperties";                                                         
                if (readBooleanElement(findTagByNameInArray(element.children, "w:b"))) {                  
                    result["isBold"] = readBooleanElement(findTagByNameInArray(element.children, "w:b"));
                }
                if (readBooleanElement(findTagByNameInArray(element.children, "w:i"))) {
                    result["isItalic"] = readBooleanElement(findTagByNameInArray(element.children, "w:i"));
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
                var fontSize = /^[0-9]+$/.test(fontSizeString) ? parseInt(fontSizeString, 10) * 96 / 144 : null;
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
                        isBold: false,
                        isItalic: false,
                        isStrikethrough: false,
                        underline: false,
                        underlineColor: false, 
                        font: null,                                    // TODO: сделать проддержку других семей шрифтов
                        fontSize: null,
                        color: null,
                        highlight: false,
                        verticalAlignment: false,
                        isDStrikethrough: false,
                        isCaps:false,
                        isSmallCaps: false
                    });
            }
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
        l = /^[0-9]+$/.test(l) ? parseInt(l, 10) / 20 : 0;
        r = /^[0-9]+$/.test(r) ? parseInt(r, 10) / 20 : 0;
        f = /^[0-9]+$/.test(f) ? parseInt(f, 10) / 20 : 0;
        h = /^[0-9]+$/.test(h) ? parseInt(h, 10) / 20 : 0;
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
     * Returns true if element's value exists and not equals false or 0. Returns false otherwise
     * @param {*} element A tag with w:val
     * @returns 
     */
    // TODO: Проверить работу в других местах, может где-то важно, что 0 => false
    //      Раньше не читало теги w:i, w:b и так далее, они пустые и без w:val
    function readBooleanElement(element) {
        if (element) {
            if (element.attributes["w:val"] == '0') {
                return("parent_override_false");
            }
            return(true);
            /*var value = element.attributes["w:val"];
            return value !== "false" && value !== "0";*/
        } else {
            return false;
        }
    }

    var unknownComplexField = {type: "unknown"};


    function readFldChar(element) {
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
    }


    function currentHyperlinkHref() {
        var topHyperlink = _.last(complexFieldStack.filter(function(complexField) {
            return complexField.type === "hyperlink";
        }));
        return topHyperlink ? topHyperlink.href : null;
    }


    function parseHyperlinkFieldCode(code) {
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
    }


    function noteReferenceReader(noteType) {
        return function(element) {
            var noteId = element.attributes["w:id"];
            return elementResult(new documents.NoteReference({
                noteType: noteType,
                noteId: noteId
            }));
        };
    }


    function readChildElements(element) {
        return readXmlElements(element.children);
    }

    
    return {
        readXmlElement: readXmlElement,
        readXmlElements: readXmlElements
    };


    function readTable(table) {

        var tblProps = findTagByNameInArray(table.children, "w:tblPr")

        var rowStyle = null;
        var cellStyle = null;
        var tableStyle = readChildTagVal(tblProps, "w:tblStyle")
        if (tableStyle) {
            var style = styles.findTableStyleById(tableStyle);
            tableStyle = style.tablePr
            rowStyle = style.rowPr
            cellStyle = style.cellPr || {}

            cellStyle["parentBorders"] = tableStyle.borders

        } else {
            tableStyle = {};
        }

        var inlineProps = readTableProperties(tblProps);
        if (inlineProps.align) {
            tableStyle["align"] = inlineProps.align;
        }
        if (inlineProps.borders) {
            tableStyle["borders"] = inlineProps.borders;
        }
        if (inlineProps.caption) {
            tableStyle["caption"] = inlineProps.caption;
        }
        if (inlineProps.cellsMarg) {
            tableStyle["cellsMarg"] = inlineProps.cellsMarg;
        }
        if (inlineProps.cellsPadd) {
            tableStyle["cellsPadd"] = inlineProps.cellsPadd;
        }
        if (inlineProps.indent) {
            tableStyle["indent"] = inlineProps.indent;
        }
        if (inlineProps.width) {
            tableStyle["width"] = inlineProps.width;
        }
        if (inlineProps.stylingFlags) {
            tableStyle["stylingFlags"] = inlineProps.stylingFlags;
        }

        var tblGrid = findTagByNameInArray(table.children, "w:tblGrid")
        if (tblGrid) {
            var result = [];
            tblGrid.children.forEach(function(col) {
                result.push(col.attributes["w:w"] / 20);
            })
            tblGrid = result;
        }

        var rows = _.filter(table.children, function(child) { return child.name == "w:tr"})
        rows.map(function(row) {
            row["rowStyle"] = rowStyle
            row["cellStyle"] = cellStyle
        })

        var result = myReadXmlElements(rows)

        return elementResult({type: "table", rows: result, style: tableStyle, grid: tblGrid})

    }

    function readTableProperties(element) {
        var result = {};

        var tblStyle = readChildTagVal(element, "w:tblStyle");
        if (tblStyle) {
            tblStyle = styles.findTableStyleById(tblStyle);
            result = tblStyle;
        }
        var props = readTableStyleProperties(element);
        if (props.align) {
            result["align"] = props.align;
        }
        if (props.borders) {
            result["borders"] = props.borders;
        }
        if (props.caption) {
            result["caption"] = props.caption;
        }
        if (props.cellsMarg) {
            result["cellsMarg"] = props.cellsMarg;
        }
        if (props.cellsPadd) {
            result["cellsPadd"] = props.cellsPadd;
        }
        if (props.indent) {
            result["indent"] = props.indent;
        }
        if (props.width) {
            result["width"] = props.width;
        }
        if (props.stylingFlags) {
            result["stylingFlags"] = props.stylingFlags;
        }

        return result
    }

    function readTableRow(element) {
        // Тут есть tblPrEx ещё http://officeopenxml.com/WPtablePropertyExceptions.php
        // но оно указано как для legacy документов. В тесте там просто 2 margin с нулями
        var style = element.rowStyle        // Значения из стиля таблицы, если оно было прописано
        if (!style) {
            style = {}
        }

        var props = findTagByNameInArray(element.children, "w:trPr")
        var props = readTableRowProperties(props)
        if (props) {
            if (props.alignment) {
                style["alignment"] = props.alignment
            }
            // TODO: проверить, нет ли тут случая, когда false может означать отсутствие тега и наслежования от родителя
            style["canSplit"] = props.canSplit
            style["cellsMarg"] = props.cellsMarg
            if (props.height) {
                style["height"] = props.height
            }
            style["isHeader"] = props.isHeader
        }

        var cells = _.filter(element.children, function(child) { return child.name == "w:tc"})
        cells.map(function(cell) {
            cell["cellStyle"] = JSON.parse(JSON.stringify(element.cellStyle))
        })

        cells = readXmlElements(cells).value

        return {
            type: "tableRow",
            style: style,
            cells: cells
        }
    }

    function readTableCell(element) {
        return readXmlElements(element.children).map(function(children) {
            var cellStyle = element.cellStyle
            if (!cellStyle) {
                cellStyle = {}
            }
            if (cellStyle.parentBorders) {
                var borders = cellStyle.parentBorders
                if (cellStyle.borders) {
                    cellStyle.borders.forEach(function(border) {
                        debugger;
                    });
                }
                cellStyle["borders"] = cellStyle["borders"] || {}
                borders.forEach(function(border) {
                    cellStyle["borders"][border.name] = border
                })
            }

            cellStyle["borders"] = cellStyle["borders"] || {}       // Может получиться так, что у родителя и у стиля свойства грани не заданы

            var properties = element.firstOrEmpty("w:tcPr")
            var inlineProps = readTableCellProperties(properties)
            if (inlineProps) {
                (inlineProps.borders || []).forEach(function(border) {
                    cellStyle["borders"][border.name] = border
                })
                // TODO: разобраться с шириной ячейки, когда она каким образом используется
            }

            var gridSpan = properties.firstOrEmpty("w:gridSpan").attributes["w:val"];
            var colSpan = gridSpan ? parseInt(gridSpan, 10) : 1;

            var cell = documents.TableCell(children, {colSpan: colSpan});
            cell._vMerge = readVMerge(properties);
            cell["cellProps"] = cellStyle
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
                read: files.read.bind(files, imagePath)     // Вот туточки и косяк. Нужны файлы. Но мб и не нужны
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
            read: file.read.bind(file, path)
        };
    }

    function readImage(imageFile, altText) {
        var image = documents.Image({
            readImage: imageFile.read,
            altText: altText,
            path: imageFile.path
        });
        return elementResult(image);
    }
}


function readNumberingProperties(element, numbering) {
    if (!element) {
        return null
    }
    var level = element.firstOrEmpty("w:ilvl").attributes["w:val"];
    var numId = element.firstOrEmpty("w:numId").attributes["w:val"];
    if (level === undefined || numId === undefined) {
        return null;
    } else {
        return numbering.findLevel(numId, level);
    }
}


var ignoreElements = {
    "office-word:wrap": true,
    "v:shadow": true,
    "v:shapetype": true,
    "w:annotationRef": true,
    "w:bookmarkEnd": true,
    "w:sectPr": true,
    "w:proofErr": true,
    "w:commentReference": true,
    "w:del": true,
    "w:footnoteRef": true,
    "w:endnoteRef": true,
    "w:tblPr": true,
    "w:tblGrid": true
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
    var val = _.pluck(results, "_result")
    var result = Result.combine(val);
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


