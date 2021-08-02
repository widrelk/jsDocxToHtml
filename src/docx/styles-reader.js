var _ = require('underscore')

exports.readStylesXml = readStylesXml
exports.Styles = Styles
exports.defaultStyles = new Styles({}, {})

exports.readRunProperties = readRunProperties
exports.findTagByNameInArray = findTagByNameInArray
exports.readChildTagVal = readChildTagVal
exports.readBooleanElement = readBooleanElement
exports.readChildTagAttr = readChildTagAttr
exports.readStyleElement = readStyleElement

exports.readTableStyleProperties = readTableStyleProperties
exports.readTableRowProperties = readTableRowProperties
exports.readTableCellProperties = readTableCellProperties

exports.readParagraphIndent = readParagraphIndent

function Styles(paragraphStyles, characterStyles, tableStyles, numberingStyles) {
    return {
        findParagraphStyleById: function(styleId) {
            if (paragraphStyles[styleId]){
                return JSON.parse(JSON.stringify(paragraphStyles[styleId]));        // Styles obtained trough this can be changed later,
            } else {                                                                //  so JSON is used to make a deep copy of required style
                return(null);
            }
        },
        findCharacterStyleById: function(styleId) {
            if (characterStyles[styleId]) {
                return JSON.parse(JSON.stringify(characterStyles[styleId]));
            } else {
                return(null);
            }
        },
        findTableStyleById: function(styleId) {
            if (tableStyles[styleId]) {
                return JSON.parse(JSON.stringify(tableStyles[styleId]));
            } else {
                return(null);
            }
        },
        findNumberingStyleById: function(styleId) {
            if (numberingStyles[styleId]) {
                return JSON.parse(JSON.stringify(numberingStyles[styleId]));
            } else {
                return(null);
            }
        }
    };
}

Styles.EMPTY = new Styles({}, {}, {}, {});


function readStylesXml(root) {
    var paragraphStyles = {};
    var characterStyles = {};
    var tableStyles = {};
    var numberingStyles = {};

    var styles = {
        "paragraph": paragraphStyles,
        "character": characterStyles,
        "table": tableStyles
    };

    // root - styles.xml, представленный объектом
    root.getElementsByTagName("w:style").forEach( function(styleElement) {
        var style = readStyleElement(styleElement);
        if (style.type === "numbering") {
            numberingStyles[style.styleId] = readNumberingStyleElement(styleElement);
        } else {
            var styleType = styles[style.type];
            if (styleType) {
                styleType[style.styleId] = style;
            }
        }
    });

    return new Styles(paragraphStyles, characterStyles, tableStyles, numberingStyles);
}

/**
 * Returns an object with styles that is more handy than raw XML
 * @param {*} styleElement 
 * @returns {props} object with props of given XML structure
 */
// TODO: По идее, коневертеры должны быть идентичными что для document, что для styles.
function readStyleElement(styleElement) {
    var type = styleElement.attributes["w:type"];
    var styleId = styleElement.attributes["w:styleId"];
    var name = readChildTagVal(styleElement, "w:name");
    var result = {type: type, styleId: styleId, styleName: name}

    result["basedOn"] = readChildTagVal(styleElement, "w:basedOn");
    
    if (type == "paragraph"){
        var rLink = readChildTagVal(styleElement, "w:link");                                // id связного стиля run

        var runPr = findTagByNameInArray(styleElement.children, "w:rPr");
        if (runPr){
            runPr = readRunProperties(runPr);
        } else {
            runPr = readRunProperties(null);
        }

        var pPr = findTagByNameInArray(styleElement.children, "w:pPr");                     // Свойства paragraph
        if (pPr) {
            var spcng = findTagByNameInArray(pPr.children,"w:spacing");                     // NOTE: По одному маленькому примеру документации, 20 spacing = 1pt
            if (spcng) {
                var a = spcng.attributes["w:after"];
                var b = spcng.attributes["w:before"];
                var l = spcng.attributes["w:line"];
                spcng = {
                    after:  a ? parseInt(a, 10) / 20 : 0,                   // Distance between next p
                    before: b ? parseInt(b, 10) / 20 : 0,                   // Distance between previous p
                    line: l ? parseInt(l, 10) / 20 : 0                      // Line height / distance between lines
                };
            } else {
                spcng = { 
                    before: 0,                                                  // Default style values
                    after: 8,
                    line: 0
                }
            }
           
            if (findTagByNameInArray(pPr.children, "w:widowControl")){
                //TODO: разобраться с тегом и сделать https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_widowControl_topic_ID0E5EEO.html
            }

            result["pPr"] = {
                alignment: readChildTagVal(pPr, "w:jc") || false,
                numbering: readNumberingProperties(findTagByNameInArray(pPr.children,"w:numPr")),
                indent: readParagraphIndent(findTagByNameInArray(pPr.children,"w:ind")) || false,
                spacing: spcng,
                rPr: runPr,
                link: rLink,
                outlineLvl: readChildTagAttrVal(pPr, "w:outlineLvl") || 0,
                keepLines: readBooleanElement(findTagByNameInArray(pPr.children,"w:keeplines")) || false,
                keepNext: readBooleanElement(findTagByNameInArray(pPr.children,"w:keepnext")) || false
            };
        } else {
            result["pPr"] = {                                               // Like default paragraph properties
                alignment: "left",
                numbering: false,
                indent: false,
                spacing: {before:0, after:8, line:0},
                rPr: readRunProperties(null),
                link: "a0",
                outlineLvl: 0,
                keepLines: false,
                keepNext: false
            };
        }

    } else if (type == "character"){
        var pLink = readChildTagVal(styleElement, "w:link");                // id связного стиля paragraph

        var rPr = styleElement.first("w:rPr");
        if (rPr){
            rPr = readRunProperties(rPr);
        } else {
            rPr = readRunProperties(null);
        }
        result["rPr"] = rPr;
        result["link"] = pLink;

    } else if (type == "table"){
        result["tablePr"] = readTableStyleProperties(findTagByNameInArray(styleElement.children, "w:tblPr"));
        var regionStyles = _.filter(styleElement.children, function(child) {return child.name == "w:tblStylePr"});
        regionStyles.forEach(function(region) {
            // TODO: могут быть ещё tblPr, tcPr, trPr. Нужно придумать тест и сделать
            var par = findTagByNameInArray(region.children, "w:pPr");
            if (par) {
                par.children.push(par);
                par.attributes["w:type"] = "paragraph";
                par = readStyleElement(par).pPr;

            }
            var run = findTagByNameInArray(region.children, "w:rPr")
            if (run) {
                run.children.push(run);
                run.attributes["w:type"] = "character";
                run = readStyleElement(run).rPr;
            }
            result[region.attributes["w:type"]] = {pPr: par, rPr: run};
        });
        result["rowPr"] = readTableRowProperties(findTagByNameInArray(styleElement.children, "w:trPr"))
        result["cellPr"] = readTableCellProperties(findTagByNameInArray(styleElement.children, "w:tcPr"))
        // TODO: могут быть ещё tcPr и trPr. Нужно придумать тест и сделать
    } else if (type == "numbering"){
        // TODO: сделать стили списков?
    }
    return(result);
}

/**
 * Reads rPr properties and returns them as an object.
 * When passing null or undefined as styleElement, the "default" style will be returned
 * @param {*} styleElement 
 * @returns {props} of given XML structure 
 */
 function readRunProperties(element) {
    try {
        if (element){
            sId = element.attributes["w:styleId"];
            sName = readChildTagVal(element, "w:name");
    
            var fontSizeString = readChildTagVal(element, "w:sz");
            // w:sz gives the font size in half points, so halve the value to get the size in points
            var fontSize = /^[0-9]+$/.test(fontSizeString) ? parseInt(fontSizeString, 10) * 96 / 144 : null;
    
            // TODO: Иногда указывается только семейство шрифтов. Что в этом сучае делать - не особо понятно
            var font = findTagByNameInArray(element.children, "w:rFonts");
            if (font) {
                font = font.attributes["w:ascii"];
            } else {
                font = null;
            }
    
            return {
                type: "runProperties",
                styleId:            sId,
                styleName:          sName,
                isBold:             readBooleanElement(findTagByNameInArray(element.children, "w:b")) || false,
                isItalic:           readBooleanElement(findTagByNameInArray(element.children, "w:i")) || false,
                isStrikethrough:    readBooleanElement(findTagByNameInArray(element.children, "w:strike")) || false,
                /** Specifies the underline pattern if one is used (single/thick/dash...) */
                underline:          readChildTagVal(element, "w:u") || false,
                underlineColor:     readChildTagAttr(element, "w:u", "w:color") || false,
                font:               font,                                                   // TODO: сделать проддержку других семей шрифтов
                fontSize:           fontSize,
                color:              readChildTagVal(element, "w:color") || false,
                highlight:          readChildTagVal(element, "w:highlight") || false,
                verticalAlignment:  readChildTagVal(element, "w:vertAlign") || false,
                isDStrikethrough:   readBooleanElement(findTagByNameInArray(element.children, "w:dstrike")) || false,
                isCaps:             readBooleanElement(findTagByNameInArray(element.children, "w:caps")) || false,
                isSmallCaps:        readBooleanElement(findTagByNameInArray(element.children, "w:smallcaps")) || false
            };
        } else {
            // TODO: Хоть и дефолты почти всегда одинаковые, их можно считать из w:docxDefaults
            return {                                                        // Some "default" run style props.
                                                                            // Need this to avoid font problems in paragraph
                    type: "runProperties",
                    styleId: "a0",                                          // "a0" is a default run style ID
                    styleName: "Default Paragraph Font",
                    isBold: false,
                    isItalic: false,
                    isStrikethrough: false,
                    underline: false,
                    underlineColor: false, 
                    font: "Times New Roman",                                        // TODO: сделать проддержку других семей шрифтов
                    fontSize: 11 * 96 / 72,
                    color: null,
                    highlight: null,
                    verticalAlignment: null,
                    isDStrikethrough: false,
                    isCaps:false,
                    isSmallCaps: false
                };
        }
    } catch(error){
            debugger;
        }
    }
    
// В одном из стилей существовало ilvl = 1, хоть и в документе не было видно, что это список. Так что на всякий случай
function readNumberingProperties(numPr) {
    if (!numPr) {
        return(null);
    }
    return({ilvl: readChildTagVal(numPr, "w:ilvl"), numId: readChildTagVal(numPr, "w:numId")});
}


function readNumberingStyleElement(styleElement) {
    var numId = styleElement
        .firstOrEmpty("w:pPr")
        .firstOrEmpty("w:numPr")
        .firstOrEmpty("w:numId")
        .attributes["w:val"];
    return {numId: numId};
}


/**
 * tblStyle нужно считывать отдельно
 * @param {*} tblPr 
 */
function readTableStyleProperties(tblPr) {
    var alignment = readChildTagVal(tblPr, "w:jc");

    var tblInd = (findTagByNameInArray(tblPr.children, "w:tblInd") || {attributes: {"w:w": 0}}).attributes["w:w"] / 20;
    /* Скорее всего бесполезные
     w:shd
     w:tblOverlap
    === мб полезные
     w:tblInd
     w:tblLayout
     w:tblpPr
     w:tblStyleColBandSize http://officeopenxml.com/WPstyleTableStylesCond.php
     w:tblStyleRowBandSize http://officeopenxml.com/WPstyleTableStylesCond.php
     
    */
    var tblLook = findTagByNameInArray(tblPr.children, "w:tblLook").attributes;
    if (tblLook) {
        tblLook = {
            firstRow:   tblLook["w:firstRow"] ? true : false,
            lastRow:    tblLook["w:lastRow"] ? true : false,
            firstColumn: tblLook["w:firstColumn"] ? true : false,
            lastColumn: tblLook["w:lastColumn"] ? true : false,
            noHband:    tblLook["w:noHBand"] ? true : false,
            noVBand:    tblLook["w:noVBand"] ? true : false
       };
    };
    
    var tblBorders = readBordersTag(findTagByNameInArray(tblPr.children, "w:tblBorders"))
    
    var tblCaption = readChildTagVal(tblPr, "w:tblCaption");
    
    var tblCellMar = readMarginsTag(findTagByNameInArray(tblPr.children, "w:tblCellMar"))
    
    var tblCellSpacing = (findTagByNameInArray(tblPr.children, "w:tblCellSpacing") || {attributes: {"w:w": 0}}).attributes["w:w"] / 20;
    // TODO: тут могут быть проценты, внезапно. Надо подобрать тест и сделать   
    var tableWidth = (findTagByNameInArray(tblPr.children, "w:tblW") || {attributes: {"w:w": 0}}).attributes["w:w"] / 20;
    // TODO: тут ещё могут быть rPr и pPr, внезапно. Но не везде, и иногда оно как дефолт чисто
    return {
        type: "tblPr",
        align: alignment,
        indent: tblInd,
        stylingFlags: tblLook,
        borders: tblBorders,
        caption: tblCaption,
        cellsPadd: tblCellMar,
        cellsMarg: tblCellSpacing,
        width: tableWidth
    };
}


function readTableRowProperties(trPr) {
    // /Есть ещё hidden, но не факт, что нужно. В word мельком глянул, не нашёл его
    if (!trPr){
        return null
    }
    var rowHeight = findTagByNameInArray(trPr.children, "w:trHeight")
    if (rowHeight) {
        rowHeight = {
            rule: rowHeight.attributes["w:hRule"],
            val: rowHeight.attributes["w:val"]
        }
    }
    return {
        type: "rowPr",
        canSplit: readBooleanElement(findTagByNameInArray(trPr.children, "w:canSplit")),
        alignment: readChildTagVal(trPr, "w:jc"),
        cellsMarg: (findTagByNameInArray(trPr.children, "w:tblCellSpacing") || {attributes: {"w:w": 0}}).attributes["w:w"] / 20,
        isHeader: readBooleanElement(findTagByNameInArray(trPr.children, "w:tblHeader")),
        height: rowHeight 
    }
}


function readTableCellProperties(tcPr) {
    if (!tcPr) {
        return null
    }
    /* TODO: Сделать тест и прогнать
         hideMark не факт, что нужно, надо подобрать тест http://officeopenxml.com/WPhideMark.php
         noWrap
         shd
         vMerge - слияния ясчеек обрабатываются как-то там сами либой
    */  
    var tcBorders = readBordersTag(findTagByNameInArray(tcPr.children, "w:tcBorders"))
    var forceTextFit = readBooleanElement(findTagByNameInArray(tcPr.children, "w:tcFitText"))
    var tblCellMar = readMarginsTag(findTagByNameInArray(tcPr.children, "w:tcMar"))
    // TODO: тут могут быть проценты, внезапно. Надо подобрать тест и сделать   
    var tcW = (findTagByNameInArray(tcPr.children, "w:tcW") || {attributes: {"w:w": 0}}).attributes["w:w"] / 20

    var vAlign = readChildTagVal(tcPr, "w:vAlign")

    return({
        type: "cellPr",
        borders: tcBorders,
        forceTextToFit: forceTextFit,
        padding: tblCellMar,
        width: tcW,
        verticalAlign: vAlign
    })
}

/**
 * В word название стилей не соответствует названиям в css
 */
var wordToOOXMLStyleNameConverter = {
    single: "solid",
    dashDotStroked: "dashed",
    dashed: "dashed",
    dashSmallGap: "dashed",
    dotDash: "dotted",
    dotDotDash:"dotted",
    dotted: "dotted",
    double: "double",
    doubleWave: "double",
    inset: "inset",
    nil: "none",
    none: "none",
    outset: "outset",
    thick: "solid",
    thickThinLargeGap: "solid",
    thickThinMediumGap: "solid",
    thickThinSmallGap: "solid",
    thinThickLargeGap: "solid",
    thinThickMediumGap: "solid",
    thinThickSmallGap: "solid",
    thinThickThinLargeGap: "solid",
    thinThickThinMediumGap: "solid",
    thinThickThinSmallGap: "solid",
    threeDEmboss: "solid",
    threeDEngrave: "solid",
    triple: "double",
    wavee: "solid"
}

function readBordersTag(borders) {
    if (!borders) {
        return null
    }
    var res = [];
    borders.children.forEach(function(child){
        color = child.attributes["w:color"]
        if (color == "auto") {
            color = '000000'
        }
        res.push({           // Получаем название границы
            name: child.name.split(':')[1],
            style: wordToOOXMLStyleNameConverter[child.attributes["w:val"]],
            width: parseInt(child.attributes["w:sz"], 10) / 8,   // See sz http://officeopenxml.com/WPtableCellProperties-Borders.php
            color: "#" + color,
            space: child.attributes["w:space"],
        })
    });
    return res;
}


function readMarginsTag(margins) {
    if (!margins) {
        return null
    }
    return {
        // Если результат fTBNIA будет null, используется foo и не выбросится ошибка с undefined
        top:    (findTagByNameInArray(margins.children, "w:top") || {attributes: {"w:w": 0}}).attributes["w:w"] / 20,
        left:   (findTagByNameInArray(margins.children, "w:left") || {attributes: {"w:w": 0}}).attributes["w:w"] / 20,
        bottom: (findTagByNameInArray(margins.children, "w:bottom") || {attributes: {"w:w": 0}}).attributes["w:w"] / 20,
        right:  (findTagByNameInArray(margins.children, "w:right") || {attributes: {"w:w": 0}}).attributes["w:w"] / 20
    }

}
// TODO: Проверить работу в других местах, может где-то важно, что 0 => false
//      Раньше не читало теги w:i, w:b и так далее, они пустые и без w:val
/**
 * Returns true if element exists, false otherwise. Used as fancy boolean reader
 * @param {*} element a given element
 * @returns {boolean}
 */
function readBooleanElement(element) {
    if (element) {
        // С некоторой вероятностью, стиль тоже может содержать значения с w:val = 0
        // TODO: проверить это
        return true;
        /*var value = element.attributes["w:val"];
        return value !== "false" && value !== "0";*/
    } else {
        return false;
    }
}
// Справедливости ради, эти самописные функции могут быть не нужны, если использовать типо element.firstOrEmpty, но оно может подходить не везде, так тчо можно и так
// Суть проблемы в том, что местами fOE применяется на массив, что вызывает ошибку, можно только на объект. Либо косяк автора либы, либо мой косяк, где вносил правки
/**
 * Returns w:val value of a given tag in the element's children array
 * @param {*} element given element
 * @param {*} tag tag value to be returned
 * @returns returns tag's  w:val if it exists, or false
 */
function readChildTagVal(element, tag) {
    var val = findTagByNameInArray(element.children, tag);
    if (val) {
        return val.attributes["w:val"];                 // Here fTBNIA() is not needed, because 'attributes' is an object
    } else {
        return null;
    }
}

/**
 * Returns attribute of a given tag in the element's children array
 * @param {*} element given element
 * @param {*} tag given tag
 * @param {*} attribute tag attribute to be returned
 * @returns returns tag's attribute if it exists, or false
 */
function readChildTagAttr(element, tag, attribute) {
    val = findTagByNameInArray(element.children, tag);
    if (val) {
        val = findTagByNameInArray(val.children, attribute);
        if (val) {
            return(val);
        }
    }
    return(null);
    
}

/**
 * Returns value of an attribute of a given tag in the element's children array.
 * Needed in case when readChildTagAttr will return a complex attribute with w:val inside and you need it
 * @param {*} element 
 * @param {*} tag 
 * @param {*} attribute 
 * @returns tag's attribute val if it exists, or false
 */
function readChildTagAttrVal(element, tag, attribute) {
    val = readChildTagAttr(element, tag, attribute);
    if (val) {
        return val["w:val"];
    } else {
        return(null);
    }
}

/**
 * Returns first full tag with given w:name from array
 * @param {*} array given array
 * @param {*} name to search
 * @returns tag with corresponding w:name
 */
function findTagByNameInArray(array, name){
    for (var index = 0; index < array.length; index++) {
        if (array[index].name == name){
            return(array[index]);
        }
    }
    return(false);
}

/**
 * Returns an object based on given w:ind element
 * @param {*} element w:ind element
 * @returns object with left, right, firstline and hanging props, in pt
 */
function readParagraphIndent(element) {
    if (!element) {
        return(false);
    }
    var l = element.attributes["w:start"] || element.attributes["w:left"];
    var r = element.attributes["w:end"] || element.attributes["w:right"];
    var f = element.attributes["w:firstLine"];
    var h = element.attributes["w:hanging"];
    // in Word distance is measured in 1/20 of the point
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