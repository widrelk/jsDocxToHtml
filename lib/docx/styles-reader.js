exports.readStylesXml = readStylesXml;
exports.Styles = Styles;
exports.defaultStyles = new Styles({}, {});

function Styles(paragraphStyles, characterStyles, tableStyles, numberingStyles) {
    return {
        findParagraphStyleById: function(styleId) {
            return paragraphStyles[styleId];
        },
        findCharacterStyleById: function(styleId) {
            return characterStyles[styleId];
        },
        findTableStyleById: function(styleId) {
            return tableStyles[styleId];
        },
        findNumberingStyleById: function(styleId) {
            return numberingStyles[styleId];
        }
    };
}

Styles.EMPTY = new Styles({}, {}, {}, {});

// root здесь это всёсодержимое styles.xml
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
    // NOTE: тут то и преобразование идёт
    // root - styles.xml, представленный структурой
    root.getElementsByTagName("w:style").forEach(function(styleElement) {   // Для каждой записи в styles.xml
        var style = readStyleElement(styleElement);
        if (style.type === "numbering") {
            numberingStyles[style.styleId] = readNumberingStyleElement(styleElement);
        } else {
            var styleSet = styles[style.type];
            if (styleSet) {
                styleSet[style.styleId] = style;
            }
        }
    });
    return new Styles(paragraphStyles, characterStyles, tableStyles, numberingStyles);
}

/**
 * Reads rPr properties and returns them as an object.
 * When passing null or undefined as styleElement, the "default" style will be returned
 * @param {*} styleElement 
 */
function readRunProperties(element) {
    if (!element){                          
        return {
            type: "runProperties",
            styleId: "a0",                                       // "a0" is a default run style
            styleName: "Default Paragraph Font",
            isBold: false,
            isItalic: false,
            isStrikethrough: false,
            /** Specifies the underline pattern if one is used (single/thick/dash...) */
            underline: false,
            underlineColor: false, 
            /** Specifies font size in standart points */
            font: "Calibri",                   // TODO: сделать проддержку других семей шрифтов
            fontSize: 11,
            color: "000000",
            highlight: false,
            verticalAlignment: false,
            isDStrikethrough: false,
            isCaps:false,
            isSmallCaps: false
        };
    }
    sId = null;
    sName = null;
    if (style) {
        sId = style.styleId;
        sname = style.Name;
    }
    var fontSizeString = element.firstOrEmpty("w:sz").attributes["w:val"];
    // w:sz gives the font size in half points, so halve the value to get the size in points
    var fontSize = /^[0-9]+$/.test(fontSizeString) ? parseInt(fontSizeString, 10) / 2 : null;
        
    return {
        type: "runProperties",
        styleId: sId,
        styleName: sName,
        isBold: readBooleanElement(element.first("w:b")),
        isItalic: readBooleanElement(element.first("w:i")),
        isStrikethrough: readBooleanElement(element.first("w:strike")),
        /** Specifies the underline pattern if one is used (single/thick/dash...) */
        underline: element.firstOrEmpty("w:u").attributes["w:val"],
        underlineColor: element.firstOrEmpty("w:u").attributes["w:val"], 
        /** Specifies font size in standart points */
        font: element.firstOrEmpty("w:rFonts").attributes["w:ascii"],                   // TODO: сделать проддержку других семей шрифтов
        fontSize: fontSize,
        color: element.firstOrEmpty("w:color").attributes["w:val"],
        highlight: element.firstOrEmpty("w:highlight").attributes["w:val"],
        verticalAlignment: element.firstOrEmpty("w:vertAlign").attributes["w:val"],
        isDStrikethrough: readBooleanElement(element.first("w:dstrike")),
        isCaps:readBooleanElement(element.first("w:caps")),
        isSmallCaps: readBooleanElement(element.first("w:smallCaps"))
    };
}
/**
 * Returns an object with styles that is more handy than raw XML
 * @param {*} styleElement 
 * @returns {props} of given XML structure
 */
// TODO: По идее, коневертеры должны быть идентичными что для document, что для styles.
function readStyleElement(styleElement) {
    var type = styleElement.attributes["w:type"];
    var styleId = styleElement.attributes["w:styleId"];
    var name = styleName(styleElement);
    var result = {type: type, styleId: styleId, styleName: name}

    result["basedOn"] = styleElement.first("w:basedOn");
    // TODO: может есть способ всё это лучше. Проблемы с тем, что всё крашится, если попробовать
    // дописать что-то после first() или []
    if (result.basedOn){
        result.basedOn = result.basedOn.attributes["w:val"];
    }
    
    if (type == "paragraph"){
        var rLink = styleElement.first("w:link");              // id связного стиля run
        if (rLink){
            rLink = rLink.attributes["w:val"];
        }

        var rPr = styleElement.first("w:rPr");                              // Стандартные свойства run для абзаца
        if (rPr){
            result["rPr"] = readRunProperties(styleElement);
            /*pStyleId = rPr.firstOrEmpty("w:rStyle");
            var color = rPr.first("w:color");
            // NOTE: у color усть и другие свойства, но вроде не особо важны.
            if (color){
                color = color.attributes["w:val"];          // цвет в hex rgb 
            } else {
                color = "000000"
            }
            result["color"] = color;
            // TODO: sz в половинном измерении. Надо делить на пополам либо дальше, либо тут
            var size = rPr.first("w:sz");
            if (size) {
                size = size.attributes["w:val"];
            } else {
                size = "11pt";
            }
            result["size"] = size;
            font = styleElement.first("w:rFonts");
            if (font){
                font = {ascii:font.attributes["w:ascii"], hAnsi:font.attributes["w:hAnsi"]};
            } else {
                font = {ascii: "Calibri", hAnsi: "Calibri"};
            }
            result["rFonts"] = font;
            // TODO: по идее, у paragraph может быть больше свойств
            */
        } else {
            rPr = readRunProperties(null);
        }
        // TODO: у pPr могут быть ещё полезные свойства
        var pPr = styleElement.first("w:pPr");                              // Свойства paragraph
        if (pPr) {
            var spcng = pPr.first("w:spacing");                           // NOTE: По одному маленькому примеру документации, 20 sapcing = 1pt
            if (spcng) {
                spcng = { after: spcng.firstOrEmpty("w:after"),        // Distance between next p
                            before: spcng.firstOrEmpty("w:before"),     // Distance between previous p
                            line: spcng.firstOrEmpty("w:line")
                        }         // Высота линии / промежуток между линиями
            } else {
                spcng = { before: "160pt",                             // Default style values
                            after: "0pt",
                            line: "0pt"
                        }
            }

            result["pPr"] = {
                alignment: element.firstOrEmpty("w:jc").attributes["w:val"],
                numbering: readNumberingProperties(element.firstOrEmpty("w:numPr"), numbering),
                indent: readParagraphIndent(element.firstOrEmpty("w:ind")),
                spacing: spcng,
                rPr: runPr,
                link: rLink, 
                };
        }
    } else if (type == "character"){
        var pLink = styleElement.first("w:link");           // id связного стиля paragraph
        if (pLink){
            pLink = pLink.attributes["w:val"];
        }

        var rPr = styleElement.first("w:rPr");
        if (rPr){
            rPr = readRunProperties(styleElement);
            /*var color = rPr.first("w:color");
            // NOTE: у color есть и другие свойства, но вроде не особо важны.
            if (color){
                color = color.attributes["w:val"];          // цвет в hex rgb 
            } else {
                color = "000000"
            }
            result["color"] = color;
            // TODO: sz в половинном измерении. Надо делить на пополам либо дальше, либо тут
            var size = rPr.first("w:sz");
            if (size) {
                size = size.attributes["w:val"] + "pt";
            } else {
                size = "11pt";
            }
            result["size"] = size;     
            /* Существует разновидность с параметром rFonts
                <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
                 В этом случае шрифт должен браться из theme(номер).xml, например. Там большая таблица
                 соответствий шрифта по языку. Можно заморочиться и сделать, конечно, но пока что не до того
                 TODO: добавить как font-family, скорее всего, это оно
            
            font = rPr.first("w:rFonts");
            if (font){
                font = {ascii:font.attributes["w:ascii"], hAnsi:font.attributes["w:hAnsi"]};
            }
            result["rFonts"] = font;

            var italic = rPr.first("w:i");
            if (italic){
                italic = true;
            }
            result["italic"] = italic;

            var bold = rPr.first("w:b");
            if (bold){
                bold = true;
            }
            result["bold"] = bold;
            // TODO: реализовать разные виды underline
            var underline = rPr.first("w:u");
            if (underline){
                debugger;
            }
            result["underline"] = underline;

            var strike  = rPr.first("w:strike");
            if (strike){
                strike = true
            }
            result["strike"] = strike;

            var dStrike = rPr.first("w:dStrike");
            if (dStrike){
                dStrike = true
            }
            result["dStrile"] = dStrike;

            // TODO: скорее всего не работает. Пофиксить
            var vertAlign = rPr.first("vertAlign");
            if (vertAlign){
                debugger;
            }
            result["vertAlign"] = vertAlign;
            */

        } else {                                // TODO: Не факт, что это требуется
            rPr = readRunProperties(null);
        }
        result["rPr"] = rPr;
        result["link"] = link;
    } else if (type == "table"){
        // TODO: сделать стили таблицы
    } else if (type == "numbering"){
        // TODO: сделать стили списков?
    }
    return(result);
}

function styleName(styleElement) {
    var nameElement = styleElement.first("w:name");
    return nameElement ? nameElement.attributes["w:val"] : null;
}

function readNumberingStyleElement(styleElement) {
    var numId = styleElement
        .firstOrEmpty("w:pPr")
        .firstOrEmpty("w:numPr")
        .firstOrEmpty("w:numId")
        .attributes["w:val"];
    return {numId: numId};
}
