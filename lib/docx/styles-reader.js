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
 * Returns an object with styles that is more handy than raw XML
 * @param {*} styleElement 
 * @returns {props} of given XML structure
 */
function readStyleElement(styleElement) {
    var type = styleElement.attributes["w:type"];                   // Тип элемента
    var styleId = styleElement.attributes["w:styleId"];
    var name = styleName(styleElement);                             // Название стиля, может пригодиться при отображении, разве что
    var result = {type: type, styleId: styleId, name: name}
    result["basedOn"] = styleElement.first("w:basedOn");
    // TODO: может есть способ всё это лучше. Проблемы с тем, что всё крашится, если попробовать
    // дописать что-то после first() или []
    if (result.basedOn){
        result.basedOn = result.basedOn.attributes["w:val"];
    }
    
    if (type == "paragraph"){
        result["link"] = styleElement.first("w:link");           // id связного стиля run
        if (result.link){
            result.link = result.link.attributes["w:val"];
        }
        // TODO: у pPr могут быть ещё полезные свойства
        var pPr = styleElement.first("w:pPr");
        if (pPr) {
            pPr_spacing = pPr.first("w:spacing");        // Как же жаль, что эта штука не возвращает undefined когда spacing undefined        
            if (pPr_spacing) {
                pPr_spacing = {before: pPr_spacing.attributes["w:before"],  // промежуток сверху абзаца
                                after: pPr_spacing.attributes["w:after"],   // Промежуток снизу абзаца
                                line: pPr_spacing.attributes["w:line"]}     // Высота линии / промежуток между линиями
            }
            result["spacing"] = pPr_spacing;
        }

        var rPr = styleElement.first("w:rPr");
        if (rPr){
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
                size = "11";
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
        } else {    // TODO: Не факт, что это требуется
            result["color"] = "000000";
            result["size"] = "11";
            result["rFonts"] = {ascii: "Calibri", hAnsi: "Calibri"};
        }

    } else if (type == "character"){
        result["link"] = styleElement.first("w:link");           // id связного стиля paragraph
        if (result.link){
            result.link = result.link.attributes["w:val"];
        }

        var rPr_styles = styleElement.first("w:rPr");
        if (rPr_styles){
            var color = rPr_styles.first("w:color");
            // NOTE: у color есть и другие свойства, но вроде не особо важны.
            if (color){
                color = color.attributes["w:val"];          // цвет в hex rgb 
            } else {
                color = "000000"
            }
            result["color"] = color;
            // TODO: sz в половинном измерении. Надо делить на пополам либо дальше, либо тут
            var size = rPr_styles.first("w:sz");
            if (size) {
                size = size.attributes["w:val"];
            } else {
                size = "11";
            }
            result["size"] = size;     
            /* Существует разновидность с параметром rFonts
                <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
                 В этом случае шрифт должен браться из theme(номер).xml, например. Там большая таблица
                 соответствий шрифта по языку. Можно заморочиться и сделать, конечно, но пока что не до того
                 TODO: добавить как font-family, скорее всего, это оно
            */
            font = rPr_styles.first("w:rFonts");
            if (font){
                font = {ascii:font.attributes["w:ascii"], hAnsi:font.attributes["w:hAnsi"]};
            }
            result["rFonts"] = font;

            var italic = rPr_styles.first("w:i");
            if (italic){
                italic = true;
            }
            result["italic"] = italic;

            var bold = rPr_styles.first("w:b");
            if (bold){
                bold = true;
            }
            result["bold"] = bold;
            // TODO: реализовать разные виды underline
            var underline = rPr_styles.first("w:u");
            if (underline){
                debugger;
            }
            result["underline"] = underline;

            var strike  = rPr_styles.first("w:strike");
            if (strike){
                strike = true
            }
            result["strike"] = strike;

            var dStrike = rPr_styles.first("w:dStrike");
            if (dStrike){
                dStrike = true
            }
            result["dStrile"] = dStrike;

            // TODO: скорее всего не работает. Пофиксить
            var vertAlign = rPr_styles.first("vertAlign");
            if (vertAlign){
                debugger;
            }
            result["vertAlign"] = vertAlign;
        } else {                                // TODO: Не факт, что это требуется
            result["color"] = null;
            result["size"] = null;
            result["rFonts"] = null;
            result["italic"] = null;
            result["bold"] = null;
            result["underline"] = null;
            result["strike"] = null;
            result["dStrile"] = null;
            result["vertAlign"] = null;
        }

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
