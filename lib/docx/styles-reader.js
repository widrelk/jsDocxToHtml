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

// Возвращает состав xml стиля в виде более удобного объекта
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
        var pPr_spacing = styleElement.first("w:pPr");
        if (pPr_spacing) {
            pPr_spacing = pPr_spacing.first("w:spacing");        // Как же жаль, что эта штука не возвращает undefined когда spacing undefined        
            if (pPr_spacing) {
                pPr_spacing = {before: pPr_spacing.attributes["w:before"],  // промежуток сверху абзаца
                                after: pPr_spacing.attributes["w:after"],   // Промежуток снизу абзаца
                                line: pPr_spacing.attributes["w:line"]}     // Высота линии / промежуток между линиями
            }
        }
        result["pPr"] = {spacing: pPr_spacing};

        var rPr_styles = styleElement.first("w:rPr");
        if (rPr_styles){
            var color = rPr_styles.first("w:color");
            // NOTE: у color усть и другие свойства, но вроде не особо важны.
            if (color){
                color = color.attributes["w:val"];          // цвет в hex rgb 
            }      
            var size = rPr_styles.first("w:sz");
            if (size) {
                size = size.attributes["w:val"];
                rPr_styles = {sz: size,
                            color: color};
            }
            result["rPr"] = rPr_styles;
            font = styleElement.first("w:rFonts");
            if (font){
                font = {ascii:font.attributes["w:ascii"], hAnsi:font.attributes["w:hAnsi"]};
            }
            result["rFonts"] = font;
        }
        result["rPr"] = rPr_styles; 

    } else if (type == "character"){
        result["link"] = styleElement.first("w:link");           // id связного стиля paragraph
        if (result.link){
            result.link = result.link.attributes["w:val"];
        }

        var rPr_styles = styleElement.first("w:rPr");

        if (rPr_styles){
            var color = rPr_styles.first("w:color");
            // NOTE: у color усть и другие свойства, но вроде не особо важны.
            if (color){
                color = color.attributes["w:val"];          // цвет в hex rgb 
            }      
            var size = rPr_styles.first("w:sz");
            if (size) {
                size = size.attributes["w:val"];
                rPr_styles = {sz: size,
                            color: color};
            }
            result["rPr"] = rPr_styles;     
            /* Существует разновидность с параметром rFonts
                <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
                 В этом случае шрифт должен браться из theme(номер).xml, например. Там большая таблица
                 соответствий шрифта по языку. Можно заморочиться и сделать, конечно, но пока что не до того
                 TODO: доделать эти стили
            */
            font = styleElement.first("w:rFonts");
            if (font){
                font = {ascii:font.attributes["w:ascii"], hAnsi:font.attributes["w:hAnsi"]};
            }
            result["rFonts"] = font;
            
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
