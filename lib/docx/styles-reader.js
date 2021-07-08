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
    debugger;
    var type = styleElement.attributes["w:type"];                   // Тип элемента
    var styleId = styleElement.attributes["w:styleId"];
    var name = styleName(styleElement);                             // Название стиля, может пригодиться при отображении, разве что
    var result = {type: type, styleId: styleId, name: name}

    if (type == "paragraph"){
        result["basedOn"] = styleElement.first("w:basedOn");
        // TODO: может есть способ это сделать лучше. После first() если дописать продолжение, оно крашится
        if (result.basedOn){
            result.basedOn = result.basedOn.attributes["w:val"];
        }
        result["link"] = styleElement.first("w:link");           // id связного стиля run
        if (result.link){
            result.link = result.link.attributes["w:val"];
        }
        // TODO: у pPr могут быть ещё полезные свойства
        var pPr_spacing = styleElement.first("w:pPr");
        if (pPr_spacing) {
            pPr_spacing = pPr_spacing.children["w:spacing"];        // Как же жаль, что эта штука не возвращает undefined когда spacing undefined        
            pPr_spacing = {before: pPr_spacing.first("w:before"),  // промежуток сверху абзаца
                            after: pPr_spacing.first("w:after"),   // Промежуток снизу абзаца
                            line: pPr_spacing.first("w:line")}     // Высота линии / промежуток между линиями
        }
        result["pPr"] = {spacing: pPr_spacing};

        result["rPr"] = {sz: styleElement.first("w:sz"),
                        color: styleElement.first("w:color")};     // цвет в hex rgb
    } else if (type == "run"){

    }
    return(result);
    //return {type: type, styleId: styleId, name: name};
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
