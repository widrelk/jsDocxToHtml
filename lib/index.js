var _ = require("underscore");

var docxReader = require("./docx/docx-reader");
var docxStyleMap = require("./docx/style-map");
var DocumentConverter = require("./document-to-html").DocumentConverter;
var readStyle = require("./style-reader").readStyle;
var readOptions = require("./options-reader").readOptions;
var unzip = require("./unzip");
var Result = require("./results").Result;

exports.convertToHtml = convertToHtml;
exports.convert = convert;

exports.images = require("./images");
exports.transforms = require("./transforms");
exports.underline = require("./underline");

/**
 * Converts given .docx file into identical html string. Styles are inline applied from the document's styles.xml
 * @param {*} input Path, ArrayBuffer or File with .docx
 * @param {*} options Deprecated
 * @returns String with HTML representation of given file
 */
function convertToHtml(input, options) {
    return convert(input, options);
}


function convert(input, options) {
    options = readOptions(options);
    
    return unzip.openZip(input)
        .then(function(docxFile) {
            docxReader.readStylesFromZipFile(docxFile, "word/styles.xml").then( function(styles) {   // Считываем стили отдельно
                // В styles функции из mammoth/lib/docx/styles-reader.js
                // docxReader как-то там считывает xml и заполняет styles сам. Функции ищут по ID структуру со стилем
                //  из styles.xml.
                return docxReader.read(docxFile, input)
                .then(function(documentResult) {
                    // TODO: Перенести это в document-to-html.js/convert-всё-на-свете, но как-то там нужно будет получить доступ к styles.xml
                    // TODO: переписать для forEach, если применимо вообще
                    documentResult.value.children = documentResult.value.children.map( function(par) {
                        var pStyle = styles.findParagraphStyleById(par.styleId);
                        if (!pStyle){
                            pStyle = styles.findParagraphStyleById('a');        // 'a' is the id of default paragraph style
                        }

                        // TODO: Я уверен, что есть способ сделать это лучше. Например, можно хранить свойства как объект, сохраняя название свойства и по нему смотреть
                        par.styleId = pStyle.styleId;
                        par.styleName = pStyle.styleName;
                        par["pFonts"] = pStyle.rFonts;
                        par["fontSize"] = pStyle.size;
                        par["color"] = pStyle.color;
                        if (pStyle.link){                   // По идее, link указывает на стиль, который идёт дефолтом для run
                            par["link"] = pStyle.link;
                        } 
                        if (pStyle.basedOn){
                            par["basedOn"] = pStyle.link;
                        }
                        // TODO: Некоторые поля могут быть заданы и стилем, и внутри document.xml в run. Модифицировать конвертер первоначальный из xml, чтобы добавить оставшиеся поля типо dStrike
                        par.children = par.children.map( function (run) {
                            var rStyle = styles.findCharacterStyleById(run.styleId);
                            if (!rStyle) {
                                rStyle = styles.findCharacterStyleById('a0');   // 'a0' is the id of default character/run style
                            }
                            // TODO: Тут просто месево вариантов, что откуда берётся. Надо бы посмотреть, можно ли переделать
                            if (!run.isBold) {
                                run.isBold = rStyle.bold;
                            } 
                            if (!run.isItalic) {
                                run.isItalic = rStyle.italic;
                            }
                            if (!run.isStrikeThrough) {
                                run.isStrikeThrough = rStyle.strike;
                            }
                            //run.isSmallCaps               TODO: посмотреть, гдe оно добавлено. Для отображения не играет роли
                            //run["isDStrikeThrough"] = rStyle.dStrike;
                            if (!run.underline) {
                                run.underline = rStyle.underline;
                            }
                            //run["vertAlign"] = rStyle.vertAlign;
                            run["pFonts"] = rStyle.rFonts;             // По умолчанию, стили character совпадают со стилем paragraph
                            if (!run.pFonts) {
                                run.pFonts = par.pFonts;
                            }
                            run["color"] = rStyle.color;
                            if (!run.color) {
                                run.color = par.color;
                            }
                            run["fontSize"] = rStyle.size;
                            if (!run.fontSize) {
                                run.fontSize = par.fontSize;
                            }
                            return(run);
                        });
                        return(par);
                    });
                    return convertDocumentToHtml(documentResult, options);
                });
            });
        });
}


function convertDocumentToHtml(documentResult, options) {
    var styleMapResult = parseStyleMap(options.readStyleMap());
    var parsedOptions = _.extend({}, options, {
        styleMap: styleMapResult.value
    });
    var documentConverter = new DocumentConverter(parsedOptions);
    
    return documentResult.flatMapThen(function(document) {
        return styleMapResult.flatMapThen(function(styleMap) {
            return documentConverter.convertToHtml(document);
        });
    });
}


function parseStyleMap(styleMap) {
    return Result.combine((styleMap || []).map(readStyle))
        .map(function(styleMap) {
            return styleMap.filter(function(styleMapping) {
                return !!styleMapping;
            });
        });
}
