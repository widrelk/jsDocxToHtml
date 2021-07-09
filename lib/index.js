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
                                                                    // Так как в стиле поля указываются не все, нужно всё заполнить "по умолчанию"
                        par.indent.start = "8pt";                   // Соответствие размеров в word и css - однозначное
                        par["font"] = {
                            fontSize: "11pt",       
                            fontFamily: "Calibri",                  // Надо бы тут добавить в font family sans/serif/monospace ещё
                        };


                        if (pStyle) {
                            debugger;

                        }
                        par.children = par.children.map( function (run) {
                            var rStyle = styles.findCharacterStyleById(run.styleId);
                            run["font"] = par.font;             // По умолчанию, стили character совпадают со стилем paragraph
                            if (rStyle){
                                debugger;
                                
                            }
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
