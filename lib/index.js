var _ = require("underscore");

var docxReader = require("./docx/docx-reader");
var docxStyleMap = require("./docx/style-map");
var DocumentConverter = require("./document-to-html").DocumentConverter;
var readStyle = require("./style-reader").readStyle;
var readOptions = require("./options-reader").readOptions;
var unzip = require("./unzip");
var Result = require("./results").Result;

exports.convertToHtml = convertToHtml;

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
    options = readOptions(options);
    
    return unzip.openZip(input)
        .then(function(docxFile) {
            docxReader.readStylesFromZipFile(docxFile, "word/styles.xml").then( function(styles) {   // Считываем стили отдельно
                // В styles функции из mammoth/lib/docx/styles-reader.js
                // docxReader как-то там считывает xml и заполняет styles сам. Функции ищут по ID структуру со стилем
                //  из styles.xml.
                options["stylesReader"] = styles;
                return docxReader.read(docxFile, input)
                    .then(function(documentResult) {    
                        return(makeHtml(documentResult)); 
                        return convertDocumentToHtml(documentResult, options);
                });
            });
        });
}


function makeHtml(documentResult) {
    var pages = splitToPages(documentResult.value.children);
    debugger;
}

/**
 * Walks trough paragraphs array and searches for lastRenderedPageBreak inside of the runs.
 * Then gropus paragraphs as a pages based on that.
 * @param {*} paragraphs paragraphs array
 * @returns {[]} array of "pages" that contains given paragraphs in the correct groups
 */
function splitToPages(paragraphs) {
    var paragraphsCpy = JSON.parse(JSON.stringify(paragraphs)); // Making a copy because sCBPB mutates data
    var result = [];
    var page = [];
    for (var i = 0; i < paragraphsCpy.length; i++) {
        var par = paragraphsCpy[i];
        var split = splitChildrenByPageBreak(par);

        while (split.after.length != 0) {
            if (split.before.length != 0) {                     // We don't need to make an empty paragraph if break is in the first run
                par.children = split.before;                    // of the paragraph
                page.push(par);
            }

            result.push({type: "page", children: JSON.parse(JSON.stringify(page))});
            page.length = 0;

            par = makeAfterbreakParagraph(par, split.after);

            split = splitChildrenByPageBreak(par);                      // It is possible to have one paragraph on multiple pages.                                                           
        }

        page.push(par);  
    }
    result.push({type: "page", children: JSON.parse(JSON.stringify(page))});                                              // We need to push the last page separately, because there is no page with 
                                                                    //  break to trigger push inside of the cycle
    return(result);                                                 
}

/**
 * Split children of a given paragraph. Split point is a first run with true lastRenderedPageBreak
 * property. Split point is a part of afterBreak array.
 * IMPORTANT: modifies lastRenderedPageBreak property of the runs: lRPB = false.
 * @param {*} paragraph
 * @returns {[], []} An object of two arrays. Paragraph children before run with break in first, other ones in second
 */
function splitChildrenByPageBreak(paragraph) {
    //TODO: реализовать keepNext/keepLines
    //TODO: возможно, есть способ сделать это лучше
    var runs = JSON.parse(JSON.stringify(paragraph.children));
    if (!runs) {
        return({before: [], after: []});
    }

    var breakIndex = -1;
    
    for (var i = 0; i < runs.length; i++) {
        if (runs[i].lastRenderedPageBreak){
            breakIndex = i;
            break;
        }
    }

    if (breakIndex == -1) {
        return({before: runs, after: []});
    }

    var afterbreak = runs.splice(breakIndex);
    afterbreak[0].lastRenderedPageBreak = false;
    
    return({before: runs, after: afterbreak});
}


function makeAfterbreakParagraph(paragraph, newChildren) {
    var paragraphCpy = JSON.parse(JSON.stringify(paragraph));

    paragraphCpy.indent.firstLine = 0;                              // Paragraph's continuation will not have any firstline or hanging indent
    paragraphCpy.indent.hanging = 0;

    paragraphCpy.children = newChildren;

    return(paragraphCpy);
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
