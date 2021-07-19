/*
- Убраны лишние шаги в конвертере
- Написан свой преобразователь в HTML
*/
var _ = require("underscore");

var docxReader = require("./docx/docx-reader");
var readOptions = require("./options-reader").readOptions;
var unzip = require("./unzip");

exports.convertToHtml = convertToHtml;

exports.images = require("./images");
exports.transforms = require("./transforms");
exports.underline = require("./underline");

exports.convertParagraphToHtml = convertParagraphToHtml;
exports.convertRunToHtml = convertRunToHtml;

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
            return docxReader.readStylesFromZipFile(docxFile, "word/styles.xml").then( function(styles) {   // Считываем стили отдельно
                options["stylesReader"] = styles;
                return docxReader.read(docxFile, input)
                    .then(function(documentResult) {
                        var html = makeHtml(documentResult)
                        return(html);

                });
            });
        });
}


function makeHtml(documentResult) {
    var pages = splitToPages(documentResult.value.children);
    var result = convertElementsToHtml(pages);
    return(result);
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
    var runs = paragraph.children;
    if (!runs) {
        return({before: [], after: []});
    }

    runs = JSON.parse(JSON.stringify(runs));
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
    // На этом этапе поля со значениями undefintd внутри run пропадают.
    // С одной стороны, это ни на что не влияет, с другой - странно
    var paragraphCpy = JSON.parse(JSON.stringify(paragraph));

    paragraphCpy.indent.firstLine = 0;                              // Paragraph's continuation will not have any firstline or hanging indent
    paragraphCpy.indent.hanging = 0;

    paragraphCpy.children = newChildren;

    return(paragraphCpy);
}


function convertElementsToHtml(elements) {
    var result = "";
try {
    elements.forEach(function(element) {
        result = result + elementToHtmlConverters[element.type](element);
    });
} catch(error){
    debugger;
}
    return(result);
}


var elementToHtmlConverters = {
    "page": convertPageToHtml,
    "paragraph": convertParagraphToHtml,
    "run": convertRunToHtml,
    "table": convertTable,
    "hyperlink": function(hyperLink) {
        var result = "<a href=\"" + toString(hyperLink.href) + "\">" + convertElementsToHtml(hyperLink.children) +"<a>";
        return(result)
    }
}


function convertPageToHtml(page) {
    // TODO: Здесь нужно вставлять значения из sectPr
try {
    var result = "<div style=\"width:595pt; height:842pt; padding: 57pt 42pt 57pt 85pt; margin-bottom: 5pt; background-color: lightblue\">";
    
    page.children.forEach( (child) => {
        var converter = elementToHtmlConverters[child.type];
        if (converter) {
            result = result + converter(child);
        }
    });

    result = result + "</div>";
    return(result);
}catch (error) {
    debugger;
}
}


function convertParagraphToHtml(paragraph) {
    //word-wrap: break-word;
    var result = "<div style=\"";

    if (paragraph.alignment) {
        if (paragraph.alignment == "right") {
            result += "text-align: right; ";
        } else if (paragraph.alignment == "left"){
            result += "dtext-align: left; ";
        }else if (paragraph.alignment == "center") {
            result += "text-align: center; ";
        } else if (paragraph.alignment == "both") {
            // TODO: text-justify поддерживается в chrome и firefox экспериментально. Надо б найти способ сделать это так, чтоб работало везде
            result += "text-align: justify; text-justify: inter-word;";
        }
    }

    if (paragraph.indent.left != 0) {
        result += "padding-left: " + paragraph.indent.left.toString() + "pt; ";
    }
    if (paragraph.indent.right != 0) {
        result += "padding-right: " + paragraph.indent.right.toString() + "pt; ";
    }

    if (paragraph.indent.hanging != 0) {
        result += "text-indent: " + (-1 * paragraph.indent.hanging).toString() + "pt; ";
    }
    if (paragraph.indent.firstLine != 0) {                                // Done this way because if both specified firstLine is in control
        result += "text-indent: " + paragraph.indent.firstLine.toString() + "pt; ";
    }
    
    if (paragraph.spacing.before != 0) {
        result += "margin-top: " + paragraph.spacing.before.toString() + "pt; ";
    }
    if (paragraph.spacing.after != 0) {
        result += "margin-bottom: " + paragraph.spacing.after.toString() + "pt; ";
    }

    if (paragraph.spacing.line != 0) {
        result += "line-height: " + paragraph.spacing.line.toString() + "pt; ";
    }

    var rPr = paragraph.rPr;
    if (rPr) {
        var color = "";
        if (rPr.color) {
            color = 'color: #' + rPr.color + '; ';
        }

        if (rPr.fontSize) {
            result += "font-size: " + rPr.fontSize.toString() + "pt; ";
        }

        if (rPr.font) {
            result += "font-family: \'" + rPr.font + "\' ";
        } 

        if (rPr.isItalic) {
            result += "font-style: italic ;";
        }
        if (rPr.isBold) {
            result += "font-weight: bold ;";
        }

        // TODO: можно разбить название на слова и обработать отдельно double и так далее
        if (rPr.underline) {  
            var type = rPr.underline;
            if (type == "single") {
                type = "";
            }                                          // TODO: underline может не сочетаться поназванию с CSS
            result += "text-decoration: underline " + type;
            if (rPr.underlineColor) {
                result += rPr.underlineColor;
            }
            result += "; ";
        }

        if (rPr.isStriketrough) {
            result += "text-decoration-line: line-trough; ";
        }
        if (rPr.isDStriketrough) {
            result += "text-decoration-line: double line-trough; text-decoration-style: double";
        }       
    }   

    result += "\"";
    
    if (color != "") {
        result += color;
    }

    result += "> ";

    result = result + convertElementsToHtml(paragraph.children);

    result += "</div>";
    return(result);

}


function convertRunToHtml(run) {
try {
    var result = "<span ";
    result += "style=\"";
    if (run.color) {
        result += "color: #" + run.color + "; ";
    }
    var fontSize = fontSize;
    if (run.fontSize) {
        result += "font-size: " + run.fontSize.toString() + "pt; ";
    }

    if (run.font) {
        result += "font-family: \'" + run.font + "\' ";
    } 

    if (run.isItalic) {
        result += "font-style: italic ;";
    }
    if (run.isBold) {
        result += "font-weight: bold ;";
    }

    if (run.underline) {                                            // TODO: underline может не сочетаться по названию с CSS
        var type = run.underline;
        if (type == "single") {
            type = "";
        }
        result += "text-decoration: underline " + type;
        if (run.underlineColor) {
            result += run.underlineColor;
        }
        result += "; ";
    }

    if (run.isStrikethrough) {
        result += "text-decoration-line: line-through; ";
    }
    if (run.isDStrikethrough) {
        result += "text-decoration-line: double line-trough; text-decoration-style: double";
    }

    if (run.highlight) {
        result += "background-color: " + run.highlight + "; ";
    }

    result += "\">";

    var addTag = null;                                        // In HTML we need separate tage for sub/sup scripts
    if (run.verticalAlignment == "subscript") {
        addTag = "sub";
    } else if (run.verticalAlignment == "superscript") {
        addTag = "sup";
    }
    if (addTag) {
        result += "<" + addTag + ">" + run.text + "</" + addTag + ">";
    } else {
        result += run.text;
    }

    result += "</span>";
    return(result);
}catch(error){
    debugger;
}
}


function convertTable(table) {
    var result = "<table style=\"border-collapse:collapse\">"
    table.children.forEach(function(tableRow) {
        var row = "<tr>";

        tableRow.children.forEach(function(tableCell) {
            var cell = "<td style=\"border: 1px solid black;\" rowspan=\"" + tableCell.rowSpan.toString() + "\" colspan=\"" + tableCell.colSpan.toString() + "\">";
            cell += convertElementsToHtml(tableCell.children);
            cell += "</td>";
            row += cell;
        });

        row += "</tr>";
        result += row;
    });

    result += "</table>";
    return(result);
}