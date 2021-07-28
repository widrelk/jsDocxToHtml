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
                        return({html: html, comments: documentResult.value["comments"]});

                });
            });
        });
}


function makeHtml(documentResult) {

    var elementToHtmlConverters = {
        "page": convertPageToHtml,
        "paragraph": convertParagraphToHtml,
        "run": convertRunToHtml,
        "table": convertTable,
        "hyperlink": function(hyperLink) {
            var result = "<a href=\"" + toString(hyperLink.href) + "\">" + convertElementsToHtml(hyperLink.children) +"<a>";
            return(result)
        },
        "bookmarkStart": function(bookmark) {return ""},
        "commentRangeStart": function(element) {
            return('<span class="commentArea" id="' + element.commentId + '" key="' + element.commentId + '">')
        },
        "commentRangeEnd": function(element) {
            return("</span>")
        },
        "sectPr": function(sectPr) {return ""}                  // Заглушка для тега w:sectPr
    }
    var headers = documentResult.value.headers
    var footers = documentResult.value.footers
    var numberings = {}                                         // Содержит вложенные объекты типа списка со счётчиками для уровней

    var pages = splitToPages(documentResult.value.xmlResult)
    pages = addSectProps(pages)
    var result = convertElementsToHtml(pages)
    return(result)

    /**
     * Walks trough paragraphs array and searches for lastRenderedPageBreak inside of the runs.
     * Then gropus paragraphs as a pages based on that.
     * @param {*} paragraphs paragraphs array
     * @returns {[]} array of "pages" that contains given paragraphs in the correct groups
     */
    // TODO: сделать так, чтобы обрабатывало ситуации с отсутствющим lastRenderedPageBreak
    function splitToPages(paragraphs) {
        var paragraphsCpy = JSON.parse(JSON.stringify(paragraphs)); // Making a copy because sCBPB mutates data

        var result = [];
        var page = [];
        var pgIndx = 0;
        for (var i = 0; i < paragraphsCpy.length; i++) {
            var par = paragraphsCpy[i];
            var split = splitChildrenByPageBreak(par);

            while (split.after.length != 0) {
                if (split.before.length != 0) {                     // We don't need to make an empty paragraph if break is in the first run
                    par.children = split.before;                    // of the paragraph
                    page.push(par);
                }

                result.push({type: "page", pageIndex: pgIndx, children: JSON.parse(JSON.stringify(page))});
                pgIndx++;
                page.length = 0;

                par = makeAfterbreakParagraph(par, split.after);

                split = splitChildrenByPageBreak(par);                      // It is possible to have one paragraph on multiple pages.                                                           
            }
            
            page.push(par);  
        }
        result.push({type: "page", pageIndex: pgIndx, children: JSON.parse(JSON.stringify(page))});                                              // We need to push the last page separately, because there is no page with 
        pgIndx++;                                                                //  break to trigger push inside of the cycle
        return(result);                                                 
    }
    // Не работает с bookmarkStart!!!
    // TODO: пофиксить
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

    /**
     * Looks through pages children for sectPr elements and then adds this element
     * to the pages roots according to the section's ranges
     * @param {*} pages 
     * @returns 
     */
    function addSectProps(pages) {
        var sections = [];
        var sectFirstPageIndex = 0;

        for (var i = 0; i < pages.length; i++) {            // перебор страниц
            pages[i].children.forEach(function(pageElem) {  // Перебор элементов страницы
                if (pageElem.sectPr) {
                    section = JSON.parse(JSON.stringify(pageElem.sectPr));
                    section["start"] = sectFirstPageIndex;
                    section["stop"] = i;
                    sections.push(section);
                    sectFirstPageIndex = i + 1
                }
                if (pageElem.type == "sectPr") {
                    section = JSON.parse(JSON.stringify(pageElem));
                    section["start"] = sectFirstPageIndex;
                    section["stop"] = i;
                    sections.push(section);
                    sectFirstPageIndex = i + 1
                }
            })
        };

        sections.forEach(function(section) {
            for (var i = section.start; i <= section.stop; i++) {
                pages[i]["sectPr"] = section;               // Возможно, для оптимизации по памяти стоит просто хранить секции где-то отдельно и на месте доставать.
                                                            // TODO: проверить
            }
        });
        return(pages);
    }


    function convertElementsToHtml(elements) {
        var result = ""
        elements.forEach(function(element) {
            result = result + elementToHtmlConverters[element.type](element);
        })
        return(result)
    }


    function convertPageToHtml(page) {
    try {
        // Эти были получены экспериментально, подгоном под длину строки и количество строк исходного документа при одинаковом шрифте
        //var result = "<div style=\"width:710px; height:1000px; padding: 76px 56px 76px 113px; margin-bottom: 5pt; background-color: lightblue;word-wrap: break-word\">";
    /*
        var result = "<div style=\"width:" + page.sectPr.pgSz.w + "pt; height: " + page.sectPr.pgSz.h + "pt; " + 
        "padding: " + page.sectPr.pgMar.left +"pt " + page.sectPr.pgMar.top + "pt " + page.sectPr.pgMar.right + "pt " + page.sectPr.pgMar.bottom + "pt;" +
        " margin-bottom: 5pt; background-color: lightblue\">";
    */
        // TODO: в styles.xml можно откопать дефолтный стиль. Его нужно применить на страницу прям, шрифт и так далее
        // padding с "position: relative" позволяет отобразить нормально содержимое страницы и добавить отдельно header И footer   
        var result = '<div style="width:' + page.sectPr.pgSz.w + 'pt; height: ' + page.sectPr.pgSz.h + 'pt; '
                    + 'padding: ' + page.sectPr.pgMar.left +'pt ' + page.sectPr.pgMar.top + 'pt '
                    + page.sectPr.pgMar.right + 'pt ' + page.sectPr.pgMar.bottom + 'pt; '
                    + 'position: relative; font-family: Times New Roman; border: 1px solid; margin-bottom: 10px; background-color:white">';

        // Несколько странно, что нужно указывать padding right заново
        var header = _.find(page.sectPr.headers, function(hdr) {return hdr.headerType == "default"});
        if (page.pageIndex == 0) {
            var firstHeader = _.find(page.sectPr.headers, function(header) {return header.headerType == "first"});
            if (firstHeader) {
                header = firstHeader;
            }
        }
        if(header) {
            header = _.find(headers, function(header) {return header.id == header.id})
            result += '<div style="position: absolute; top: ' + page.sectPr.pgMar.header + "pt; "
                    + "padding-right: " + page.sectPr.pgMar.right + "pt"  
                    + ' ">' + convertElementsToHtml(header.children) + "</div>";
        }

        var footer = _.find(page.sectPr.footers, function(footer) {return footer.footerType == "default"});
        if (page.pageIndex == 0) {
            var firstFooter = _.find(page.sectPr.footers, function(footer) { return footer.footerType == "first" });
            if (firstFooter) {
                footer = firstFooter;
            }
        }
        if (footer) {
            footer = _.find(footers, function(ftr) { return ftr.id == footer.id });
            result += "<div style=\"position: absolute; bottom: " + page.sectPr.pgMar.footer + "pt; " 
                    + "padding-right: " + page.sectPr.pgMar.right + "pt\">"
                    + convertElementsToHtml(footer.children) + "</div>";
        }
        

        page.children.forEach( function(child) {
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
                result += "text-align: left; ";
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
                result += "font-size: " + rPr.fontSize.toString() + "px; ";
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
                }
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

        if (color != "") {
            result += color;
        }

        result += "\">";

        if (paragraph.numbering) {

            var currentNumbering = numberings[paragraph.numbering.numberingId]
            if (!currentNumbering) {
                numberings[paragraph.numbering.numberingId] = {}
                currentNumbering = numberings[paragraph.numbering.numberingId]
            }

            if (!currentNumbering[paragraph.numbering.level]) {
                currentNumbering[paragraph.numbering.level] = 0
            }
            currentNumbering[paragraph.numbering.level]++
            var pattern = paragraph.numbering.lvlText
            // TODO: сделать нормальную замену на основе numFmt для римских чисел и прочего
            for (var i = 0; i <= paragraph.numbering.level; i++) {
                pattern = pattern.replace("%" + (i + 1).toString(), currentNumbering[i])
            }

            result += pattern + " "             // TODO: тут нужно как-то сделать отступ, но это нельзя задать через "красную строку",
                                                //  потому что она задаётся для всего абзаца
        }

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
            result += "font-size: " + run.fontSize.toString() + "px; ";
        }

        if (run.font) {
            result += "font-family: \'" + run.font + "\' ";
        } 

        if (run.isItalic) {
            if (run.isItalic == "parent_override_false") {
                result += "font-style: normal; "
            } else {
                result += "font-style: italic; ";
            }
        }
        if (run.isBold) {
            if (run.isBold == "parent_override_false") {
                result += "font-weight: normal; ";
            } else {
                result += "font-weight: bold; ";
            }

        }

        if (run.underline) {
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

        run.children.forEach(function(child) {
            switch (child.type) {
                case "text":
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
                    break
                case "image":
                    var imgSrc = window.sessionStorage.getItem(child.id)
                    result += "<img src=\"" + imgSrc + "\" alt=\"image from document\" width=\"" 
                            + child.cx + "\"pt height=\"" + child.cy + "pt\">"
                    break
                case "break":
                    result += "<br>"
                    break
                case "commentReference":
                    result += '<a name="' + child.commentId+ '"/>'
                    break
                }
        })
        
        result += "</span>";
        return(result);
    }catch(error){
        debugger;
    }
    }


    function convertTable(table) {
        var result = "<table style=\"border-collapse:collapse; "
        if (table.style) {
            if (table.style.align == "center") {
                result += "margin-left: auto; margin-right: auto; "
            }
        }
        result += "\">"
        var tableTotalWidth = table.grid.reduce(function(sum, col) { return sum + col}, 0)
        var columnsWidth = table.grid.map(function(col) { return col / tableTotalWidth * 100})
        table.rows.forEach(function(tableRow) {
            var row = "<tr>";

            tableRow.cells.forEach(function(tableCell, cellIndex) {
                var cell = "<td"
                            + " width=\"" + table.grid[cellIndex]+ "\""
                            + " rowspan=\"" + tableCell.rowSpan.toString() 
                            + "\" colspan=\"" + tableCell.colSpan.toString()
                            + "\" style=\"";
                if (tableCell.cellProps.borders) {
                    //if (!table.style.stylingFlags.noHBand) {          // Это свойство должно убирать все границы, но почему-то оно активно всегда, даже когда границы есть
                        if (tableCell.cellProps.borders.top) {
                            cell += "border-top: " + tableCell.cellProps.borders.top.width + "pt "
                                    + tableCell.cellProps.borders.top.style + " " + tableCell.cellProps.borders.top.color + "; "
                        }
                        if (tableCell.cellProps.borders.bottom) {
                            cell += "border-bottom: " + tableCell.cellProps.borders.bottom.width + "pt "
                                    + tableCell.cellProps.borders.bottom.style + " " + tableCell.cellProps.borders.bottom.color + "; "
                        }
                    //}
                    //if (!table.style.stylingFlags.noVBand) {
                        if (tableCell.cellProps.borders.left) {
                            cell += "border-left: " + tableCell.cellProps.borders.left.width + "pt "
                                    + tableCell.cellProps.borders.left.style + " " + tableCell.cellProps.borders.left.color + "; "
                        }
                        if (tableCell.cellProps.borders.right) {
                            cell += "border-right: " + tableCell.cellProps.borders.right.width + "pt "
                                    + tableCell.cellProps.borders.right.style + " " + tableCell.cellProps.borders.right.color + "; "
                        }
                    //}
                }
                    if (table.style.cellsPadd) {
                        cell += "padding: " + table.style.cellsPadd.left + "pt "
                                + table.style.cellsPadd.top + "pt " + table.style.cellsPadd.right + "pt "
                                + table.style.cellsPadd.bottom + "pt; "
                    }
                    cell += "\">";
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
}