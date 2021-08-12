let _ = require("underscore");

let docxReader = require("./docx/docx-reader");
let unzip = require("./unzip");

/**
 * Converts given .docx file into identical html string. Styles applied inline to the html tags from the document's styles.xml
 * Additional tuning for headers is required
 * @param {*} input Path, ArrayBuffer or File with .docx
 * @returns String with HTML representation of given file
 */
const convertToHtml = (input) => {
    if (!input) {
        return new Promise((resolve, reject) => {
            resolve({html: "", comments: []})
        })
    }

    return input.arrayBuffer()
            .then(arrayBuffer => unzip.openZip(arrayBuffer))
            .then(docxFile => docxReader.read(docxFile))
            .then(documentResult => {
                let html = makeHtml(documentResult)
                htmlCopy = html
                return({html: html, comments: documentResult.value["comments"]});
            })
}


const makeHtml = (documentResult) => {
    let headers = documentResult.value.headers
    let footers = documentResult.value.footers
    let numberings = {}                                         // Содержит вложенные объекты типа списка со счётчиками для уровней
                                                                // По сути, обычный массив, но сделано так чтобы избежать ошибок с indexOutOfReach или как оно там
    let pages = splitToPages(documentResult.value.xmlResult)
    pages = addSectProps(pages)

    /**
     * Walks trough paragraphs array and searches for lastRenderedPageBreak inside of the runs.
     * Then gropus paragraphs as a pages based on that.
     * @param {*} paragraphs paragraphs array
     * @returns {[]} array of "pages" that contains given paragraphs in the correct groups
     */
    // TODO: сделать так, чтобы обрабатывало ситуации с отсутствющим lastRenderedPageBreak
    //      Для этого скорее всего необходимо использовать element.clientHeight() https://developer.mozilla.org/en-US/docs/Web/API/Element/clientHeight
    function splitToPages(paragraphs) {
        let paragraphsCpy = JSON.parse(JSON.stringify(paragraphs)); // Making a copy because sCBPB mutates data

        let result = [];
        let page = [];
        let pgIndx = 0;
        for (let i = 0; i < paragraphsCpy.length; i++) {
            let par = paragraphsCpy[i];

            if (par.type == "table") {
                let split = splitTableByPageBreak(par)
                
                while (split.after) {
                    if (split.before) {
                        page.push(split.before)
                    }
                    result.push({type: "page", pageIndex: pgIndx, children: JSON.parse(JSON.stringify(page))})
                    pgIndx++
                    page.length = 0
    
                    par = split.after
    
                    split = splitTableByPageBreak(par)                                                   
                }
            } else {
                let split = splitChildrenByPageBreak(par);

                while (split.after.length != 0) {
                    if (split.before.length != 0) {                     // We don't need to make an empty paragraph if break is in the first run
                        par.children = split.before;                    // of the paragraph
                        page.push(par);
                    }

                    result.push({type: "page", pageIndex: pgIndx, children: JSON.parse(JSON.stringify(page))});
                    pgIndx++;
                    page.length = 0;

                    par = makeAfterbreakParagraph(par, split.after);    // We use props of the originap paragraph and
                                                                        // just reassigning children

                    split = splitChildrenByPageBreak(par);              // It is possible to have one paragraph on multiple pages.                                                           
                }
            }   
            page.push(par);  
        }
        result.push({type: "page", pageIndex: pgIndx, children: JSON.parse(JSON.stringify(page))});                                              // We need to push the last page separately, because there is no page with 
        pgIndx++;                                                                //  break to trigger push inside of the cycle
        return(result);                                                 
    }
    // TODO: Не работает с bookmarkStart? Вроде и работает, надо протестить
    /**
     * Splits children of a given paragraph. Split point is a first run with true lastRenderedPageBreak
     * property. Split point is a part of afterBreak array.
     * IMPORTANT: modifies lastRenderedPageBreak property of the runs: lRPB = false.
     * TODO: lRPB is a child of the run. Additional property is not needed
     * @param {*} paragraph
     * @returns {[], []} An object of two arrays. Paragraph children before run with break in first, others in second
     */
    function splitChildrenByPageBreak(element) {
        //TODO: реализовать keepNext/keepLines
        let children = _.filter(element.children, (child) => {
            return child.type != "bookmarkStart"
        });
        if (!children) {
            return({before: [], after: []});
        }

        children = JSON.parse(JSON.stringify(children));
        let breakIndex = -1;
        
        for (let i = 0; i < children.length; i++) {
            if (children[i].lastRenderedPageBreak){
                breakIndex = i;
                break;
            }
        }

        if (breakIndex == -1) {
            return({before: children, after: []});
        }

        let afterbreak = children.splice(breakIndex);
        afterbreak[0].lastRenderedPageBreak = false;
        
        return({before: children, after: afterbreak});
    }

    /**
     * Splits given table by lastRenderedPageBreak inside of a row.
     * Creates a new row for afterbreak table with paragraphs of a row after lastRenderedPageBreak tag.
     * These paragraphs are removed from the original row
     * TODO: copy table header to the afterbreak table
     * @param {*} table 
     * @returns {table, table} An object of two tables. First contains a table before break, second contains remaining table
     */
    function splitTableByPageBreak(table) {
            let tableHeader = table.rows[0]
            let foundBreak = false
            let rowCopy = null                                                                  // Used later if break is present

            for(let row = 0; row < table.rows.length; row++) {
                if ((table.rows[row].cells[0].children[0] || {type: ""}).type == "lastRenderedPageBreak") {     // If the first child of the first cell in a row is lRPB
                    let afterbreak = table.rows.splice(row)                                     //  then this row is entirely on the new page. No other way of 
                    
                    let tableCpy = JSON.parse(JSON.stringify(table))
                    tableCpy.rows = afterbreak
                    if (table.style.stylingFlags["firstRow"]) {
                        tableCpy.rows = after.splice(0,0, tableHeader)
                    }   
                    return({before: table, after: tableCpy})        
                }

                rowCopy = JSON.parse(JSON.stringify(table.rows[row]))                                 // Making a copy of the row, but with "empty" cells.
                rowCopy.cells = rowCopy.cells.map((cell) => {                                   
                    cell.children = []
                    return cell
                })

                table.rows[row].cells = table.rows[row].cells.map((cell, cellIndex) => {
                    let breakInCell = false
                    cell.children = cell.children.map((cellChild) => {
                        if (breakInCell) {
                            rowCopy.cells[cellIndex].children.push(cellChild)
                            return null
                        }
                        let breakResult = splitChildrenByPageBreak(cellChild)
                        if (breakResult.after.length != 0) {
                            foundBreak = true
                            breakInCell = true
                            rowCopy.cells[cellIndex].children = rowCopy.cells[cellIndex].children.concat(makeAfterbreakParagraph(cellChild, breakResult.after))
                            return(makeAfterbreakParagraph(cellChild, breakResult.before))
                        }
                        return cellChild
                    })
                    cell.children = _.filter(cell.children, (child) => {return child})
                    return cell
                })

                if (foundBreak) {
                    let afterbreak = table.rows.splice(row + 1)

                    let tableCpy = JSON.parse(JSON.stringify(table))
                    tableCpy.rows = afterbreak
                    tableCpy.rows.splice(0,0, rowCopy)
                    /*if (table.style.stylingFlags["firstRow"]) {                             // Inserting table header to the next page if needed
                        tableCpy.rows.splice(0,0, tableHeader)
                    }*/
                    return({before: table, after: tableCpy}) 
                }
            }
            
            return({before: table, after: null})
    }

    function makeAfterbreakParagraph(paragraph, newChildren) {
        // На этом этапе поля со значениями undefintd внутри run пропадают.
        // С одной стороны, это ни на что не влияет, с другой - странно
        let paragraphCpy = JSON.parse(JSON.stringify(paragraph));

        paragraphCpy.indent.firstLine = 0;                              // Paragraph's continuation will not have any firstline or hanging indent
        paragraphCpy.indent.hanging = 0;

        paragraphCpy.children = newChildren;

        return(paragraphCpy);
    }

    /**
     * Looks through pages children for sectPr elements and then adds this element
     * to the pages roots according to the section's ranges
     * @param {*} pages 
     * @returns {*} modified pages with added section properties elements
     */
    function addSectProps(pages) {
        let sections = [];
        let sectFirstPageIndex = 0;

        for (let i = 0; i < pages.length; i++) {            // перебор страниц
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
            for (let i = section.start; i <= section.stop; i++) {
                pages[i]["sectPr"] = section;               // Возможно, для оптимизации по памяти стоит просто хранить секции где-то отдельно и на месте доставать.
                                                            // TODO: проверить
            }
        });
        return(pages);
    }


    function convertElementsToHtml(elements) {
        if (elements.length == 0) {
            return "&nbsp;"
        }
        let result = ""
        elements.forEach(function(element) {
            result = result + elementToHtmlConverters[element.type](element);
        })
        return(result)
    }


    function convertPageToHtml(page) {
        // TODO: в styles.xml можно откопать дефолтный стиль. Его нужно применить на страницу прям, шрифт и так далее
        // Creating page itself
        let result = `<div id="pg${page.pageIndex}" style="width: ${page.sectPr.pgSz.w / 3 * 4}px; `
                        + `height: ${page.sectPr.pgSz.h / 3 * 4}px; `
                        + `padding: 0 ${page.sectPr.pgMar.right / 3 * 4}px `
                        + `${page.sectPr.pgMar.bottom / 3 * 4}px ${page.sectPr.pgMar.left / 3 * 4}px; `
                        + "position: relative; font-family: Times New Roman; border: 1px solid; "
                        + 'box-sizing: border-box; margin-bottom: 10px; background-color:white">';
        // Adding page footer
        let footer = _.find(page.sectPr.footers, function(footer) {return footer.footerType == "default"});
        if (page.pageIndex == 0) {
            let firstFooter = _.find(page.sectPr.footers, function(footer) { return footer.footerType == "first" });
            if (firstFooter) {
                footer = firstFooter;
            }
        } else if (page.pageIndex + 1 % 2 == 0) {
            let evenFooter = _.find(page.sectPr.footers, function(footer) {return footer.footerType == "even"});
            if (evenFooter) {
                footer = evenFooter;
            }
        }
        if (footer) {
            footer = _.find(footers, function(ftr) { return ftr.id == footer.id });
            result += `<div id="footer_pg${page.pageIndex}" style="position: absolute; `
                        + `bottom: 0; margin-bottom: ${page.sectPr.pgMar.footer}pt">`
                        + convertElementsToHtml(footer.children) + "</div>";
        }
        
        // Adding page header
        let header = _.find(page.sectPr.headers, function(hdr) {return hdr.headerType == "default"});
        if (page.pageIndex == page.sectPr.start) {
            let firstHeader = _.find(page.sectPr.headers, function(header) {return header.headerType == "first"});
            if (firstHeader) {
                header = firstHeader;
            }
        } else if (page.pageIndex + 1 % 2 == 0) {
            let evenHeader = _.find(page.sectPr.headers, function(header) {return header.headerType == "even"});
            if (evenHeader) {
                header = evenHeader;
            }
        }
        if(header) {
            header = _.find(headers, function(hdr) {return hdr.id == header.id})
            result += `<div id="header_pg${page.pageIndex}" style="position: absolute; top: 0; margin-top:${page.sectPr.pgMar.header}pt; `
                    + `padding-right:${page.sectPr.pgMar.right}pt;">`  
                    + convertElementsToHtml(header.children) + "</div>";
        }
        result = result + `<div id="content_pg${page.pageIndex}">`;
        page.children.forEach( function(child) {
            let converter = elementToHtmlConverters[child.type];
            if (converter) {
                result = result + converter(child);
            }
        });
        result = result + "</div></div>";
        return(result);
    }


    const convertParagraphToHtml = (paragraph) => {
        //word-wrap: break-word;
        let result = "<div style=\"";

        if (paragraph.alignment) {
            if (paragraph.alignment == "right") {
                result += "text-align: right; ";
            } else if (paragraph.alignment == "left"){
                result += "text-align: left; ";
            }else if (paragraph.alignment == "center") {
                result += "text-align: center; ";
            } else if (paragraph.alignment == "both") {
                // TODO: text-justify поддерживается в chrome и firefox экспериментально. Надо б найти способ сделать это так, чтоб работало везде
                // text-justify: inter-word;
                result += "text-align: justify; ";
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
        if (paragraph.indent.firstLine != 0) {                                // Done in this order because if both specified firstLine is in control
            result += "text-indent: " + paragraph.indent.firstLine.toString() + "pt; ";
        }
        
        if (paragraph.spacing.before != 0) {
            result += "margin-top: " + paragraph.spacing.before.toString() + "pt; ";
        }
        if (paragraph.spacing.after != 0) {
            result += "margin-bottom: " + paragraph.spacing.after.toString() + "pt; ";
        }

        if (paragraph.spacing.line != 0) {      // TODO: сейчас работает не совсем как надо. Надо поправить http://officeopenxml.com/WPspacing.php
            result += "line-height: " + paragraph.spacing.line + "pt; ";
        }

        let rPr = paragraph.rPr;
        if (rPr) {
            if (rPr.color) {
                result += 'color: #' + rPr.color + '; ';
            }

            if (rPr.fontSize) {
                result += "font-size: " + rPr.fontSize.toString() + "px; ";
            }

            if (rPr.font) {
                result += "font-family: \'" + rPr.font + "\'; ";
            } 

            if (rPr.isItalic) {
                result += "font-style: italic ;";
            }
            if (rPr.isBold) {
                result += "font-weight: bold ;";
            }

            // TODO: можно разбить название на слова и обработать отдельно double и так далее
            if (rPr.underline) {  
                let type = rPr.underline;
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

        if (paragraph.numbering) {
            if (paragraph.numbering.indent) {
                if (paragraph.numbering.indent.left != 0) {
                    result += "padding-left: " + paragraph.numbering.indent.left.toString() + "pt; ";
                }
                if (paragraph.numbering.indent.right != 0) {
                    result += "padding-right: " + paragraph.numbering.indent.right.toString() + "pt; ";
                }
        
                if (paragraph.numbering.indent.hanging != 0) {
                    result += "text-indent: " + (-1 * paragraph.numbering.indent.hanging).toString() + "pt; ";
                }
                if (paragraph.numbering.indent.firstLine != 0) {                                // Done in this order because if both specified firstLine is in control
                    result += "text-indent: " + paragraph.numbering.indent.firstLine.toString() + "pt; ";
                }
            }
            if (paragraph.numbering.spacing) {
                if (paragraph.numbering.spacing.before != 0) {
                    result += "margin-top: " + paragraph.numbering.spacing.before.toString() + "pt; ";
                }
                if (paragraph.numbering.spacing.after != 0) {
                    result += "margin-bottom: " + paragraph.numbering.spacing.after.toString() + "pt; ";
                }
        
                if (paragraph.numbering.spacing.line != 0) {
                    result += "line-height: " + paragraph.numbering.spacing.line.toString() + "pt; ";
                }
            }
            let currentNumbering = numberings[paragraph.numbering.numberingId]
            if (!currentNumbering) {
                numberings[paragraph.numbering.numberingId] = {"0": 0}
                currentNumbering = numberings[paragraph.numbering.numberingId]
            }

            if (currentNumbering[paragraph.numbering.level] == 0) {                         // Это несколько костыльно, но пришлось. Инкремент нужно делать до
                currentNumbering[paragraph.numbering.level] = paragraph.numbering.start - 1 // начала замены, а не после, т.к. текущее значение используется в подпунктах.
            }                                                                               // Если сделать инкремент после, то будет картина "пункт 1; подпункт 2.1"
            currentNumbering[paragraph.numbering.level]++

            currentNumbering[paragraph.numbering.level + 1] = 0         // "зануляем" подуровень текущего уровня, чтобы не копилась пунумерация
            var pattern = paragraph.numbering.lvlText
            // TODO: сделать нормальную замену на основе numFmt для римских чисел и прочего
            for (let i = 0; i <= paragraph.numbering.level; i++) {
                pattern = pattern.replace("%" + (i + 1).toString(), currentNumbering[i])
            }

            switch (paragraph.numbering.suff) {             // Отступ после номера идёт как наибольшая длина номера + символ отступа. 
                case "tab":                                 //  Проблема в том, что наибольшую длину нельзя высчитать сразу, т.к. даже в формате
                    pattern += "&nbsp;&nbsp;&nbsp;&nbsp;"   //  теоретически могут быть многозначные числа, что влияет на длину.
                    break                                   // TODO: попробовать вставлять спецсимвол и для каждого списка высчитывать 
                case "space":                               //  наибольшую длину и потом заменить на нужное количество nbsp
                    pattern += "&nbsp;"                    
                    break
            }
        }

        result += "\">" + (pattern ? pattern : "") + convertElementsToHtml(paragraph.children) + "</div>"
        return(result);
    } 

    /**
     * Converts given run into html code
     * @param {*} run 
     * @returns {string} html
     */
    const convertRunToHtml = (run) => {
        let result = '<span style="';

        if (run.color) {
            result += `color: #${run.color}; `;
        }
        if (run.fontSize) {
            result += `font-size: ${run.fontSize.toString()}px; `;
        }

        if (run.font) {
            result += `font-family:'${run.font}'; `;
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
            let type = run.underline;
            if (type == "single") {
                type = "";
            }
            result += `text-decoration: underline ${type}`;
            if (run.underlineColor) {
                result += run.underlineColor;
            }
            result += "; ";
        }

        if (run.isStrikethrough) {
            result += "text-decoration-line: line-through; ";
        }
        if (run.isDStrikethrough) {
            result += "text-decoration-line: double line-trough; text-decoration-style: double; ";
        }

        if (run.highlight) {
            result += `background-color: ${run.highlight}; `;
        }

        result += '">';

        run.children.forEach(function(child) {
            switch (child.type) {
                case "text":
                    let addTag = null;                                        // In HTML we need separate tage for sub/sup scripts
                    if (run.verticalAlignment == "subscript") {
                        addTag = "sub";
                    } else if (run.verticalAlignment == "superscript") {
                        addTag = "sup";
                    }
                    
                    if (addTag) {
                        result += `<${addTag}>${child.value}</${addTag}>`;
                    } else {
                        result += child.value;
                    }
                    break
                case "image":
                    let imgSrc = window.sessionStorage.getItem(child.id)
                    result += `<img src=" ${imgSrc}" alt="image from document" width="` 
                            + `${child.cx}pt" height="${child.cy}pt">`
                    break
                case "break":
                    result += "<br>"
                    break
                case "commentReference":
                    result += `<a name="${child.commentId}"/>`
                    break
                case "symbol":
                    let code = parseInt(child.char, 16) - parseInt("f000", 16)
                    result += `<span style="font-family: ${ child.font }">&#${ code }</span>`
                    break
                }
        })
        
        result += "</span>";
        return(result);
    }

    /**
     * Converts table element into HTML equivalent
     * @param {*} table 
     * @returns {string} conversion result
     */
    const convertTable = (table) => {
        let result = '<table style="border-collapse:collapse; '
        if (table.style) {
            if (table.style.align == "center") {
                result += "margin-left: auto; margin-right: auto; "
            }
        }
        result += '">'

        table.rows.forEach((tableRow) => {
            let row = "<tr>";

            tableRow.cells.forEach((tableCell, cellIndex) => {
                let cell = "<td"
                            + ` width="${table.grid[cellIndex] / 3 * 4}px" rowspan="${tableCell.rowSpan.toString()}" `
                            + ` colspan="${tableCell.colSpan.toString()}" style="`;
                if (tableCell.cellProps.borders) {
                        if (tableCell.cellProps.borders.top) {
                            cell += "border-top: " + tableCell.cellProps.borders.top.width + "pt "
                                    + tableCell.cellProps.borders.top.style + " " + tableCell.cellProps.borders.top.color + "; "
                        }
                        if (tableCell.cellProps.borders.bottom) {
                            cell += "border-bottom: " + tableCell.cellProps.borders.bottom.width + "pt "
                                    + tableCell.cellProps.borders.bottom.style + " " + tableCell.cellProps.borders.bottom.color + "; "
                        }

                        if (tableCell.cellProps.borders.left) {
                            cell += "border-left: " + tableCell.cellProps.borders.left.width + "pt "
                                    + tableCell.cellProps.borders.left.style + " " + tableCell.cellProps.borders.left.color + "; "
                        }
                        if (tableCell.cellProps.borders.right) {
                            cell += "border-right: " + tableCell.cellProps.borders.right.width + "pt "
                                    + tableCell.cellProps.borders.right.style + " " + tableCell.cellProps.borders.right.color + "; "
                        }
                }
                    if (table.style.cellsPadd) {
                        cell += "padding: " + table.style.cellsPadd.top + "pt "
                                + table.style.cellsPadd.right + "pt " + table.style.cellsPadd.bottom + "pt "
                                + table.style.cellsPadd.left + "pt; "
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

    const elementToHtmlConverters = {
        "page": convertPageToHtml,
        "paragraph": convertParagraphToHtml,
        "run": convertRunToHtml,
        "table": convertTable,
        "hyperlink": (hyperLink) => {
            return `<a href="${toString(hyperLink.href)}">${convertElementsToHtml(hyperLink.children)}<a>`
        },
        "bookmarkStart": (bookmark) => {return ""},             // Stub for w:bookmarkStart tag
        "commentRangeStart": (element) => {
            return(`<span class="commentArea" id="comment${element.commentId}" key="${element.commentId}">`)
        },
        "commentRangeEnd": (element) => {
            return("</span>")
        },
        "sectPr": (sectPr) => {return ""}                       // Stub for w:sectPr tag
    }
    
    return convertElementsToHtml(pages)
}

exports.convertToHtml = convertToHtml;