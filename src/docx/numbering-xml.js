const readParagraphIndent = require("./styles-reader.js").readParagraphIndent

exports.readNumberingXml = readNumberingXml;
exports.Numbering = Numbering;
exports.defaultNumbering = new Numbering({});

function Numbering(nums, abstractNums, styles) {
    function findLevel(numId, level) {
        var num = nums[numId];
        if (num) {
            var abstractNum = abstractNums[num.abstractNumId];
            if (abstractNum.numStyleLink == null) {
                return abstractNums[num.abstractNumId].levels[level];
            } else {
                var style = styles.findNumberingStyleById(abstractNum.numStyleLink);
                return findLevel(style.numId, level);
            }
        } else {
            return null;
        }
    }

    return {
        findLevel: findLevel
    };
}

function readNumberingXml(root, options) {
    if (!options || !options.styles) {
        throw new Error("styles is missing");
    }

    var abstractNums = readAbstractNums(root);
    var nums = readNums(root, abstractNums);
    return new Numbering(nums, abstractNums, options.styles);
}

function readAbstractNums(root) {
    var abstractNums = {};
    root.getElementsByTagName("w:abstractNum").forEach(function(element) {
        var id = element.attributes["w:abstractNumId"];
        abstractNums[id] = readAbstractNum(element);
    });
    return abstractNums;
}

function readAbstractNum(element) {
    const levels = {};
    // Element description: http://officeopenxml.com/WPnumberingLvl.php
    element.getElementsByTagName("w:lvl").forEach(function(levelElement) {
        const numFmt = levelElement.first("w:numFmt").attributes["w:val"];
        const indent = findTagByNameInArray((levelElement.first("w:pPr") || {children: []}).children, "w:ind")
        let spcng = findTagByNameInArray((levelElement.first("w:pPr") || {children: []}).children, "w:spacing")
        if (spcng) {
            var a = spcng.attributes["w:after"];
            var b = spcng.attributes["w:before"];
            var l = spcng.attributes["w:line"];
            spcng = {
                after:  a ? parseInt(a, 10) / 20 : 0,
                before: b ? parseInt(b, 10) / 20 : 0,
                line: l ? parseInt(l, 10) / 20 : 0
            };
        } 
        levels[levelElement.attributes["w:ilvl"]] = {
            isOrdered: numFmt !== "bullet",
            numberingId: element.attributes["w:abstractNumId"],
            level: parseInt(levelElement.attributes["w:ilvl"], 10),
            start: parseInt((levelElement.first("w:start") || {attributes:{"w:val": null}}).attributes["w:val"], 10),
            numFmt: (levelElement.first("w:numFmt") || {attributes:{"w:val": null}}).attributes["w:val"],
            lvlText: (levelElement.first("w:lvlText") || {attributes:{"w:val": null}}).attributes["w:val"],           // TODO: пофиксить не работающие спецсимволы
            lvlJc: (levelElement.first("w:lvlJc") || {attributes:{"w:val": null}}).attributes["w:val"],
            indent: readParagraphIndent(indent),
            spacing: spcng, 
            //isLgl
            //lvlPicBulletId
            //lvlRestart
            //pStyle

            //pPr В принципе, идентичны styles.xml, но на практике всегда буквально несколько полей для стандартных списков
            //rPr   TODO: сделать полный функционал
            suff: (levelElement.first("w:suff") || {attributes:{"w:val": "space"}}).attributes["w:val"]       // Символ между нумерацией и текстом
                                                                                                            // Пока что не ясно, как оно определяется. В тесте было только space в одном из стилейб
                                                                                                            // но на практике были немного разные отступы. indent тоже был разный при этом
                                                                                                            // TODO: разобраться, от чего оно зависит
        };
    });

    var numStyleLink = element.firstOrEmpty("w:numStyleLink").attributes["w:val"];

    return {levels: levels, numStyleLink: numStyleLink};
}

function readNums(root) {
    var nums = {};
    root.getElementsByTagName("w:num").forEach(function(element) {
        var numId = element.attributes["w:numId"];
        var abstractNumId = element.first("w:abstractNumId").attributes["w:val"];
        nums[numId] = {abstractNumId: abstractNumId};
    });
    return nums;
}

function findTagByNameInArray(array, name){
    for (var index = 0; index < array.length; index++) {
        if (array[index].name == name){
            return(array[index]);
        }
    }
    return(false);
}