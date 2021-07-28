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
    var levels = {};
    // Element description: http://officeopenxml.com/WPnumberingLvl.php
    element.getElementsByTagName("w:lvl").forEach(function(levelElement) {
        var numFmt = levelElement.first("w:numFmt").attributes["w:val"];
        levels[levelElement.attributes["w:ilvl"]] = {
            isOrdered: numFmt !== "bullet",
            numberingId: element.attributes["w:abstractNumId"],
            level: parseInt(levelElement.attributes["w:ilvl"], 10),
            start: (levelElement.first("w:start") || {attributes:{"w:val": null}}).attributes["w:val"],
            numFmt: (levelElement.first("w:numFmt") || {attributes:{"w:val": null}}).attributes["w:val"],
            lvlText: (levelElement.first("w:lvlText") || {attributes:{"w:val": null}}).attributes["w:val"],           // TODO: пофиксить не работающие спецсимволы
            lvlJc: (levelElement.first("w:lvlJc") || {attributes:{"w:val": null}}).attributes["w:val"],
            //isLgl
            //lvlPicBulletId
            //lvlRestart
                                // TODO: сделать считывание стилей и их применение
                                //pPr:,
                                //pStyle:,
                                //rPr,
            suff: (levelElement.first("w:suff") || {attributes:{"w:val": null}}).attributes["w:val"]
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
