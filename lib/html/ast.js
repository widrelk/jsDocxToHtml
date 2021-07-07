var htmlPaths = require("../styles/html-paths");


function nonFreshElement(tagName, attributes, children) {
    return elementWithTag(
        htmlPaths.element(tagName, attributes, {fresh: false}),
        children);
}

function freshElement(tagName, attributes, children) {
    var tag = htmlPaths.element(tagName, attributes, {fresh: true});
    return elementWithTag(tag, children);
}

// Видимо тут и отбрасывается потом ещё дополнительный раз. Нафига оно так было надо делать - вопрос
// TODO: возможно, бесполезная ветка вызовов
function elementWithTag(tag, children) {
    return {
        type: "element",
        tag: tag,
        children: children || []
    };
}

function text(value) {
    return {
        type: "text",
        value: value
    };
}

var forceWrite = {
    type: "forceWrite"
};

exports.freshElement = freshElement;
exports.nonFreshElement = nonFreshElement;
exports.elementWithTag = elementWithTag;
exports.text = text;
exports.forceWrite = forceWrite;

var voidTagNames = {
    "br": true,
    "hr": true,
    "img": true
};

function isVoidElement(node) {
    // NOTE: Модифицировано: || voidTagNames[node.tag]
    return (node.children.length === 0) && (voidTagNames[node.tag.tagName] || voidTagNames[node.tag]);
}

exports.isVoidElement = isVoidElement;
