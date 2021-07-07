var ast = require("./ast");

exports.freshElement = ast.freshElement;
exports.nonFreshElement = ast.nonFreshElement;
exports.elementWithTag = ast.elementWithTag;
exports.text = ast.text;
exports.forceWrite = ast.forceWrite;

exports.simplify = require("./simplify");


function write(writer, nodes) {
    nodes.forEach(function(node) {
        writeNode(writer, node);
    });
}

function writeNode(writer, node) {
    toStrings[node.type](writer, node);
}

var toStrings = {
    element: generateElementString,
    text: generateTextString,
    forceWrite: function() { }
};


// NOTE: Изменены параметры: note.tag.tagName -> note.tag
function generateElementString(writer, node) {
    if (ast.isVoidElement(node)) {
        writer.selfClosing(node.tag, node.tag.attributes);
    } else {
        // Вот тут и пишутся типо строчки
        writer.open(node.tag, node.tag.attributes);
        write(writer, node.children);
        writer.close(node.tag);
    }
}

function generateTextString(writer, node) {
    writer.text(node.value);
}

exports.write = write;
