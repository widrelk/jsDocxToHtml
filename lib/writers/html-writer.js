var util = require("util");
var _ = require("underscore");


exports.writer = writer;


function writer(options) {
    options = options || {};
    if (options.prettyPrint) {
        return prettyWriter();
    } else {
        return simpleWriter();
    }
}


var indentedElements = {
    div: true,
    p: true,
    ul: true,
    li: true
};


function prettyWriter() {
    var indentationLevel = 0;
    var indentation = "  ";
    var stack = [];
    var start = true;
    var inText = false;
    
    var writer = simpleWriter();
    
    //open - write - close
    function open(tagName, attributes) {
        if (indentedElements[tagName]) {
            indent();
        }
        stack.push(tagName);
        writer.open(tagName, attributes);
        if (indentedElements[tagName]) {
            indentationLevel++;
        }
        start = false;
    }
    
    function close(tagName) {
        if (indentedElements[tagName]) {
            indentationLevel--;
            indent();
        }
        stack.pop();
        writer.close(tagName);
    }
    
    function text(value) {
        startText();
        var text = isInPre() ? value : value.replace("\n", "\n" + indentation);
        writer.text(text);
    }
    
    function selfClosing(tagName, attributes) {
        indent();
        writer.selfClosing(tagName, attributes);
    }
    
    function insideIndentedElement() {
        return stack.length === 0 || indentedElements[stack[stack.length - 1]];
    }
    
    function startText() {
        if (!inText) {
            indent();
            inText = true;
        }
    }
    
    // Видимо, разбивает одну строку в нормальный "вертикальный" файл
    function indent() {
        inText = false;
        if (!start && insideIndentedElement() && !isInPre()) {
            writer._append("\n");
            for (var i = 0; i < indentationLevel; i++) {
                writer._append(indentation);
            }
        }
    }
    
    function isInPre() {
        return _.some(stack, function(tagName) {
            return tagName === "pre";
        });
    }
    
    return {
        asString: writer.asString,
        open: open,
        close: close,
        text: text,
        selfClosing: selfClosing
    };
}

// Преобразует информацию из стилей OOXML в inlint стиль css
function generateDivProps(node){
    var result = "";

    // TODO: А может лучше Switch?
    if (node.attributes.alignment){                                // Выставляем выравнивание
        if (node.attributes.alignment === "right"){
            result.append("text-align=\"right\"; ")
        } else if (node.attributes.alignment === "center"){
            result.append("text-align=\"center\"; ")
        } else if (node.attributes.alignment === "both"){
            result.append("text-align=\"justify\"; ")
        }
    }
    // TODO: поправить indent где-то в xml-reader, а потом здесь. Проверить, где идёт значение для начала параграфа
    // Когда тест сдвинут весь, есть поле left. Скорее всего, если есть ограничение справа, будет right.
    if (node.attributes.indent) {
        if (node.attributes.indent.start) {
            result += "margin-top=" + (node.attributes.indent.start / 2) + "pt; ";
        }
        if (node.attributes.indent.end) {
            result += "margin-bottom=" + (node.attributes.indent.end / 2) + "pt; ";
        }
        if (node.attributes.indent.firstLine) {
            result += "text-indent=" + (node.attributes.indent.firstLine / 2) + "pt; ";
        }
    }
    
    return(result);
}


function generateSpanProps(node){
    var result = "";
    debugger;

    return(result);
}


function simpleWriter() {
    var fragments = [];
    
    var attributesGenerator = {
        div: generateDivProps,
        span: generateSpanProps
    }

    // Открывает тег
    function open(node) {
        //var attributeString = generateAttributeString(node.attributes);
        var attributeString = attributesGenerator[node.tag](node);

        fragments.push(util.format("<%s style=\"%s\">", node.tag, attributeString));
    }
    
    function close(tagName) {
        fragments.push(util.format("</%s>", tagName));
    }
    
    function selfClosing(tagName, attributes) {
        var attributeString = generateAttributeString(attributes);
        fragments.push(util.format("<%s style=\"%s\"/>", tagName, attributeString));
    }
    
    function generateAttributeString(attributes) {
        return _.map(attributes, function(value, key) {
            if (value == null)
                return("");
            return util.format(' %s="%s"', key, escapeHtmlAttribute(value));    // Зачем тут replace - не понятно
        }).join("");
    }
    
    function text(value) {
        fragments.push(escapeHtmlText(value));
    }
    
    function append(html) {
        fragments.push(html);
    }
    
    function asString() {
        return fragments.join("");
    }
    
    return {
        asString: asString,
        open: open,
        close: close,
        text: text,
        selfClosing: selfClosing,
        _append: append
    };
}

function escapeHtmlText(value) {
    return value
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function escapeHtmlAttribute(value) {
    return value
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}
