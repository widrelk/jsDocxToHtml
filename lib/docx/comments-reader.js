var documents = require("../documents");
var Result = require("../results").Result;

function createCommentsReader(bodyReader) {
    function readCommentsXml(element) {
        var elements = element.getElementsByTagName("w:comment")
        .map(readCommentElement)
        return elements
    }

    function readCommentElement(element) {
        var id = element.attributes["w:id"];

        function readOptionalAttribute(name) {
            return (element.attributes[name] || "").trim() || null;
        }

        var commentContent = bodyReader.readXmlElements(element.children)

        var dateTime = readOptionalAttribute("w:date")

        var result = {
            type:"comment",
            linkTo: id,
            authorName: readOptionalAttribute("w:author"),
            authorInitials: readOptionalAttribute("w:initials"),
            date: dateTime.split('T')[0],
            time: dateTime.split('T')[1].slice(0, -1),
            content: []
        }

        commentContent.value.forEach(function(paragraph) {
            paragraph.children.forEach(function(run) {
                run.children.forEach(function(runContent) {
                    switch (runContent.type) {
                        case "text":
                            result.content.push(runContent.value)
                            break
                    }
                })
            })
        })
        return result
    }
    
    return readCommentsXml;
}

exports.createCommentsReader = createCommentsReader;
