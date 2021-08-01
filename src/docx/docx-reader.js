exports.read = read;
exports._findPartPaths = findPartPaths;
exports.readStylesFromZipFile = readStylesFromZipFile;

var path = require("path");

var promises = require("../promises");
var documents = require("../documents");
var Result = require("../results").Result;
var zipfile = require("../zipfile");

var readXmlFromZipFile = require("./office-xml-reader").readXmlFromZipFile;
var createBodyReader = require("./body-reader").createBodyReader;
var DocumentXmlReader = require("./document-xml-reader").DocumentXmlReader;
var relationshipsReader = require("./relationships-reader");
var contentTypesReader = require("./content-types-reader");
var numberingXml = require("./numbering-xml");
var stylesReader = require("./styles-reader");
var notesReader = require("./notes-reader");
var commentsReader = require("./comments-reader");

var DOCUMENT_XML_RELS = "word/_rels/document.xml.rels";
var _ = require("underscore");

function read(docxFile) {
    return promises.props({
        contentTypes: readContentTypesFromZipFile(docxFile),
        partPaths: findPartPaths(docxFile),
        docxFile: docxFile,         // Содержит zip файл документа
    }).also(function(result) {
        return {
            styles: readStylesFromZipFile(docxFile, result.partPaths.styles)
        };     
    }).also(function(result) {
        return {
            numbering: readNumberingFromZipFile(docxFile, result.partPaths.numbering, result.styles)
        };
    }).also(function(result) {
        return {
            // Во премя отладки наткнулся на то, что это уже это уже где-то считывается, на самом деле, но где - не понятно
            relationships: xmlFileReader({
                filename: DOCUMENT_XML_RELS,
                readElement: function(element) {
                    var result = {type: "documentXmlRels", children: []};
                    
                    element.children.forEach(function(child) {
                        result.children.push({type: "relationship",
                                                id: child.attributes["Id"],
                                                target: child.attributes["Target"]});
                    });

                    return(result);
                },
                defaultValue: {type: "documentXmlRels", children: []}
            })(docxFile)
        }
    }).also(function(result) {
        var ftrs = [];

        for (var i = 1;;i++) {
            var name = "footer" + i + ".xml";
            if (!docxFile.exists("word/" + name)) {
                break;
            }
            
            var id = null;
            result.relationships.children.forEach(function(child) {
                if (child.target == name) {
                    id = child.id;
                }
            });

            ftrs.push({
                type: "footer",
                target: name,
                id: id
            });

            xmlFileReader({
                filename: "word/footer" + i + ".xml",
                readElement: function(element, target) {
                    var ftrsCount = ftrs.length;
                    for (var i = 0; i < ftrsCount; i++) {
                        if (ftrs[i].target == target) {
                            // bodyReader "обрезанный", но поди норм
                            var bodyReader = new createBodyReader({
                                docxFile: result.docxFile,
                                styles: result.styles,
                            });
                            try{
                            var val =  bodyReader.readXmlElements(element.children);
                            } catch (error) {
                                debugger;
                            }
                            ftrs[i]["children"] = val.value;
                            break;
                        }
                    }
                },

                defaultValue: {},
                addParam: name
            })(docxFile);       // Что странно, если поставить там then, вход будет сразу, хоть и promise не выполнено
        }

        return {
            footers: ftrs
        }
    }).also(function(result) {
        var hdrs = [];

        for (var i = 1;;i++) {
            var name = "header" + i + ".xml";
            if (!docxFile.exists("word/" + name)) {
                break;
            }
            
            var id = null;
            result.relationships.children.forEach(function(child) {
                if (child.target == name) {
                    id = child.id;
                }
            });

            hdrs.push({
                type: "header",
                target: name,
                id: id
            });

            xmlFileReader({
                filename: "word/header" + i + ".xml",
                readElement: function(element, target) {
                    var hdrsCount = hdrs.length;
                    for (var i = 0; i < hdrsCount; i++) {
                        if (hdrs[i].target == target) {
                            // bodyReader "обрезанный", но поди норм
                            var bodyReader = new createBodyReader({
                                docxFile: result.docxFile,
                                styles: result.styles,
                            });

                            var val =  bodyReader.readXmlElements(element.children);
                            hdrs[i]["children"] = val.value;
                            break;
                        }
                    }
                },

                defaultValue: {},
                addParam: name
            })(docxFile);
        }

        return {
            headers: hdrs
        }
    }).also(function(result) {
        return {
            footnotes: readXmlFileWithBody(result.partPaths.footnotes, result, function(bodyReader, xml) {
                if (xml) {
                    return notesReader.createFootnotesReader(bodyReader)(xml);
                } else {
                    return new Result([]);
                }
            }),
            endnotes: readXmlFileWithBody(result.partPaths.endnotes, result, function(bodyReader, xml) {
                if (xml) {
                    return notesReader.createEndnotesReader(bodyReader)(xml);
                } else {
                    return new Result([]);
                }
            }),
            comments: readXmlFileWithBody(result.partPaths.comments, result, function(bodyReader, xml) {
                if (xml) {
                    return commentsReader.createCommentsReader(bodyReader)(xml);
                } else {
                    return new Result([]);
                }
            })
        };
    }).also(function(result) {
        return {
            notes: result.footnotes.flatMap(function(footnotes) {
                return result.endnotes.map(function(endnotes) {
                    return new documents.Notes(footnotes.concat(endnotes));
                });
            })
        };
    }).then(function(result) {
        // function тут это типо callback
        return readXmlFileWithBody(result.partPaths.mainDocument, result, function(bodyReader, xml) {
            var reader = new DocumentXmlReader({
                bodyReader: bodyReader,
                notes: result.notes,
                comments: result.comments,
                styles:result.styles,
                file: result.docxFile
            });
            return new Result({xmlResult: reader.convertXmlToDocument(xml).value.children, footNotes: result.footNotes, comments: result.comments, headers: result.headers, footers: result.footers}, {});

        });
    });
}

function findPartPaths(docxFile) {
    return readPackageRelationships(docxFile).then(function(packageRelationships) {
        var mainDocumentPath = findPartPath({
            docxFile: docxFile,
            relationships: packageRelationships,
            relationshipType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            basePath: "",
            fallbackPath: "word/document.xml"
        });

        if (!docxFile.exists(mainDocumentPath)) {
            throw new Error("Could not find main document part. Are you sure this is a valid .docx file?");
        }

        return xmlFileReader({
            filename: relationshipsFilename(mainDocumentPath),
            readElement: relationshipsReader.readRelationships,
            defaultValue: relationshipsReader.defaultValue
        })(docxFile).then(function(documentRelationships) {
            function findPartRelatedToMainDocument(name) {
                return findPartPath({
                    docxFile: docxFile,
                    relationships: documentRelationships,
                    relationshipType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + name,
                    basePath: zipfile.splitPath(mainDocumentPath).dirname,
                    fallbackPath: "word/" + name + ".xml"
                });
            }

            return {
                mainDocument: mainDocumentPath,
                comments: findPartRelatedToMainDocument("comments"),
                endnotes: findPartRelatedToMainDocument("endnotes"),
                footnotes: findPartRelatedToMainDocument("footnotes"),
                numbering: findPartRelatedToMainDocument("numbering"),
                styles: findPartRelatedToMainDocument("styles")
            };
        });
    });
}

function findPartPath(options) {
    var docxFile = options.docxFile;
    var relationships = options.relationships;
    var relationshipType = options.relationshipType;
    var basePath = options.basePath;
    var fallbackPath = options.fallbackPath;

    var targets = relationships.findTargetsByType(relationshipType);
    var normalisedTargets = targets.map(function(target) {
        return stripPrefix(zipfile.joinPath(basePath, target), "/");
    });
    var validTargets = normalisedTargets.filter(function(target) {
        return docxFile.exists(target);
    });
    if (validTargets.length === 0) {
        return fallbackPath;
    } else {
        return validTargets[0];
    }
}

function stripPrefix(value, prefix) {
    if (value.substring(0, prefix.length) === prefix) {
        return value.substring(prefix.length);
    } else {
        return value;
    }
}

// return function мб можно убрать
// TODO: проверить это
function xmlFileReader(options) {
    return function(zipFile) {
        return readXmlFromZipFile(zipFile, options.filename)
            .then(function(element) {
                // Костыльно, но по другому в коллбэк не передать значение.
                // Вероятно, если переписать на promise - можно обойтись без этого
                // TODO: переписать на promise
                if (options.addParam) {
                    return element ? options.readElement(element, options.addParam) : options.defaultValue;
                }
                return element ? options.readElement(element) : options.defaultValue;
            });
    };
}

function readXmlFileWithBody(filename, options, func) {

    var readRelationshipsFromZipFile = xmlFileReader({
        filename: relationshipsFilename(filename),
        readElement: relationshipsReader.readRelationships,
        defaultValue: relationshipsReader.defaultValue
    });

    return readRelationshipsFromZipFile(options.docxFile).then(function(relationships) {
        var bodyReader = new createBodyReader({
            relationships: relationships,
            numbering: options.numbering,
            styles: options.styles,
            docxFile: options.docxFile
        });
        return readXmlFromZipFile(options.docxFile, filename)
            .then(function(xml) {
                return func(bodyReader, xml);
            });
    });
}

function relationshipsFilename(filename) {
    var split = zipfile.splitPath(filename);
    return zipfile.joinPath(split.dirname, "_rels", split.basename + ".rels");
}

var readContentTypesFromZipFile = xmlFileReader({
    filename: "[Content_Types].xml",
    readElement: contentTypesReader.readContentTypesFromXml,
    defaultValue: contentTypesReader.defaultContentTypes
});

function readNumberingFromZipFile(zipFile, path, styles) {
    return xmlFileReader({
        filename: path,
        readElement: function(element) {
            return numberingXml.readNumberingXml(element, {styles: styles});
        },
        defaultValue: numberingXml.defaultNumbering
    })(zipFile);
}

function readStylesFromZipFile(zipFile, path) {
    return xmlFileReader({
        filename: path,
        readElement: stylesReader.readStylesXml,
        defaultValue: stylesReader.defaultStyles
    })(zipFile);
}

var readPackageRelationships = xmlFileReader({
    filename: "_rels/.rels",
    readElement: relationshipsReader.readRelationships,
    defaultValue: relationshipsReader.defaultValue
});
