/*
- Добавлены новые поля в Paragraph и Run
*/

var _ = require("underscore");

var types = exports.types = {
    document: "document",
    paragraph: "paragraph",
    run: "run",
    text: "text",
    tab: "tab",
    hyperlink: "hyperlink",
    noteReference: "noteReference",
    image: "image",
    note: "note",
    commentReference: "commentReference",
    comment: "comment",
    tableCell: "tableCell",
    break: "break",
    bookmarkStart: "bookmarkStart"
};


function Document(children, options) {
    options = options || {};
    return {
        type: types.document,
        children: children,
        notes: options.notes || new Notes({}),
        comments: options.comments || []
    };
}

/**
 * Returns an object from paragraph children and properties.
 * Makes it a bit more handy than it was in body-reader
 * @param {*} children Paragraph children (runs, etc)
 * @param {*} properties pPr object
 * @returns a refined P object
 */
function Paragraph(children, properties) {
    properties = properties || {};

    return {
        type: types.paragraph,
        children: children,
        styleId: properties.styleId || 'a',             // TODO: Отладить, почему тут может не быть 'a' - не понятно
        styleName: properties.styleName || null,
        numbering: properties.numbering || null,
        alignment: properties.alignment || "left",
        indent: properties.indent || {left:0, right:0, hanging:0,firstLine:0},
        spacing: properties.spacing || { before: 0, after: 8, line: 0 },
        keepLines: properties.keepLines || false,
        keepNext: properties.keepNext || false,
        outlineLvl: properties.outlineLvl || "0",
        rPr: properties.rPr || null,
        sectPr: properties.sectPr
    };
}

/**
 * Returns an object from run children and properties.
 * Makes it a bit more handy than it was in body-reader
 * @param {*} children Run children (a single text object most of the time)
 * @param {*} properties rPr object
 * @returns a refined R object
 */
function Run(children, properties) {
    properties = properties || {};

    var lRPB = false;
    for (var i = 0; i < children.length; i++) {
        if (children[i].type == "lastRenderedPageBreak") {
            lRPB = true;
        }
    }
    return {
        type: types.run,
        children: children,
        styleId: properties.styleId,
        font: properties.font,
        fontSize: properties.fontSize,
        color:properties.color,
        isBold: properties.isBold,
        isItalic: properties.isItalic,
        underline: properties.underline,
        underlineColor: properties.underlineColor,
        isStrikethrough: properties.isStrikethrough,
        isDStrikethrough: properties.isDStrikethrough,
        verticalAlignment: properties.verticalAlignment || verticalAlignment.baseline,
        highlight: properties.highlight,
        isCaps: properties.isCaps,
        isSmallCaps: properties.isSmallCaps,
        lastRenderedPageBreak: lRPB
    };
}

var verticalAlignment = {
    baseline: "baseline",
    superscript: "superscript",
    subscript: "subscript"
};

function Text(value) {
    return {
        type: types.text,
        value: value
    };
}

function Tab() {
    return {
        type: types.tab
    };
}

function Hyperlink(children, options) {
    return {
        type: types.hyperlink,
        children: children,
        href: options.href,
        anchor: options.anchor,
        targetFrame: options.targetFrame
    };
}

function NoteReference(options) {
    return {
        type: types.noteReference,
        noteType: options.noteType,
        noteId: options.noteId
    };
}

function Notes(notes) {
    this._notes = _.indexBy(notes, function(note) {
        return noteKey(note.noteType, note.noteId);
    });
}

Notes.prototype.resolve = function(reference) {
    return this.findNoteByKey(noteKey(reference.noteType, reference.noteId));
};

Notes.prototype.findNoteByKey = function(key) {
    return this._notes[key] || null;
};

function Note(options) {
    return {
        type: types.note,
        noteType: options.noteType,
        noteId: options.noteId,
        body: options.body
    };
}

function commentReference(options) {
    return {
        type: types.commentReference,
        commentId: options.commentId
    };
}

function comment(options) {
    return {
        type: types.comment,
        commentId: options.commentId,
        body: options.body,
        authorName: options.authorName,
        authorInitials: options.authorInitials
    };
}

function noteKey(noteType, id) {
    return noteType + "-" + id;
}

function Image(options) {
    return {
        type: types.image,
        read: options.readImage,
        altText: options.altText,
        path: options.path
    };
}

function TableCell(children, options) {
    options = options || {};
    return {
        type: types.tableCell,
        children: children,
        colSpan: options.colSpan == null ? 1 : options.colSpan,
        rowSpan: options.rowSpan == null ? 1 : options.rowSpan
    };
}


function lastRenderedPageBreak(){
    return({
        type: "lastRenderedPageBreak",
        breakType: "afterPage"
    });
}

function Break(breakType) {
    return {
        type: types["break"],
        breakType: breakType
    };
}

function BookmarkStart(options) {
    return {
        type: types.bookmarkStart,
        name: options.name
    };
}

exports.document = exports.Document = Document;
exports.paragraph = exports.Paragraph = Paragraph;
exports.run = exports.Run = Run;
exports.Text = Text;
exports.tab = exports.Tab = Tab;
exports.Hyperlink = Hyperlink;
exports.noteReference = exports.NoteReference = NoteReference;
exports.Notes = Notes;
exports.Note = Note;
exports.commentReference = commentReference;
exports.comment = comment;
exports.Image = Image;
exports.TableCell = TableCell;
exports.lineBreak = Break("line");
exports.pageBreak = Break("page");
exports.lastRenderedPageBreak = lastRenderedPageBreak;
exports.columnBreak = Break("column");
exports.BookmarkStart = BookmarkStart;

exports.verticalAlignment = verticalAlignment;
