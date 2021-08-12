# jsDocxToHtml .docx to HTML converter

This library converts given blob of a .docx file created in Microsoft Word
(other word processors are not tested) to raw HTML. It designed to generate
as accurate as possible representation of the .docx file content.

You just need to pass a blob - all styles and document structure is handled inside.
As a result, you will get a string with HTML that will look just like .docx file when rendered.
Alongside with HTML you will find additional data, such as comments with
internal links to the document segments. 

Currently supported:
* Almost full character styling (except for non-standart underlining)
* Almost full list support (With custom patterns like "Chapter 3.2.1:")
* Full paragraphs styling (Text alignment, margins)
* Table support (Custom borders, merged cells)
* Inline picures
* Division by pages (If manual or "auto" pagebreaks presented, see known issues)
* External links

Currently not supported:
* Table headers
* Non-standart underlining
* "ankered" pictures
* Excel and other documents cut-ins
* Numberings with "number maps" (Like Roman numerals)

# Installation
    npm i @ree_n/jsdocxtohtml
# Usage
Import the library and then call its convertToHtml() function.
It will return a promise, that will contain a string with HTML and an array of document's comments when resolved.
Here is an example of library usage indide of a React component

    import jsDocxToHtml from "@ree_n/jsDocxToHtml"

    const [html, setHtml] = useState("");
    const [comments, setComments] = useState([]);

    useEffect(
        () => {
            jsDocxToHtml.convertToHtml(props.blob)
                .then((result) => {
                    setHtml(result.html);
                    setComments(result.comments.map((comment) => CommentElement(comment)))  
            })
        }, [props.blob])


- result.html
A string with all generated HTML

- result.comments
An array of objects with a content of the document's comments.
{\n
  type: 'comment',\n
  authorInitials - A string with comment's author initials\n
  authorName  - A string with a comment's author name. An email address most of the time;\n
  content - An array of comment's contents. Most of the time contains a single string;\n
  date - Comment's date in YYYY-MM-DD format;\n
  linkTo - An ID of the corresponding <span> inside of generated HTML;\n
  time - Comment's time in HH-MM-SS format\n
}
# These ID's and classes are implemented into HTML result and can be useful
## ID's
* Page: "pg" + page index
* Page's content: "content_pg" + page index
* Header: "header_pg" + page index
* Footer: "footer_pg" + page index
* Comment: "comment" + comment ID
## Classes
* Comment: commentArea

# Known issues
- Division by pages

In OOXML, division by pages is passed to the application that works with the document
and not directly presented in the .docx file. At the moment, division is made based on the
lastRenderedPageBreak tag, that you can find inside of a run. This works in most cases,
but sometimes this tag is simply not presented.
This can be solved with calculation of page content's height inside of the library and later adjustions,
but it is very difficult from my perspective.

The simpler solution is to use [Element.clientHeight](https://developer.mozilla.org/en-US/docs/Web/API/Element/clientHeight) after you render the html, but at the moment, I don't know how to correctly implement it here.
- Headers display

For correct display of the headers some additional work is required.
Header's height is stated nowhere inside of the document, just like with division by pages, all work
is done by the application.

After HTML's render finished, find headers and page contents by their id, use clientHeight with header
and add margin-top for the content. Here is an example for use with React:
    useEffect(() => {
            let pagesCount = document.getElementById("DocxContainer").childElementCount
            if (pagesCount > 0) {
                for (let page = 0; page < pagesCount; page++) {
                    let header = document.getElementById(`header_pg${page}`)
                    if (header) {
                        let headerTop = window.getComputedStyle(header).marginTop
                        let margin = parseFloat(header.clientHeight) + parseFloat(headerTop.substring(0, headerTop.length - 2))
                        document.getElementById(`content_pg${page}`).style.marginTop = `${margin}px`
                    }
                }
            }
        })

Also, sometimes there is a blank header with w:std tag in xml, but in Word header is in place. Looks like
it is somehow imported or inherited. Currently not supported.
- Tables width in footers of landscape pages

In OOXML, [table width](http://officeopenxml.com/WPtable.php) defenition is a little bit a mess. In short, it defined in table
properties, but also can be overwritten by column width sometimes, and so on.
In my tests, landscape tables works just fine, but when it comes to tables in footers, if rendered with
column width, stated in xml file, the resulted table often does not match with what you can see in Word.
Correct behaviour in this case is unknown for me.
- Table page breaks

Sometimes, pagebreaks inside of a table row are inconsistent. This can cause some content of a row to stay
on the previous page, while in Word it is transfered to the next page and so on.
- Underline styles

OOXML has a lot more underline options than HTML. At the moment, not all options are supported.
- Page numbering

According to [documentation](http://officeopenxml.com/WPSectionPgNumType.php), page numbering stated in the section properties,
but in my tests, a document with numberings in Microsoft Word did not have this tag stated anywhere.
Therefore, correct behavior in this case is unknown to me. Page numberings are not supported at the moment.
- Line height

Looks like line height is calculated differently for OOXML and HTML. Need further research.
At the moment results in footer overlapping with page content (basically just takes up more space), and similar issues.
## Acknowledgements

This library is inspired by [Mammoth](https://www.npmjs.com/package/pammoth) and somewhat based on it,
so thanks to Mammoth's author and contributors.