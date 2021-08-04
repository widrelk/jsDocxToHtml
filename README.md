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
* Full table support (Custom borders, merged cells)
* Inline picures
* Division by pages (If pagebreaks presented, work in progress)
* External links

# Installation
    npm install jsDocxToHtml
# Usage
Import the library and then call its convertToHtml function.
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


  result.html
A string with all generated HTML

  result.comments
An array of objects with a content of the document's comments.
{
  type: 'comment',
  authorInitials - A string with comment's author initials
  authorName - A string with a comment's author name. An email address most of the time;
  content - An array of comment's contents. Most of the time contains a single string;
  date - Comment's date in YYYY-MM-DD format;
  linkTo - An ID of the corresponding <span> inside of generated HTML;
  time - Comment's time in HH-MM-SS format
}

## Acknowledgements

This library is inspired by [Mammoth](https://www.npmjs.com/package/pammoth) and somewhat based on it,
so thanks to Mammoth's author and contributors.