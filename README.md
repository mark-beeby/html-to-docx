html-to-docx
============

[![NPM Version][npm-image]][npm-url]

html-to-docx is a js library for converting HTML documents to DOCX format supported by Microsoft Word 2007+, LibreOffice Writer, Google Docs, WPS Writer etc.

It was inspired by [html-docx-js] project but mitigates the problem of documents generated being non-compatiable with word processors like Google Docs and libreOffice Writer that doesn't support [altchunks] feature.

html-to-docx earlier used to use [libtidy] to clean up the html before parsing, but had to remove it since it was causing so many dependency issues due to node-gyp.

### Disclaimer

Even though there is an instance of html-to-docx running in production, please ensure that it covers all the cases that you might be encountering usually, since this is not a complete solution.

Currently it doesn't work with browser directly, but it was tested against React.

## Installation

Use the npm to install foobar.

```bash
npm install @mark-beeby/html-to-docx
```

## Usage

```js
await HTMLtoDOCX(htmlString, headerHTMLString, documentOptions, footerHTMLString, headerConfig)
```

## Example

```js
await HTMLtoDOCX(htmlString, headerHTMLString, documentOptions, footerHTMLString, headerConfig)
```

Here's an example of how to use the headerConfig:
```js
await HTMLtoDOCX(
  htmlString,
  headerHTMLString,
  documentOptions,
  footerHTMLString,
  {
    backgroundImage: {
      url: 'https://example.com/background.jpg',
      width: '100%',
      height: '150px'
    },
    logos: [
      {
        url: 'https://example.com/logo1.png',
        width: '100px',
        height: '50px',
        alignment: 'left'
      },
      {
        url: 'https://example.com/logo2.png',
        width: '80px',
        height: '40px',
        alignment: 'right'
      }
    ]
  }
)
```

Full fledged examples can be found under `example/`

### Parameters

- `htmlString` <[String]> clean html string equivalent of document content.
- `headerHTMLString` <[String]> clean html string equivalent of header. Defaults to `<p></p>` if header flag is `true`.
- `documentOptions` <?[Object]>
  - `orientation` <"portrait"|"landscape"> defines the general orientation of the document. Defaults to portrait.
  - `pageSize` <?[Object]> Defaults to U.S. letter portrait orientation.
    - `width` <[Number]> width of the page for all pages in this section in [TWIP]. Defaults to 12240. Maximum 31680. Supports equivalent measurement in [pixel], [cm] or [inch].
    - `height` <[Number]> height of the page for all pages in this section in [TWIP]. Defaults to 15840. Maximum 31680. Supports equivalent measurement in [pixel], [cm] or [inch].
  - `margins` <?[Object]>
    - `top` <[Number]> distance between the top of the text margins for the main document and the top of the page for all pages in this section in [TWIP]. Defaults to 1440. Supports equivalent measurement in [pixel], [cm] or [inch].
    - `right` <[Number]> distance between the right edge of the page and the right edge of the text extents for this document in [TWIP]. Defaults to 1800. Supports equivalent measurement in [pixel], [cm] or [inch].
    - `bottom` <[Number]> distance between the bottom of text margins for the document and the bottom of the page in [TWIP]. Defaults to 1440. Supports equivalent measurement in [pixel], [cm] or [inch].
    - `left` <[Number]> distance between the left edge of the page and the left edge of the text extents for this document in [TWIP]. Defaults to 1800. Supports equivalent measurement in [pixel], [cm] or [inch].
    - `header` <[Number]> distance from the top edge of the page to the top edge of the header in [TWIP]. Defaults to 720. Supports equivalent measurement in [pixel], [cm] or [inch].
    - `footer` <[Number]> distance from the bottom edge of the page to the bottom edge of the footer in [TWIP]. Defaults to 720. Supports equivalent measurement in [pixel], [cm] or [inch].
    - `gutter` <[Number]> amount of extra space added to the specified margin, above any existing margin values. This setting is typically used when a document is being created for binding in [TWIP]. Defaults to 0. Supports equivalent measurement in [pixel], [cm] or [inch].
  - `title` <?[String]> title of the document.
  - `subject` <?[String]> subject of the document.
  - `creator` <?[String]> creator of the document. Defaults to `html-to-docx`
  - `keywords` <?[Array]<[String]>> keywords associated with the document. Defaults to ['html-to-docx'].
  - `description` <?[String]> description of the document.
  - `lastModifiedBy` <?[String]> last modifier of the document. Defaults to `html-to-docx`.
  - `revision` <?[Number]> revision of the document. Defaults to `1`.
  - `createdAt` <?[Date]> time of creation of the document. Defaults to current time.
  - `modifiedAt` <?[Date]> time of last modification of the document. Defaults to current time.
  - `headerType` <"default"|"first"|"even"> type of header. Defaults to `default`.
  - `header` <?[Boolean]> flag to enable header. Defaults to `false`.
  - `footerType` <"default"|"first"|"even"> type of footer. Defaults to `default`.
  - `footer` <?[Boolean]> flag to enable footer. Defaults to `false`.
  - `font` <?[String]> default font name to be used. Defaults to `Times New Roman`.
  - `fonts` <?[Array]<[Object]>> array of font configurations for embedding custom fonts.
    - `name` <[String]> name of the font to be used in the document.
    - `path` <[String]> absolute path to the TTF font file.
    - `weight` <?[String]> font weight (e.g., 'normal', 'bold'). Defaults to 'normal'.
    - `style` <?[String]> font style (e.g., 'normal', 'italic'). Defaults to 'normal'.
  - `fontSize` <?[Number]> size of font in HIP(Half of point). Defaults to `22`. Supports equivalent measure in [pt].
  - `complexScriptFontSize` <?[Number]> size of complex script font in HIP(Half of point). Defaults to `22`. Supports equivalent measure in [pt].
  - `table` <?[Object]>
    - `row` <?[Object]>
      - `cantSplit` <?[Boolean]> flag to allow table row to split across pages. Defaults to `false`.
  - `pageNumber` <?[Boolean]> flag to enable page number in footer. Defaults to `false`. Page number works only if footer flag is set as `true`.
  - `skipFirstHeaderFooter` <?[Boolean]> flag to skip first page header and footer. Defaults to `false`.
  - `lineNumber` <?[Boolean]> flag to enable line numbering. Defaults to `false`.
  - `lineNumberOptions` <?[Object]>
    - `start` <[Number]> start of the numbering - 1. Defaults to `0`.
    - `countBy` <[Number]> skip numbering in how many lines in between + 1. Defaults to `1`.
    - `restart` <"continuous"|"newPage"|"newSection"> numbering restart strategy. Defaults to `continuous`.
  - `numbering` <?[Object]>
    - `defaultOrderedListStyleType` <?[String]> default ordered list style type. Defaults to `decimal`.
  - `decodeUnicode` <?[Boolean]> flag to enable unicode decoding of header, body and footer. Defaults to `false`.
  - `lang` <?[String]> language localization code for spell checker to work properly. Defaults to `en-US`.
  - `spacing` <?[Object]> configuration for document-wide paragraph and line spacing.
    - `defaultLineSpacing` <?[Number]> default line spacing in points (pt). Defaults to normal line spacing based on font size.
    - `defaultParagraphSpacing` <?[Object]> configuration for space before and after paragraphs.
      - `before` <?[Number]> space before paragraphs in line units (e.g., 0.17 = 0.17 lines). Defaults to 0.17.
      - `after` <?[Number]> space after paragraphs in line units (e.g., 0.17 = 0.17 lines). Defaults to 0.17.
- `footerHTMLString` <[String]> clean html string equivalent of footer. Defaults to `<p></p>` if footer flag is `true`.
- `headerConfig` <?[Object]> configuration for header images and logos. 
  - `backgroundImage` <?[Object]> configuration for the header background image.
    - `url` <[String]> URL of the background image.
    - `width` <[String]> width of the background image (e.g., '100%').
    - `height` <[String]> height of the background image (e.g., '100px'). 
  - `logos` <?[Array]<[Object]>> array of logo configurations. 
    - `url` <[String]> URL of the logo image.
    - `width` <[String]> width of the logo (e.g., '50px'). 
    - `height` <[String]> height of the logo (e.g., '50px'). 
    - `alignment` <"left"|"center"|"right"> alignment of the logo.
- `footerConfig` <?[Object]> configuration for footer images and logos. 
  - `backgroundImage` <?[Object]> configuration for the footer background image.
    - `url` <[String]> URL of the background image.
    - `width` <[String]> width of the background image (e.g., '100%').
    - `height` <[String]> height of the background image (e.g., '100px'). 
  - `logos` <?[Array]<[Object]>> array of logo configurations. 
    - `url` <[String]> URL of the logo image.
    - `width` <[String]> width of the logo (e.g., '50px'). 
    - `height` <[String]> height of the logo (e.g., '50px'). 
    - `alignment` <"left"|"center"|"right"> alignment of the logo.

### Returns

<[Promise]<[Buffer]|[Blob]>>

## Column support

You can introduce columns into the layout using a basic 12 grid structure similar to this:-

```html
<div data-docx-column-group="true">
  <div data-docx-column="8">
    <p>left content</p>
  </div>
  <div data-docx-column="4">
    <p>right column</p>
  </div>
</div>
```
Native columns within docx are not appropriate for HTML conversion, as they overflow left to
right, not onto the next page. To achieve a html-like column division tables are used to control
the layout within the resulting document.

## Page sections

When you want to print sections of content logically separately, with an optional header / footer
control you can use data-section and data-show-header and data-show-footer on content to achieve this.

```html
<div data-section="true" data-show-header="false" data-show-footer="false">
  <div>
    <p>Content on its own page with header and footer suppressed</p>
  </div>
</div>
<div data-section="true" data-show-header="true" data-show-footer="true">
  <div>
    <p>Content starting on a new page with header and footer intact</p>
  </div>
</div>
```
Native columns within docx are not appropriate for HTML conversion, as they overflow left to
right, not onto the next page. To achieve a html-like column division tables are used to control
the layout within the resulting document.

## Notes

Currently page break can be implemented by having div with classname "page-break" or style "page-break-after" despite the values of the "page-break-after", and contents inside the div element will be ignored. `<div class="page-break" style="page-break-after: always;"></div>`

Due to a bug when converting the HTML to a vTree, spaces between spans can be lost. 
For spans with space between it is possible to preprocess a span with a data-force-space='true' 
attribute to reinject that space like so:-

```html
<!-- the space between pre and post is lost -->
<span>pre</span> <span>post</span>

<!-- preprocess your html to force the following -->
<span>pre</span><span data-force-space="true"> </span><span>post</span>
```

CSS list-style-type for `<ol>` element are now supported. Just do something like this in the HTML:
```
  <ol style="list-style-type:lower-alpha;">
    <li>List item</li>
    ...
  </ol>
```
List of supported list-style-types:
- upper-alpha, will result in `A. List item`
- lower-alpha, will result in `a. List item`
- upper-roman, will result in `I. List item`
- lower-roman, will result in `i. List item`
- lower-alpha-bracket-end, will result in `a) List item`
- decimal-bracket-end, will result in `1) List item`
- decimal-bracket, will result in `(1) List item`
- decimal, **(the default)** will result in `1. List item`

Also you could add attribute `data-start="n"` to start the numbering from the n-th.

`<ol data-start="2">` will start the numbering from ( B. b. II. ii. 2. )


Font family doesnt work consistently for all word processor softwares

- Word Desktop work as intended
- LibreOffice ignores the fontTable.xml file, and finds a font by itself
- Word Online ignores the fontTable.xml file, and finds closest font in their font library

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to branch new branches off of develop for contribution.

## Support

[!["Buy Me A Coffee"](https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png)](https://www.buymeacoffee.com/privateOmega)

## License

MIT

[npm-image]: https://img.shields.io/npm/v/html-to-docx.svg
[npm-url]: https://npmjs.org/package/html-to-docx
[html-docx-js]: https://github.com/evidenceprime/html-docx-js "html-docx-js"
[altchunks]: https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.altchunk?view=openxml-2.8.1 "altchunks"
[libtidy]: https://github.com/jure/node-libtidy "libtidy"
[String]: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type "String"
[Object]: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object "Object"
[Number]: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type "Number"
[TWIP]: https://en.wikipedia.org/wiki/Twip "TWIP"
[Array]: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array "Array"
[Date]: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Date "Date"
[Boolean]: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type "Boolean"
[Promise]: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise "Promise"
[Buffer]: https://nodejs.org/api/buffer.html#buffer_buffer "Buffer"
[Blob]: https://developer.mozilla.org/en-US/docs/Web/API/Blob "Blob"
[pixel]: https://en.wikipedia.org/wiki/Pixel#:~:text=Pixels%2C%20abbreviated%20as%20%22px%22,what%20screen%20resolution%20views%20it. "pixel"
[cm]: https://en.wikipedia.org/wiki/Centimetre "cm"
[inch]: https://en.wikipedia.org/wiki/Inch "inch"
[pt]: https://en.wikipedia.org/wiki/Point_(typography) "pt"

## Contributors

<a href="https://github.com/privateomega/html-to-docx/graphs/contributors">
  <img src="https://contrib.rocks/image?repo=privateomega/html-to-docx" />
</a>

Made with [contrib.rocks](https://contrib.rocks).