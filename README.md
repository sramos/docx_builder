# docx_builder

Programatically build Microsoft Word documents from ruby code, using the Office Open XML format.

## Overview

docx\_builder lets you create a new Word document from a template document. Then you can make method calls to add content to the end of the document, building it from the top down just as if you were writing it in Word.

  Templates for docx_builder are just regular documents, not Word or Excel templates. If your template document contains text, this text will be in every new document that you create from it. The easiest way to create a new template document is to write it with Word, adjust the styling to your requirements, and then delete all the contents.

## Details

Word documents are represented by `Word::WordDocument` objects. The following methods (there are others) are available to work with Word documents:

`.blank_document(*options={}*)` - Create a new document. If you supply a `:base_document`option then the new document is based on the document you supply. Otherwise it is initialized from an internal blank document.

`.add_heading(*text*)` - Add the specified text using the Heading1 style. Returns the paragraph.

`.add_sub_heading(*text*)` - Add the specified text using the Heading2 style. Returns the paragraph.

`.add_paragraph(*text*, *options={}*)` - Adds the specified text as a new paragraph. If you supply a `:style` option sets the style of the text accordingly. Returns the paragraph.

`.add_image(*image*, *options={}*)` - Inserts the supplied `Magick::Image` or `ImageList` object as a new image. If you supply a `:style` option sets the style of the text accordingly.

`.add_table(*hash*, *options={}*)` - Inserts a table. By default, the keys of the hash provide the first row of data (column headings) in the table, and the values provide successive rows. If you supply a `:table_style` option it will be applied to the created table. If you supply a `:column_widths` option (as an array of twip measurements, 1440 twips = 1 inch) then the columns will be supplied accordingly. If you set the `:skip_header` option to true then the keys will not be rendered as a table row. If you supply a `:column_styles` option (as an array of strings) then the styles will be applied to the corresponding columns in the table.

`.replace_all(*source_text*, *replacement*)` - Searches and replaces in the document.

## Acknowledgements

docx\_builder is based on office\_docs by [mwelham](https://github.com/mwelham). office_docs supports creation of both Word and Excel documents, and has been developed in a slightly different direction.

This work was supported by [Labrador Omnimedia](http://labradorom.com).


