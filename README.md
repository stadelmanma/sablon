# Sablon

[![Gem Version](https://badge.fury.io/rb/sablon.svg)](http://badge.fury.io/rb/sablon) [![Build Status](https://travis-ci.org/senny/sablon.svg?branch=master)](https://travis-ci.org/senny/sablon)

Is a document template processor for Word `docx` files. It leverages Word's
built-in formatting and layouting capabilities to make template creation easy
and efficient.

*Note: Sablon is still in early development. Please report if you encounter any issues along the way.*

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'sablon'
```


## Usage

```ruby
require "sablon"
template = Sablon.template(File.expand_path("~/Desktop/template.docx"))
context = {
  title: "Fabulous Document",
  technologies: ["Ruby", "HTML", "ODF"]
}
template.render_to_file File.expand_path("~/Desktop/output.docx"), context
```


### Writing Templates

Sablon templates are normal Word documents (`.docx`) sprinkled with MailMerge fields
to perform operations. The following section uses the notation `«=title»` to
refer to [Word MailMerge](http://en.wikipedia.org/wiki/Mail_merge) fields.

A detailed description about how to create a template can be found [here](misc/TEMPLATE.md)

#### Content Insertion

The most basic operation is to insert content. The contents of a context
variable can be inserted using a field like:

```
«=title»
```

It's also possible to call a method on a context object using:

```
«=post.title»
```

NOTE: The dot operator can also be used to perform a hash lookup.
This means that it's not possible to call methods on a hash instance.
Sablon will always try to make a lookup instead.

This works for chained method calls and nested hash lookup as well:

```
«=buyer.address.street»
```

##### WordProcessingML

Generally Sablon tries to reuse the formatting defined in the template. However,
there are situations where more fine grained control is needed. Imagine you need
to insert a body of text containing different formats. If you can't decide the
format ahead of processing time (in the template) you can insert
[WordProcessingML](http://en.wikipedia.org/wiki/Microsoft_Office_XML_formats)
directly.

It's enough to use a simply insertion operation in the template:

```
«=long_description»
```

To insert WordProcessingML prepare the context accordingly:

```ruby
word_processing_ml = <<-XML.gsub("\n", "")
<w:p>
<w:r w:rsidRPr="00B97C39">
<w:rPr>
<w:b />
</w:rPr>
<w:t>this is bold text</w:t>
</w:r>
</w:p>
XML

context = {
  long_description: Sablon.content(:word_ml, word_processing_ml)
}
template.render_to_file File.expand_path("~/Desktop/output.docx"), context
```

IMPORTANT: This feature is very much *experimental*. Currently, the insertion
    will replace the containing paragraph. This means that other content in the same
    paragraph is discarded.

##### Images [experimental]

For inserting images into a document, you have to follow some rules within the `.docx` template:

* Create a MERGEFIELD called `@image:start`
* Create an image placeholder
* Create a MERGEFIELD called `@image:end`

Note that the context variable name is arbitrary, we can use any name like `@profile.photo:start` and so on.


A special naming convention must be used when defining the context so the gem knows to read the actual content at the path provided. Alternatively you can process the image imediately using `Sablon.content`
```
{
  'image:my_image' => '/path/to/image.jpg',
  my_image2: Sablon.content(:image, '/path/to/image.jpg')
}

template.render_to_file()output_path, context)
```
For a complete example see the test file located on `test/image_test.rb`.

This functionality was inspired in the [kubido fork](https://github.com/kubido/sablon) for this project - kubido/sablon

##### HTML

Similar to WordProcessingML it's possible to use html as input while processing the template. You don't need to modify your templates, a simple insertion operation is sufficient:

```
«=article»
```

To use HTML insertion prepare the context like so:

```ruby
html_body = <<-HTML
<div>
  This text can contain <em>additional formatting</em>
  according to the <strong>HTML</strong> specification.
</div>

<p style="text-align: right; background-color: #FFFF00">
  Right aligned content with a yellow background color
</p>

<div>
  <span style="color: #123456">Inline styles</span> are possible as well
</div>

<h3>HTML tables can be parsed</h3>
<table>
  <tr>
    <th>First Name</th><th>Last name</th>
  </tr>
  <tr>
    <td>Foo</td><td>Bar<td>
  </tr>
</table>

<h3>Footnotes</h3>
<p>
  It is possible to add new footnote references<footnoteref placeholder="test"/>
  to your document as well. You will need at least one footnote in the template document so the docx archive contains the proper files and references. This pre-existing footnote can be wrapped in a comment block and does not have to be present in the final output.
</p>
<footnote placeholder="test">
  My new footnote reference
</footnote>

<h3>Special Fields can be added</h3>
<p>
  There is support for the creation of new fields by using the "ins" tag. The content between the tags is directly used as the field instructions visible when "Toggle Field Codes" is clicked on. To output a literal backslash in some cases (i.e. ruby code) you will need to use two instead of one. To update all fields in the document use Ctrl+A, F9
  <br>
  Current Timestamp: <ins>DATE \\@ "yyyy-dd-MM hh:mm:ss"</ins>
  <br>
  You can create new merge fields this way, however they will not be processed by Sablon prior to insertion.
  <ins placeholder="new_field">MERGEFIELD new_field \\* MERGEFORMAT</ins>
</p>

<h3>Bookmarks and Captions</h3>
<p>
  You can create new bookmarks in the document using 'bookmark' tags containing the desired content. Use dashes or underscores instead of spaces in the name attribute.
  <br>
  <bookmark name="test-new_name">Bookmarked Content</bookmark>
  <br>
  Existing and newly created bookmarks can be referenced using the previously mentioned 'ins' tag.
  <ins>REF test-new_name \\h</ins>
  <br>
  <br>
  Additionally, you can easily create captions using the 'caption' element. The type of caption and it's name are defined by attributes while the content of the caption goes between the tags, omitting the "Figure/Table/etc. ##" portion as it is added automatically.
</p>
<caption type="figure" name="fig-cap">My caption's content</caption>
HTML
context = {
  article: Sablon.content(:html, html_body) }
  # alternative method using special key format
  # 'html:article' => html_body
}
template.render_to_file File.expand_path("~/Desktop/output.docx"), context
```

It is recommended that the block level tags are not nested within each other, otherwise the final document may not generate as anticipated. List tags (`ul` and `ol`) and inline tags (`span`, `b`, `em`, etc.) can be nested as deeply as needed, except for the `<ins>` tag, it can only contain plain text. Additionally, for best results it is best to define all styles being used inside a comment block so they are sure to be included in the final template. For an example see the insertion_template.docx file in test/fixtures.

Not all tags are supported. Currently supported tags are defined in [converter.rb](lib/sablon/html/converter.rb) for paragraphs in method `prepare_paragraph` and for text runs in `prepare_run`.

Basic conversion of CSS inline styles into matching WordML properties in supported through the `style=" ... "` attribute in the HTML markup. Not all possible styles are supported and only a small subset of CSS styles have a direct WordML equivalent. Styles are passed onto nested elements. The currently supported styles are defined in [ast.rb](lib/sablon/html/ast.rb) for each WordML node type. Simple toggle properties that aren't directly supported can be added using the `text-decoration: ` style attribute with the proper WordML tag name as the value. Paragraph and Run property reference can be found at:
  * http://officeopenxml.com/WPparagraphProperties.php
  * http://officeopenxml.com/WPtextFormatting.php

If you wish to write out your HTML code in an indented human readable fashion, or you are pulling content from the ERB templating engine in rails the following regular expression can help eliminate extraneous whitespace in the final document.
```ruby
# define block level tags
blk_tags = 'h\d|div|p|br|ul|ol|li|table|tr|th|td'
# combine all white space
html_str = html_str.gsub(/\s+/, ' ')
# clear any white space between block level tags and other content
html_str.gsub(%r{\s*<(/?(?:#{blk_tags}).*?)>\s*}, '<\1>')
```

IMPORTANT: Currently, the insertion will replace the containing paragraph. This means that other content in the same paragraph is discarded.


#### Conditionals

Sablon can render parts of the template conditionally based on the value of a
context variable. Conditional fields are inserted around the content.

```
«technologies:if»
    ... arbitrary document markup ...
«technologies:endIf»
```

This will render the enclosed markup only if the expression is truthy.
Note that `nil`, `false` and `[]` are considered falsy. Everything else is
truthy.

For more complex conditionals you can use a predicate like so:

```
«body:if(present?)»
    ... arbitrary document markup ...
«body:endIf»
```

#### Loops

Loops repeat parts of the document.

```
«technologies:each(technology)»
    ... arbitrary document markup ...
    ... use `technology` to refer to the current item ...
«technologies:endEach»
```

Loops can be used to repeat table rows or list enumerations. The fields need to
be placed in within table cells or enumeration items enclosing the rows or items
to repeat. Have a look at the
[example template](test/fixtures/cv_template.docx) for more details.


#### Nesting

It is possible to nest loops and conditionals.

#### Comments

Sometimes it's necessary to include markup in the template that should not be
visible in the rendered output. For example when defining sample numbering
styles for HTML insertion.

```
«comment»
    ... arbitrary document markup ...
«endComment»
```

### Executable

The `sablon` executable can be used to process templates on the command-line.
The usage is as follows:

```
cat <context path>.json | sablon <template path> <output path>
```

If no `<output path>` is given, the document will be printed to stdout.


Have a look at [this test](test/executable_test.rb) for examples.

### Examples

#### Using a Ruby script

There is a [sample template](test/fixtures/cv_template.docx) in the
repository, which illustrates the functionality of sablon:

<p align="center">
  <img
  src="https://raw.githubusercontent.com/senny/sablon/master/misc/cv_template.png"
  alt="Sablon Template"/>
</p>

Processing this template with some sample data yields the following
[output document](test/fixtures/cv_sample.docx).
For more details, check out this [test case](test/sablon_test.rb).

<p align="center">
  <img
  src="https://raw.githubusercontent.com/senny/sablon/master/misc/cv_sample.png"
  alt="Sablon Output"/>
</p>

#### Using the sablon executable

The [executable test](test/executable_test.rb) showcases the `sablon`
executable.

The [template](test/fixtures/recipe_template.docx)

<p align="center">
  <img
  src="https://raw.githubusercontent.com/senny/sablon/master/misc/recipe_template.png"
  alt="Sablon Output"/>
</p>

is rendered using a [json context](test/fixtures/recipe_context.json) to provide
the data. Following is the resulting [output](test/fixtures/recipe_sample.docx):

<p align="center">
  <img
  src="https://raw.githubusercontent.com/senny/sablon/master/misc/recipe_sample.png"
  alt="Sablon Output"/>
</p>

## Contributing

1. Fork it ( https://github.com/senny/sablon/fork )
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request


## Inspiration

These projects address a similar goal and inspired the work on Sablon:

* [ruby-docx-templater](https://github.com/jawspeak/ruby-docx-templater)
* [docx_mailmerge](https://github.com/annaswims/docx_mailmerge)
