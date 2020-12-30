<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../header_popup.asp" -->

<!-- SECTION_START:Overview -->
<h3>Documentation Module Help</h3>

<p>
The documentation module is a unified system for writing
reference documentation for many different areas of study. It
contains a book authoring system that allows you to create long
intensive documentation or stand-alone (individual) documents.
</p>

<p>
This came about because I wanted to write a lot of technical
reference materials for the ASP Nuke project.  There is so much
material that can be written so I thought it would be best if
the module could organize everything as a book.
</p>

<p>
With this system we can store single page documents or entire
books online and publish them at will.  The visitor to the site
can browse through the documents or do a keyword search to locate
information they are interested in.
</p>

<!-- SECTION_END:Overview -->

<!-- SECTION_START:Document -->

<h3>Document Help</h3>

<p>
Use the document tab to administer the documents you have created.  A
document may be a stand-alone reference article or part of a book.
When creating a document, be aware that it is nothing more than an
HTML page with some attributes (meta-information) associated with it.
</p>

<h4>Document Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Book</td>
	<td class="formd">Indicate the book for which the documentation is to be associated with.  Leave the drop-list blank if you wish to create a stand-alone document.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Parent Section</td>
	<td class="formd">Content sections of your book are related in a hierarchy, just like the table-of-contents for a book.  Choose the parent section under which this document will be placed in the structure of your book.  Leave this blank if you are creating a stand-alone document.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Author</td>
	<td class="formd">Select the author who has written this document.  The options in this list are defined by going to the <b>Author</b> administration tab.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Title</td>
	<td class="formd">Enter a title for this document.  If this document is a section of a book, then the title will be the section heading.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Short Description</td>
	<td class="formd">Enter your administrative comments here.  These comments will not be included in the content of the book and are purely for your reference.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Body</td>
	<td class="formd">This is the body of your document. You may include all types of HTML formatting in this box.  Additionally, you may enter a carriage return to create a line break or two carriage returns to create a paragraph break.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Section Name</td>
	<td class="formd">The section name is used for embedding the contents of one section of a book into another.  You can reference another section "by name" by embedding a macro in the parent content.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Is Inline Content?</td>
	<td class="formd">This is only for documents which are sections in a book.  Answer yes, to embed the contents within the parent content (or append it to the end.)</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Author Notes</td>
	<td class="formd">Enter any author notes that you want to make available to the reader of the section.  Instead of placing these notes inline with the content, they will appear when the user clicks a note icon next to the section.  This is similar to the footnotes in a book.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Document -->

<!-- SECTION_START:Book -->

<h3>Book Help</h3>

<p>
You should create a book if you have a large number of related
documents that would be better organized if they could be placed
in a book and organized with a table-of-contents.  Otherwise,
you should just create individual documents.
</p>

<p>
You can create as many books as you like for your web site.  The
books are browsable on the public side through the book module capsule
which will appear on the left or right-hand side.  You may also create
a menu item (using the Menu module) that takes the user to an index of
all available books.
</p>

<p>
Please be aware that although we allow you to specify versions of
your book, we don't keep a history of previous versions of your book.
Although this may change in the future, the capability to do so does
not exist yet.
</p>

<p>
In the future, we may provide tools for exporting the contents
of your book into common formats (such as DocBook and maybe even
Microsoft Word&trade;)

<h4>Book Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Author</td>
	<td class="formd">Select the author of this book.  Items from this list are taken from the author list which is defined under the <b>Author</b> tab.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Title</td>
	<td class="formd">Indicate the main title for your book how it should appear to visitors to your site.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Sub Title</td>
	<td class="formd">Optionally, provide a sub-title for your book.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Version</td>
	<td class="formd">This is analagous to the print edition and indicates major revisions to your book.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Publish Date</td>
	<td class="formd">When you change the version number, you should also indicate when the book version was published.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Show Section No's</td>
	<td class="formd">This is mainly for technical documentation.  If <i>yes</i> then a section number will be appended to all sections in the content hierarchy.  So a section in chapter 3, major section 2, minor section 5 would be numbered <kbd>3.2.5</kbd></td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Author Notes</td>
	<td class="formd">Provide any author notes that you want to attach to the book, but DO NOT want included with the book.  This is useful for release notes (when a new version is published.)</td>
</tr>
</table>
</p>

<!-- SECTION_END:Book -->

<!-- SECTION_START:Author -->

<h3>Document Author Help</h3>

<p>
The document authors are used to give credit to the people
who have written the documentation.  Before you can assign
an author to a Book or Document, you must first create an
author.  Once you do, that author will show up in the
<kbd>Author</kbd> drop-list.
</p>

<p>
In most cases, the author corresponds to a single individual
who wrote or contributed the documentation.  In some special
cases, you may want to create a special author such as
"Anonymous Contributor" or "Various Authors" to indicate that
the document was not the work of just one person.
</p>

<h4>Document Author Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Title</td>
	<td class="formd">Prefix or honorary title to prepend to the name such as "Dr." or "Sir"</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>First Name</td>
	<td class="formd">This is the first name for the author.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Middle Name</td>
	<td class="formd">This is the middle name for the author.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Last Name</td>
	<td class="formd">This is the last name for the author.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Surname</td>
	<td class="formd">Enter the surname for the author which is appended after the name (like "Jr." or "Esq.")</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>E-Mail Address</td>
	<td class="formd">The E-mail address where the author may be reached (optional)</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Biography</td>
	<td class="formd">Gives a background on the history of the author and explains their credentials.  Enter a carriage return for a line break and two carriage returns for a paragraph break.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Author -->

<!-- SECTION_START:Type -->

<h3>Document Type Help</h3>

<p>
The document types are created so that you can assign major types to
each document you create.  This is especially useful for stand-alone
documents when you have a lot of documents to deal with.
</p>

<p>
Don't confuse the types with categories which denote a particular area
of interest.  Types are used to denot the major document type and more
commonly refer to the structure or layout of content you will find in
the document.  Some example types would be "Tutorial", "API Document"
or "Lookup Chart".
</p>

<h4>Document Type Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Type Name</td>
	<td class="formd">A title to label the documents with.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Type -->

<!-- #include file="../../../../footer_popup.asp" -->