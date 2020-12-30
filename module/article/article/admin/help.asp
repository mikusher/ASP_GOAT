<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../header_popup.asp" -->

<!-- SECTION_START:Overview -->
<h3>Article Help</h3>

<p>
The articles are the centerpiece of the site and are intended to be
displayed in the center content area of the main site.  They are
loosely modeled after the <a href="http://www.slashdot.org" target="_NEW">SlashDot</a>
web site.
</p>

<h4>Article Components</h4>

<ul>
<li><B>Author</B> - Person who wrote or entered the article
<li><B>Category</B> - Primary category where article is listed
<li><B>Title</B> - Main heading for the article
<li><B>Article Lead-In</B> - Synopsis of article (1st paragraph)
<li><B>Article Body</B> - Main content of the article
</ul>

<p>
The <i>Article Lead-In</i> is a short paragraph that provides a synopsis of
the content of the article.  If you practice good writing style, the first
paragraph should contain a relatively good summary that you can use.
</p>

<p>
The <i>Category</i> is the same concept as departments on Slash Code&trade;.
This allows us to create graphic icons that appear next to articles on the
home page that represent the major section (like the sections of a newspaper)
where the article was placed.
</p>

<h4>Article Content</h4>

<p>
Since the article are completely under the control of the site administrator,
they can enter any HTML code into an article.  So if you are a professional
web designer, you can make your article like a custom web page.  In most cases
it is best to just use little text formatting tags such as bold, italics,
underline, lists, simple tables and hyperlinks.
</p>

<p>
If you do not know much about HTML, you can enter your articles in plain text
and they will look great.  No indenting or text formatting will be preserved
with the exception of the carriage return.  The carriage return will insert a
line break (&lt;br&gt;) into your article.  Using two carriage returns will
effectively make a paragraph break.
</p>

<p>
If you wish to learn HTML, search <a href="http://www.google.com" target="_NEW">Google</a>
for "Beginner's HTML tutorial" and I'm sure you'll find some excellent help.
</p>

<h4>Article Organization</h4>

<p>
On the home page, we display a list of current articles along with a short
summary and links to view comments or read the entire article.  The module
configuration lets you define what constitutes a "current article".  You can
configure it so articles expire after one day or after one month.
</p>

<h4>Archiving Articles</h4>

<p>
You can also configure how to archive your older articles.  You can choose to
archive them daily, weekly, monthly or yearly.  A link to the article archives
is provided on the home page after the last article listed.
</p>

<p>
When a user goes to the aritcle archive, they will see a summary of all of the
time periods which contain articles displayed in chronological order.  The user
may then click on a time period to view all of the articles in that archive.
</p>

<!-- SECTION_END:Overview -->

<!-- SECTION_START:Comments -->

<h3>Article Comments Help</h3>

<p>
Visitors to your site may enter comments regarding your article posts by
simply hitting the "comments" link at the bottom of the article.  Comments
do not appear in the main article listing.  Only a link labeled "comments"
along with a count of the number of comments received.
</p>

<p>
Comments are listed in a threaded discussion format (like a tree hierarchy.)
The indentation of the comments gives visitors a visual representation of
how the comments relate to each other.  Users may insert replies to a
comment anywhere they like.  Siblings are always ordered sequentially by
date.
</p>

<p>
The comments tab allows you to create new replies, edit existing comments
and delete comments that are inappropriate.
</p>

<h4>Article Comments Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Subject</td>
	<td class="formd">Title of the article comment</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Body</td>
	<td class="formd">Main body (content) for the article comment</td>
</tr>
</table>
</p>

<!-- SECTION_END:Comments -->

<!-- SECTION_START:Author -->

<h3>Article Author</h3>

<p>
When you post a new article to your web site, you must assign it an author.
Before you can assign it an author, you must define the list of available
authors using the Article Author administration.
</p>

<p>
The article authors are separate from the members and users who have the
ability to log onto the site.  Authors are only be used for the article module
to attribute the source of the contribution.  You may enter as many authors as
you need and even create special cases such as "Anonymous Contributor" or even
the name of your web site if you don't want to attribute an article to one
specific person.
</p>

<h4>Article Author Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Title</td>
	<td class="formd">This is a title or honarary title that precedes the name such as "Mr.", "Mrs." or "Dr."</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>First Name</td>
	<td class="formd">The first name for the author</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Middle Name</td>
	<td class="formd">The middle name or middle initial for the author</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Last Name</td>
	<td class="formd">The last name for the author</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Surname</td>
	<td class="formd">This is a suffix for the name or honarary title that should follow the name such as "Jr.", "Esq."</td>
</tr>
</table>
</p>

<!-- SECTION_END:Author -->

<!-- SECTION_START:Category -->

<h3>Article Categories Help</h3>

<p>
Article categories are used to group similar articles together
much like the sections of a newspaper are used to organize content.
These categories work in a similar fashion and help you to organize
a large collection of articles.
</p>

<p>
Associated with each article category is an icon image.  If you are
familiar with <A href="http://www.slashdot.org" target="_NEW">SlashDot.org</A>
then you will understand how the category icons are placed next to the article
to create a quick visual aid for the reader who is scanning a long list of
articles.
</p>

<p>
Image uploads must be done in <kbd>JPEG</kbd> or <kbd>GIF</kbd> format.
Generally speaking, you should use JPEG format for photo-realistic images
and GIF format for line drawings.
</p>

<p>
In order to upload images to your ASP Nuke application using this form, you will need to
make sure that your upload directory has been configured to allow write
permission for the web user.  More information about this can be found in
the <kbd>README.txt</KBD> file that comes with the ASP Nuke package.
</p>

<h4>Article Category Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Name</td>
	<td class="formd">Title of the category</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Icon Image</td>
	<td class="formd">If you have loaded the category icons on your server manually (through FTP or some other method), indicate the path and name to the icon image here (like:  <kbd>/images/article/something.gif</kbd>)</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Upload Icon Image</td>
	<td class="formd">To upload a category icon image from your local machine, click on the <kbd>browse</kbd> button to select the file.  When the form is submitted, the icon image will be uploaded and associated with the category.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Comments</td>
	<td class="formd">Include some short comments about the category.  This synopsis will be publically visible when a visitor goes to browse an article category.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Category -->

<!-- #include file="../../../../footer_popup.asp" -->