<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../header_popup.asp" -->

<!-- SECTION_START:Overview -->
<h3>Language Help</h3>

<p>
The languages administration allows you to manage the various language
translations that allow others to use ASP Nuke using a different
language.  Most of our language translations are submitted by various
volunteers around the world.
</p>

<p>
Using the language admin, you can create new languages that you want
to make available for translation and also control the publishing
of new translations.  The language administration only manages the
interface elements such as labels, buttons and menu items that are
used to navigate and administer the site.
</p>

<!-- SECTION_END:Overview -->

<!-- SECTION_START:Language -->

<h3>Language Help</h3>

<p>
Before any translations can be done, you first need to define translation
languages in the database.  Basically, each language has a two-letter
code which matches the country of origin for that language.   You can
define as many langauges as you like and have translations ongoing in
those languages.
</p>

<p>
You may decide to remove a language translation if you feel it is
incomplete by "unpublishing" it.  Generally, you will not publish a
translation until it is somewhat complete.
</p>

<h4>Language Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Language Code</td>
	<td class="formd">Two letter code that uniquely identifies the language (2-letter country code for the nation of origin).</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Country Name</td>
	<td class="formd">The name of the country for the language origin (for example: "Spain").</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Native Language</td>
	<td class="formd">The name of the language as spoken in the native tongue (for Spanish, this would be Espanol)</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Flag Icon</td>
	<td class="formd">An image icon displaying the flag for the country of origin</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Published</td>
	<td class="formd">Has this language been published (made available) for visitors or administrators of the site?</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Maintained by User</td>
	<td class="formd">The username of the person who is maintaining the language translation.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Language -->

<!-- SECTION_START:Text -->

<h3>Text Help</h3>

<p>
The language text defines the various labels, button text and menu items
which require translation.  These are usually words or short phrases that
are used on a navigation item or other interface element necessary to use
or operate the ASP Nuke application.
</p>

<p>
Unless you are developing a module on your own, you won't really need
to add new language texts.  In order for these to be used by the ASP
Nuke application, they must be referenced within a web page using the
"<kbd>steTxt</kbd>" library function.  For the same reason, you should
not need to update language texts unless you are developing your own
modules.
</p>

<p>
Note that this is only the English words that need to be translated
and not the actual translation itself.  Using a "one-to-many" relationship
in the database, we can provide translations in multiple languages for
each unique English phrase.
</p>

<h4>Language Text Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>English Text</td>
	<td class="formd">English word or phrase that should be translated (255 chars max length).</td>
</tr>
</table>
</p>

<!-- SECTION_END:Text -->

<!-- SECTION_START:Translation -->

<h3>Translation Help</h3>

<p>
The translations area is where you can review the translations that
have been made for each language.  Unlike the public area of the site
where members can submit translations for the various interface elements,
the administration side currently doesn't allow you to do this yet.
</p>

<p>
Select the language you would like to view translations for from the
drop-list at the top of the screen.  The system will display a list of
all translations for the language.  You can then edit or delete the
translation text as needed.
</p>

<h4>Language Translation Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Translation Language</td>
	<td class="formd">Native language name and country of origin for the translation being done.</td>
</tr><tr class="list0">
	<Td class="forml" valign="top" nowrap>English Text</td>
	<td class="formd">English word or phrase that should be translated (255 chars max length).</td>
</tr><tr class="list0">
	<Td class="forml" valign="top" nowrap>Translation</td>
	<td class="formd">Translation of the English word or phrase in the native language.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Translation -->

<!-- #include file="../../../../footer_popup.asp" -->