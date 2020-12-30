<!-- #include file="../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' site_lib.asp
'	Display module development instructions
'
' Copyright (C) 2002 Orvado Technologies (http://www.orvado.com)
'
' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'--------------------------------------------------------------------
%>
<!-- #include file="../../header.asp" -->

<p>
<font class="maintitle">Site Library</font><br>
<font class="articleauthor">Kenneth W. Richards</font><br>
<a href="http://www.orvado.com" class="tinytext">Orvado Technologies</a>
</p>

<h3>Introduction</h3>

<p>
The site library contains all of the common functions that are used
on all of the pages on the site.  These "global" functions may be
used on any page on the site.  All pages must include the site library
using a call to:
</p>

<kbd>&lt;!-- #include virtual="/lib/site_lib.asp" --&gt;</kbd>

<p>
This library will include the database library for you (<a href="ado_lib.asp">ado_lib.asp</a>)
In most cases, this will be the only file you will include containing strictly
code.
</p>

<h4>steForm (function)</h4>

<blockquote>
	<p>
	Retrieve the value of a querystring or form variable.  The querystring
	is checked first, and if a value is found there, it is returned.  Otherwise,
	the form collection is checked for the value.  This is equivalent to:
	</p>

	<kbd>
	If Request.QueryString(sName) <> "" Then<br>
		sValue = Request.QueryString(sName)<br>
	Else<Br>
		sValue = Request.Form(sName)<br>
	End If
	</kbd>

	<p><B>Parameters</B><br><br>
		<dl>
		<dt>sName</dt>
		<dd>Name of the form variable to retrieve</dd>
		</dl>
	</p>

	<p><B>Returns</B><br><br>
		String containing the value of the querystring or form variable.
	</p>
</blockquote>

<h4>steNForm (function)</h4>

<blockquote>
	<p>
	Retrieve the value of a querystring or form variable and convert it to an
	integer value.  If a numeric value cannot be deciphered from the form data,
	a value of 0 (zero) is returned.
	</p>

	<p><B>Parameters</B><br><br>
		<dl>
		<dt>sName</dt>
		<dd>Name of the form variable to retrieve</dd>
		</dl>
	</p>

	<p><B>Returns</B><br><br>
		Integer parsed from the value of the querystring or form variable.
	</p>
</blockquote>

<h4>steEncForm (function)</h4>

<blockquote>
	<p>
	Retrieve the value of a querystring or form variable and perform an HTML
	encoding on it.  HTML encoding is used to escape special HTML characters
	such as "less than" (&lt;) and "greater than" (&gt;) so that text can be
	shown on a web page properly.
	</p>

	<p><B>Parameters</B><br><br>
		<dl>
		<dt>sName</dt>
		<dd>Name of the form variable to retrieve</dd>
		</dl>
	</p>

	<p><B>Returns</B><br><br>
		HTML-encoded string from the value of the querystring or form variable.
	</p>
</blockquote>

<h4>steRecordValue (function)</h4>

<blockquote>
	<p>
	Retrieve a value from the recordset object if the recordset is not
	empty and the field exists, otherwise, grab the form or querystring
	value.
	</p>

	<p>
	This function (and </kbd>steRecordEncValue</kbd>) are extremely useful
	for building web forms.  It allows you to pull a value from a recordset
	and display it on the page.  This function is used extensively on the
	ASP Nuke site.
	</p>

	<p><B>Parameters</B><br><br>
		<dl>
		<dt>rs</dt>
		<dd>ADODB.Recordset object containing the field</dd>
		<dt>sField</dt>
		<dd>Name of the field which holds the value</dd>		</dl>
	</p>

	<p><B>Returns</B><br><br>
		A field value from the ADODB.Recordset parameter or the querystring/form collection.
	</p>
</blockquote>

<H4>steNRecordValue (function)</h4>

<blockquote>
	<p>
	Retrieve a value from the database and convert it to an integer.
	This method is rarely (if ever) used because most fields that are
	required to be an integer will be defined as such in the database
	schema.  
	</p>

	<p>
	This may be used to eliminate null values from a field since it
	will convert null values to the integer value 0 (zero.)
	</p>

	<p><B>Parameters</B><br><br>
		<dl>
		<dt>rs</dt>
		<dd>ADODB.Recordset object containing the field</dd>
		<dt>sField</dt>
		<dd>Name of the field which holds the value</dd>
		</dl>
	</p>

	<p><B>Returns</B><br><br>
		A field value from the ADODB.Recordset parameter or the querystring/form
		collection converted to an integer.
	</p>
</blockquote>

<h4>steRecordEncValue (function)</h4>

<blockquote>
	<p>
	Retrieve a value from the database with HTML-encoding for placement
	within a form element such as a TEXT or TEXTAREA input element.  If
	the field value is null, this function will convert the value to the
	empty string.
	</p>

	<p><B>Parameters</B><br><br>
		<dl>
		<dt>rs</dt>
		<dd>ADODB.Recordset object containing the field to encode</dd>
		<dt>sField</dt>
		<dd>Name of the field that we want to retrieve</dd>
		</dl>
	</p>

	<p><B>Returns</B><br><br>
		HTML-encoded field value from the ADODB.Recordset object or the
		querystring/form collection.
	</p>
</blockquote>

<h4>steDateValue (function)</h4>

<blockquote>
	<p>
	Returns a formatted date value from the database using the <kbd>vbGeneralDate</kbd>
	format.  If a valid date is not found (like when the value is null,) the function
	will return a value of <i>n/a</i>.  This allows you to do a safe date format
	</p>

	<p><B>Parameters</B><br><br>
		<dl>
		<dt>rs</dt>
		<dd>ADODB.Recordset object containing the date field</dd>
		<dt>sField</dt>
		<dd>Name of the database field containing the date to format</dd>
		</dl>
	</p>

	<p><B>Returns</B><br><br>
		Formatted date value from the ADODB.Recordset object or <i>n/a</i> if
		the value is NULL.
	</p>
</blockquote>

<h4>steStripHTML</h4>

<blockquote>
	<p>
	Strip out all HTML from the content parameter (passed by reference)
	and strip out any unnecessary whitespace.  This does an "intelligent"
	stripping of the HTML tags defined by the HTML 4.0 specification.
	</p>

	<p>
	For certain block-level elements like HTML comments and <kbd>SCRIPT</kbd>
	blocks, all enclosed content will be stripped as well.  Other block level
	elements like TABLE will try to preserve any content between the open and
	close tags.
	</p>

	<p>
	This will strip the leading and trailing whitespace and also replace more
	than two repeated blank lines (containing only whitespace and a carriage
	return) with two line breaks.
	</p>

	<p>
	Uses the Regular Expression object provided in IIS 5.0 for efficiencies
	sake.
	</p>

	<p><B>Parameters</B><br><br>
		<dl>
		<dt>sContent (ByRef)</dt>
		<dd>Text to be stripped of HTML and whitespace</dd>
		</dl>
	</p>

	<p><B>Returns</B><br><br>
		String containing the cleaned content.
	</p>
</blockquote>

<h3>Summary</h3>

<p>
The goal with the <kbd>site_lib.asp</kbd> library is to define only the most
generic functions that are used throughout the site for basic functionality.
It also includes all other common include files (currently only <a href="ado_lib.asp">ado_lib.asp</a>)
that are needed.
</p>

<p>
Make sure all of your main scripts (everything except module code) always include site_lib.
Try to avoid using <kbd>Request.Form</kbd> and <kbd>Request.QueryString</kbd> for
anything except basic text values.
</p>

<p>
<font class="tinytext">Last Updated: Sep 29, 2003</font>
</p>

<!-- #include file="../../footer.asp" -->
