<!-- #include file="../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' codestandard.asp
'	Display coding standards and practices
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
<!-- #include file="../header.asp" -->

<p>
<font class="maintitle">ASP Coding Standards</font><br>
<font class="articleauthor">Kenneth W. Richards</font><br>
<a href="http://www.orvado.com" class="tinytext">Orvado Technologies</a>
</p>

<h3>Introduction</h3>

<p>
This document outlines the coding conventions we will be using when writing server-side scripting for Orvado and its clients.  Specifically, it details coding standard and practiced pertaining to Microsoft's scripting language Active Server Pages or ASP for short.
</p>

<h3>Variables</h3>

<p>
Microsoft's standard coding practices encourage you to use three letter prefixes for your variable names.  For our purposes, we will only use a single character prefix with no underscores in the name.  The table of prefixes is shown below:
</p>

<table border=0 cellpadding=2 cellspacing=0 class="list">
<tr>
	<td class="listhead">Prefix</td><td class="listhead">&nbsp;&nbsp;</td>
	<td class="listhead">Example</td><td class="listhead">&nbsp;&nbsp;</td>
	<td class="listhead">Uses</td>
</tr>
<tr class="list0">
	<td>a</td><td></td>
	<td><kbd>aMonth</kbd></td><td></td>
	<td>Array of month names</td>
</tr><tr class="list0">
	<td>b</td><td></td>
	<td><kbd>bEditMode</kbd></td><td></td>
	<td>Boolean value may be "True" or "False"</td><td></td>
</tr><tr class="list0">
	<td>c</td><td></td>
	<td><kbd>cPriceCode</kbd></td><td></td>
	<td>Single character</td>
</tr><tr class="list0">
	<td>d</td><td></td>
	<td><kbd>dStart</kbd></td><td></td>
	<td>Date or time value</td>
</tr><tr class="list0">
	<td>f</td><td></td>
	<td><kbd>fAmount</kbd></td><td></td>
	<td>Floating point object (referred to as a double)</td>
</tr><tr class="list0">
	<td>n</td><td></td>
	<td><kbd>nMonth</kbd></td><td></td>
	<td>Numeric integer type</td>
</tr><tr class="list0">
	<td>o</td><td></td>
	<td><kbd>oDict</kbd></td><td></td>
	<td>Object created with Server.CreateObject including dictionary objects</td>
</tr><tr class="list0">
	<td>s</td><td></td>
	<td><kbd>sFirstName</kbd></td><td></td>
	<td>All string and character data (more than one character)</td>
</tr>
</table>


<p>
For Visual Basic programmers, it is important to note that all variables in
ASP have the "variant" type.  It is the variant sub-type that ASP uses to manage all of its "type checking".  It is this sub-type we reference with the Hungarian
notation shown above.
</P>

<P>
Many ASP developers prefer to use the Microsoft recommended three letter prefix
for their variable names, but I find this is often overkill for the size of the
average script in ASP.  There are not many sub-types commonly used in ASP
programming.
</P>

<h3>Formatting</h3>

<p>
Formatting of the source code is important to ensure readability of the code.  This includes how lines of code are indented relative to one another and also how whitespace is used.  What follows are guidelines to use when writing or editing code.
</p>

<ol>
<li>Tab size is 4 characters, no more and no less
<li>Use whitespace between all operators (<kbd>1 + x</kbd>) Not (<kbd>1+x</kbd>).  Also include spaces around all Boolean operators (<kbd>x &lt;&gt; 5</kbd>) not (<kbd>x&lt;&gt;5</kbd>).
<li>Use whitespace between assignment operator (<kbd>x = 1</kbd>) not (<kbd>x=1</kbd>)
<li>Use string concatenation operator (&) to split long lines.  This is especially useful when building long SQL statements.
</ol>

<h3>HTML Formatting</h3>

<p>
Along with formatting the ASP code, we also need to format the HTML code that we build since it is integrated with ASP.  The following standards should be applied when creating or editing files:
</p>

<ol>
<li>Indent all tags between &lt;HEAD&gt; and &lt;/HEAD&gt;
<li>Indent all tags between &lt;TR&gt; and &lt;/TR&gt;
<li>Put all tags between &lt;TD&gt; and &lt;/TD&gt; including the "TD" tags on a single line (only if the line is not too long &gt; 80 chars)
<li>Put tags &lt;TD&gt; and &lt;/TD&gt; on a line by themselves only if the content of the tags is longer than 80 characters.
</ol>

<h3>Conclusion</h3>

<p>
It is hard to get a grasp of all of the coding conventions and practices we are using.  If you are new to our programming practices, you should try browsing some of our scripts to understand how the code is written and how the formatting is done.  You may also find it easier to take an existing script and modify it to your own purposes.  I do this all the time and it saves a lot of time.  We also have templates that will help you to build web scripts.
</p>

<p>
<font class="tinytext">Last Updated: Sep 29, 2003</font>
</p>
<!-- #include file="../footer.asp" -->
