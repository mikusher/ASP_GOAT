<!-- #include file="../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' themetemplates.asp
'	Display documentation for the software updater.
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
<font class="maintitle">Theme Templates</font><br>
<font class="articleauthor">Kenneth W. Richards</font><br>
<a href="http://www.orvado.com" class="tinytext">Orvado Technologies</a>
</p>

<h4>Introduction</h4>

<p>
Here we are going to discuss the use of theme templates to control the look and feel of pages on the ASP Nuke web site.  Templates will allow people with no knowledge of programming (but a rudimentary understanding of HTML) to quickly and easily build web page templates.
</p>

<h4>Separating Style and Layout</h4>

<p>
One of the big advances that Cascading Style Sheets gives HTML developers, is that it allows them to separate the layout (done with HTML) with the design (Cascading Style Sheets).  This keeps the code clean and allows the designer and the developer to keep their physical work files separate (though they still need to integrate with each other.)
</p>

<p>
In a similar fashion, we would like to keep the look of our themes separate from the layout.  Keeping with the tradition of HTML, we will use HTML-based templates to control the layout and CSS files to control the "look".
</p>

<h4>HTML Templates</h4>

<p>
Our HTML templates will control the layout for the pages on the Nuke site.  When building a traditional web site, you would create a site header and footer which would be included using a Server-Side Include (SSI) directive.  You would then "wrap" your content pages with the header and footer in this way.
</p>

<p>
The ASP Nuke templates will use a single HTML template for each unique "layout" created for a theme.  This template will have a special XML macro indicating where the main content should be placed within the template. It looks like the following:
</p>

<blockquote>
<pre>&lt;MAINCONTENT/&gt;</pre>
</blockquote>

<p>
All macros in the HTML templates will use XML format because it is a standardized format and it will be hidden if no substitution is performed on the macro as a page is built.
</p>

<p>
Each ASP Nuke theme will come with a set of HTML Templates.  Each template must be designed to work with the stylesheet and graphics that make up the look of the theme.  While other systems only allow for one layout for every single page on the site, we will provide several standardized layout templates along with special templates which may be selected for special cases.  At the bare minimum, each theme must provide at least one theme.
</p>

<p>
The standard HTML templates are listed below.  You are required to define the following template for every theme that you create.  If you wish, you can simply copy the same file to create all of the required templates.
</p>

<dl>
<dt><pre>default.html</pre></dt>
	<dd>The default template used throughout the site where no other template has been configured</dd>
<dt><pre>admin.html</pre></dt>
	<dd>The administration (control panel) template to use when a user logs into the site administration area.
</dl>

<p>
As you can see, there is not much that is required to build a theme. You simply need one HTML template and a stylesheet and any graphics you require.  The reason we have and "admin" template separate from the "default" is that some site developers are forced to work under small resolutions, and a custom admin template will allow them to make the most out of the full width of their screen.
</p>

<h4>ASP Nuke Modules</h4>

<p>
You manage modules in ASP Nuke by creating "named" module groups which organize a series of modules together. Because every theme is different, it will have a different set of groups (since the layout will be changing.)  We need a way to transition the module configuration from one theme to another if we wish to provide smooth switching of themes on the site.
</p>

<p>
<pre>+------------------------------------------+
|                LOGOBAR                   |
+------------------------------------------+
|                MENUBAR                   |
+------------------------------------------+
|       |                          |       |
| LEFT  |       CENTERCOL          | RIGHT |
| COL   |                          | COL   |
|       |                          |       |
|       |                          |       |
|       |                          |       |
+------------------------------------------+
|                FOOTER                    |
+------------------------------------------+</pre>
</p>

<p>
We have standardized sizes for the modules which may be used in the default layout types shown above.

<table border=0 cellpadding=2 cellspacing=0 class="list">
<tr>
	<td class="listhead">Module Type</td>
	<td class="listhead">Width</td>
	<td class="listhead">Height</td>
	<td class="listhead">Multiple</td>
	<td class="listhead">Direction</td>
	<td class="listhead">Hide?</td>
</tr><tr class="list0">
	<td>LOGOBAR</td>
	<td>*</td>
	<td>*</td>
	<td>Y</td>
	<td>Horiz</td>
	<td>N</td>
</tr><tr class="list0">
	<td>MENUBAR</td>
	<td>*</td>
	<td>*</td>
	<td>Y</td>
	<td>Horiz</td>
	<td>Y</td>
</tr><tr class="list0">
	<td>LEFTCOL</td>
	<td>140px</td>
	<td>*</td>
	<td>Y</td>
	<td>Vert</td>
	<td>Y</td>
</tr><tr class="list0">
	<td>CENTERCOL</td>
	<td>*</td>
	<td>*</td>
	<td>Y</td>
	<td>Vert</td>
	<td>N</td>
</tr><tr class="list0">
	<td>RIGHTCOL</td>
	<td>140px</td>
	<td>*</td>
	<td>Y</td>
	<td>Vert</td>
	<td>Y</td>
</tr><tr class="list0">
	<td>FOOTER</td>
	<td>*</td>
	<td>*</td>
	<td>Y</td>
	<td>Vert</td>
	<td>N</td>
</tr>
</table>

<h4>Packaging Themes</h4>

<p>
When building a theme for distribution, you need to keep the elements of your theme in a standardized folder structure.  The structure for the themes is outlined below:
</p>

<pre>
root
+-- images
+-- style
+-- js
</pre>


</p>