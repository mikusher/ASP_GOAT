<!-- #include file="../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' addinmodules.asp
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
<font class="maintitle">Language Translations</font><br>
<font class="articleauthor">Kenneth W. Richards</font><br>
<a href="http://www.orvado.com" class="tinytext">Orvado Technologies</a>
</p>

<h4>What Needs to be Translated</h4>

<p>
The translation of the web site into different languages will take place in three stages.  Naturally, not all languages will be developed at the same pace so some languages will be finished before others.  Also, we will divide the translation work into three main areas:
</p>

<ol>
<li>ASP Nuke framework
<li>ASP Nuke modules
<li>Reference & Technical documentation
</ol>

<p>
Each of these three areas may be developed independently.  We will track in the database the various levels of translation and keep a list of translations that need to be done.
</p>

<p>
Within each area, we will group components that need to be translated.  These are basically categories of items requiring translation.  They are as follows:
</p>

<ul>
<li>Form Labels</li>
<li>Button Text</li>
<li>Action Link</li>
<li>Help Text</li>
<li>Web Content</li>
<li>Error Messages</li>
<li>Database Records</li>
</ul>

<p>
<b>Form Labels</b> are simply the heading that appear before form inputs such as the <kbd>TEXT</kbd>, <kbd>TEXTAREA</kbd> and <kbd>SELECT</kbd> input controls.  These indicate what type of data should be entered into the field.
</p>

<p>
<b>Button Text</b> is the text which appears on buttons which normally appear on a web form.  Similarily, <b>Action Links</b> are the operations which you may perform on an object such as "edit" or "delete".
</b>

<p>
<b>Help Text</b> is an optional support feature which gives helpful instructions on how to use a form or use the ASP Nuke application in general.  These can popup for an individual form input or they can appear as a general purpose tutorial for a specific application.
</p>

<p>
<b>Web Content</b> is a "catch-all" area for any content that does not fall under any of the other areas.  Basically, this includes all content besides form elements and help text that is displayed on a web page.
</p>

<p>
<b>Error Messages</b> are the warnings and errors which appear when a person uses the web site incorrectly.  They only appear after the user makes a mistake and usually appear in a red font color to stand out more.
</p>

<p>
<b>Database Records</b> are supporting records that must be created at install-time for the module or ASP Nuke framework.  The database sometimes holds categories or other classification and configuration settings that must be translated into different languages.
</p>

<p>
<font class="tinytext">Last Updated: Nov 13, 2003</font>
</p>

<!-- #include file="../footer.asp" -->
