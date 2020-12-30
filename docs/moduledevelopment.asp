<!-- #include file="../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' moduledevelopment.asp
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
<!-- #include file="../header.asp" -->

<p>
<font class="maintitle">Module Development</font><br>
<font class="articleauthor">Kenneth W. Richards</font><br>
<a href="http://www.orvado.com" class="tinytext">Orvado Technologies</a>
</p>

<h3>Introduction</h3>

<p>
This document describes how to write a new module for the ASP Nuke
Content Management system.  This instructional document is intended
for a medium to advanced-level ASP developer.  You will need to have
a firm understanding of Active Server Pages, SQL Server database
programming (Transact-SQL) and Cascading Style Sheets (CSS).
</p>

<p>
Note that this document does not contain information about the
automatic module updater / installer and the theme manager, both of which
have yet to be built.  As soon as these components are complete, this
documentation will be updated with instructions on how to add support
for these features.
</p>

<h3>Module Framework</h3>

<p>
There is no real magic to the way our module framework is built.  As you
might imagine, we make use of Server-Side Includes for some modular elements.
But we do so using an IIS feature that allows us to dynamically include files.
</p>

<p>
All modular code is placed in a directory named "module".  Each unique
application is given a unique folder name where we place all of the module code.
This folder name is also used to create an administration feature for your
module.
</p>

<p>
Most style elements which affect how your application looks and feels should
be defined in the Cascading Style Sheets (CSS) file (<kbd>/css/style.css</kbd>.)
You should try to use the pre-defined styles that already exist.  This will
ensure that all themes developed for ASP Nuke will apply their look completely
to your module.
</p>

<h3>Major Components</h3>

The module is the main feature of the ASP Nuke project.  A module, in this case,
has two different meanings.  The first is the modular visible components which make up
the content on the web site.   The other is the code module which can be easily
"plugged" into the framework, setup quickly and integrate smoothly.
</p>

<p>
The small boxes which appear on the left and right-hand column are called
"capsules".  These encapsulate a synopsis of content that you want displayed to
visitors on all of the pages of your site.  These will be configurable through
the admin area so you can add or remove capsules and reorder them however you
like.
</p>

<p>
Sometimes, like in the case of the home page, the center column also is a module (thought
it may not appear with the same style box surrounding it.  For the most part,
the modules will have a common look and feel to make it easier for the visitor
to navigate and understand the site.
</p>

<h3>Dynamic Module Includes</h3>

<p>
Introduced in IIS (Internet Information Server) 5.0, the <kbd>Server.Execute</kbd>
method allows us to dynamically include scripts (physical files) which will
be evaluated and placed within another web page.
</p>

<P>
The one difference between <kbd>Server.Execute</kbd> and server-side includes,
which look like <kbd>&lt;!-- #include file=... --&gt;</kbd>, is that Server.Execute
will evaluate a script and then include the output on your page (where you called
the <kbd>Execute</kbd> method.)  A server-side include, will include the code in
your script BEFORE it gets evaluated.
</p>

<p>
The effect of this, is that you need to make your modules complete "stand-alone"
scripts which are capable of functioning on their own.  This helps us, rather
than hinders us because we can easily test new modules that we develop. It also
means that we can't share code between module code and the main content pages.
</p>

<h3>Database Calls</h3>

<p>
All of the database calls on your page should go through our common database
library code.  It is contained in the common code library folder (<kbd>/lib</kbd>.)
Place the following at the top of your page:
</p>

<kbd>&lt;!-- #include virtual="/lib/ado_lib.asp" --&gt;</kbd>

<p>
Any SQL statement that returns a recordset such as a <kbd>SELECT</kbd> or an
<kbd>EXECUTE</kbd> statement should use the <kbd>adoOpenRecordset</kbd> method,
otherwise, you should use the <kbd>adoExecute</kbd> method.
</p>

<kbd>
&lt;!-- #include virtual="/lib/ado_lib.asp" --&gt;<br><br>
sStat = "SELECT * FROM tblCountry"<br>
Set rsCountry = adoOpenRecordset(sStat)
</kbd>

<kbd>
&lt;!-- #include virtual="/lib/ado_lib.asp" --&gt;<br><br>
sStat = "UPDATE tblCountry SET Active = 1"<br>
Call adoExecute(sStat)
</kbd>

<blockquote>
<a href="lib/ado_lib.asp" class="footerlink">ADO Library Documentation</a>
</blockquote>

<h3>Software License</h3>

<p>
At the beginning of your page, you should place a comment block containing
the name of the script, a short description of the purpose of the file and
the GNU General Public License (GPL) which will look as follows:
</p>

<p><kbd>
&lt;%<br>
'--------------------------------------------------------------------<br>
' moduledevelopment.asp<br>
'	Display module development instructions<br>
'<br>
' Copyright (C) 2003 ASP Nuke (http://www.aspnuke.com)<br>
'<br>
' This program is free software; you can redistribute it and/or<br>
' modify it under the terms of the GNU General Public License<br>
' as published by the Free Software Foundation; either version 2<br>
' of the License, or (at your option) any later version.<br>
'<br>
' This program is distributed in the hope that it will be useful,<br>
' but WITHOUT ANY WARRANTY; without even the implied warranty of<br>
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the<br>
' GNU General Public License for more details.<br>
'<br>
' You should have received a copy of the GNU General Public License<br>
' along with this program; if not, write to the Free Software<br>
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.<br>
'--------------------------------------------------------------------<br>
%&gt;<br>
</kbd></p>

<p>
Except for trivial pages (anything with 10 or more lines of code not counting
server-side includes lines), we require that you add this license to your
module if you want it to be officially included and recognized by ASP Nuke.
This is recommended by the GNU organization to protect the freedom of the
code and ensure the source is open.
</p>

<h3>Creating Capsules</h3>

<h3>Creating Admin</h3>

<p>
The process of creating the admin portion of your module is rather complex so
we've decided to create another separate document for this purpose.
</p>

<p>
<font class="tinytext">Last Updated: Sep 29, 2003</font>
</p>
<!-- #include file="../footer.asp" -->