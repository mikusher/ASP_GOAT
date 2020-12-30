<!-- #include file="../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' moduleframework.asp
'	Display module framework documentation
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
<font class="maintitle">Module Framework</font><br>
<font class="articleauthor">Kenneth W. Richards</font><br>
<a href="http://www.orvado.com" class="tinytext">Orvado Technologies</a>
</p>

<h3>Introduction</h3>

<p>
This document outlines the structure and layout of modules which are
written for the ASP Nuke Content Management System.  These guidelines
will dictate how modules should be built and managed by the web portal
software.
</p>

<p>
For information on developing and customizing existing modules, please
see the <a href="moduledevelopment.asp">module development</a> documentation.
This document contains the second revision of the module framework and is
subject to change in the future.
</p>

<h3>Folder Structure</h3>

<p>
The "folder" or directory structure is used to organize the various modules
which will be used by the ASP Nuke system.  There can be anywhere from 10 to
100 different modules being used on your site.  We need an intelligent way
to organize the various modules.
</p>

<p>
A base folder named "module" contains all of the module code and resource
files (such as images specific to the module).  Underneath this folder are
two more levels which serve to organize modules even further.  The first
level contains the category name (eg: "support" and "account")  The next
level beneath this contains the module folder (eg: "FAQ" or "register")
</p>

<p>
The top-level folder structure containing the categories will also be used
for site operators who are browsing all of the available modules.  It is better
pick a category of modules rather than viewing all of the modules all lumped
into one big list.  Shown below is an example of the current module category
folder structure:
</p>

<pre>
Module
+-- Account
+-- Article
+-- Discuss
+-- Menu
+-- Other
+-- Support
+-- Survey
</pre>

<p>
Remember that the structure above shows only the module categories.  To create
a registration module, you might create a "module folder" named "register"
under the "Account" folder.  For this reason, all modules will be placed three
folders deep from the root of the ASP Nuke installation.
</p>

<h3>Module Contents</h3>

<p>
Within each module, we will define all of the necessary scripts that will be
used to install, configure and execute the module.  In order to make the
modules truly modular, we need to define some naming conventions that will
standardize how the module is used.
</p>

<p>
First let's discuss the folder structure.  An outline of the folders within
the "module folder" are listed below:
</p>

<ul>
<li><kbd>admin</kbd> - Administration scripts
<li><kbd>css</kbd> - Cascading Style-Sheet library code
<li><kbd>images</kbd> -  Graphics and other necessary media
<li><kbd>js</kbd> - Javascript library code
<li><kbd>lib</kbd> - Library code (ASP include files)
<li><kbd>setup</kbd> - Setup and configuration scripts
<li><kbd>sql</kbd> - Database schema and data records
</ul>

<p>
Also within the module folder, we will have some standardized scripts which
are used to integrate the module with the site layout.  These are discussed
in detail below:
</p>

<p>
<B><kbd>capsule.asp</kbd></b> - Used to display the narrow (~140 pixels wide) boxes which
appear in the left-hand or right-hand column of the pages.  Typically, these
boxes are consistant throughout all of the pages on the site.
</p>

<p>
<b><kbd>content.asp</kbd></b> - This script is used to fill the main content area (middle
column) for the ASP Nuke layout. For now, this is only used on the home page and only by
the article module.
</p>

<p>
<b><kbd>index.asp</kbd></b> - This should be the main entry point for your module.  In
many cases, this will be the main menu, an overview of the section, a search tool or an
introduction to your module.
</p>

<p>
Remember that these scripts and folders are only the standard files and folders.  A
module developer may create additional files and folder to support their application.
Some modules may be so complex, they may be like a web site of their own.
</p>

<h3>Navigation Elements</h3>
  
<p>
Typically, a visitor to your ASP Nuke site will get to your module scripts through
links found in the capsule.  The capsule will be placed on your web pages when you
configure the module group layouts.  Alternatively, you can place links to the
module in your dynamic menu or the site links module.
</p>

<p>
Basically, each module should be completely independant of the others.  You should
never have two modules so closely related and intertwined that the operation of one
is totally dependent on the operation of another.  For this reason, all navigation
for a module should only link to pages within the "module folder".
</p>

<h3>Using Shared Resources</h3>

<p>
One thing to keep in mind about modules, is that you are not totally restricted to the
resources defined within the "module folder".  A module can also draw upon the excellent
features of the ASP Nuke portal.  Examples of this are the CSS style sheet (defined under
<kbd>/css/style.asp</kbd> and the standard images and common code libraries.
</p>

<p>
Any module author should study the library code (found under the <kbd>/lib</kbd> folder
to learn how they can use this code to improve their applications.  Documentation will
be forthcoming describing all of the libraries and how they should be used.
</p>

<h3>Dependencies</h3>

<p>
As discussed before in the <i>Navigation Elements</i> section, one module should not
depend on another module.  We will, however, allow module code to be dependent on a
specific version of ASP Nuke.  This way, when a user goes to install a particular
module, the installer can check the required version of ASP Nuke and make sure the
module will work correctly before installing.
</p>

<p>
We will not make any guarantees about backward compatibility.  This means that if you
have a module that works with version 0.50, we are not guaranteeing that it will always
work as new versions (0.55 and 0.6) are released.  The beauty of open source code is
that any incompatible modules will be patched quickly and made available to everybody.

<h3>Database Schema</h3>

<p>
We are currently working on a standard way of defining a database schema and then
loading data records into the database.  As soon as documentation is available for
this, we will make it available on the site.  Here is what we know so far:
</p>

<p>
Under the <kbd>sql</kbd>, you will define two control files.  A folder named
<B>schema.sql</b> will contain the database schema definition and a folder named
<b>data.sql</b> will contain the data records that need to be inserted into the
database.
</p>

<p>
<font class="tinytext">Last Updated: Nov 18, 2003</font>
</p><br>

<!-- #include file="../footer.asp" -->