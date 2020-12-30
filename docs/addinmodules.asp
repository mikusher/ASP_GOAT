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
<font class="maintitle">Add-in Modules</font><br>
<font class="articleauthor">Kenneth W. Richards</font><br>
<a href="http://www.orvado.com" class="tinytext">Orvado Technologies</a>
</p>

<h4>Introduction</h4>

<p>
This document outlines the module scheme that will be used in the ASP Nuke web portal project.  This scheme represents guidelines for all developers who create add-in modules for ASP Nuke and should be followed "to the letter" in order to guarentee module compatibility.
</p>

<p>
What won't be covered in this document are coding standards and database schema design.  Please be aware that we are NOT using any stored procedures or user-defined functions in any of the ASP Nuke code.  This will ensure that the process of porting the code to another database will be a managable task.
</p>


<h4>Organizing Modules</h4>

<p>
All of the module code will be placed in a folder under the "module" root path.  This path is typically the path to the web site plus "/module".  We also add a module category path between the module directory and the module folder.
</p>

<h4>Module Parameters</h4>

<p>
I have already setup a dynamic module parammeter system which allows us to define module parameters as-needed.  No new database tables need to be created and no schema changes are required.  Simply add your variables using the module parameters interface and you are done.
</p>

<p>
If your module requires any type of configuration at all, you should seriously consider storing the values in the "module parameters" area.  This will handle the simple entry of basic types including text, numbers, dates & times, yes/no fields and drop-lists.
</p>

<h4>Templates</h4>

<p>
In order to make the layout and design of ASP Nuke sites flexible and as configurable as possible, I propose we use templated layout pages to control the positioning of elements on the site.

<p>
<font class="tinytext">Last Updated: Oct 03, 2003</font>
</p>

<!-- #include file="../footer.asp" -->
