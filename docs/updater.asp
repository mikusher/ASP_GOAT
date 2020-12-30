<!-- #include file="../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' updater.asp
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
<font class="maintitle">Web Site Software Updater</font><br>
<font class="articleauthor">Kenneth W. Richards</font><br>
<a href="http://www.orvado.com" class="tinytext">Orvado Technologies</a>
</p>

<h4>Introduction</h4>

<p>
Our goal with ASP Nuke is to provide software updates and patches in a quick and efficient matter.  Ideally, we would like to see an update "site" which will allow users to pull code (ASP scripts) and database updates (Stored Procedures,
Functions, and Table changes).
</p>

<p>
There are three major components to this sytem which we will describe in detail below.  These are: a packaging system to group all files necessary for an update; a transfer mechanism to deliver the updates to the web sites; a versioning system to track differnt versions of a feature; an installer to manage installing/de-installing updates and the update browser to view all of the available updates.
</p>

<h4>1. Packaging System</h4>

<p>
In order to provide software updates over the web, we need to create a simple update package which will allow us to deliver the following elements:
</p>

<ol>
<li>Web page updates (<kbd>.asp, .html, .css, .js</kbd>)
<li>Graphic image updates (<kbd>.gif, .jpg, .png</kbd>)
<li>Database schema updates (<kbd>.sql</kbd>)
<li>Data updates (<kbd>.dat</kbd>)
</ol>

<p>
If we had very large propellor beanies, we could try to package a huge set of updates into a "zip" archive for delivery.  But, I think it's better to stay true to the nature of the web (which uses lots of little files to build a single web page.)
</p>

<p>
Instead, I propose we just use the filesystem to organize our packages.  We will do this by using major categories for the packages.  Under each category, we will have our update packages which will have a unique folder name.  Then, under each package, we will have individual updates which are organized by date.
</p>
<pre>
CategoryName
+--- PackageName
       +--- VersionName
</pre>
<p>
Thus, we would produce something that looks like the following:
</p>
<pre>
Articles
+--- NewsArticles
     +--- upd20031003
</pre>
<p>
The prefix <kbd>upd</kbd> indicates that this folder contains an update package.  In the future we may define other types of packages such as splitting a package into two packages, migrating one package into another, or rolling back to a previous package.
</p>

<p>
Included in the update folder will be an information file containing details about what is included in the update.  This will be an HTML file that may be displayed within the update browser.  The name for this file will be <kbd>changes.html</kbd> and will be a complete stand-alone HTML file so users may view it manually if needed.
</p>

<h4>2. Transfer Mechanism</h4>

<p>
We need to define a deliver mechanism to send web updates between our server and all of the clients.  The most obvious way is to use the web server itself to send these updates.  And that is just what we are going to do with our update agent.
</p>

<p>
With the advent of XML, Microsoft has provided an excellent XML component for retrieving web pages.  This component is called "<kbd>ServerXMLHTTP</kbd>" and it is what we will use on the client side to pull updates from our server.
</p>

<p>
One of the things we need to be careful of on the server side is that scripts will be interpreted (executed) before the contents are sent back to the client.  We want to send the actual source code and not the interpreted results.  To do this, we need to either encrypt all of the server-side script elements or we need to change the name of the file.
</p>

<p>
I propose that we use a known "mime type" that comes back uninterpreted.  The best would probably be ".txt".  We just have to make sure that we never need to deliver an actual <kbd>.txt</kbd> file to the client.  There really isn't much need to have these types of files on your web server, so we should be okay.
</p>

<h4>3. Versioning System</h4>

<p>
One of the things on my wish list is to be able to track what version a package is WITHOUT relying on the database.  In fact, all of the software should be independant of the database.  That way, if someone loses their entire database, the system can reconstruct the status of the updater and know exactly where (what version) the software is at.
</p>

<p>
We need to store all of the update packages installed by the user since their web site was created.  This way, if their database dies, and they need to create a new database, the updater application can go back through all of the previous updates that were applied and re-apply the database and data updates.
</p>

<p>
In order for this to work properly, they should also store the initial database setup file in their "update" directory as well.  If the web server user is given full control to a new database that is created, the updater application can rebuild the entire schema from scratch.
</p>

<p>
We need to be careful when we build a package that we keep all of the correct version numbers for all of the updates.  Within the "update" directory, the version number will be the folder the package files are found in.  When the update is applied, each file that is moved out of the update folder should be stamped with the version number in the page header as follows:
</p>

<kbd>
' $VERSION 2003.10.03.1
</kbd>

<p>
This way we can have a validator that goes through and checks each file to make sure the correct version has been installed.  It will also make sure that we have an up-to-date package by checking all of the contents of the package.
</p>

<h4>4. Installer / Deinstaller</h4>

<p>
In order for updates to be applied to a web site, an updater needs to go through all of the files and perform the updates.  The process of doing this is rather complicated so we will only touch on it briefly here.  
</p>

<p>
The update files will have to be applied in a certain order.  The most effective order to apply the updates is shown below (although it would be nice if we could control the order of the updates using a "roadmap" file.
</p>

<ol>
<li>Database schema updates (<kbd>.sql</kbd>)
<li>Web page updates (<kbd>.asp, .html, .css, .js</kbd>)
<li>Data updates (<kbd>.dat</kbd>)
<li>Graphic image updates (<kbd>.gif, .jpg, .png</kbd>)
</ol>

<p>
Ideally, we would like to disable a feature on the Nuke site while the updates are being applied.  This way nobody sees the site "in transaction" which in most cases will mean the site is broken.  One saving grace of using <kbd>Server.Execute</kbd> to include our modules, is that it won't break (stop processing) of the entire page and will only show the error in the particular area where the update is being made.
</p>

<p>
Deinstalling an update is a more involved process and will be handled at a later date.  Basically, it involves traversing back through the update history and finding previous versions of files.  Undoing database changes is a whole 'nother can of worms.
</p>

<h4>5. Update Browser</h4>

<p>
The update browser will make a request to the main server (ASP Nuke) using the ServerXMLHTTP component to get the available updates.  These updates will be returned as an HTML file containing a synopsis of all of the available updates.
</p>

<p>
There are two components to the update browser.  One part shows you all of the updates that are available for your existing code.  The other part shows you all of the new modules which can be added to ASP Nuke (that aren't already installed.)  The user must do the update and new module installs separately (meaning two batchs - one batch for all the updates and one for all the new features to add.)
</p>

<p>
For each update, the adminstrator will see the name of the feature, the date it was posted, a link to the changes.html file and a short synopsis describing the change.  Next to each update will be a checkbox which the user must click to mark it for installation. When the user clicks the "Update" button, all of the changes will be downloaded and the system will begin the process of updating the web site.
</p>

<p>
After the installation has been completed, the user will be shown the status of all of the updates.  From here, the user may go back to the list of available updates (which will reflect the changes just made.)
</p>

<p>
A review page will also be useful showing a summary of all of the updates that have been applied to the ASP Nuke installation.  This will be very useful for bug reports and debugging purposes.
</p>

<h4>Summary</h4>

<p>
With an update agent, we greatly simplify the process of upgrading or downgrading features on ASP Nuke.  It provides a simple web-based interface for browsing available updates and downloading and installing updates as needed.  Through a notification system, the Nuke administrator will be alerted when critical updates become available.
</p>

<p>
This update agent may soon become a project of its own.  Its applications will apply to all sorts of web-based projects, not just ASP Nuke.  Our focus will only be to support the database and scripting languages that ASP Nuke uses which are, for now, Active Server Pages and SQL Server 2000.
</p>

<p>
<font class="tinytext">Last Updated: Oct 03, 2003</font>
</p>
<!-- #include file="../footer.asp" -->