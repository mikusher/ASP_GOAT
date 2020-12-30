<!-- #include file="../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' ado_lib.asp
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
<font class="maintitle">ADO Library</font><br>
<font class="articleauthor">Kenneth W. Richards</font><br>
<a href="http://www.orvado.com" class="tinytext">Orvado Technologies</a>
</p>

<h3>Introduction</h3>

<p>
The ADO library handles all of the database calls the website makes using
library functions.  The only two methods you should ever use from this
library are <kbd>adoExecute</kbd> and <kbd>adoOpenRecordset</kbd>.  These
two methods execute a statement or procedure.  adoOpenRecordset will return
an <kbd>ADODB.Recordset</kbd> object whereas the other does not.
</p>

<p>
This file is included in <kbd>site_lib.asp</kbd> which, conversely, is
included on nearly every page on the site since database calls are needed
on nearly every page.  The library encapsulates the minutia of creating
ADODB objects and opening a database connection so all you need to worry
about is writing proper SQL queries and passing the right parameters to
your stored procedures.
</p>

<p>
The connection string for the database is stored in the <kbd>global.asa</kbd>
file at the root of your web server.  This is the only place the database
string is stored.  The library will read the connection string and use this
to make new connections to the database server.
</p>

<h4>adoRecordsetErrors (sub)</h4>

<blockquote>
	<p>
	Displays all of the errors in the Errors collection of the
	recordset.  This function is called automtically when opening a recordset via the
	<kbd>adoOpenRecordset</kbd> method creates an error.
	</p>

	<p><b class="error">For Internal Use Only</b></p>
</blockquote>


<h4>adoConnect (sub)</h4>

<blockquote>
	<p>
	Opens a new connection to the database (configured in global.asa)
	unless a connection has already been opened.  No need to do pooling
	since that is handled by IIS internally.
	</p>

	<p><b class="error">For Internal Use Only</b></p>
</blockquote>

<h4>adoExecute (sub)</h4>

<blockquote>
	<p>
	Executes a query without returning a recordset.  This method will
	return a number indicating the number of rows that were affected
	by the query.
	</p>

	<p><B>Parameters</B>
		<dl>
		<dt>sStat</dt>
		<dd>Transact-SQL statement to be run by the database server.</dd>
		</dl>
	</p>

	<p><B>Returns</B><br><br>
		<i>Nothing</i>
	</p>

</blockquote>

<h4>adoOpenRecordset (function)</h4>

<blockquote>
	<p>
	Opens a forward-only recordset from the database using the
	supplied query (sQuery).
	</p>

	<p><B>Parameters</B>
		<dl>
		<dt>sStat</dt>
		<dd>Transact-SQL statement to be run by the database server.</dd>
		</dl>
	</p>

	<p><B>Returns</B><br><br>
		ADODB.Recordset object containing the results of the SQL statement.
	</p>
</blockquote>

<h4>adoDisconnect (sub)</h4>

<blockquote>
	<p>Disconnect from the database here</p>

	<p><B>Parameters</B><br><br>
		<i>None</i>
	</p>

	<p><B>Returns</B><br><br>
		<i>Nothing</i>
	</p>
</blockquote>

<h3>Summary</h3>

<p>
All of your database calls should go through the <kbd>ado_lib.asp</kbd> library.
There is no reason to create your own database connection.  By keeping all database
calls using the same library, it will allow us to easily upgrade and improve the
performance of ASP nuke as needed.
</p>

<p>
It will also allow us to write custom error handlers which will appropriately
log any errors that occur due to our database calls.  If you have a need to create
a database connection on your own, you should have this approved before you complete
your module, otherwise it may not be accepted as an approved ASP Nuke module.
</p>

<p>
<font class="tinytext">Last Updated: Sep 29, 2003</font>
</p>

<!-- #include file="../../footer.asp" -->