<!-- #include file="../lib/ado_lib.asp" -->
<!-- #include file="setup_lib.asp" -->
<%
' ---------------------------------------------------------------
' setup2.asp
'	Load the Transact-SQL setup script and setup the database
'
' AUTH:	Ken Richards
' DATE:	07/25/2001
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

Const FSO_FORREADING = 1
Const FSO_FORWRITING = 2

Dim sStat
Dim sAction
Dim sContents
Dim dtModified
Dim sErrorMsg
Dim bWasError
Dim bAllowForward

sAction = LCase(Request.Form("action"))
If sAction <> "" Then
	If Request.Form("database") = "" Then
		sErrorMsg = "You must select a database type below"
	End If
End If
%>
<html>
<head>
	<title>ASP Nuke Setup</title>
	<meta name="author" content="Ken Richards">
	<meta http-equiv="content-type" content="text/html; charset=iso-8859-1">
	<meta name="robots" content="all">
	<meta http-equiv="Pragma" content="no-cache">
	<meta http-equiv="Expires" content="-1">
	<style>
	BODY { font: 10pt/10pt tahoma,helvetica,sans-serif; font-weight: normal; }
	P { font: 10pt/10pt tahoma,helvetica,sans-serif; font-weight: normal; }
	B { font: 10pt/10pt tahoma,helvetica,sans-serif; font-weight: bold; }
	H2 { font: 14pt/14pt tahoma,helvetica,sans-serif; font-weight:bold; color: #805030 }
	H3 { font: 12pt/12pt tahoma,helvetica,sans-serif; font-weight:bold; color: #805030 }
	H4 { font: 10pt/10pt tahoma,helvetica,sans-serif; font-weight:bold; color: #805030 }
	HR { height: 1px; margin-top: 2px; margin-bottom: 2px; padding: 0px; width:100%; color: #F0C0A0; }
	.error { font-family: Tahoma,Helvetica,sans-serif; font-size: 10pt; font-weight: bold; color: #FF0000 }
	A { font: 10pt/10pt Tahoma,Helvetica,sans-serif; color: #A06030; font-weight: bolder; text-decoration: none; }
	A:hover { font: 10pt/10pt Tahoma,Helvetica,sans-serif; color: #8080F0;; font-weight: bolder; text-decoration: none; }
	</style>
</head>
<body>

<% If sAction = "" Or sErrorMsg <> "" Then %>

<h2>Build Database Schema</h2>

<p>
This script will perform the setup of the SQL database schema from the
installation scripts found on the web server.  Should this script fail,
it may mean that your web site is configured incorrectly, or your server
doesn't have access to the database.
</p>

<% If sErrorMsg <> "" Then %>
<p><b class="error"><%= sErrorMsg %></b></p>
<% End If %>

<form method="post" action="setup2.asp">
<input type="hidden" name="action" value="run">
<Table border=0 cellpadding=5 cellspacing=0>
<tr>
	<td nowrap><b>Database Type</b>&nbsp;&nbsp;</td>
	<td>
		<select name="database">
		<option value=""> -- Choose One --
		<option value="MySQL"> MySQL
		<option value="sqlserver7"> SQL Server 7
		<option value="sqlserver2000"> SQL Server 2000
		</select>
	</td>
</tr>
</table>

<p>
<input type="submit" name="_submit" value=" Run Setup Script ">
</p>
</form>

<% ElseIf sAction = "run" Then %>

<h2>Build Database Schema</h2>
<%
	' check for any existing user tables
	Dim bTestSuccess
	Dim bTablesExist
	bTestSuccess = setDBConnTest2(bTablesExist)
	If (Not bTestSuccess) Or bTablesExist Then
		bWasError = True %>

<% If bTablesExist Then %>
<P><B class="{color:red}"><%= setStatusMsg %></B></P>
<% Else %>
<P><B class="{color:red}"><%= steErrorMsg %></B></P>
<% End If %>

<P>
The database you chose to run setup on is not empty, as a security
precaution, we don't allow the ASP Nuke database to be setup on a database
that is not empty.  Please create a new database or empty the current one
in order to proceed with the setup process.
</p>

<P>
Or, if you prefer, you may run the script manually using the SQL Query Analyzer
tool which comes with Microsoft SQL Server.  Running the script on an existing
ASP Nuke database, will wipe out all existing data on the site.
</P>

<P>
You may be seeing this error message because you ran the setup script twice.
If this is the case, you should be able to go to your ASP Nuke site now to
see if it works or not.
</P>
<%
	Else
		' database is empty - first setup the database schema
		setDatabaseType = Request.Form("database")
		bAllowForward = setSetupSchemaSQL
		' if successful - import the table data
		If Not setWasError Then
			bAllowForward = setSetupDataSQL
		End if
		' display any errors or status here
		Call setDisplayStatus
	End If 
	If setWasError Then %>

<P>
If you are having problems communicating to the database, you should
make sure that you have configured the database connection string in
your <B>global.asa</B> configuration file.  This should be in the format:
</p>

<p>
<kbd>Provider=SQLOLEDB;server=yourserverip;driver={SQL Server};uid=yourusername;pwd=yourpassword;database=yourdbname;</kbd>
</P>

<P>
If you cannot solve the problem, it may be that you need to define the
database schema by running the schema creation and data initialization
scripts manually for your database.  These can be found at the following
location (where your ASP Nuke is installed:)
</P>

<UL>
<LI> <kbd>/setup/schema.sql</kbd>
<LI> <kbd>/setup/data.sql</kbd>
</UL>

<P>
If you are using SQL Server 2000 as your database server, then you should run the <kbd>schema.sql</kbd> script.
If you are using MySQL, you should run <kbd>schema_mysql.sql</kbd>.
script under the SQL Server Query Analyzer tool.  If you have any questions,
please e-mail <a href="mailto:support@aspnuke.com">support@aspnuke.com</a>.
</p>

<%	End If ' sErrorMsg <> ""
  End If

' show the wizard buttons only after no errors occurred
Call setWizardButtons(bAllowForward) 
%>
</body>
</html>