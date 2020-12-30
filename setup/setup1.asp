<!-- #include file="../lib/ado_lib.asp" -->
<!-- #include file="setup_lib.asp" -->
<%
' -------------------------------------------------------------------
' setup1.asp
'	Support routines for building the dynamic ASP Nuke application
'	configuration forms (/admin/configure.asp)
'
' AUTH:	Ken Richards
' DATE:	10/13/03
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
Dim bTablesExist
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
	<h2>ASP Nuke Setup</h2>

	<p>
	Welcome to the ASP Nuke setup script.  This will help guide you in setting up
	your ASP Nuke installation.  This setup script follows a "wizard" type interface
	consisting of a series of configuration screens.  After you have finished the
	process of configuring your server, you may begin further installation and
	configuration of the modules.
	</p>

	<p>
	Our first order of business is to check for a database connection.  We will do
	that first and then (if necessary) will provide you with instructions on how to
	setup and configure your Microsoft SQL Server database.
	</p>

	<h3>Database Configuration Test</h3>

	<p><B>This script must be run on your web server where ASP Nuke is installed and
	must be stored within the folder where your web site was installed.</b></p>

<%
	Call setDBConnTest
	Call setDisplayStatus
	If Not setWasError Then

	Else
%>
	<h3>Database Setup</h3>

	<p>
	In order to setup ASP Nuke, you will need to make sure that your Microsoft
	SQL Server database is setup properly and your database connection string
	is set in the global.asa.
	</p>

	<h4>Creating a New Database</h3>

	<p>
	To create a new database, you will need to open enterprise manager and
	<kbd>right-click</kbd> on the database server node.  From this context menu,
	you need to select <kbd>New</kbd> and then <kbd>Database...</kbd>.  Give your
	database a name (something like <kbd>aspnuke</kbd> and click <kbd>OK</kbd>.
	</p>

	<p>
	You also have the option of using an existing database.  Be forewarned if you
	choose this route: your tables will be all lumped together and it is possible
	that the ASP Nuke database will need to use the same table name as your
	pre-existing database.  The one guarantee ASP Nuke makes is that all tables
	will be prefixed with the letters "<kbd>tbl</kbd>".
	</p>

	<h4>Creating a Database User</h4>

<%	End If

	If Not setWasError Then %>

	<h3>Database Connection Test</h3>

	<P>
	Connecting to the database and making sure that the database is empty.
	Please be careful about installing ASP Nuke onto a pre-existing
	database server.
	</p>
<%
		Call setDBConnTest2(bTablesExist)
		Call setDisplayStatus
	End If


	If Not setWasError Then
		Call setWizardButtons(True)
	End If %>
</body>
