<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' table_drop.asp
'	Drop a mysql table from the database
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

Dim sAction
Dim aType		' all of the mysql data types
Dim aSize		' fixed size for each data type
Dim aUseSize	' use the size when creating/altering the table
Dim aUsePrec	' use the precision when creating/altering the table
Dim aAutoInc	' allow the auto_increment modifier on the column def
Dim rsTab		' table to create
Dim sErrorMsg
Dim sql
Dim I, J

sAction = steForm("action")

If sAction = "drop" Then
	If Trim(steForm("table")) = "" Then
		sErrorMsg = "The table to drop was not specified"
	Else
		' make sure the table doesn't already exist
		If Not locTableExists(steForm("table")) Then
			sErrorMsg = "The table that was selected (""" & steForm("table") & """) doesn't exist!"
		End If
	End If

	If sErrorMsg = "" Then
		If steForm("confirm") <> "drop table" Then
			sErrorMsg = "You must confirm that you want to drop this table by typing ""drop table"""
		Else
			' create the SQL that creates the database table
			sql = "drop table if exists " & steForm("table")
			' If sErrorMsg = "" Then Response.Write Replace(Server.HTMLEncode(sql), vbCrLf, "<br>") : Response.End
			Call adoExecute(sql)
		End If
	End If
End If
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Tables" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If sAction <> "drop" Or sErrorMsg <> "" Then %>


<h4>Drop MySQL Table</h4>

<p>
Please confirm that you would like to drop the table shown below by typing
the words "drop table" in the form below.  Once a table has been dropped,
it can not be recovered unless you have restore from a backup.
</p>

<% If sErrorMsg <> "" Then %>
<p><b class="error"><%= sErrorMsg %></b></p>
<% End If %>

<form method="post" action="table_drop.asp">
<input type="hidden" name="table" value="<%= steEncForm("table") %>">
<input type="hidden" name="action" value="drop">

<table border=0 cellpadding=2 cellspacing=0 class="list" align="center">
<tr>
	<td align="right" class="forml">Type "drop table" here to confirm:</td><td>&nbsp;&nbsp;</td>
	<td><input type="text" class="form" name="confirm" value="<%= steEncForm("confirm") %>" style="width:100px"></td>
	<td><input type="submit" class="form" name="_confirm" value="DROP"></td>
</tr>
</table><br>

</form>

<table border=0 cellpadding=2 cellspacing=0 class="list" align="center">
<tr>
	<td align="right" class="forml">Table Name</td><td>&nbsp;&nbsp;</td>
	<td class="formd"><%= steEncForm("table") %></td>
</tr>
</table><br>

<% ' show the column definitions here
Dim rsCol
Set rsCol = adoOpenRecordset("describe " & steForm("table"))
If Not rsCol.EOF Then %>
<table border=0 cellpadding=2 cellspacing=0 class="list">
<tr>
	<td class="listhead">Field Name</td>
	<td class="listhead">Data Type</td>
	<td class="listhead">Null</td>
	<td class="listhead">Key</td>
	<td class="listhead">Default</td>
	<td class="listhead">Extra</td>
</tr>
<% I = 1
   Do Until rsCol.EOF %>
<tr class="list<%= I Mod 2 %>">
	<td><%= rsCol.Fields("Field").Value %></td>
	<td><%= rsCol.Fields("Type").Value %></td>
	<td><%= rsCol.Fields("Null").Value %></td>
	<td><%= rsCol.Fields("Key").Value %></td>
	<td><%= rsCol.Fields("Default").Value %></td>
	<td><%= rsCol.Fields("Extra").Value %></td>
</tr>
<%	rsCol.MoveNext
	I = I + 1
   Loop
	rsCol.Close
	Set rsCol = Nothing %>
</table>
<% Else %>

<p><b class="error">Unable to load the column definitions for table "<%= steForm("table") %>"</b></p>

<% End If %>

<% Else %>

<h4>Table Dropped Successfully</h4>

<p>
The table was successfully dropped from the database.  Please click on
the link below to continue administering the MySQL database.
</p>

<% End If %>

<p align="center">
	<a href="table_list.asp" class="adminlink"><% steTxt "Table List" %></A>
</p>

<!-- #include file="../../../footer.asp" -->
<%
Function locTableExists(sTable)
	Dim rsTab
	Set rsTab = adoOpenRecordset("show tables;")
	Do Until rsTab.EOF
		If StrComp(rsTab.Fields(0).Value, sTable, vbTextCompare) = 0 Then
			locTableExists = True
			rsTab.Close
			Exit Function
		End If
		rsTab.MoveNext
	Loop
	rsTab.Close
	locTableExists = False
End Function
%>
