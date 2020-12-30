<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' member_list.asp
'	Displays a list of the registered members for the site
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

Dim sStat
Dim rsList
Dim rsCount
Dim rsState
Dim rsCountry
Dim sAction
Dim sWhere		' where clause for search
Dim sCriteria	' criteria we are searching on
Dim I

sAction = LCase(steForm("action"))

If sAction = "search" Then
	If Trim(steForm("FirstName")) <> "" Then
		sWhere = sWhere & " AND tblMember.FirstName LIKE '%" & Replace(steForm("FirstName"), "'", "''") & "%'"
		sCriteria = sCriteria & "<I>First Name</I> like <B>" & steForm("FirstName") & "</B><BR>"
	End If
	If Trim(steForm("LastName")) <> "" Then
		sWhere = sWhere & " AND tblMember.LastName LIKE '%" & Replace(steForm("LastName"), "'", "''") & "%'"
		sCriteria = sCriteria & "<I>Last Name</I> like <B>" & steForm("LastName") & "</B><BR>"
	End If
	If Trim(steForm("Username")) <> "" Then
		sWhere = sWhere & " AND tblMember.Username LIKE '%" & Replace(steForm("Username"), "'", "''") & "%'"
		sCriteria = sCriteria & "<I>Username</I> like <B>" & steForm("Username") & "</B><BR>"
	End If
	If Trim(steNForm("StateCode")) <> 0 Then
		sWhere = sWhere & " AND tblMember.StateCode = '" & Replace(steForm("StateCode"), "'", "''") & "'"
		sCriteria = sCriteria & "<I>State</I> is <B>" & steForm("StateCode") & "</B><BR>"
	End If
	If Trim(steNForm("CountryID")) <> 0 Then
		sWhere = sWhere & " AND tblMember.CountryID = " & steForm("CountryID")
		' retrieve the name of this country
		sStat = "SELECT	CountryName " &_
				"FROM	tblCountry " &_
				"WHERE	CountryID = " & steForm("CountryID")
		Set rsCountry = adoOpenRecordset(sStat)
		If Not rsCountry.EOF Then sCriteria = sCriteria & "<I>Country</I> is <B>" & rsCountry.Fields("CountryName").Value & "</B><BR>"
	End If
End If

' retrieve the list of members to display here
sStat = "SELECT	" & adoTop(30) & " tblMember.MemberID, tblMember.FirstName, tblMember.LastName, tblMember.Username, " &_
		"		tblMember.Modified, tblState.StateName, tblCountry.CountryName " &_
		"FROM	tblMember " &_
		"LEFT JOIN	tblState ON tblMember.StateCode = tblState.StateCode " &_
		"LEFT JOIN	tblCountry ON tblMember.CountryID = tblCountry.CountryID " &_
		"WHERE	tblMember.Archive = 0 " & sWhere & adoTop2(30)
Set rsList = adoOpenRecordset(sStat)

' count the total number of records matched here
sStat = "SELECT COUNT(*) AS MemberCount " &_
		"FROM	tblMember " &_
		"WHERE	Archive = 0 " & sWhere
Set rsCount = adoOpenRecordset(sStat)

' build the selection list for the states
query = "SELECT StateCode, StateName " &_
		"FROM	tblState " &_
		"WHERE	Active <> 0 " &_
		"AND	Archive = 0 " &_
		"ORDER BY StateName"
Set rsState = adoOpenRecordset(query)

' build the selection ist for the countries
query = "SELECT CountryID, CountryName " &_
		"FROM	tblCountry " &_
		"WHERE	Active <> 0 " &_
		"AND	Archive = 0 " &_
		"ORDER BY CountryName"
Set rsCountry = adoOpenRecordset(query)
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Members" %>
<!-- #include file="pagetabs_inc.asp" -->

<!-- form for member search -->

<FORM METHOD="post" ACTION="member_list.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="search">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml"><% steTxt "First Name" %><BR>
	<INPUT TYPE="text" NAME="FirstName" VALUE="<%= steEncForm("FirstName") %>" SIZE=16 MAXLENGTH=32 class="form">
	</TD>
	<TD CLASS="forml"><% steTxt "Last Name" %><BR>
	<INPUT TYPE="text" NAME="LastName" VALUE="<%= steEncForm("LastName") %>" SIZE=16 MAXLENGTH=32 class="form">
	</TD>
	<TD CLASS="forml"><% steTxt "Username" %><BR>
	<INPUT TYPE="text" NAME="Username" VALUE="<%= steEncForm("Username") %>" SIZE=16 MAXLENGTH=32 class="form">
	</TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "State" %><BR>
	<SELECT NAME="StateCode" class="form">
	<OPTION VALUE=""> -- Choose --
	<% Do Until rsState.EOF %>
	<OPTION VALUE="<%= rsState.Fields("StateCode").Value %>"<% If steForm("StateCode") = rsState.Fields("StateCode").Value Then Response.Write(" SELECTED") %>> <%= rsState.Fields("StateName").Value %>
	<%	rsState.MoveNext
	   Loop %>
	</SELECT>
	</TD>
	<TD CLASS="forml"><% steTxt "Country" %><BR>
	<SELECT NAME="CountryID" class="form">
	<OPTION VALUE="0"> -- Choose --
	<% Do Until rsCountry.EOF %>
	<OPTION VALUE="<%= rsCountry.Fields("CountryID").Value %>"<% If steNForm("CountryID") = rsCountry.Fields("CountryID").Value Then Response.Write(" SELECTED") %>> <%= rsCountry.Fields("CountryName").Value %>
	<%	rsCountry.MoveNext
	   Loop %>
	</SELECT>
	</TD><TD VALIGN="bottom">
		<INPUT TYPE="submit" NAME="_submit" ACTION=" <% steTxt "Search" %> " class="form">
	</TD>
</TR>
</TABLE>
</FORM>

<!-- display all of the members here (max of 30) -->

<H3><% steTxt "Member List" %></H3>

<BLOCKQUOTE>
<%= sCriteria %>
</BLOCKQUOTE>

<P ALIGN="center">
<I>Total of <%= rsCount.Fields("MemberCount").Value & " " & steGetText("Members Found") %></I>
</P>
<% If Not rsList.EOF Then %>

<P>
<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 class="list">
<TR>
	<TD class="listhead"><% steTxt "Full Name" %></TD>
	<TD class="listhead"><% steTxt "Username" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "State" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Country" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Modified" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Action" %></TD>
</TR>
<% I = 0
   Do Until rsList.EOF %>
<TR CLASS="list<%= I mod 2 %>">
	<TD><%= rsList.Fields("FirstName").Value & " " & rsList.Fields("LastName").Value %></TD>
	<TD><%= rsList.Fields("Username").Value %></TD>
	<TD ALIGN="right"><%= rsList.Fields("StateName").Value %></TD>
	<TD ALIGN="right"><%= rsList.Fields("CountryName").Value %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsList.Fields("Modified").Value, vbShortDate) %></TD>
	<TD ALIGN="right">
		<A HREF="member_edit.asp?memberid=<%= rsList.Fields("MemberID").Value %>" class="actionlink"><% steTxt "edit" %></A> .
		<A HREF="member_delete.asp?memberid=<%= rsList.Fields("memberid").Value %>" class="actionlink"><% steTxt "delete" %></A>
	</TD>
</TR>
<%	rsList.MoveNext
	I = I + 1
   Loop %>
</TABLE>
</P>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No members found to display here" %></B></P>

<% End If %>

<P ALIGN="center">
	<A HREF="member_add.asp" CLASS="adminlink"><% steTxt "Add New Member" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
