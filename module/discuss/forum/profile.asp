<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' profile.asp
'	Creates a  profile for a member who is using our message forums.
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

Dim query
Dim sUsername		' username for profile
Dim nTopicID		' topic the user was browsing
Dim nMemberID

sUsername = Request.QueryString("username")
nTopicID = Request.QueryString("topicid")

If IsNumeric(nTopicID) And CStr(nTopicID) <> "" Then nTopicID = CInt(nTopicID) Else nTopicID = 0

query = "SELECT	mbr.MemberID, mbr.Firstname, mbr.Lastname, mp.Email, " &_
		"		mp.Location, mp.HomePage, mp.ForumIcon, mp.Biography " &_
		"FROM	tblMember mbr " &_
		"LEFT JOIN	tblMessageProfile mp ON mp.MemberID = mbr.MemberID " &_
		"WHERE	mbr.Username = '" & Replace(sUsername, "'", "''") & "' " &_
		"AND	mbr.Active <> 0 " &_
		"AND	mbr.Archive = 0"
Set rsMember = adoOpenRecordset(query)

%>
<!-- #include file="../../../header.asp" -->

<% If Not rsMember.EOF Then
	nMemberID = rsMember.Fields("MemberID").Value %>

<table border=0 cellpadding=2 cellspacing=0 width="100%">
<tr>
	<td><H3><% steTxt "Profile for" %> <%= sUsername %></H3></td>
	<% If Trim(rsMember.Fields("ForumIcon").Value & "") <> "" Then %>
	<td align="right">
		<img src="<%= rsMember.Fields("ForumIcon").Value %>" alt="<% steTxt "Avatar Icon for" %> <%= Server.HTMLEncode(sUsername) %>">
	</td>
	<% End If %>
</tr>
</table>

<table border=0 cellpadding=0 cellspacing=10>
<tr>
	<td valign="top">
	<table border=0 cellpadding="10" cellspacing=0 align="left">
	<tr>
		<td>
		<TABLE BORDER=0 CELLPADDING=6 CELLSPACING=0 CLASS="list">
		<TR CLASS="list0">
			<TD class="forml"><% steTxt "Full Name" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
			<TD class="formd"><%= rsMember.Fields("FirstName").Value & " " & rsMember.Fields("LastName").Value %></TD>
		</TR><TR CLASS="list1">
			<TD class="forml"><% steTxt "Location" %></TD><TD></TD>
			<TD class="formd"><% If Trim(rsMember.Fields("Location").Value) <> "" Then %><%= rsMember.Fields("Location").Value %><% Else %><i>n/a</i><% End If %></TD>
		</TR><TR CLASS="list0">
			<TD class="forml"><% steTxt "E-mail" %></TD><TD></TD>
			<TD class="formd"><% If Trim(rsMember.Fields("Email").Value & "") <> "" Then %><a href="mailto:<%= rsMember.Fields("Email").Value %>"><%= rsMember.Fields("Email").Value  %></A><% Else %><i>n/a</i><% End If %></TD>
		</TR><TR CLASS="list1">
			<TD class="forml"><% steTxt "Home Page" %></TD><TD></TD>
			<TD class="formd"><% If Trim(rsMember.Fields("HomePage").Value & "") <> "" Then %><a href="<%= rsMember.Fields("HomePage").Value %>"><%= rsMember.Fields("HomePage").Value %></A><% Else %><i>n/a</i><% End If %></TD>
		</TR>
		</TABLE>
		</td>
	</tr>
	</table>

	<h4><% steTxt "Forum Member Biography" %></h4>

	<% If Trim(rsMember.Fields("Biography").Value & "") <> "" Then %>
	<p>
	<%= Replace(rsMember.Fields("Biography").Value, vbCrLf, "<BR>") %>
	</p>
	<% Else %>
	<p>
	<% steTxt "No biography available for this member." %>
	</p>
	<% End If %>
	</td>
</tr>
</table>

<% Else %>

<h3><% steTxt "Profile for" %>&nbsp;<%= sUsername %></h3>

<P><B CLASS="error"><% steTxt "Sorry, no profile available for member" %> <I><%= sUsername %></I></B></P>

<% End If %>

<P ALIGN="center">
	<A HREF="index.asp" CLASS="footerlink"><% steTxt "Forum Overview" %></A>
<% If nTopicID > 0 Then %>
	&nbsp; <A HREF="topic.asp?topicid=<%= nTopicID %>" CLASS="footerlink"><% steTxt "Topic Overview" %></A>
<% End If %>
<% If steNForm("ThreadID") > 0 Then %>
	&nbsp; <a href="thread.asp?topicid=<%= nTopicID %>&threadid=<%= steNForm("ThreadID") %>" class="footerlink"><% steTxt "Back to Thread" %></a> &nbsp;
<% End If %>
	&nbsp; <A HREF="profile_edit.asp?topicid=<%= nTopicID %>" CLASS="footerlink"><% steTxt "Edit Your Profile" %></A>
</P>

<!-- #include file="../../../footer.asp" -->