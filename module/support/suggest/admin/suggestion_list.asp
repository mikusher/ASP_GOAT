<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/class/adminlist.asp" -->
<%
'--------------------------------------------------------------------
' suggestion_list.asp
'	Displays a list of the suggestions for the site
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
Dim I

%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Suggestions" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "Suggestion List" %></H3>

<P>
<% steTxt "Shown below are all of the current suggestions defined in the database." %>
</P>

<%
sStat = "SELECT	tblSuggestion.SuggestionID, tblSuggestion.FromName, tblSuggestion.FromEmail, " &_
		"		tblSuggestion.Subject, tblSuggestion.Created, tblSuggestion.Modified " &_
		"FROM	tblSuggestion " &_
		"LEFT JOIN	tblMember ON tblSuggestion.MemberID = tblMember.MemberID " &_
		"ORDER BY tblSuggestion.Created DESC"
Set oList = New clsAdminList
oList.Query = sStat
oList.AddColumn "Subject", steGetText("Subject"), ""
oList.AddColumn "<a href=""mailto:##FromEmail##"">##FromName##</A>", steGetText("From Name"), ""
oList.AddColumn "Created", steGetText("Submitted"), ""
oList.AddColumn "Modified", steGetText("Modified"), ""
oList.ActionLink = "<A HREF=""suggestion_edit.asp?SuggestionID=##SuggestionID##"" class=""actionlink"">edit</A> . <A HREF=""suggestion_delete.asp?SuggestionID=##SuggestionID##"" class=""actionlink"">delete</A>"
' oList.QueryString = "topicid=" & nTopicID
Call oList.Display
%>

<!-- #include file="../../../../footer.asp" -->