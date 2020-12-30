<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' topic_edit.asp
'	Edit an existing forum topic to the database
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

Dim sErrorMsg
Dim sStat
Dim rsTopic
Dim nTopicID

nTopicID = steNForm("topicid")

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("Title")) = ""	Then
		sErrorMsg = steGetText("Please enter the Title for this forum topic")
	ElseIf Trim(steForm("shortcomments")) = "" Then
		sErrorMsg = steGetText("Please enter the Short Description for this new forum topic")
	Else
		' update the forum topic in the database
		sStat = "UPDATE tblMessageTopic " &_
				"SET	Title = " & steQForm("Title") & "," &_
				"		ShortComments = " & steQForm("ShortComments") & "," &_
				"		Modified = " & adoGetDate & " " &_
				"WHERE	TopicID = " & nTopicID
		Call adoExecute(sStat)
	End If
End If

' retrieve the forum topic to edit
sStat = "SELECT	* FROM tblMessageTopic WHERE TopicID = " & nTopicID
Set rsTopic = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Topic" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Forum Topic" %></H3>

<P>
<% steTxt "Please enter the properties for the forum topic using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="topic_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="topicid" VALUE="<%= nTopicID %>">
<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD class="forml" nowrap><% steTxt "Title" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="Title" VALUE="<%= steRecordEncValue(rsTopic, "Title") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Short Description" %></TD><TD></TD>
	<TD><TEXTAREA NAME="ShortComments" COLS="48" ROWS="6" CLASS="form"><%= steRecordEncValue(rsTopic, "ShortComments") %></TEXTAREA></TD>
</TR><TR>
	<TD COLSPAN=3><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update Forum Topic" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Forum Topic Updated" %></H3>

<P>
<% steTxt "The forum topic was successfully updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="topic_list.asp" CLASS="adminlink"><% steTxt "Topic List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->
