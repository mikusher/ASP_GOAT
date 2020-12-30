<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' topic_delete.asp
'	Delete an existing forum topic from the database
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

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If steNForm("confirm") <> 1 Then
		sErrorMsg = steGetText("Please confirm the deletion of this forum topic")
	Else
		' delete the topic from the database
		sStat = "DELETE FROM tblMessageTopic " &_
				"WHERE	TopicID = " & nTopicID
		Call adoExecute(sStat)
	End If
End If

If nTopicID > 0 Then
	sStat = "SELECT * FROM tblMessageTopic " &_
			"WHERE TopicID = " & nTopicID
	Set rsTopic = adoOpenRecordset(sStat)
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Topic" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Forum Topic" %></H3>

<P>
<% steTxt "Please confirm the deletion of the forum topic by clicking the yes button next to <I>Confirm</I>." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="topic_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="topicid" VALUE="<%= nTopicID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml"><% steTxt "Title" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsTopic, "Title") %></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "Short Description" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsTopic, "ShortComments") %></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "Confirm Delete" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="confirm" VALUE="1" CLASS="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="confirm" VALUE="0" CLASS="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Forum Topic" %> " CLASS="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Forum Topic Deleted" %></H3>

<P>
<% steTxt "The forum topic was successfully deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<!-- #include file="../../../../footer.asp" -->
