<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/tab_lib.asp" -->
<%
'--------------------------------------------------------------------
' feed_edit.asp
'	Edit an RSS feed to the database
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
Dim rsFeed
Dim rsOrder
Dim nOrderNo

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("feedname")) = ""	Then
		sErrorMsg = "Please enter the Feed Name for this new RSS feed"
	ElseIf Trim(steForm("title")) = "" Then
		sErrorMsg = "Please enter the Title for this new RSS feed"
	ElseIf Trim(steForm("feedurl")) = "" Then
		sErrorMsg = "Please enter the Feed URL for this new RSS feed"
	ElseIf steNForm("maxitems") < 0 Or steNForm("maxitems") > 100 Then
		sErrorMsg = "Please enter a valid Max Items (between 0 and 100) for this new RSS feed"
	ElseIf steNForm("cachehours") < 1 Or steNForm("cachehours") > 100 Then
		sErrorMsg = "Please enter a valid Cache Hours (between 1 and 100) for this new RSS feed"
	Else
		' create the new RSS feed in the database
		sStat = "UPDATE tblRSSFeed SET " &_
				"	OrderNo = " & nOrderNo & "," &_
				"	FeedName = " & steQForm("FeedName") & "," &_
				"	Title= " & steQForm("Title") & "," &_
				"	FeedURL = " & steQForm("FeedURL") & "," &_
				"	MaxItems = " & steNForm("MaxItems") & "," &_
				"	ShowDescription = " & steNForm("ShowDescription") & "," &_
				"	CacheHours = " & steNForm("CacheHours") & " " &_
				"WHERE	FeedID = " & steNForm("FeedID")
		Call adoExecute(sStat)
	End If
End If

' retrieve the feed to edit
If steNForm("FeedID") > 0 Then
	Set rsFeed = adoOpenRecordset("SELECT * FROM tblRSSFeed WHERE FeedID = " & steNForm("FeedID"))
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->
<% ' tabShow "Article,Author,Category", "article_list.asp,feed_list.asp,category_list.asp", "Author" %>

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3>Edit New RSS Feed</H3>

<P>
Please enter the new properties for the new RSS feed using the form below.
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="feed_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="feedid" VALUE="<%= steEncForm("feedid") %>">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
	<TD CLASS="forml">Feed Name</TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="FeedName" VALUE="<%= steRecordEncValue(rsFeed, "FeedName") %>" SIZE="32" MAXLENGTH="100" CLASS="form"></TD>
</TR><TR>
	<TD CLASS="forml">Title</TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="Title" VALUE="<%= steRecordEncValue(rsFeed, "Title") %>" SIZE="32" MAXLENGTH="100" CLASS="form"></TD>
</TR><TR>
	<TD CLASS="forml">Feed URL</TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="FeedURL" VALUE="<%= steRecordEncValue(rsFeed, "FeedURL") %>" SIZE="32" MAXLENGTH="255" CLASS="form"></TD>
</TR><TR>
	<TD CLASS="forml">Max Items Shown</TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="MaxItems" VALUE="<%= steRecordEncValue(rsFeed, "MaxItems") %>" SIZE="12" MAXLENGTH="10" CLASS="form"></TD>
</TR><TR>
	<TD CLASS="forml">Show Description</TD><TD></TD>
	<TD>
		<INPUT TYPE="radio" NAME="ShowDescription" VALUE="1"<% If steRecordBoolValue(rsFeed, "ShowDescription") Then Response.Write " CHECKED" %> CLASS="formradio"> Yes
		<INPUT TYPE="radio" NAME="ShowDescription" VALUE="0"<% If Not steRecordBoolValue(rsFeed, "ShowDescription") Then Response.Write " CHECKED" %> CLASS="formradio"> No
	</TD>
</TR><TR>
	<TD CLASS="forml">Cache Hours</TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="CacheHours" VALUE="<%= steRecordEncValue(rsFeed, "CacheHours") %>" SIZE="12" MAXLENGTH="10" CLASS="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" Update RSS Feed " CLASS="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3>RSS Feed Updated</H3>

<P>
The RSS feed has been updated in the database.  Please proceed administering
the site by using the menu shown at the top of the screen.
</P>

<% End If %>

<p align="center">
	<a href="feed_list.asp" class="adminlink">RSS Feed List</a>
</p>

<!-- #include file="../../../../footer.asp" -->
