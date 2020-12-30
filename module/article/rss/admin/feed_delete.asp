<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/tab_lib.asp" -->
<%
'--------------------------------------------------------------------
' feed_delete.asp
'	Delete an existing RSS feed to the database
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
Dim nFeedID

nFeedID = steNForm("feedid")

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If steNForm("Confirm") <> 1	Then
		sErrorMsg = "Please confirm the deletion of this RSS feed"
	Else
		' create the new RSS feed in the database
		sStat = "DELETE FROM tblRSSFeed WHERE FeedID = " & nFeedID
		Call adoExecute(sStat)
	End If
End If

sStat = "SELECT * FROM tblRSSFeed WHERE FeedID = " & nFeedID
Set rsFeed = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->
<% ' tabShow "Article,Author,Category", "article_list.asp,author_list.asp,feed_list.asp", "Category" %>

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3>Delete RSS Feed</H3>

<P>
Please confirm that you would like to delete this RSS feed by
clicking <I>Yes</I> next to <B>Confirm</B> below.  Once the feed
has been deleted, it can not be recovered.
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="feed_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="feedid" VALUE="<%= nFeedID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml">Feed Name</TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsFeed, "FeedName") %></TD>
</TR><TR>
	<TD CLASS="forml">Title</TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsFeed, "Title") %></TD>
</TR><TR>
	<TD CLASS="forml">Feed URL</TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsFeed, "FeedURL") %></TD>
</TR><TR>
	<TD CLASS="forml">Max Items Shown</TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsFeed, "MaxItems") %></TD>
</TR><TR>
	<TD CLASS="forml">Show Description</TD><TD></TD>
	<TD CLASS="formd">
		<% If steRecordBoolValue(rsFeed, "ShowDescription") Then Response.Write "Yes" Else Response.Write "No" %>
	</TD>
</TR><TR>
	<TD CLASS="forml">Cache Hours</TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsFeed, "CacheHours") %></TD>
</TR><TR>
	<TD class="forml">Confirm Delete</TD><TD></TD>
	<TD CLASS="formd"><INPUT TYPE="radio" NAME="confirm" VALUE="1" class="formradio"> Yes
		<INPUT TYPE="radio" NAME="confirm" VALUE="0" class="formradio"> No
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" Delete RSS Feed " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3>RSS Feed Deleted</H3>

<P>
The RSS feed was permanently deleted from the database.
Please proceed administering
the site by using the menu shown at the top of the screen.
</P>

<% End If %>

<p align="center">
	<a href="feed_list.asp" class="adminlink">RSS Feed List</a>
</p>

<!-- #include file="../../../../footer.asp" -->
