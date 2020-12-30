<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/tab_lib.asp" -->
<%
'--------------------------------------------------------------------
' feed_list.asp
'	Displays a list of the current RSS feeds
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
Dim rsFeed
Dim I

sAction = LCase(steForm("action"))

Select Case sAction
	Case "moveup"
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblRSSFeed " &_
					"SET	OrderNo = OrderNo + 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & steNForm("OrderNo") - 1
			Call adoExecute(sStat)

			sStat = "UPDATE	tblRSSFeed " &_
					"SET	OrderNo = " & (steNForm("OrderNo") - 1) & ", Modified = " & adoGetDate & " " &_
					"WHERE	FeedID = " & steNForm("FeedID")
			Call adoExecute(sStat)
			modRefresh True
	Case "movedown"
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblRSSFeed " &_
					"SET	OrderNo = OrderNo - 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & steNForm("OrderNo") + 1
			Call adoExecute(sStat)

			sStat = "UPDATE	tblRSSFeed " &_
					"SET	OrderNo = " & (steNForm("OrderNo") + 1) & ", Modified = " & adoGetDate & " " &_
					"WHERE	FeedID = " & steNForm("FeedID")
			Call adoExecute(sStat)
			modRefresh True
End Select

sStat = "SELECT	FeedID, FeedName, OrderNo, Modified " &_
		"FROM	tblRSSFeed " &_
		"ORDER BY OrderNo"
Set rsFeed = adoOpenRecordset(sStat)

%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->
<% ' tabShow "Article,Author,Category", "article_list.asp,feed_list.asp,category_list.asp", "Author" %>

<H3>RSS Feed List</H3>

<P>
Shown below are all of the RSS feeds defined in the database.
</P>

<% If Not rsFeed.EOF Then %>

<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 class="list">
<TR>
	<TD CLASS="listhead">Order</TD>
	<TD CLASS="listhead">Feed Name</TD>
	<TD CLASS="listhead" ALIGN="right">Modified</TD>
	<TD CLASS="listhead" ALIGN="right">Action</TD>
</TR>
<% I = 0
Do Until rsFeed.EOF %>
<TR CLASS="list<%= I mod 2 %>">
	<TD><%= rsFeed.Fields("OrderNo").Value %></TD>
	<TD><%= Server.HTMLEncode(rsFeed.Fields("FeedName").Value) %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsFeed.Fields("Modified").Value, vbShortDate) %></TD>
	<TD>
		<A HREF="feed_list.asp?FeedID=<%= nFeedID %>&orderno=<%= rsFeed.Fields("OrderNo").Value %>&action=moveup" class="actionlink">up</A> .
		<A HREF="feed_list.asp?FeedID=<%= nFeedID %>&orderno=<%= rsFeed.Fields("OrderNo").Value %>&action=movedown" class="actionlink">down</A> .
		<A HREF="feed_edit.asp?feedid=<%= rsFeed.Fields("FeedID").Value %>" class="actionlink">edit</A> . <A HREF="feed_delete.asp?feedid=<%= rsFeed.Fields("FeedID").Value %>" class="actionlink">delete</A>
	</TD>
</TR>
<%	rsFeed.MoveNext
	I = I + 1
   Loop %>
</TABLE>

<% Else %>

<P><B CLASS="error">Sorry, No RSS feeds exist in the database</B></P>

<% End If %>

<P ALIGN="center">
	<A HREF="feed_add.asp" class="adminlink">Add New RSS Feed</A>
</P>

<!-- #include file="../../../../footer.asp" -->