<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/tab_lib.asp" -->
<%
'--------------------------------------------------------------------
' feed_add.asp
'	Add a new RSS feed to the database
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

If steForm("action") = "add" Then
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
		' get the new order no
		sStat = "SELECT Coalesce(Max(OrderNo) + 1, 1) As OrderNo FROM tblRSSFeed"
		Set rsOrder = adoOpenRecordset(sStat)
		If Not rsOrder.EOF Then nOrderNo = rsOrder.Fields("OrderNo").Value Else nOrderNo = 1
		rsOrder.Close
		Set rsOrder = Nothing

		' create the new RSS feed in the database
		sStat = "INSERT INTO tblRSSFeed (" &_
				"	OrderNo, FeedName, Title, FeedURL, MaxItems, ShowDescription, CacheHours, Created" &_
				") VALUES (" &_
				nOrderNo & "," &_
				steQForm("FeedName") & "," &_
				steQForm("Title") & "," &_
				steQForm("FeedURL") & "," & steNForm("MaxItems") & "," &_
				steNForm("ShowDescription") & "," & steNForm("CacheHours") & "," &_
				adoGetDate &_
				")"
		Call adoExecute(sStat)
	End If
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->
<% ' tabShow "Article,Author,Category", "article_list.asp,feed_list.asp,category_list.asp", "Author" %>

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3>Add New RSS Feed</H3>

<P>
Please enter the new properties for the new RSS feed using the form below.
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="feed_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
	<TD CLASS="forml">Feed Name</TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="FeedName" VALUE="<%= steEncForm("FeedName") %>" SIZE="32" MAXLENGTH="100" CLASS="form"></TD>
</TR><TR>
	<TD CLASS="forml">Title</TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="Title" VALUE="<%= steEncForm("Title") %>" SIZE="32" MAXLENGTH="100" CLASS="form"></TD>
</TR><TR>
	<TD CLASS="forml">Feed URL</TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="FeedURL" VALUE="<%= steEncForm("FeedURL") %>" SIZE="32" MAXLENGTH="255" CLASS="form"></TD>
</TR><TR>
	<TD CLASS="forml">Max Items Shown</TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="MaxItems" VALUE="<%= steEncForm("MaxItems") %>" SIZE="12" MAXLENGTH="10" CLASS="form"></TD>
</TR><TR>
	<TD CLASS="forml">Show Description</TD><TD></TD>
	<TD>
		<INPUT TYPE="radio" NAME="ShowDescription" VALUE="1"<% If steNForm("ShowDescription") = 1 Then Response.Write " CHECKED" %> CLASS="formradio"> Yes
		<INPUT TYPE="radio" NAME="ShowDescription" VALUE="0"<% If steNForm("ShowDescription") = 0 Then Response.Write " CHECKED" %> CLASS="formradio"> No
	</TD>
</TR><TR>
	<TD CLASS="forml">Cache Hours</TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="CacheHours" VALUE="<%= steEncForm("CacheHours") %>" SIZE="12" MAXLENGTH="10" CLASS="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" Add RSS Feed " CLASS="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3>New RSS Feed Added</H3>

<P>
The new RSS feed has been added to the database.  Please proceed administering
the site by using the menu shown at the top of the screen.
</P>

<% End If %>

<p align="center">
	<a href="feed_list.asp" class="adminlink">RSS Feed List</a>
</p>

<!-- #include file="../../../../footer.asp" -->
