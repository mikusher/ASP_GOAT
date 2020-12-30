<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' article_delete.asp
'	Delete an existing article from the database
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
Dim rsArt
Dim rsAuth		' list of authors to choose from
Dim rsCat		' list of categories to choose from
Dim sCatList	' list of currently selected categories
Dim sErrorMsg	' error message to display to user

sAction = LCase(steForm("action"))

If sAction = "delete" Then
	' check for required fields here
	If steNForm("confirm") = 0 Then
		sErrorMsg = steGetText("Please confirm the deletion of this article")
	Else
		' update the existing article in the database
		sStat = "DELETE FROM tblArticle WHERE ArticleID = " & steNForm("articleID")
		Call adoExecute(sStat)
	End If
End If

' retrieve the article to delete
sStat = "SELECT * FROM tblArticle WHERE ArticleID = " & steNForm("articleID")
Set rsArt = adoOpenRecordset(sStat)

' retrieve the related article properties
If Not rsArt.EOF Then
	' build the list of authors to choose from
	sStat = "SELECT	Title, FirstName, MiddleName, LastName, Surname " &_
			"FROM	tblArticleAuthor " &_
			"WHERE	AuthorID = " & rsArt.Fields("AuthorID").Value
	Set rsAuth = adoOpenRecordset(sStat)

	' build the list of categories assigned to this article
	sStat = "SELECT	ac.CategoryName " &_
			"FROM	tblArticleCategory ac " &_
			"INNER JOIN	tblArticleToCategory atc ON atc.CategoryID = ac.CategoryID " &_
			"WHERE	atc.ArticleID = " & steNForm("ArticleID") & " " &_
			"ORDER BY ac.CategoryName"
	Set rsCat = adoOpenRecordset(sStat)
	Do Until rsCat.EOF
		sCatList = sCatList & Server.HTMLEncode(rsCat.Fields("CategoryName").Value) & "<BR>" & vbCrLf
		rsCat.MoveNext
	Loop
	rsCat.Close
	Set rsCat = Nothing
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Article" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "Delete Article" %></H3>

<P>
<% steTxt "Please confirm that you would like to delete the article shown below." %>&nbsp;
<% steTxt "You must check <I>Yes</I> next to <B>Confirm Delete</B> in order to delete this record permanently." %>
</P>

<% If sAction <> "delete" Or sErrorMsg <> "" Then %>

<FORM METHOD="post" ACTION="article_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="ArticleID" VALUE="<%= steNForm("ArticleID") %>">

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Author" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD CLASS="formd"><%= rsAuth.Fields("Title").Value & " " & rsAuth.Fields("FirstName").Value & " " & rsAuth.Fields("MiddleName").Value & " " & rsAuth.Fields("LastName").Value & " " & rsAuth.Fields("Surname").Value %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Categories" %></TD><TD></TD>
	<TD CLASS="formd"><%= sCatList %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Title" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsArt, "title") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Article Lead In" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsArt, "LeadIn") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Article Body" %></TD><TD></TD>
	<TD CLASS="formd"><%= Replace(steRecordEncValue(rsArt, "ArticleBody"), vbCrLf, "<BR>") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Confirm Delete" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="confirm" VALUE="1" CLASS="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="confirm" VALUE="0" CLASS="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Article" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Article Deleted" %></H3>

<P>
<% steTxt "The article was deleted successfully." %>&nbsp;
<% steTxt "You may use the admin links above to continue administering the site." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="article_list.asp" class="adminlink"><% steTxt "Article List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->