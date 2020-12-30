<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' article_add.asp
'	Displays a list of the current articles for the site
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
Dim nArticleID	' articleID for new article just added
Dim rsAuth		' list of authors to choose from
Dim rsCat		' list of categories to choose from
Dim sCatList	' list of currently selected categories
Dim sErrorMsg	' error message to display to user

sCatList = Replace(steForm("CategoryID"), " ", "")

If UCase(steForm("action")) = "ADD" Then
	' check for required fields here
	If steNForm("authorid") = 0 Then
		sErrorMsg = steGetText("Please select the author for this article")
	ElseIf Trim(sCatList) = "" Then
		sErrorMsg = steGetText("Please check at least one category for this article")
	ElseIf Trim(steForm("Title")) = "" Then
		sErrorMsg = steGetText("Please enter the title for this article")
	ElseIf Trim(steForm("LeadIn")) = "" Then
		sErrorMsg = steGetText("Plese enter the Lead-In for this article")
	ElseIf Trim(steForm("ArticleBody")) = "" Then
		sErrorMsg = steGetText("Please enter the Body for this article")
	Else
		' insert the new article into the database
		sStat = "INSERT INTO tblArticle (" &_
				"	AuthorID, Title, LeadIn, ArticleBody, Created" &_
				") VALUES (" &_
				steNForm("AuthorID") & "," &_
				steQForm("Title") & "," & steQForm("LeadIn") & "," &_
				steQForm("ArticleBody") & "," & adoGetDate &_
				")"
		Call adoExecute(sStat)

		nArticleID = 0
		Set rsArtID = adoOpenRecordset("SELECT Max(ArticleID) As ArticleID FROM tblArticle WHERE AuthorID = " & steNForm("AuthorID"))
		If Not rsArtID.EOF Then
			nArticleID = rsArtID.Fields("ArticleID").Value
		End If
		If nArticleID > 0 Then
			' assign the categories for this article
			Call adoExecute("DELETE FROM tblArticleToCategory WHERE ArticleID = " & steNForm("articleID"))
			Dim aCatID
			aCatID = Split(sCatList, ",")
			For I = 0 To UBound(aCatID)
				sStat = "INSERT INTO tblArticleToCategory (" &_
						"	ArticleID, CategoryID" &_
						") VALUES (" &_
						nArticleID & "," & Trim(aCatID(I)) &_
						")"
				Call adoExecute(sStat)
			Next
		Else
			sErrorMsg = steGetText("Article was added, but failed to retrieve the Article ID to assign the categories")
		End If
	End If
End If

' build the list of authors to choose from
sStat = "SELECT	AuthorID, Title, FirstName, MiddleName, LastName, Surname " &_
		"FROM	tblArticleAuthor " &_
		"WHERE	Active <> 0 " &_
		"AND	Archive = 0 " &_
		"ORDER BY LastName, FirstName, MiddleName"
Set rsAuth = adoOpenRecordset(sStat)

' build the list of categories to choose from
sStat = "SELECT	CategoryID, CategoryName " &_
		"FROM	tblArticleCategory " &_
		"WHERE	Active <> 0 " &_
		"AND	Archive = 0 " &_
		"ORDER BY CategoryName"
Set rsCat = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Article" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If UCase(steForm("action")) <> "ADD" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New Article" %></H3>

<P>
<% steTxt "Please enter the information for the new article in the form below." %>
</P>

<FORM METHOD="post" ACTION="article_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Author" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><SELECT NAME="AuthorID" class="form">
	<OPTION VALUE="0"> -- Choose --
	<% Do Until rsAuth.EOF %>
	<OPTION VALUE="<%= rsAuth.Fields("AuthorID").Value %>"<% If CStr(steForm("AuthorID")) = CStr(rsAuth.Fields("AuthorID").Value) Then Response.Write " SELECTED" %>> <%= rsAuth.Fields("Title").Value & " " & rsAuth.Fields("FirstName").Value & " " & rsAuth.Fields("MiddleName").Value & " " & rsAuth.Fields("LastName").Value & " " & rsAuth.Fields("Surname").Value %>
	<%	rsAuth.MoveNext
	   Loop %>
	</SELECT>
	</TD>
</TR><TR>
	<TD VALIGN="top" class="forml"><% steTxt "Categories" %></TD><TD></TD>
	<TD>
	<table border="0" cellpadding="4" width="100%">
<%	' build the list of categories to choose from
	sStat = "SELECT	CategoryID, CategoryName " &_
			"FROM	tblArticleCategory " &_
			"WHERE	Active <> 0 " &_
			"AND	Archive = 0 " &_
			"ORDER BY CategoryName"
	Set rsCat = adoOpenRecordset(sStat)
	Dim nCol
	nCol = 0
	Do Until rsCat.EOF
		If nCol Mod 4 = 0 Then Response.Write "<tr>" & vbCrLf %>
		<td><input type="checkbox" name="CategoryID" value="<%= rsCat.Fields("CategoryID").Value %>"<% If InStr(1, ","&sCatList&",", ","&rsCat.Fields("CategoryID").Value&",") > 0 Then Response.Write " CHECKED" %>> <%= rsCat.Fields("CategoryName").Value %></td>
	<%	If nCol Mod 4 = 3 Then Response.Write "</tr>" & vbCrLf
		nCol = nCol + 1
		rsCat.MoveNext
	Loop
	rsCat.Close
	Set rsCat = Nothing
	If nCol Mod 4 <> 0 Then Response.Write "</tr>" & vbCrLf %>
	</table>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Title" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Title" VALUE="<%= steEncForm("title") %>" SIZE="32" MAXLENGTH="100" class="form" style="width:100%"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Article Lead In" %></TD><TD></TD>
	<TD><TEXTAREA NAME="LeadIn" COLS="42" ROWS="10" WRAP="Virtual" class="form" style="width:100%"><%= steEncForm("LeadIn") %></TEXTAREA></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Article Body" %></TD><TD></TD>
	<TD><TEXTAREA NAME="ArticleBody" COLS="42" ROWS="10" WRAP="Virtual" class="form" style="width:100%"><%= steEncForm("ArticleBody") %></TEXTAREA></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Article" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Article Added" %></H3>

<P>
<% steTxt "The new article was successfully added." %>&nbsp;
<% steTxt "You may use the admin links above to continue administering the site." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="article_list.asp" class="adminlink"><% steTxt "Article List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->