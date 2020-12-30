<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' article_edit.asp
'	Update an existing article in the database
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
Dim I

sAction = LCase(steForm("action"))

If sAction = "edit" Then
	' check for required fields here
	If steNForm("authorid") = 0 Then
		sErrorMsg = steGetText("Please select the author for this article")
	ElseIf steForm("categoryid") = "" Then
		sErrorMsg = steGetText("Please check at least one category for this article")
	ElseIf Trim(steForm("Title")) = "" Then
		sErrorMsg = steGetText("Please enter the title for this article")
	ElseIf Trim(steForm("LeadIn")) = "" Then
		sErrorMsg = steGetText("Plese enter the Lead-In for this article")
	ElseIf Trim(steForm("ArticleBody")) = "" Then
		sErrorMsg = steGetText("Please enter the Body for this article")
	Else
		' update the existing article in the database
		sStat = "UPDATE tblArticle SET " &_
				"AuthorID = " & steNForm("AuthorID") & "," &_
				"Title = " & steQForm("Title") & "," &_
				"LeadIn = " & steQForm("LeadIn") & "," &_
				"ArticleBody = " & steQForm("ArticleBody") &_
				"WHERE	ArticleID = " & steNForm("articleID")
		Call adoExecute(sStat)

		' assign the categories for this article
		Call adoExecute("DELETE FROM tblArticleToCategory WHERE ArticleID = " & steNForm("articleID"))
		Dim aCatID
		aCatID = Split(steForm("CategoryID"), ",")
		For I = 0 To UBound(aCatID)
			sStat = "INSERT INTO tblArticleToCategory (" &_
					"	ArticleID, CategoryID" &_
					") VALUES (" &_
					steNForm("articleID") & "," & Trim(aCatID(I)) &_
					")"
			Call adoExecute(sStat)
		Next
	End If
End If

' build the list of authors to choose from
sStat = "SELECT	AuthorID, Title, FirstName, MiddleName, LastName, Surname " &_
		"FROM	tblArticleAuthor " &_
		"WHERE	Active <> 0 " &_
		"AND	Archive = 0 " &_
		"ORDER BY LastName, FirstName, MiddleName"
Set rsAuth = adoOpenRecordset(sStat)

' retrieve the list of selected categories
sStat = "SELECT	CategoryID FROM tblArticleToCategory WHERE ArticleID = " & steNForm("ArticleID")
Set rsCat = adoOpenRecordset(sStat)
Do Until rsCat.EOF
	sCatList = sCatList & "," & rsCat.Fields("CategoryID").Value
	rsCat.MoveNext
Loop
rsCat.Close

' retrieve the article we are editing
sStat = "SELECT * FROM tblArticle WHERE ArticleID = " & steNForm("ArticleID")
Set rsArt = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Article" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If sAction <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Article" %></H3>

<P>
<% steTxt "Please enter the changes for the article in the form below." %>
</P>

<FORM METHOD="post" ACTION="article_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="ArticleID" VALUE="<%= steNForm("ArticleID") %>">

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD VALIGN="top" class="forml"><% steTxt "Author" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><SELECT NAME="AuthorID" class="form">
	<OPTION VALUE="0"> -- Choose --
	<% Do Until rsAuth.EOF %>
	<OPTION VALUE="<%= rsAuth.Fields("AuthorID").Value %>"<% If steRecordEncValue(rsArt, "AuthorID") = CStr(rsAuth.Fields("AuthorID").Value) Then Response.Write " SELECTED" %>> <%= rsAuth.Fields("Title").Value & " " & rsAuth.Fields("FirstName").Value & " " & rsAuth.Fields("MiddleName").Value & " " & rsAuth.Fields("LastName").Value & " " & rsAuth.Fields("Surname").Value %>
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
	<TD VALIGN="top" class="forml"><% steTxt "Title" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Title" VALUE="<%= steRecordEncValue(rsArt, "title") %>" SIZE="32" MAXLENGTH="100" style="width:100%" class="form"></TD>
</TR><TR>
	<TD class="forml" VALIGN="top"><% steTxt "Article Lead In" %></TD><TD></TD>
	<TD><TEXTAREA NAME="LeadIn" COLS="58" ROWS="10" WRAP="Virtual" class="form" style="width:100%"><%= steRecordEncValue(rsArt, "LeadIn") %></TEXTAREA></TD>
</TR><TR>
	<TD class="forml" VALIGN="top"><% steTxt "Article Body" %></TD><TD></TD>
	<TD><TEXTAREA NAME="ArticleBody" COLS="58" ROWS="10" WRAP="Virtual" class="form" style="width:100%"><%= steRecordEncValue(rsArt, "ArticleBody") %></TEXTAREA></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update Article" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Article Updated" %></H3>

<P>
<% steTxt "The changes to the article were made successfully." %>&nbsp;
<% steTxt "You may use the admin links above to continue administering the site." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="article_list.asp" class="adminlink"><% steTxt "Article List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->