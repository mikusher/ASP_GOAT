<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/graphics/gfx_lib.asp" -->

<%
'--------------------------------------------------------------------
' category_edit.asp
'	Update existing article category in the database
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
Dim rsCat
Dim nCategoryID
Dim sIconImage
Dim sPrefix

' process the form post
If InStr(1, Replace(Application("SiteRoot"), "//", ""), "/") > 0 Then
	sPrefix = Mid(Application("SiteRoot"), InStrRev(Application("SiteRoot"), "/"))
	If Right(sPrefix, 1) = "/" Then sPrefix = Left(sPrefix, Len(sPrefix) - 1)
End If
Call steFormPost(sPrefix & "/img/articles/category/", sErrorMsg)

nCategoryID = steNForm("categoryid")

' validate the image uploaded (if any)
If steForm("action") = "edit" Then
	Dim nWidth, nHeight, nColors, sImgType

	' define the category icon
	If steForm("iconimagefile") <> "" Then
		sIconImage = Application("ASPNukeBasePath") & "img/articles/category/" & steForm("iconimagefile")
		' check the size of the icon image
		If gfxSpex(Server.MapPath(sIconImage), nWidth, nHeight, nColors, sImgType) Then
			If nWidth <> modParam("Articles", "IconImageWidth") Or nHeight <> modParam("Articles", "IconImageHeight") Then
				sErrorMsg = "Invalid Icon Image Size (" & nWidth & "x" & nHeight &_
					") - Should be (" & modParam("Articles", "IconImageWidth") & "x" & modParam("Articles", "IconImageHeight") & ")<br>"
			End If
		Else
			sErrorMsg = steGetText("File is corrupt") & " (" & steForm("iconimagefile") & ") - " & steGetText("Expected GIF, JPG or PNG image") & "<br>"
		End If
	Else
		sIconImage = steForm("iconimage")
	End If
End If

If steForm("action") = "edit" And sErrorMsg = "" Then
	' make sure the required fields are present
	If Trim(steForm("CategoryName")) = ""	Then
		sErrorMsg = steGetText("Please enter the name for this category")
	Else
		' create the new article category in the database
		sStat = "UPDATE tblArticleCategory SET " &_
				"	CategoryName = " & steQForm("CategoryName") & "," &_
				"	IconImage = '" & Replace(sIconImage, "'", "''") & "', " &_
				"	Comments = " & steQForm("Comments") & " " &_
				"WHERE	CategoryID = " & nCategoryID
		Call adoExecute(sStat)
	End If
End If

sStat = "SELECT * FROM tblArticleCategory WHERE CategoryID = " & nCategoryID
Set rsCat = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Category" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Article Category" %></H3>

<P>
<% steTxt "Please make your changes to the article category using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="category_edit.asp" enctype="multipart/form-data">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="categoryid" VALUE="<%= nCategoryID %>">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="CategoryName" VALUE="<%= steRecordEncValue(rsCat, "CategoryName") %>" SIZE="32" MAXLENGTH="40" class="form"></TD>
</TR><tr>
	<td class="forml"><% steTxt "Icon Image" %><br><font class="tinytext">(GIF, JPG, PNG Image)</font></td><td></td>
	<td><input type="text" name="iconimage" value="<%= steRecordEncValue(rsCat, "iconimage") %>" size="42" maxlength="255" class="form"></td>
</tr><tr>
	<td class="forml"><% steTxt "Upload Icon Image" %><br><font class="tinytext">(GIF, JPG, PNG Image)</font></td><td></td>
	<td><input type="file" name="iconimagefile" size="42" maxlength="255" class="form"></td>
</tr><TR>
	<TD CLASS="forml" VALIGN="top">Comments</TD><TD></TD>
	<TD><TEXTAREA NAME="Comments" COLS="52" ROWS="10" WRAP="virtual" class="form"><%= steRecordEncValue(rsCat, "Comments") %></TEXTAREA></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" Update Category " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Article Category Updated" %></H3>

<P>
<% steTxt "The article category was successfully updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<!-- #include file="../../../../footer.asp" -->
