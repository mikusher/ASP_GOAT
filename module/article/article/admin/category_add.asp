<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/graphics/gfx_lib.asp" -->
<%
'--------------------------------------------------------------------
' category_add.asp
'	Add a new article category to the database
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
Dim sIconImage

' process the form post
Call steFormPost("/img/articles/category/", sErrorMsg)

' validate the image uploaded (if any)
If steForm("action") = "add" Then
	Dim nWidth, nHeight, nColors, sImgType

	' define the category icon
	If steForm("iconimagefile") <> "" Then
		sIconImage = Application("ASPNukeBasePath") & "img/articles/category/" & steForm("iconimagefile")
		' check the size of the icon image
		If gfxSpex(Server.MapPath(sIconImage), nWidth, nHeight, nColors, sImgType) Then
			If nWidth <> modParam("Articles", "IconImageWidth") Or nHeight <> modParam("Articles", "IconImageHeight") Then
				sErrorMsg = steGetText("Invalid Icon Image Size") & " (" & nWidth & "x" & nHeight &_
					") - " & steGetText("Should be") & " (" & modParam("Articles", "IconImageWidth") & "x" & modParam("Articles", "IconImageHeight") & ")<br>"
			End If
		Else
			sErrorMsg = steGetText("File is corrupt") & " (" & steForm("iconimagefile") & ") - " & steGetText("Expected GIF, JPG or PNG image") & "<br>"
		End If
	Else
		sIconImage = steForm("iconimage")
	End If
End If

If steForm("action") = "add" And sErrorMsg = "" Then

	' make sure the required fields are present
	If Trim(steForm("CategoryName")) = ""	Then
		sErrorMsg = steGetText("Please enter the name for this category")
	Else
		' create the new article category in the database
		sStat = "INSERT INTO tblArticleCategory (" &_
				"	CategoryName, Comments, IconImage, Created " &_
				") VALUES (" &_
				steQForm("CategoryName") & "," &_
				steQForm("Comments") & "," &_
				"'" & Replace(sIconImage, "'", "''") & "'," & adoGetDate &_
				")"
		Call adoExecute(sStat)
	End If
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Category" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New Article Category" %></H3>

<P>
<% steTxt "Please enter the new properties for the new article category using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="category_add.asp" enctype="multipart/form-data">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Name" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="CategoryName" VALUE="<%= steEncForm("CategoryName") %>" SIZE="32" MAXLENGTH="40" class="form"></TD>
</TR><tr>
	<td class="forml"><% steTxt "Icon Image" %><br><font class="tinytext">(<% steTxt "GIF or JPG Image" %>)</font></td><td></td>
	<td><input type="text" name="iconimage" value="<%= steEncForm("iconimage") %>" size="42" maxlength="255" class="form"></td>
</tr><tr>
	<td class="forml"><% steTxt "Upload Icon Image" %><br><font class="tinytext">(<% steTxt "GIF or JPG Image" %>)</font></td><td></td>
	<td><input type="file" name="iconimagefile" size="42" maxlength="255" class="form"></td>
</tr><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Comments" %></TD><TD></TD>
	<TD><TEXTAREA NAME="Comments" COLS="52" ROWS="10" WRAP="virtual" class="form"><%= steEncForm("Comments") %></TEXTAREA></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Category" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "New Article Category Added" %></H3>

<P>
<% steTxt "The new article category has been added to the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="category_list.asp" class="adminlink"><% steTxt "Category List" %></A> |
	<A HREF="article_list.asp" class="adminlink"><% steTxt "Article List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->
