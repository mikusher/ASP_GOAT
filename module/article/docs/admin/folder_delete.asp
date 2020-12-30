<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/class/form/listinput.asp" -->
<%
'--------------------------------------------------------------------
' folder_add.asp
'	Displays a list of the folders for organizing documents.
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
Dim rsFolder	' folder to edit
Dim sCatList	' list of currently selected categories
Dim sErrorMsg	' error message to display to user
Dim nFolderID	' folder to delete

nFolderID = steNForm("FolderID")

If UCase(steForm("action")) = "DELETE" Then
	If steNForm("confirm") <> 1 Then
		sErrorMsg = steGetText("Please confirm the deletion of this document folder")
	Else
		' insert the new folder into the database
		sStat = "DELETE FROM tblDocFolder " &_
				"WHERE FolderID = " & nFolderID
		Call adoExecute(sStat)
	End If
End If

' retrieve the folder to edit here
Dim sParentFolder		' name of the parent folder
If nFolderID > 0 Then
	Set rsFolder = adoOpenRecordset("SELECT * FROM tblDocFolder WHERE FolderID = " & nFolderID)
	If Not rsFolder.EOF Then
		Dim rsParent
		Set rsParent = adoOpenRecordset("SELECT FolderName FROM tblDocFolder WHERE FolderID = " & rsFolder.Fields("ParentFolderID").Value)
		If Not rsParent.EOF Then sParentFolder = rsParent.Fields("FolderName").Value
		rsParent.Close
		Set rsParent = Nothing
	End If
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Folder" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If UCase(steForm("action")) <> "DELETE" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Document Folder" %></H3>

<P>
<% steTxt "Please confirm that you would like to delete the document folder shown below." %>
</P>

<FORM METHOD="post" ACTION="folder_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="folderid" VALUE="<%= nFolderID %>">

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Parent Folder" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD CLASS="formd"><%= sParentFolder %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Folder Name" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsFolder, "foldername") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Short Description" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsFolder, "ShortDescription") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Confirm Delete" %></TD><TD></TD>
	<TD CLASS="formd">
		<input type="radio" name="confirm" value="1" class="form"> Yes
		<input type="radio" name="confirm" value="0" class="form"> No
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Document Folder" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Document Folder Deleted" %></H3>

<P>
<% steTxt "The document folder was successfully deleted." %>&nbsp;
<% steTxt "You may use the admin links above to continue administering the site." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="folder_list.asp" class="adminlink"><% steTxt "Folder List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->