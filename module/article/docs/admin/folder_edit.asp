<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/class/form/listinput.asp" -->
<%
'--------------------------------------------------------------------
' folder_edit.asp
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
Dim nFolderID
Dim sErrorMsg	' error message to display to user

nFolderID = steNForm("folderid")

If UCase(steForm("action")) = "EDIT" Then
	' check for required fields here
	If Trim(steForm("FolderName")) = "" Then
		sErrorMsg = steGetText("Please enter the Folder Name for this folder")
	ElseIf steForm("ShortDescription") = "" Then
		sErrorMsg = steGetText("Plese enter the Short Description for this folder")
	Else
		' determine the new order no
		Set rsOrder= adoOpenRecordset("SELECT Coalesce(Max(OrderNo) + 1, 1) As OrderNo FROM tblDocFolder WHERE ParentFolderID = " & steNForm("ParentFolderID"))
		If Not rsOrder.EOF Then nOrderNo = rsOrder.Fields("OrderNo").Value Else nOrderNo = 1
		rsOrder.Close
		Set rsOrder = Nothing

		' attempt to find the user from the database
		Dim rsUser, nUserID
		'nUserID = 0
		'set rsUser = adoOpenRecordset("SELECT UserID FROM tblUser WHERE Username = " & steQForm("Username"))
		'If Not rsUser.EOF Then
		'	nUserID = rsUser.Fields("UserID").Value
		'End If
		nUserID = CInt(Request.Cookies("AdminUserID"))

		If nUserID = 0 Then
			sErrorMsg = "Unrecognized username: """ & steForm("Username") & """"
		Else
			' insert the new folder into the database
			sStat = "UPDATE tblDocFolder SET " &_
					"	ParentFolderID = " & steNForm("ParentFolderID") & ", " &_
					"	FolderName = " & steQForm("FolderName") & ", " &_
					"	ShortDescription = " &  steQForm("ShortDescription") & " " &_
					"WHERE FolderID = " & nFolderID
			Call adoExecute(sStat)
		End If
	End If
End If

' retrieve the folder to edit here
Set rsFolder = adoOpenRecordset("SELECT * FROM tblDocFolder WHERE FolderID = " & nFolderID)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Folder" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If UCase(steForm("action")) <> "EDIT" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Folder" %></H3>

<P>
<% steTxt "Please enter your changes for the document folder in the form below." %>
</P>

<FORM METHOD="post" ACTION="folder_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="folderid" VALUE="<%= nFolderID %>">

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Parent Folder" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><%
	Dim oList
	Set oList = New clsListInput
	oList.ChooseOptionLabel = "TOP-LEVEL FOLDER"
	oList.TreeListInput "ParentFolderID", "tblDocFolder", "FolderID", "ParentFolderID", "FolderID <> " & nFolderID, _
		"OrderNo", "FolderID", "FolderName", steRecordEncValue(rsFolder, "ParentFolderID"), "", False
	%>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Folder Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="FolderName" VALUE="<%= steRecordEncValue(rsFolder, "foldername") %>" SIZE="32" MAXLENGTH="50" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Short Description" %></TD><TD></TD>
	<TD><TEXTAREA NAME="ShortDescription" COLS="42" ROWS="10" WRAP="Virtual" class="form"><%= steRecordEncValue(rsFolder, "ShortDescription") %></TEXTAREA></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update Document Folder" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Document Folder Updated" %></H3>

<P>
<% steTxt "The document folder was successfully updated." %>&nbsp;
<% steTxt "You may use the admin links above to continue administering the site." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="folder_list.asp" class="adminlink"><% steTxt "Folder List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->