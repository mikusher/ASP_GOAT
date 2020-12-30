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
Dim rsCat		' list of categories to choose from
Dim sCatList	' list of currently selected categories
Dim sErrorMsg	' error message to display to user

If UCase(steForm("action")) = "ADD" Then
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
			sStat = "INSERT INTO tblDocFolder (" &_
					"	CreatedByUserID, ParentFolderID, FolderName, ShortDescription, OrderNo, Created" &_
					") VALUES (" &_
					nUserID & "," & steNForm("ParentFolderID") & "," &_
					steQForm("FolderName") & "," & steQForm("ShortDescription") & "," &_
					nOrderNo & "," & adoGetDate &_
					")"
			Call adoExecute(sStat)
		End If
	End If
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Folder" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If UCase(steForm("action")) <> "ADD" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New Folder" %></H3>

<P>
<% steTxt "Please enter the information for the new document folder in the form below." %>
</P>

<FORM METHOD="post" ACTION="folder_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

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
	oList.TreeListInput "ParentFolderID", "tblDocFolder", "FolderID", "ParentFolderID", "", _
		"OrderNo", "FolderID", "FolderName", steNForm("ParentFolderID"), "", False
	%>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Folder Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="FolderName" VALUE="<%= steEncForm("foldername") %>" SIZE="32" MAXLENGTH="50" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Short Description" %></TD><TD></TD>
	<TD><TEXTAREA NAME="ShortDescription" COLS="42" ROWS="10" WRAP="Virtual" class="form"><%= steEncForm("ShortDescription") %></TEXTAREA></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Document Folder" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Document Folder Added" %></H3>

<P>
<% steTxt "The new document folder was successfully added." %>&nbsp;
<% steTxt "You may use the admin links above to continue administering the site." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="folder_list.asp" class="adminlink"><% steTxt "Folder List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->