<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' status.asp
'	Administer the task status
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

strTitle = steGetText("Task Status")
strObjectName = steGetText("Status")
strPrimaryKey = "StatusID"
strPrimaryKeyValue = steNForm("StatusID")
strTableName = "tblTaskStatus"
strEditFields = "StatusName,Comments"
strEditLabels = steGetText("Status Name,Comments")
strEditSizes = "50,0"
strEditTypes = "T,A"
strDisplayFields = "StatusName"
strDisplayLabels = steGetText("Status")
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Status" %>
<!-- #include file="pagetabs_inc.asp" -->

<!-- #include file="../../../../lib/wizard/admin_list.asp" -->

<!-- #include file="../../../../footer.asp" -->