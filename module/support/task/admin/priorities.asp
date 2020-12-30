<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' priorities.asp
'	Administer the task priorities
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

strTitle = steGetText("Task Priorities")
strObjectName = steGetText("Priority")
strPrimaryKey = "PriorityID"
strPrimaryKeyValue = steNForm("PriorityID")
strTableName = "tblTaskPriority"
strEditFields = "PriorityName,ColorCode,Comments"
strEditLabels = steGetText("Priority Name,HTML Color Code,Comments")
strEditSizes = "32,16,0"
strEditTypes = "T,T,A"
strDisplayFields = "OrderNo,PriorityName,ColorCode"
strDisplayLabels = steGetText("Order,Priority,Color")
strOrderField = "OrderNo"
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Priority" %>
<!-- #include file="pagetabs_inc.asp" -->

<!-- #include file="../../../../lib/wizard/admin_list.asp" -->

<!-- #include file="../../../../footer.asp" -->