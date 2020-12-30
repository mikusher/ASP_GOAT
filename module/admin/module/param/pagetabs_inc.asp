<!-- #include file="../../../../lib/tab_lib.asp" -->
<%
Dim sCurrentTab
Dim sTabQuery
If Request.QueryString("ModuleID") <> "" Then
	sTabQuery = "?ModuleID=" & Request.QueryString("ModuleID")
End If
%>
<% tabShow "Param,Options,Types", "param_list.asp"&sTabQuery&",option_list.asp"&sTabQuery&",type_list.asp"&sTabQuery, sCurrentTab %>
