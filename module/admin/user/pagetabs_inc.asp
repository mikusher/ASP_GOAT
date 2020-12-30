<!-- #include file="../../../lib/tab_lib.asp" -->
<%
Dim sCurrentTab
%>
<% tabShow "Users,Rights", "user_list.asp,userright_edit.asp?userid=" & steNForm("UserID"),  sCurrentTab %>
