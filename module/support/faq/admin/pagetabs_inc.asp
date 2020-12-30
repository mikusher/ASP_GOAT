<!-- #include file="../../../../lib/tab_lib.asp" -->
<%
Dim sCurrentTab
%>
<% tabShow "Document,Questions,Author", "document_list.asp,question_list.asp?documentid=" & Request("DocumentID") & ",author_list.asp", sCurrentTab %>
