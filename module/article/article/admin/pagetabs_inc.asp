<!-- #include file="../../../../lib/tab_lib.asp" -->
<%
Dim sTabList
Dim sTabLinkList
Dim sCurrentTab

sTabList = "Article"
sTabLinkList = "article_list.asp"
If Request.QueryString("articleid") <> "" Then
	sTabList = sTabList & ",Comments"
	sTabLinkList = sTabLinkList & ",comments_list.asp?articleid=" & Request.QueryString("articleid")
End If
sTabList = sTabList & ",Author,Category"
sTabLinkList = sTabLinkList & ",author_list.asp,category_list.asp"
%>
<% tabShow sTabList, sTabLinkList, sCurrentTab %>

