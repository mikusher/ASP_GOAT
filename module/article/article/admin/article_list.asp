<% Option Explicit %>
<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/class/adminlist.asp" -->
<%
'--------------------------------------------------------------------
' article_list.asp
'	Displays a list of the current articles for the site
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
Dim nCategoryID
Dim oList			' admin list object
Dim sWhere
Dim sJoin
Dim I

nCategoryID = steNForm("CategoryID")

%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<script language="javascript" type="text/javascript">
function pickCat(nCatID) {
	location.href="article_list.asp?categoryid=" + nCatID;
}
</script>

<% sCurrentTab = "Article" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "Article List" %></H3>

<p>
<form method="post" action="article_list.asp">
<table border=0 cellpadding=2 cellspacing=0>
<tr>
	<td class="forml" nowrap><% steTxt "Displaying Articles in Category:" %> </td><Td>&nbsp;&nbsp;</td>
	<td nowrap><select name="CategoryID" class="form" onChange="pickCat(this.options[this.selectedIndex].value)">
	<option value="0"> <% steTxt "All Categories" %>
<%	' build the list of categories to choose from
	Dim rsCat
	sStat = "SELECT	ac.CategoryID, ac.CategoryName, Count(*) As ArticleCount " &_
			"FROM	tblArticleCategory ac " &_
			"INNER JOIN	tblArticleToCategory atc ON atc.CategoryID = ac.CategoryID " &_
			"INNER JOIN	tblArticle art ON art.ArticleID = atc.ArticleID " &_
			"WHERE	ac.Active <> 0 " &_
			"AND	ac.Archive = 0 " &_
			"AND	art.Active <> 0 " &_
			"AND	art.Archive = 0 " &_
			"GROUP BY ac.CategoryID, ac.CategoryName " &_
			"ORDER BY ac.CategoryName"
	Set rsCat = adoOpenRecordset(sStat)
	Do Until rsCat.EOF %>
	<option value="<%= rsCat.Fields("CategoryID").Value %>"<% If nCategoryID = rsCat.Fields("CategoryID").Value Then Response.Write " SELECTED" %>><%= Server.HTMLEncode(rsCat.Fields("CategoryName").Value) & " (" & rsCat.Fields("ArticleCount").Value & ")" %>
	<%	rsCat.MoveNext
	Loop
	rsCat.Close
	Set rsCat = Nothing %>
	</select>
</tr>
</table>
</form>
</p>

<%
' retrieve the list of articles to display
If nCategoryID > 0 Then sJoin = "INNER JOIN tblArticleToCategory atc ON atc.ArticleID = art.ArticleID AND atc.CategoryID = " & nCategoryID & " "
sStat = "SELECT	art.ArticleID, art.Title, auth.FirstName, auth.MiddleName, " &_
		"		auth.LastName, art.Created, art.Modified " &_
		"FROM	tblArticle art " &_
		"INNER JOIN	tblArticleAuthor auth ON art.AuthorID = auth.AuthorID " & sJoin &_
		"WHERE	art.Active <> 0 " &_
		"AND	art.Archive = 0 " &_
		"ORDER BY art.Created DESC"
Set oList = New clsAdminList
oList.query = sStat
Call oList.AddColumn("<b>##Title##</b><br><font class=""tinytext"">##FirstName## ##MiddleName## ##LastName##</font>", "Title / Author", "")
Call oList.AddColumn("Created", steGetText("Published"), "right")
Call oList.AddColumn("Modified", steGetText("Modified"), "right")
oList.ActionLink = "<A HREF=""article_edit.asp?articleid=##ArticleID##"" class=""actionlink"">" & steGetText("edit") & "</A> . <A HREF=""article_delete.asp?articleid=##ArticleID##"" class=""actionlink"">" & steGetText("delete") & "</A>"
Call oList.Display
%>

<P ALIGN="center">
<A HREF="../../../admin/configure.asp?module=Articles" class="adminlink"><% steTxt "Configure" %></A> &nbsp;
<A HREF="rss_publish.asp" class="adminlink"><% steTxt "Publish RSS" %></A> &nbsp;
<A HREF="article_add.asp" class="adminlink"><% steTxt "Add New Article" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->