<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="rss/rsspub_lib.asp" -->
<%
'--------------------------------------------------------------------
' rss_publish.asp
'	Publish the RSS news feed file to the local filesystem.
'	Requires write access to the folder/file defined by the module
'	parameter "RSSFeedFile".
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
Dim oRDF		' RDF class
Dim sErrorMsg	' error to report to the user


Set oRDF = New clsRDF
oRDF.RSSTemplate = "rss/rsspub_template.xml"
If InStrRev(oRDF.RSSTemplate, "/") > 0 Then
	Dim sRSS
	' get  the starting path from the SiteRoot configuration
	If InStr(1, Replace(Application("SiteRoot"), "//", ""), "/") > 0 Then
		sRSS = Mid(Application("SiteRoot"), InStrRev(Application("SiteRoot"), "/"))
		If Right(sRSS, 1) = "/" Then sRss = Left(sRSS, Len(sRSS) - 1)
	End If
	sRSS = sRSS & modParam("Articles", "RSSFeedFile")
	oRDF.RSSPath = Left(sRSS, InStrRev(sRSS, "/") - 1)
	oRDF.RSSFile = Mid(sRSS, InStrRev(sRSS, "/") + 1)
End If

oRDF.RDF_About = Application("SiteRoot")
oRDF.RDF_Title = Application("CompanyName")
oRDF.RDF_Link = Application("SiteRoot")
oRDF.RDF_ResourceImage = Application("SiteRoot") & "/img/logo.gif"
oRDF.DC_Rights = Application("CompanyName") & " - Copyright " & Year(Now()) & " - All Rights Reserved"
oRDF.DC_Creator = "Orvado Technologies RSS Generator"
oRDF.DC_Publisher = Application("SupportEmail")
oRDF.DC_Subject = Application("CompanyName") & " Articles"
oRDF.Image_RDFAbout = Application("SiteRoot") & "/img/logo.gif"
oRDF.Image_Title = Application("CompanyName")
oRDF.Image_URL = Application("SiteRoot") & "/img/logo.gif"
oRDF.Image_Link = Application("CompanyName")

' retrieve the list of RSS articles to publish here
Dim rsArt, sArticleURL
sStat = "SELECT	" & adoTop(10) & " art.ArticleID, art.Title, art.LeadIn, auth.FirstName, auth.MiddleName, " &_
		"		auth.LastName, ac.CategoryName, art.Created, art.Modified " &_
		"FROM	tblArticle art " &_
		"INNER JOIN	tblArticleAuthor auth ON art.AuthorID = auth.AuthorID " &_
		"INNER JOIN	tblArticleToCategory atc ON atc.ArticleID = art.ArticleID " &_
		"INNER JOIN	tblArticleCategory ac ON ac.CategoryID = atc.CategoryID " &_
		"WHERE	art.Active <> 0 " &_
		"AND	art.Archive = 0 " & sWhere &_
		"ORDER BY art.Created DESC" & adoTop2(10)
Set rsArt = adoOpenRecordset(sStat)
Do Until rsArt.EOF
	sArticleURL = Application("SiteRoot") & "/module/article/article/article.asp?articleid=" & rsArt.Fields("ArticleID").Value
	oRDF.AddArticle sArticleURL, rsArt.Fields("Title").Value, sArticleURL, rsArt.Fields("LeadIn").Value, _
		rsArt.Fields("FirstName").Value & " " & rsArt.Fields("MiddleName").Value & " " & rsArt.Fields("LastName").Value, _
		rsArt.Fields("CategoryName").Value, _
		UTFDateTime(rsArt.Fields("Modified").Value), "", "", ""
	rsArt.MoveNext
Loop
rsArt.Close
Set rsArt = Nothing

If Not oRDF.BuildDynamic Then
	sErrorMsg = oRDF.ErrorMsg
End If
%>
<!-- #include file="../../../header.asp" -->

<% sCurrentTab = "Article" %>


<h3><% steTxt "Build Dynamic RSS Feed" %></h3>

<p>
<% steTxt "There was an error trying to build the dynamic RSS news feed." %>&nbsp;
</p>

<% If sErrorMsg <> "" Then %>
	<p><b class="error"><%= sErrorMsg %></b></p>
<% End If %>

<!-- #include file="../../../footer.asp" -->
<%
'------------------------------------------------------------------------
' Generate the current date and time in UTF (Universal Time Format)
' like: "2003-11-26T21:13:06-08:00" for Pacific Standard Time

Function UTFDateTime(dtDate)
	Dim nMonth, nDay, nHour, nMinute, nSecond

	If Month(dtDate) < 10 Then nMonth = "0" & Month(dtDate) Else nMonth = Month(dtDate)
	If Day(dtDate) < 10 Then nDay = "0" & Day(dtDate) Else nDay = Day(dtDate)
	If Hour(dtDate) < 10 Then nHour = "0" & Hour(dtDate) Else nHour = Hour(dtDate)
	If Minute(dtDate) < 10 Then nMinute = "0" & Minute(dtDate) Else nMinute = Minute(dtDate)
	If Second(dtDate) < 10 Then nSecond = "0" & Second(dtDate) Else nSecond = Second(dtDate)

	UTFDateTime = Year(dtDate) & "-" & nMonth & "-" & nDay &_
		"T" & nHour & ":" & nMinute & ":" & nSecond & "-08:00"
End Function
%>