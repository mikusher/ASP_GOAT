<%
'--------------------------------------------------------------------
' rss_lib.asp
'	This library will retrieve RSS (really simple synidication) feeds
'	from a web site and display the information on the page.
'
' AUTH:	Ken Richards
' DATE:	07/25/2001
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

'------------------------------------------------------------------
' Parse all of the items in an RDF file using a regular expression
' search pattern.

Function rssParseItems(sRDF)
	Dim regex, oMatch, oMatches

	Set regex = New RegExp
	regex.Pattern = "<item rdf:about=""([^""]*)"">(.*?)</item>"
	regex.Global = True
	regex.IgnoreCase = True
	Set rssParseItems = regex.Execute(sRDF)
End Function

'------------------------------------------------------------------
' Display the top "nMax" items from an RDF file

Function rssBuildList(sURL, nMax, sListBegin, sListItem, sListEnd)
	Dim regex, oMatch, oMatches, oMatch2, oMatches2, oHTTP, sRDF
	Dim nCount, sTitle, sLink, sDesc, sHTML

	' get the XML content for the RSS feed (simple ain't it?)
	Set oHTTP = Server.CreateObject("MSXML.ServerXMLHTTP")
	oHTTP.Open"GET", url
	oHTTP.Send ""
	sRDF = oHTTP.ResponseText

	' build the regex component for parsing individual list items
	Set regex = New RegExp
	regex.Global = True
	regex.IgnoreCase = True

	' get the regex match collection for all of the items in the feed
	Set oMatches = rssParseItems(sRDF)
	' display the beginning of the RDF list here
	With Response
	If sListItem = "" Then
		sHTML = sHTML & "<p>" & vbCrLf
	Else
		sHTML = sHTML & sListBegin
	End If

	nCount = 0
	For Each oMatch In oMatches
		sTitle = ""
		sLink = ""
		sDesc = ""
		' extract the title for this item
		regex.Pattern = "<title>(.*?)</title>"
		Set oMatches2 = regex.Execute(oMatch.SubMatches(1))
		For Each oMatch2 In oMatches2
			sTitle = oMatch2.SubMatches(0)
		Next
		' extract the link (url) for this item
		regex.Pattern = "<link>(.*?)</link>"
		Set oMatches2 = regex.Execute(oMatch.SubMatches(1))
		For Each oMatch2 In oMatches2
			sLink = oMatch2.SubMatches(0)
		Next
		' extract the description (synopsis) for this item
		regex.Pattern = "<description>(.*?)</description>"
		Set oMatches2 = regex.Execute(oMatch.SubMatches(1))
		For Each oMatch2 In oMatches2
			sDesc = oMatch2.SubMatches(0)
		Next
		' display using template OR default layout
		If sTitle <> "" And sLink <> "" Then
			If sListItem = "" Then			
				sHTML = sHTML & "<p><a href=""" & sLink & """ class=""rsslink"">"
				sHTML = sHTML & sTitle
				sHTML = sHTML & "</a>" & vbCrLf
				If sDesc <> "" Then sHTML = sHTML & "<br>" & sDesc
				sHTML = sHTML & "</p>"
				' sHTML = sHTML & oMatch.SubMatches(0)
			Else
				sHTML = sHTML & Replace(Replace(Replace(sListItem, "<Title/>", sTitle), "<Link/>", sLink), "<Description/>", sDesc)
			End If
			nCount = nCount + 1
			If nMax <> 0 And nCount >= nMax Then Exit For
		End If
	Next	
	' display the end of the RDF list here
	If sListItem = "" Then
		sHTML = sHTML & "</p>" & vbCrLf
	Else
		sHTML = sHTML & sListEnd
	End If
	End With
	rssBuildList = sHTML
End Function

'------------------------------------------------------------------
' Display the top "nMax" items from an RDF file and cache the
' results in the Application object

Function rssCapsule(nFeedID, sTitle)
	Dim sHTML
	If Application("RSS" & nFeedID & "_CACHEHOURS") <> "" And _
		Application("RSS" & nFeedID & "_UPDATED") <> "" Then
		If DateDiff("h", Application("RSS" & nFeedID & "_UPDATED"), Now()) < Application("RSS" & nFeedID & "_CACHEHOURS") Then
			rssCapsule = Application("RSS" & nFeedID & "_CACHE")
			Exit Function
		End If
	End If 
		
	' retrieve the configuration for the feed
	sStat = "SELECT	Title, FeedURL, MaxItems, ShowDescription, CacheHours " &_
			"FROM	tblRSSFeed " &_
			"WHERE	FeedID = " & nFeedID
	Set rsFeed = adoOpenRecordset(sStat)
	If Not rsFeed.EOF Then
		sHTML = rssBuildList(rsFeed.Fields("FeedURL").Value, rsFeed.Fields("MaxItems").Value, _
			"", "", "")
	End If
	' cache this feed in the application object
	Application("RSS" & nFeedID & "_TITLE") = rs.Fields("Title").Value
	Application("RSS" & nFeedID & "_CACHE") = sHTML
	Application("RSS" & nFeedID & "_CACHEHOURS") = rs.Fields("CacheHours").Value
	Application("RSS" & nFeedID & "_UPDATED") = Now()
	Response.Write sHTML
End Function
%>