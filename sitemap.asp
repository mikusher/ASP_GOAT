<!-- #include file="lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' sitemap.asp
'	Build a site map for the items stored in the menu module.
'	(useful for spiders to crawl to those elements)
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
%>
<!-- #include file="header.asp" -->

<h3><%= Application("CompanyName") %> Site Map</h3>

<p>
A complete list of the different areas on our web site.
</p>

<% locSiteMap %>
<% If Application("SITEMAP") <> "" Then %>

<table border=0 cellpadding=0 cellspacing=0 width="80%" align="Center">
<tr>
<% SplitWrite Application("SITEMAP"), 2, "<p>", "<td valign=""top"">", "</td>" %>
</tr>
</table>

<% Else %>
<p><b class="error">Sorry, Site map is currently empty.</b></p>
<% End If %>

<!-- #include file="footer.asp" -->
<%
' build a version of the site map based on the dynamic menu system
Sub locSiteMap
	Dim sStat, rs, oItem, sName, sURL, I
	Dim aField, aChild, aMenu, sHTML

	If IsDate(Application("SITEMAPUPDATED")) Then
		If DateDiff("n", Application("SITEMAPUPDATED"), Now()) < 30 Then Exit Sub
	End If

	sStat = "SELECT	ItemID, ParentItemID, MenuName, URL " &_
			"FROM	tblMenuItem " &_
			"WHERE	Active = 1 " &_
			"AND	Archive = 0 " &_
			"ORDER BY OrderNo"
	Set rs = adoOpenRecordset(sStat)
	Set oItem = Server.CreateObject("Scripting.Dictionary")
	Do Until rs.EOF
		' add the child list
		oItem("CHILD" & rs.Fields("ParentItemID").Value) = oItem("CHILD" & rs.Fields("ParentItemID").Value) &_
			"," & rs.Fields("ItemID").Value
		' build the URL for the menu item
		If Trim(rs.Fields("URL").Value & "") = "" Then
			sURL = Application("ASPNukeBasePath") & "content.asp?ID=" & Server.URLEncode(rs.Fields("MenuName").Value)
		ElseIf InStr(1, rs.Fields("URL").Value & "", "://") > 0 Then
			sURL = rs.Fields("URL").Value
		Else
			sURL = Replace(Application("ASPNukeBasePath") & rs.Fields("URL").Value, "//", "/")
		End If
		If rs.Fields("ParentItemID").Value <> 0 Then
			oItem("MENU" & rs.Fields("ParentItemID").Value) = oItem("MENU" & rs.Fields("ParentItemID").Value) &_
				"<a href=""" & sURL & """ class=""normal"">" & rs.Fields("MenuName").Value & "</a><br>" & vbCrLf
		Else
			sName = sName & "^" & rs.Fields("MenuName").Value & "|" & sURL
		End If
		rs.MoveNext
	Loop
	rs.Close
	rs = Empty

	aMenu = Split(Mid(sName, 2), "^")
	aChild = Split(Mid(oItem.Item("CHILD0"), 2), ",")
	With Response
		For I = 0 To UBound(aMenu)
			aField = Split(aMenu(I), "|")
			' top-level items are not links for now
			' sHTML = sHTML & "<p><a href=""" & aField(1) & """ class=""big"">" & aField(0) & "</a></p>" & vbCrLf
			sHTML = sHTML & "<p><a href=""javascript:void(0)"" class=""big"">" & aField(0) & "</a></p>" & vbCrLf
			If (oItem.Exists("CHILD" & aChild(I))) Then
				sHTML = sHTML & "<BLOCKQUOTE>" & vbCrLf
				sHTML = sHTML & oItem.Item("MENU" & aChild(I)) & vbCrLf
				sHTML = sHTML & "</BLOCKQUOTE>" & vbCrLf
			End If
		Next
	End With
	Application("SITEMAP") = sHTML
	Application("SITEMAPUPDATED") = Now()
End Sub

Sub SplitWrite(sText, nColumns, sSplitOn, sBefore, sAfter)
	Dim nStart, nEnd, nPos, nLeft, nRight, nLen

	nLen = Len(sText)
	nStart = 1
	With Response
	For I = 1 To nColumns
		.Write sBefore
		If I = nColumns Then
			.Write Mid(sText, nStart)
		Else
			nPos = CInt((I * nLen) / (nColumns))
			nRight = InStr(nPos, sText, sSplitOn, vbTextCompare)
			nLeft = InStrRev(sText, sSplitOn, nPos, vbTextCompare)
			If (nRight > 0 And nLeft > 0 And Abs(nPos - nRight) < Abs(nPos - nLeft)) _
			Or nLeft = 0 Then
				.Write Mid(sText, nStart, (nRight - nStart))
				nStart = nRight
			ElseIf nLeft > 0 Then
				.Write Mid(sText, nStart, (nLeft - nStart))
				nStart = nLeft
			Else
				.Write Mid(sText, nStart, (nPos - nStart))
				nStart = nPos
			End If
		End If
		.Write sAfter
	Next
	End With
End Sub
%>