<!-- #include file="../../../lib/ado_lib.asp" -->
<!-- #include file="../../../lib/lang_lib.asp" -->
<%
'--------------------------------------------------------------------
' capsule.asp
'	Create the FAQ capsule which will appear on
'	all pages of the site.
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

<%= Replace(Application("ModCapHeader"), "<CAPSULETITLE/>", steGetText("FAQ")) %>
<%= Application("ModCapLeft") %>

<DIV class="forumcapsule">
<% Call locCacheFAQ
	If Application("FAQCAPSULE") <> "" Then %>

<%= Application("FAQCAPSULE") %>

<% Else %>

<P><B CLASS="Error"><% steTxt "No FAQs are defined yet" %></B></P>

<% End If %>
</DIV>

<%= Application("ModCapRight") %>
<%= Application("ModCapFooter") %>
<%
Sub locCacheFAQ
	Dim rsDoc
	Dim sStat, bFirst, sHTML, I
	
	' check to see if we need to refresh (every 15 mins)
	If IsDate(Application("FAQCAPSULEREFRESH")) Then
		If DateDiff("n", Application("FAQCAPSULEREFRESH"), Now()) < 15 Then Exit Sub
	End If

	sStat = "SELECT	" & adoTop(4) & " fd.DocumentID, fd.Title " &_
			"FROM	tblFaqDocument fd " &_
			"INNER JOIN	tblFaqAuthor fa on fd.AuthorID = fa.AuthorID " &_
			"WHERE	fd.Active <> 0 " &_
			"AND	fd.Archive = 0 " &_
			"ORDER BY fd.OrderNo DESC" & adoTop2(4)
			'		coalesce(fa.Title + ' ', '') + fa.FirstName + ' ' + coalesce(fa.MiddleName + ' ', '') + fa.LastName As AuthorName " &_
	Set rsDoc = adoOpenRecordset(sStat)

	bFirst = False
	For I = 1 To 3
		If rsDoc.EOF Then Exit For
		If bFirst Then sHTML = sHTML & "<hr class=""forumcapsulesep"">" & vbCrLf
		sHTML = sHTML & "<a href=""" & Application("ASPNukeBasePath") & "module/support/faq/document.asp?documentid=" & rsDoc.Fields("DocumentID").Value & """ class=""forumtopic"">" &_
			rsDoc.Fields("Title").Value & "</a>" & vbCrLf
		rsDoc.MoveNext
		bFirst = True
	Next

	' add the archive link if more than 3 documents are available
	If Not rsDoc.EOF Then
		sHTML = sHTML & "<p align=""center"">" & vbCrLf &_
			"<A href=""" & Application("ASPNukeBasePath") & "module/support/faq/index.asp"" class=""actionlink"">Archive</a>" & vbCrLf &_
			"</p>"
	End If
	rsDoc.Close
	Set rsDoc = Nothing
	Application("FAQCAPSULEREFRESH") = Now()
	Application("FAQCAPSULE") = sHTML
End Sub
%>