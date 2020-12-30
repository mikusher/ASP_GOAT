<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' archive.asp
'	Create the poll archive page containing a list of the past
'	50 or so polls
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
Dim nPollID

Call locPollArchiveCache
%>
<!-- #include file="../../../header.asp" -->

<h3><% steTxt "Poll Archive" %></h3>

<p>
<% steTxt "The past 50 poll questions presented by" %>&nbsp;<%= Application("CompanyName") %>&nbsp;
<% steTxt "are listed below." %>&nbsp;
<% steTxt "Click on a question to see the results and any comments that were added." %>
</p>

<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<td class="listhead"><% steTxt "Question" %></td>
	<td class="listhead"><% steTxt "Votes" %></td>
	<td class="listhead"><% steTxt "Start Date" %></td>
</tr>
<%= Application("POLLARCHCACHE") %>
</table>

<p align="center">
	<a href="detail.asp?pollid=<%= Application("POLLID") %>" class="footerlink"><% steTxt "Current Poll" %></a>
</p>

<!-- #include file="../../../footer.asp" -->
<%
Sub locPollArchiveCache
	Dim rsPoll, sStat, sHTML, nCount
	Dim sQuest		' question for the poll

	' refresh the poll archive every 24 hours
	If IsDate(Application("POLLARCHCACHEREFRESH")) Then
		If DateDiff("h", Application("POLLARCHCACHEREFRESH"), Now()) < 24 Then Exit Sub
	End If

	sStat = "SELECT	" & adoTop(50) & " tblPoll.PollID, tblPoll.Question, tblPoll.Created, " &_
				"		SUM(tblPollAnswer.Votes) Responses " &_
				"FROM	tblPoll " &_
				"JOIN	tblPollAnswer ON tblPollAnswer.PollID = tblPoll.PollID AND tblPollAnswer.Archive = 0 " &_
				"WHERE	tblPoll.Active <> 0 " &_
				"AND	tblPoll.Archive = 0 " &_
				"AND	tblPoll.PollID <> " & Application("POLLID") & " " &_
				"GROUP BY tblPoll.PollID, tblPoll.Question, tblPoll.Created " &_
				"ORDER BY tblPoll.Created DESC" & adoTop2(50)
	Set rsPoll = adoOpenRecordset(sStat)

	' build the list of archived poll questions
	nCount = 0
	Do Until rsPoll.EOF
		sHTML = sHTML & "<TR CLASS=""list" & (nCount Mod 2) & """>" & vbCrLf &_
			"	<TD><a href=""detail.asp?pollid=" & rsPoll.Fields("PollID").Value & """>" & Server.HTMLEncode(rsPoll.Fields("Question").Value & "") & "</A></TD>" & vbCrLf &_
			"	<TD ALIGN=center>" & rsPoll.Fields("Responses").Value & "</TD>" & vbCrLf &_
			"	<TD>" & adoFormatDateTime(rsPoll.Fields("Created").Value, vbShortDate) & "</TD>" & vbCrLf &_
			"</TR>" & vbCrLf
		rsPoll.MoveNext
		nCount = nCount + 1
	Loop
	rsPoll.Close
	Set rsPoll = Nothing
	Application("POLLARCHCACHE") = sHTML
	Application("POLLARCHCACHEREFRESH") = Now
End Sub
%>