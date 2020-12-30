<%
'--------------------------------------------------------------------
' cache.asp
'	Cache various components of the poll application.  This library
'	is included by the admin areas and used to rebuild the cache in
'	the event that the current poll is changed.
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

Sub modCapsuleCache(bForceReload)
	Dim rsPoll, sStat, sHTML, nPollID
	Dim rsAns		' answers for the poll question
	Dim sQuest		' question for the poll

	' refresh the poll questions every 30 minutes
	If IsDate(Application("POLLCACHEREFRESH")) And Not bForceReload Then
		If DateDiff("n", Application("POLLCACHEREFRESH"), Now()) < 30 Then Exit Sub
	End If

	nPollID = 0
	sStat = "SELECT	" & adoTop(1) & " tblPoll.PollID, tblPoll.Question " &_
				"FROM	tblPoll " &_
				"WHERE	tblPoll.Active <> 0 " &_
				"AND	tblPoll.Archive = 0 " &_
				"ORDER BY tblPoll.Created DESC" & adoTop2(1)
	Set rsPoll = adoOpenRecordset(sStat)

	' retrieve the list of possible answers for the poll
	If Not rsPoll.EOF Then
		nPollID = rsPoll.Fields("PollID").Value
		sQuest = rsPoll.Fields("Question").Value
		sStat = "SELECT tblPollAnswer.AnswerID, tblPollAnswer.Answer " &_
					"FROM	tblPollAnswer " &_
					"WHERE	PollID = " & nPollID & " " &_
					"AND	Active <> 0 " &_
					"AND	Archive = 0 " &_
					"ORDER BY OrderNo"
		Set rsAns = adoOpenRecordset(sStat)
	Else
		Exit Sub
	End If

	sHTML = "<P><B>" & sQuest & "</B></P>"
	Do Until rsAns.EOF
		sHTML = sHTML & "<INPUT TYPE=""radio"" NAME=""answerid"" VALUE=""" &_
			rsAns.Fields("AnswerID").Value & """ class=""formradio""> <font class=""tinytext"">" &_
			rsAns.Fields("Answer").Value & "</font><BR>"
		rsAns.MoveNext
	Loop
	Application("POLLID") = nPollID
	Application("POLLCACHE") = sHTML
	Application("POLLCACHEREFRESH") = Now
End Sub
%>