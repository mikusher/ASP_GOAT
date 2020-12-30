<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' graemlins.asp
'	This library of functions is useful for creating html edit boxes
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

<html>
<head>
	<script language="Javascript">
		function insertgraemlin(graemlin){
			window.opener.document.formedit.htmledit.value+=graemlin;
			window.close();
		}
	</script>
	<link rel="stylesheet" href="<%= Application("ASPNukeBasePath") %>css/style.css" TYPE="text/css">
	<meta name="author" content="Ken Richards" />
	<meta http-equiv="content-type" content="text/html; charset=iso-8859-1" />
	<meta name="robots" content="all" />
	<meta http-equiv="Pragma" content="no-cache" />
	<meta http-equiv="Expires" content="-1" />
</head>
<body>

<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
<TR>
	<TD>
	<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="100%">
	<TR>
		<TD VALIGN="top" ALIGN="left" STYLE="width:8px"><IMG SRC="<%= Application("ASPNukeBasePath") %>img/uleft.gif" WIDTH=8 HEIGHT=8 ALT=""></TD>
		<TD class="white"><IMG SRC="<%= Application("ASPNukeBasePath") %>img/pixel.gif" WIDTH=1 HEIGHT=1 ALT=""></TD>
		<TD VALIGN="top" ALIGN="right" STYLE="width:8px"><IMG SRC="<%= Application("ASPNukeBasePath") %>img/uright.gif" WIDTH=8 HEIGHT=8 ALT=""></TD>
	</TR>
	</TABLE>
	</TD>
</TR><TR>
	<TD class="white">

	<table border=0 cellpadding=8 cellspacing=0
	<tr>
		<td>
		<h4><% steTxt "Add a Graemlin to Your Post" %></h4>

		<p><% steTxt "Click on a graemlin icon below to add it to your post" %></p>

		<table border="0" cellpadding="3" cellspacing="0" width="100%">
		<tr>
			<td><a href="Javascript:insertgraemlin('[:)]');"><img border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Smile" %></p></td>
			<td><p class="bodytext">[:)]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:D]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile_big.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Big Smile" %></p></td>
			<td><p class="bodytext">[:D]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[8D]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile_cool.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Cool" %></p></td>
			<td><p class="bodytext">[8D]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:I]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile_blush.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Blush" %></p></td>
			<td><p class="bodytext">[:I]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:p]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile_tongue.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Tongue" %></p></td>
			<td><p class="bodytext">[:P]</p></td>
		 </tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[}:)]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile_evil.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Evil" %></p></td>
			<td><p class="bodytext">[}:)]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[;)]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile_wink.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Wink" %></p></td>
			<td><p class="bodytext">[;)]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:o)]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile_clown.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Clown" %></p></td>
			<td><p class="bodytext">[:o)]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[B)]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile_blackeye.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Black Eye" %></p></td>
			<td><p class="bodytext">[B)]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[8]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile_8ball.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Eightball" %></p></td>
			<td><p class="bodytext">[8]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:(]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile_sad.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Frown" %></p></td>
			<td><p class="bodytext">[:(]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[8)]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile_shy.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Shy" %></p></td>
			<td><p class="bodytext">[8)]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:0]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile_shock.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Shocked" %></p></td>
			<td><p class="bodytext">[:O]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:(!]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile_angry.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Angry" %></p></td>
			<td><p class="bodytext">[:(!]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[xx(]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile_dead.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Dead" %></p></td>
			<td><p class="bodytext">[xx(]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[|)]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile_sleepy.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Sleepy" %></p></td>
			<td><p class="bodytext">[|)]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:X]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile_kisses.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Kisses" %></p></td>
			<td><p class="bodytext">[:X]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[^]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile_approve.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Approve" %></p></td>
			<td><p class="bodytext">[^]</p></td>
		 </tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[V]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile_dissapprove.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Disapprove" %></p></td>
			<td><p class="bodytext">[V]</p></td>
		 </tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[?]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/icon_smile_question.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Question" %></p></td>
			<td><p class="bodytext">[?]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:boxing:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/boxing.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Boxing" %></p></td>
			<td><p class="bodytext">[:boxing:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:crash:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/crash.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Crash" %></p></td>
			<td><p class="bodytext">[:crash:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:drool:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/drool.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Drool" %></p></td>
			<td><p class="bodytext">[:drool:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:drunk:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/drunk.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Drunk" %></p></td>
			<td><p class="bodytext">[:drunk:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:mwink:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/mwink.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Wink (Animated)" %></p></td>
			<td><p class="bodytext">[:mwink:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:nono:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/nono.gif"></a></td>
			<td><p class="bodytext"><% steTxt "No-No" %></p></td>
			<td><p class="bodytext">[:nono:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:pimp:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/pimp.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Pimp" %></p></td>
			<td><p class="bodytext">[:pimp:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:spank:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/spank.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Spank" %></p></td>
			<td><p class="bodytext">[:spank:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:sweat:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/sweat.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Sweat" %></p></td>
			<td><p class="bodytext">[:sweat:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:thefinger:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/thefinger.gif"></a></td>
			<td><p class="bodytext"><% steTxt "The Finger" %></p></td>
			<td><p class="bodytext">[:thefinger:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:2gunsfiring:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/2gunsfiring.gif"></a></td>
			<td><p class="bodytext"><% steTxt "2 Guns Firing" %></p></td>
			<td><p class="bodytext">[:2gunsfiring:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:angel:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/angel.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Angel" %></p></td>
			<td><p class="bodytext">[:angel:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:angry2:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/angry2.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Angry 2" %></p></td>
			<td><p class="bodytext">[:angry2:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:banana:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/banana.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Dancing Banana" %></p></td>
			<td><p class="bodytext">[:banana:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:beerchug:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/beerchug.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Beer Chug" %></p></td>
			<td><p class="bodytext">[:beerchug:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:birthday:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/birthday.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Birthday" %></p></td>
			<td><p class="bodytext">[:birthday:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:square:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/square.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Square" %></p></td>
			<td><p class="bodytext">[:square:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:bigeyes:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/bigeyes.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Big Eyes" %></p></td>
			<td><p class="bodytext">[:bigeyes:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:waving:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/waving.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Waving" %></p></td>
			<td><p class="bodytext">[:waving:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:eek:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/eek.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Eek" %></p></td>
			<td><p class="bodytext">[:eek:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:finger:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/finger.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Finger" %></p></td>
			<td><p class="bodytext">[:finger:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:freak:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/freak.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Freak" %></p></td>
			<td><p class="bodytext">[:freak:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:frustrated:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/frustrated.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Frustrated" %></p></td>
			<td><p class="bodytext">[:frustrated:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:hammer:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/hammer.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Hammer" %></p></td>
			<td><p class="bodytext">[:hammer:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:idea:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/idea.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Idea" %></p></td>
			<td><p class="bodytext">[:idea:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:looney:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/looney.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Looney" %></p></td>
			<td><p class="bodytext">[:looney:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:machinegun:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/machinegun.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Machinegun" %></p></td>
			<td><p class="bodytext">[:machinegun:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:newconfuse:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/newconfuse.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Confuse (New)" %></p></td>
			<td><p class="bodytext">[:newconfuse:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:nut:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/nut.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Nut" %></p></td>
			<td><p class="bodytext">[:nut:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:peek:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/peek.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Peek" %></p></td>
			<td><p class="bodytext">[:peek:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:pukey:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/pukey.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Pukey" %></p></td>
			<td><p class="bodytext">[:pukey:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:rocketlauncher:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/rocketlauncher.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Rocket Launcher" %></p></td>
			<td><p class="bodytext">[:rocketlauncher:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:rolleyes2:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/rolleyes2.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Roll Eyes 2" %></p></td>
			<td><p class="bodytext">[:rolleyes2:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:s:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/s.gif"></a></td>
			<td><p class="bodytext">S</p></td>
			<td><p class="bodytext">[:s:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:scared:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/scared.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Scared" %></p></td>
			<td><p class="bodytext">[:scared:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:sleep:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/sleep.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Sleep" %></p></td>
			<td><p class="bodytext">[:sleep:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:swear:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/swear.gif"></a></td>
			<td><p class="bodytext"><% steTxt "Swear" %></p></td>
			<td><p class="bodytext">[:swear:]</p></td>
		</tr>
		<tr>
			<td><a href="Javascript:insertgraemlin('[:what:]');"><img alt border="0" hspace="10" src="<%= Application("ASPNukeBasePath") %>img/graemlins/what.gif"></a></td>
			<td><p class="bodytext"><% steTxt "What" %></p></td>
			<td><p class="bodytext">[:what:]</p></td>
		</tr>
		</table>
		</td>
	</tr>
	</table>
	</td>
</tr><TR>
	<TD>
	<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="100%">
	<TR>
		<TD VALIGN="bottom" ALIGN="left" STYLE="width:8px"><IMG SRC="<%= Application("ASPNukeBasePath") %>img/bleft.gif" WIDTH=8 HEIGHT=8 ALT=""></TD>
		<TD class="white"><IMG SRC="<%= Application("ASPNukeBasePath") %>img/pixel.gif" WIDTH=100 HEIGHT=1 ALT=""></TD>
		<TD VALIGN="bottom" ALIGN="right" STYLE="width:8px"><IMG SRC="<%= Application("ASPNukeBasePath") %>img/bright.gif" WIDTH=8 HEIGHT=8 ALT=""></TD>
	</TR>
	</TABLE>
	</TD>
</TR>
</TABLE>

</body>
</html>