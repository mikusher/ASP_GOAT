<%
'--------------------------------------------------------------------
' capsule.asp
'	Create the login capsule which will appear on all pages of
'	the site.
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

<FORM METHOD="post" ACTION="<%= Application("ASPNukeBasePath") %>account/login.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="login">
<%= Replace(Application("ModCapHeader"), "<CAPSULETITLE/>", "Member Login") %>
<%= Application("ModCapLeft") %>
<% If Request.Cookies("Username") = "" Then %>

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 ALIGN="center" WIDTH="100%">
<TR>
	<TD CLASS="forml" ALIGN="left">Username<BR>
	<INPUT TYPE="text" NAME="username" VALUE="<%= steEncForm("username") %>" SIZE="12" MAXLENGTH="16" class="form" style="width:120px">
	</TD>
</TR><TR>
	<TD CLASS="forml" ALIGN="left">Password<BR>
	<INPUT TYPE="password" NAME="password" VALUE="" SIZE="12" MAXLENGTH="16" class="form" style="width:120px">
	</TD>
</TR><TR>
	<TD ALIGN="center"><INPUT TYPE="submit" NAME="_submit" VALUE=" Login " class="form"></TD>
</TR>
</TABLE>

<P ALIGN="center" CLASS="tinytext">
Don't have an account yet, <A HREF="../../account/register.asp" CLASS="tinytext">register now</A> it's free!<br>
<A HREF="../../account/forgot_password.asp" CLASS="tinytext">I forgot my password</A>
</P>

<% Else %>
<P ALIGN="center" CLASS="tinytext">
You are currently logged in as <B><%= Request.Cookies("Username") %></B><BR><BR>
<INPUT TYPE="button" NAME="_logoff" VALUE=" Logoff " onClick="location.href='/account/login.asp?action=logoff'" class="form">
</P>
<% End If %>
<%= Application("ModCapRight") %>
<%= Application("ModCapFooter") %>
</FORM>