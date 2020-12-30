		<!-- END OF CONTENT -->
		</TD>
		<TD VALIGN="top" style="width:140px">
		<% modShowGroup "RGHT" %>

		<p align="center">
		 <a href="http://jigsaw.w3.org/css-validator/">
		  <img style="border:0;width:88px;height:31px" src="<%= Application("ASPNukeBasePath") %>img/vcss.gif" alt="Valid CSS!"></a><br>
    	  <a href="http://validator.w3.org/check/referer">
		  <img border="0" src="<%= Application("ASPNukeBasePath") %>img/valid-html401.gif" alt="Valid HTML 4.01!" height="31" width="88"></a>
		</p>
		</TD>
	</TR>
	</TABLE>
	</TD>
</TR><TR>
	<TD class="orange"><IMG SRC="<%= Application("ASPNukeBasePath") %>img/pixel.gif" WIDTH=1 HEIGHT=1 ALT=""></TD>
</TR><TR>
	<TD class="white"><% Server.Execute(Application("ASPNukeBasePath") & "module/other/randomquote/content.asp") %></TD>
</TR><TR>
	<TD class="orange"><IMG SRC="<%= Application("ASPNukeBasePath") %>img/pixel.gif" WIDTH=1 HEIGHT=1 ALT=""></TD>
</TR><TR>
	<TD class="white" align="center">
		<FONT CLASS="copytext">&copy; 2002 <A HREF="http://www.orvado.com">Orvado Technologies</A>, All Rights Reserved</FONT>
		- <a href="<%= Application("ASPNukeBasePath") %>sitemap.asp" class="commentlink">Site Map</a>
		- <a href="http://jigsaw.w3.org/css-validator/check/referer" class="commentlink">CSS</a>
		- <a href="http://validator.w3.org/check/referer" class="commentlink">HTML</a>
		<% If steTimer <> Empty Then %>- <FONT CLASS="copytext"><%= FormatNumber((Timer - steTimer) * 1000, 4) %> msec</font><% End If %>
	</TD>
</TR><TR>
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

</BODY>
</HTML>