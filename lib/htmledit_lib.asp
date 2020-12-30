<%
'--------------------------------------------------------------------
' htmledit_lib.asp
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
<script language="Javascript" type="text/javascript">
<!-- // hide
	function ValidatePostComment(formedit)
		{
			if(formedit.PostedBy.value =="")
				{
				alert("Please Enter Your Name.");
				formedit.PostedBy.focus();
				return(false);
				}
			if(formedit.ReplySubject.value =="")
				{
				alert("Subject Cannot Be Empty.");
				formedit.ReplySubject.focus();
				return(false);
				}
			if(formedit.htmledit.value =="")
				{
				alert("The Comment Field Cannot Be Empty. ");
				formedit.htmledit.focus();
				return(false);
				}              
			return (true);  
		}
	function ValidateLinkCat(formedit)
		{
			if(formedit.LinkCatTitle.value =="")
				{
				alert("Please A Link Category");
				formedit.LinkCatTitle.focus();
				return(false);
				}            
			return (true);  
		}
	function ValidateEnterLink(formedit)
		{
			if(formedit.LinkCatID.value =="")
				{
				alert("Please Enter a Link Category");
				formedit.LinkCatID.focus();
				return(false);
				}
			if(formedit.LinkTitle.value =="")
				{
				alert("Please Enter a Link Title");
				formedit.LinkTitle.focus();
				return(false);
				}
			if(formedit.LinkURL.value =="")
				{
				alert("Please A Link URL");
				formedit.LinkURL.focus();
				return(false);
				}           
			return (true);  
		}
	function ValidateContentCat(formedit)
		{
			if(formedit.ContentCatName.value =="")
				{
				alert("Please Enter A Content Category");
				formedit.ContentCatName.focus();
				return(false);
				}            
			return (true);  
		}
	function ValidatePicCat(formedit)
		{
			if(formedit.PicCatName.value =="")
				{
				alert("Please Enter A Pic Category");
				formedit.PicCatName.focus();
				return(false);
				}            
			return (true);  
		}
	function ValidateLogin(formedit)
		{
			if(formedit.UserName.value =="")
				{
				alert("Please Enter Your User Name");
				formedit.UserName.focus();
				return(false);
				}
			if(formedit.UserPass.value =="")
				{
				alert("Please Enter Your Password");
				formedit.UserPass.focus();
				return(false);
				}
			return (true);  
		}
	function ValidateUserReg(formedit)
		{
			if(formedit.UserName.value =="")
				{
				alert("Please Enter A User Name");
				formedit.UserName.focus();
				return(false);
				}
			if(formedit.UserPass.value =="")
				{
				alert("Please Enter A Password");
				formedit.UserPass.focus();
				return(false);
				}
			if(formedit.ConfirmUserPass.value =="")
				{
				alert("Please Confirm Password");
				formedit.ConfirmUserPass.focus();
				return(false);
				}
			if(formedit.UserPass.value != formedit.ConfirmUserPass.value)
				{
				alert("Passwords Do Not Match");
				formedit.UserPass.focus();
				return(false);
				}
			return (true);  
		}
	function DoSmilie(addSmilie) {
		var revisedMessage;
		var currentMessage = document.formedit.htmledit.value;
		revisedMessage = currentMessage+addSmilie;
		document.formedit.htmledit.value=revisedMessage;
		document.formedit.htmledit.focus();
		return;
	}
	
	function CheckAll()
	  {
	  for (var i=0;i<document.SendEmail.elements.length;i++)
	    {
	    var e = document.SendEmail.elements[i];
	    if (e.name != 'allbox')
	      e.checked = document.SendEmail.allbox.checked;
	    }
	  }
	function rollon(a) {
		a.style.backgroundColor='ECECEC';
		a.style.border = '#663333 solid 1px';
		a.style.cursor = 'default';
	}	
	function rolloff(a) {
		a.style.backgroundColor='#FFFFFF';	
		a.style.border = '#FFFFFF solid 1px'; 
	}
	
	function getText() {
		if (document.formedit.htmledit.createTextRange && document.formedit.htmledit.caretPos) {
			return document.formedit.htmledit.caretPos.text;
		} else {
			return '';
		}
	}
	
	function storeCaret(ftext) {
		if (ftext.createTextRange) {
			ftext.caretPos = document.selection.createRange().duplicate();
		}
	}
	
	function AddText(NewCode) {
		if (document.formedit.htmledit.createTextRange && document.formedit.htmledit.caretPos) {
			var caretPos = document.formedit.htmledit.caretPos;
			caretPos.text = NewCode;
		} else {
			document.formedit.htmledit.value+=NewCode;
		}
		document.formedit.htmledit.focus();
	}
	
	function bold() {
		var text = getText();
		if (text) {
		txt=prompt("Text to be made BOLD.",text);
		} else {
			txt=prompt("Text to be made BOLD.","Text");
		}
		if (txt!=null) {
			AddTxt="[b]"+txt+"[/b]";
			AddText(AddTxt);
		}
	}
	
	function italicize() {
		var text = getText();
		if (text) {
			txt=prompt("Text to be italicized",text);
		} else {
			txt=prompt("Text to be italicized","Text");
		}
		if (txt!=null) {
			AddTxt="[i]"+txt+"[/i]";
			AddText(AddTxt);
		}
	}
	
	function underline() {
		var text = getText();
		if (text) {
			txt=prompt("Text to be Underlined.",text);
		} else {
			txt=prompt("Text to be Underlined.","Text");
		}
		if (txt!=null) {
			AddTxt="[u]"+txt+"[/u]";
			AddText(AddTxt);
		}
	}
	
	function strike() {
		var text = getText();
		if (text) {
			txt=prompt("Text to be stricken.",text);
		} else {
			txt=prompt("Text to be stricken.","Text");
		}
		if (txt!=null) {
			AddTxt="[s]"+txt+"[/s]";
			AddText(AddTxt);
		}
	}
	
	function hr() {
		var text = getText();
		AddTxt="[hr]" + text;
		AddText(AddTxt);
	}
	
	function hyperlink() {
		var text = getText();
		txt2=prompt("Text to be shown for the link.\nLeave blank if you want the url to be shown for the link.","");
		if (txt2!=null) {
			txt=prompt("URL for the link.","http://");
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[url="+txt+"]"+txt+"[/url]";
					AddText(AddTxt);
				} else {
					AddTxt="[url="+txt+"]"+txt2+"[/url]";
					AddText(AddTxt);
				}
			}
		}
	}
	
	function email() {
		txt2=prompt("Enter the complete email address.","");
		AddTxt="[email]" + txt2 + "[/email]";
		AddText(AddTxt);
	}
	
	function image() {
		var text = getText();
		txt=prompt("URL for graphic","http://");
		if(txt!=null) {
			AddTxt="[image]"+txt+"[/image]";
			AddText(AddTxt);
		}
	}
	var win= null;
	function NewWindow(mypage,myname,w,h,scroll,rs){
	  	var winl = (screen.width-w)/2;
		var wint = (screen.height-h)/2;
		var settings  ='height='+h+',';
		    settings +='width='+w+',';
		    settings +='top='+wint+',';
		    settings +='left='+winl+',';
		    settings +='scrollbars='+scroll+',';
		    settings +='resizable='+rs+'';
		win=window.open(mypage,myname,settings);
		if(parseInt(navigator.appVersion) >= 4){win.window.focus();}
	}
// unhide -->
</script>
<%
'----------------------------------------------------------------------------
' HTMLCommandButtons
'	Output the command buttons for editting HTML content

Sub HTMLCommandButtons
	With Response
		'format buttons
		.write("<i>Format:</i>&nbsp;&nbsp;&nbsp;")
		.write("<a href=""Javascript:bold();""><img src=""" & Application("ASPNukeBasePath") & "img/htmleditor/icon_editor_bold.gif"" width=""22"" height=""22"" alt=""Bold"" border=""0""></a>" & vbcrlf)
		.write("<a href=""Javascript:italicize();""><img src=""" & Application("ASPNukeBasePath") & "img/htmleditor/icon_editor_italicize.gif"" width=""23"" height=""22"" alt=""Italicized"" border=""0""></a>" & vbcrlf)
		.write("<a href=""Javascript:hyperlink();""><img src=""" & Application("ASPNukeBasePath") & "img/htmleditor/icon_editor_url.gif"" width=""22"" height=""22"" alt=""Insert Hyperlink"" border=""0""></a>" & vbcrlf)
		.write("<a href=""Javascript:email();""><img src=""" & Application("ASPNukeBasePath") & "img/htmleditor/icon_editor_email.gif"" width=""23"" height=""22"" alt=""Insert Email"" border=""0""></a>" & vbcrlf)
		.write("<a href=""Javascript:underline();""><img src=""" & Application("ASPNukeBasePath") & "img/htmleditor/icon_editor_underline.gif"" width=""23"" height=""22"" alt=""Underline Text"" border=""0""></a>" & vbcrlf)
		.write("<a href=""Javascript:image();""><img src=""" & Application("ASPNukeBasePath") & "img/htmleditor/icon_editor_image.gif"" width=""23"" height=""22"" alt=""Insert Image"" border=""0""></a>" & vbcrlf)
		.write("<a href=""Javascript:hr();""><img src=""" & Application("ASPNukeBasePath") & "img/htmleditor/icon_editor_hr.gif"" width=""23"" height=""22"" alt=""Horizontal Rule"" border=""0""></a>" & vbcrlf)
		.write("<a href=""Javascript:strike();""><img src=""" & Application("ASPNukeBasePath") & "img/htmleditor/icon_editor_strike.gif"" width=""23"" height=""22"" alt=""Strike Through"" border=""0""></a>" & vbcrlf)
		.write("<a href=""JavaScript:NewWindow('" & Application("ASPNukeBasePath") & "module/discuss/forum/graemlins.asp','name','550','550','yes','1')""><img src=""" & Application("ASPNukeBasePath") & "img/htmleditor/icon_editor_smilie.gif"" width=""22"" height=""22"" alt=""Insert Smilie"" border=""0""></a>" & vbcrlf)
	End With
End Sub
%>
