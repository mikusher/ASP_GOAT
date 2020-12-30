// dynamicmenu.js
// Copyright (C) 2002 Orvado Technologies (http://www.orvado.com)

var timerID;
var menuName;

// retrieve an element from the DOM by name (for IE 5)
function getItem(sName) {
	var elm;
	if (document.getElementById) {
		elm = document.getElementById(sName);
	} else if (eval('document.' + sName)) {
		elm = eval('document.' + sName);
	} else {
		elm = eval('document.all.' + sName);
	}	
	return elm;
}

// build all of the menu items from the string sMenu
function menuBuildAll(sMenu) {
	var aMenu, aField, sMenu, sID, sMenuAll;
	var aMain = sMenu.split('~');

	sMenuAll = '';
	for (var i=0; i < aMain.length; i++) {
		aMenu = aMain[i].split('^');
		sMenu = '';
		for (var j=1; j < aMenu.length; j++) {
			aField = aMenu[j].split('|');
			if (aField.length == 2) {
				sMenu = sMenu + '<TR><TD class="menuoption" onclick="location.href=\'' + aField[1] + '\'" onMouseOver="this.className=\'menuoptionsel\'" onMouseOut="this.className=\'menuoption\'"><a href="' + aField[1] + '" class="menuoption">' + aField[0] + '</a></TD></TR>\n';
			}
		}
		// add the menu to the menu group
		aField = aMenu[0].split('|');
		sID = aField[0].replace(' ', '');
		//  onMouseOut="setTimeout(\'menuHide(\\\'' + sID + '\\\')\', 1000);"
		sMenuAll = sMenuAll + '<TD><DIV CLASS="mainmenu" onMouseOver="menuShow(\'' + sID + '\');">'
			+ aField[0] + '&nbsp;&nbsp;&nbsp;&nbsp;'
			+ '<DIV ID="'+sID+'" CLASS="menu" STYLE="visibility:hidden" onMouseOver="clearTimeout(timerID);" onMouseOut="menuHide(\'' + sID + '\');">\n'
			+ '<table border=0 cellpadding=0 cellspacing=0 width="100%">\n'
			+ sMenu
			+ '</table>\n'
			+ '</DIV>\n'
			+ '</DIV></TD>\n';
		// onMouseOut="menuHide(\'' + sID + '\')
	}
	document.write('<table border=0 cellpadding=0 cellspacing=0><tr>'
		+ sMenuAll
		+ '</tr></table>');
	return '';
}

// display a menu (when user rolls onto a menu)
function menuShow(sName) {
	var oMenu = getItem(sName);

	if (oMenu) {
		if ((oMenu.style) && (oMenu.style.visibility)) {
  			oMenu.style.visibility = 'visible';
			if (menuName != sName)
				menuHide(menuName);
			// timerID = setTimeout('menuHide(\'' + sName + '\')', 1000);
			menuName = sName;
		}
	}
	clearTimeout(timerID);
}

// hide a menu (when user rolls out of menu & submenu)
function menuHide(sName) {
	var oMenu = getItem(sName);

	if (oMenu) {
		if ((oMenu.style) && (oMenu.style.visibility))
  			oMenu.style.visibility = 'hidden';
	}
	clearTimeout(timerID);
}
