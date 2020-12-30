<% Option Explicit %>
<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' help.asp
'	Display the help file for the random quotes administration.
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
<!-- #include file="../../../../header_popup.asp" -->

<!-- SECTION_START:Overview -->
<h3>Random Quotes Help</h3>

<p>
The random quotes administration allows you to administer all of the
random quotes which appear on the site.  Usually, these will be shown on
the bottom of the page near the footer information.
</p>

<p>
You can put as many quotes as you like in the random quotes system and
ASP Nuke will cache all of the quotes in the application object saving
a big performance hit on the database.  Each time the user loads a new
page on the site, a new quote will be selected randomly to be displayed
at the bottom of the page.
</p>

<!-- SECTION_END:Overview -->

<!-- SECTION_START:Quote -->

<h3>Quote Help</h3>

<p>
The quote administration tab is where you define all of the quotes which
should appear randomly on each page throughout the site.  Please define
as many quotes as you would like using this administration.
</p>

<p>
It is safe to add and delete quotes as the ASP Nuke application is running.
The random quotes module will refresh its cache of quotes every so often
(time period is defined in the <kbd>capsule.asp</kbd> script) and update the
list of available quotes.
</p>

<h4> Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Quote</td>
	<td class="formd">This is the quote which are words that some person has spoken.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Attributed To</td>
	<td class="formd">The person who is attributed with saying the quote.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Quote -->

<!-- #include file="../../../../footer_popup.asp" -->