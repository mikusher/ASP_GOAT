<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../header_popup.asp" -->

<!-- SECTION_START:Overview -->
<h3>Members Help</h3>

<p>
Members are the people who visit your site and register using the member
registration module.  This is not to be confused with Users who are site
administrators.  Through this administration interface, you can add, edit,
delete and search for members registered through ASP Nuke.
</p>

<p>
The list of members is broken up into pages and sorted based on the date
they registed with ASP Nuke.  In the future we may add the ability to sort
on other criteria.
</p>

<p>
Until we finish the internationalization of the site, the member registration
will not require members to enter the <kbd>State</kbd> or <kbd>Country</kbd>
fields.  For this reason, you may see a lot of blank entries in the member
listing.
</p>

<h4>Searching for Members</h4>

<p>
To search for members, simply enter the keywords in the search fields above
the member listing.  When you click <kbd>Submit Query</kbd>, the system will
bring back a list of all members matching your criteria (paging the results
if necessary.)
</p>

<p>
These search criteria will combine so that if you enter
"John" for first name and "Smith" for last name, it will only match members
whose first name matches "John" AND last name matches "Smith".  Since this is
a keyword search, it would match the name "LaJohnson Mcsmithy".  Also note
that this keyword searching is not case sensative.
</p>

<h4>Member Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>First Name</td>
	<td class="formd">The first name for the member</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Last Name</td>
	<td class="formd">The last name for the member</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Username</td>
	<td class="formd">This is the username the member uses to login</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Password</td>
	<td class="formd">This is the password the member uses to login</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Confirm Password</td>
	<td class="formd">You must confirm the password entered to ensure that it is correct</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Address (Line 1)</td>
	<td class="formd">The mailing (street) address for this member</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Address (Line 2)</td>
	<td class="formd">Secondary line for additional mailing address information</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>City</td>
	<td class="formd">Town or city where the member lives</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>State</td>
	<td class="formd">State / Province where the member lives</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Zip Code</td>
	<td class="formd">Zip or Postal code where the member lives</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Country</td>
	<td class="formd">Nation where the member lives</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>E-Mail Address</td>
	<td class="formd">E-mail address (required for the account activation e-mail)</td>
</tr>
</table>
</p>

<!-- SECTION_END:Overview -->

<!-- SECTION_START:Country -->

<h3>Country Administration</h3>

<p>
Countries are used in registration to indicate the nation of origin
for your members.  This can be used for a mailing address or purely
to track the location of your valued members.
</p>

<p>
The ASP Nuke application comes pre-seeded with a list of civilized
countries which are the most likely to have Internet access.  You
should have all of the developed countries in the list from which you
can delete those that you don't need.
</p>

<h4>Article Author Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Country Name</td>
	<td class="formd">The official name of the country</td>
</tr>
</table>
</p>

<!-- SECTION_END:Country -->

<!-- SECTION_START:State -->

<h3>State Administration</h3>

<p>
States are used in registration to indicate the state or province of origin
for your members.  This can be used for a mailing address or purely
to track the location of your valued members.
</p>

<p>
Some countries are associated with their own unique set of states or provinces
while other countries don't use states at all.  This will be handled later as
we add internationalization to the ASP Nuke application.
</p>

<h4>State Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>State Code</td>
	<td class="formd">A short abbreviation for the state that is commonly used in mailing addresses.  In the U.S. this is a two letter code such as <kbd>NY</kbd>.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>State Name</td>
	<td class="formd">The full name for the state or province</td>
</tr>
</table>
</p>

<!-- SECTION_END:State -->

<!-- #include file="../../../../footer_popup.asp" -->