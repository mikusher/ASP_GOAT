<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../header_popup.asp" -->

<!-- SECTION_START:Overview -->
<h3>Users Help</h3>

<p>
Users are the people who administer your ASP Nuke application by logging
into the site administration area.  This is not to be confused with Members
who are visitors who come to your site and register using the registration
module.  Through this administration interface, you can add, edit,
delete and assign access permissions for these users.
</p>

<p>
Once you have created a user, you can assign him access rights by clicking
on the <i>rights</i> action link or by clicking on the <i>Rights</i> tab
when you go into to edit a user.  You can grant and revoke rights using
the simple user rights administration page just by clicking a button. 
</p>

<p>
For now, the access permissions are fairly basic, you just indicate
whether or not a user can access a module administration.  In the future,
we may add "finer grained" control to permit viewing of items, but not
deletion and so forth.
</p>

<!-- SECTION_END:Overview -->

<!-- SECTION_START:Users -->

<h3>Users Help</h3>

<p>
Each user represents an individual person who has access to the ASP Nuke
administration area.  Creating a user account allows them to login.  By
using the <i>Rights</i> tab you can assign permissions to a user which
will allow them access to administer the various modules.
</p>

<p>
You may create as many users as you like for your ASP Nuke web site.
Also, you can grant permissions and revoke them as easy as checking a
box in the user rights administration.
</p>

<h4>User Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Username</td>
	<td class="formd">The username needed to login to the ASP Nuke administration area.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Password</td>
	<td class="formd">The password required for login to the ASP Nuke administration area.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Confirm Password</td>
	<td class="formd">Repeat the password typed in the <kbd>Password</kbd> field to verify that it is correct.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>First Name</td>
	<td class="formd">The first name for this user</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Middle Name</td>
	<td class="formd">The middle name for this user</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Last Name</td>
	<td class="formd">The last name for this user</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Email Address</td>
	<td class="formd">E-mail address for this user (in case we need to contact them.)</td>
</tr>

<tr class="list1">
	<Td class="forml" valign="top" nowrap>Day Phone</td>
	<td class="formd">Daytime phone number where this person can be reached.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Evening Phone</td>
	<td class="formd">Evening phone number where this person can be reached.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Users -->

<!-- SECTION_START:Rights -->

<h3>User Rights Help</h3>

<p>
In this area, you define the user rights which will be assigned
to the user you selected.  You select a user by clicking on the
<i>rights</i> action link next to a user in the user list.  You can
also click on the <i>Rights</i> tab after you have gone in to edit
an existing user.
</p>

<p>
Simply check the box next to the access right to assign or revoke
permissons to the module.  The menu system that each administrative
user sees when they login to the ASP Nuke admin, depends on the
rights they have been assigned.  Also, a user will be forbidden to
go into an admin area if they type in the URL (without using the
menu.)
</p>

<p>
You might want to be careful about granting permissions to the <i>Users</i>
or <i>Access</i> module since they could basically take over your system and
lock out the main site operator.  There are ways to recover from a mistake
such as this but it requires knowledge of the ASP Nuke security system and
some manual database work.
</p>

<p>
If a box is disabled ("grayed out"), then it means that the access right
does not have the sub-permission assigned.  In other words - it is "not
applicable" to this access right.  If you are developing code and need to
enable this, you can do so in the "Access" area of the nuke admin.
</p>
<!-- SECTION_END:Rights -->

<!-- #include file="../../../../footer_popup.asp" -->