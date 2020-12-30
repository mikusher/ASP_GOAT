<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' help.asp
'	Display the help information for the discussion forums.
'	THIS FILE WILL AUTOMATICALLY BE PARSED INTO THE DATABASE
'	CHANGES TO THE FILE WILL BE DETECTED AND RELOADED AS NEEDED
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

<h3>Forum Help</h3>

<p>
The discussion forums provide a public bulletin board where people can
post message and discuss various topics.  This is similar to a real-time conversation
except that it takes place at a slower pace.
</p>

<h4>Forum Organization</h4>

<p>
The discussion forums are really a series of message boards, each one
centering around a different topic of conversation.  You can create as
many message boards as you need for your discussion forums.  This is
the top-most level of the forums area.
</p>

<p>
Underneath each topic is a list of message threads.  A thread is a top-level
message which can have zero or more replies attached to it.  When you go in
to view a topic, you will see a listing of all of the threads that have been
posted ordered in reverse chronological order based on when they were posted.
</p>

<p>
When the user goes in to view a message thread, they will see a list of
messages attached to that thread.  Although the messages are stored as a
hierarchical (tree-structure) list, they will be flattened out when displayed
to make reading of the messages easier.  In the future, the user will have
the option of viewing the messages indented (to reflect the tree structure)
or in a flat list.
</p>

<!-- SECTION_END:Overview -->
<!-- SECTION_START:Topic -->

<h3>Forum Topic Help</h3>

<p>
The discussion forums are really a series of message boards, each one
centering around a different topic of conversation.  You can create as
many message boards as you need for your discussion forums.  This is
the top-most level of the forums area.
</p>

<p>
Through the Forum Topic admin area, you can add, create, modify or delete
topics which will appear in your discussion forums.  The changes to a topic
will take effect immediately.  The threads and messages that are posted to
a forum topic will not be affected when changes to a topic are made.
</p>

<p>
Using the <B>Up</B> and <B>Down</B> links in the topic list, you can change
the order that the forum topics are displayed in.  This will affect both the
capsule and the main forum index page.
</p>

<p>
Be careful about deleting a topic.  If you delete a topic, it will be very
difficult to recover the messages.  The messages will still exist in the
database, but it will take some manual effort to repair the database
(because no such functionality exists in the administration area yet.)
</p>

<h4>Topic Properties</h4>

<p>
<table border=0 cellpadding=2 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Property</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top">Title</td>
	<td class="formd">The title for the forum as shown in the topic listing (both the capsule and the forum overview)</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top">Short Description</td>
	<td class="formd">Short explanation of the forum topic's intended usage.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Topic -->

<!-- SECTION_START:Threads -->

<h3>Forum Threads Help</h3>

<p>
Threads are the top-level posts to a forum topic that are displayed when
you got in and view a topic.  Basically, all the messages posted to a
forum topic that are NOT replies are top-level posts.  These are displayed
in reverse chronological order so that the most recent posts are shown
at the top of the page.
</p>

<p>
As far as the database is concerned, there is almost no difference between a
top-level post and a reply.  We do create a separate admin area for each to
make it easier for the administrator to understand the message structure.
</p>

<p>
You cannot change the order of the message threads yet because it is based on
the date the message was posted which cannot be modified through the admin
screens.
</p>

<h4>Thread Properties</h4>

<p>
<table border=0 cellpadding=2 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Property</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top">Parent Message</td>
	<td class="formd">Indicates the message that this thread is in response to.  Being a top-level thread, this must be set to the special value <i>New Thread</i></td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top">Username</td>
	<td class="formd">Indicates the username of the person who posted the thread.  You can actually change the author of the thread by typing a new username.  In the admin area, no password is required for the username.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top">Subject</td>
	<td class="formd">This is the main subject or title of the message that appears in the topic overview.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top">Message Body</td>
	<td class="formd">This is the content of the forum message.  Include two carriage returns to create a paragraph break or one carriage return to create a line break.  You may also use the special UBB codes through the icon toolbar to create special formatting of the message.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Threads -->

<!-- SECTION_START:Message -->

<h3>Forum Messages Help</h3>

<p>
Forum Messages include both the top-level threads and any associated
replies to a message.  The information in this help section applies to
both areas despite the fact that we have a separate section written to
document the forum topic administration.
</p>

<p>
You will see the list of messages for a forum topic when you click on
a thread to view it's messages.  The messages are listed in a hierarchical
(tree-like) fashion to indicate the relationship between messages and
their replies.
</p>

<p>
At the top is a drop list labeled <b>Thread Displayed</b> which allows you
to choose for which thread you want to view messages.  This list will only
show the threads for the forum topic you selected previously.  The threads
are listed in chronological order from newest to oldest.  Selecting a new
thread will automatically refresh the page and display the messages for
the selected thread.
</p>

<p>
When you go to edit a forum message, you can change the parent message for
a reply.  If you move a message in this manner, not only will the message
be moved, but any child replies will be moved along with it and stay
assigned to the message.  For this reason, changing the parent message is
more like moving the entire branch of messages than just editting the one
message.
</p>

<h4>Message Properties</h4>

<p>
<table border=0 cellpadding=2 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Property</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top">Parent Message</td>
	<td class="formd">Indicates the message that this message is in response to.  By changing the parent message, you are effectively moving the branch to a new location in the hierarchy.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top">Username</td>
	<td class="formd">Indicates the username of the person who posted the message.  You can actually change the author of the message by typing a new username.  In the admin area, no password is required for the username.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top">Subject</td>
	<td class="formd">This is the main subject or title of the message that appears in the thread overview.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top">Message Body</td>
	<td class="formd">This is the content of the forum message.  Include two carriage returns to create a paragraph break or one carriage return to create a line break.  You may also use the special UBB codes through the icon toolbar to create special formatting of the message.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Message -->

<!-- #include file="../../../../footer_popup.asp" -->