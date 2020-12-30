<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../header_popup.asp" -->

<!-- SECTION_START:Overview -->
<h3>Poll Help</h3>

<p>
The polls are quick little surveys which appear in a capsule on the side
of the page. They ask users to answer a simple question and then report
on how people have responded to the questions using a simple "bar chart".
</p>

<p>
Our polls also allow registered members to post comments to a poll question
by simply entering a subject and message body.  The comment system follows
a hierarchical list so that replies are indented and appears like a "tree
control" with folders, nodes and documents.
</p>

<p>
You do not need to be registered to enter comments regarding a poll.  Once
a poll comment has been posted, it will show up immediately.  No HTML or
scripting code is allowed in the comment.  Any attempt to input this code
will be removed automatically when the user goes to post the comment.
</p>

<!-- SECTION_END:Overview -->

<!-- SECTION_START:Poll -->

<h3>Poll Help</h3>

<p>
Polls are the survey questions which appear on the side of the screen
and ask the user to answer a simple question.  Or, if the user prefers,
they can bypass the answering of the question and go straight to 
viewing the results.
</p>

<p>
You can create as many polls as you like in the system.  Only the most
recent poll will be displayed in the poll capsule.  As soon as you post
a new poll, the previous one will be archived and no further responses
will be accepted for it.
</p>

<p>
Your question should be short and to-the-point and relevant to the
content of your site.  If you make your poll question too long then it
will be hard to read and may take up too much space in the capsule areas.
</p>

<h4>Poll Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Question</td>
	<td class="formd">The poll question that will be put to the visitors of your ASP Nuke site.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Poll -->

<!-- SECTION_START:Answers -->

<h3>Poll Answers Help</h3>

<p>
Associated with each poll question are a series of answers.  These are the
multiple choice responses a user may choose from when answering the poll.
All of the choices are mutually exclusive meaning that the user can only
choose one option from the list.
</p>

<p>
There is no limit to the number of answers you can associate with a poll
question.  Generally, you will want to keep the number between 4 and 8 to
give a good selection of choices.  You will also want to keep the answers
relatively short so that they render well within many different browsers.
</p>

<p>
When a visitor responds to a poll question, their vote is recorded along with
their referring IP (Internet Protocol) address.  This is used to prevent
someone from "stuffing" the ballot box by repeatedly clicking an answer.  The
IP address restriction prevents a visitor from casting more than one vote
per day from any given machine.
</p>

<h4>Answer Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Answer</td>
	<td class="formd">This is the text for the answer that is displayed in the capsule and the "bar chart" results.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Answers -->

<!-- #include file="../../../../footer_popup.asp" -->