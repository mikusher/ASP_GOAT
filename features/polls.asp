<!-- #include file="../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' polls.asp
'	Describes the poll feature for ASP Nuke.
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
<!-- #include file="../header.asp" -->

<p>
<font class="maintitle">Poll Feature</font><br>
<font class="articleauthor">Kenneth W. Richards</font><br>
<a href="http://www.orvado.com" class="tinytext">Orvado Technologies</a>

<p>
Polls are a fun feature that allows visitors to your site to vote on
a particular subject.  In the newspaper world, polls were made popular
by the USA Today newspaper.  They allow you to measure a general
consensus about a topic or issue that affects our world (and your
web site.
</p>

<h3>Creating Polls</h3>

<p>
Creating polls is a simple matter of defining a question and then adding
a selection of responses.  You may have as many responses as you like,
but as a general rule of thumb it is good to have between 4 and 6 choices.
</p>

<p>
By default, the poll module will display the most recent poll on the home
page (if ASP Nuke is configured to display the poll).  In fact, as you are
entering a new poll, you will actually see the poll being built within the
capsule.
</p>

<p>
You may edit both the poll question or any of the possible responses at any
time.  This is regardless of whether or not the poll has received any votes
or not.  You should be careful about doing this because it will invalidate
the authenticity of your poll.
</p>

<h3>Poll Results</h3>

<p>
When you view the poll results, you will see a "bar chart" created entirely
from HTML.  This gives you a visual representation of the votes received so
far.  Beside each bar, you will see a percentage.  This indicates the
percentage of votes received for the response with respect to the total
number of votes cast for the poll.
</p>

<p>
Below the results, There is a comment system very similar to the article
comment system.  It allows visitors to make a comment on the poll question,
any of the responses or the results so far.
</p>

<p>
Users may enter comments whether they are logged onto the site or not.  If
you enter a comment while logged in, your comment will be shown with your
registered username attributed to the remark.  Otherwise, your comment will
appear as being posted by user "Anonymous".
</p>

<p>
Comments will be moderated in the same fashion as the article comments.
This means that users with good karma (moderation points) will be able
to "mod" items up or down.  These "mod" points will combine to form a score
for the comment which determines whether the comment should be shown or
hidden.  The moderation system has not yet been completed, so this feature is not
yet functional.
</p>

<h3>Voting Restrictions</h3>

<p>
To ensure fair and balanced voting for the poll, we employ IP checks to
make sure that each person can only cast one vote.  You can configure
whether or not a person may cast one vote per poll, one vote per day or
one vote per week.
</p>

<p>
Each time a person votes, their IP is logged to the database.  Before a
vote is cast, the system will check the log and make sure the person is
allowed to cast a vote.  If they are not allowed, the system will simply
show the poll results with a message on the top indicating that their vote
has been cast already.
</p>

<p>
<font class="tinytext">Last Updated: Oct 13, 2003</font>
</p>
<!-- #include file="../footer.asp" -->
