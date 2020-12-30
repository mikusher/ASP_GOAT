<!-- #include file="../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' articles.asp
'	Describes the article feature for ASP Nuke.
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
<font class="maintitle">Articles Feature</font><br>
<font class="articleauthor">Kenneth W. Richards</font><br>
<a href="http://www.orvado.com" class="tinytext">Orvado Technologies</a>
</p>

<h3>Introduction</h3>

<p>
The articles or news stories are the heart of the nuke application.  They
appear on the home page in the main (center) content area.  Generally speaking,
these are news stories which your readers can review and then comment on.
</p>

<p>
Whether you want to post partial stories and link to an external site containing
the full story (as Slashdot does) or if you want to post a whole article, you may
do so with the article feature.  It allows you to manage a large collection of
articles and will provide archiving and searching for visitors to your site.
</p>

<h3>Creating Articles</h3>

<p>
Before you start creating articles, you will want to create categories to help
you organize your articles.  When you read a large newspaper, you will notice
that it is divided into many sections.  Each section contains a specific "type"
or category of information.  You have the national news, local news, sports,
entertainment and classifieds.
</p>

<p>
By breaking your news articles down into categories, you allow your visitors
to view the specific information they are looking for and also avoid certain
types of articles they are not interested in.
</p>

<p>
Also, like other popular article-based sites, the home page has a sampling
of all different types of articles together.  The user must navigate to a
specific category to find news specific to that type.
</p>

<h3>Defining Authors</h3>

<p>
Every article in your system must have an author assigned to it.  This is the
person who wrote and/or submitted the article.  The administration interface
allows you to define as many articles as you like. Before you go to add a new
article, you must make sure the author exists and create it if you need to do
so.
</p>

<P>
The database will track as many authors as you want.  To assign an author to
a new article, you simply select the author from a drop-list control on the
web form.  The author name will show up directly under the article title
so your readers will know immediately the source of the story.
</p>

<h3>Archiving Articles</h3>

<p>
In order to keep a historic record of past news articles that were posted,
the article archive allows visitors to browse previous "issues".  We don't
actually keep track of issues in the sytem, instead, we just organize
articles on a month-by-month basis.
</p>

<p>
When viewing the article archive, a person can see an overview of all of the
previous months when articles where posted along with a count indicating how
many articles were posted during the month.  To view all articles posted
during the month, the user would simply click on the month name.
</p>

<p>
For some sites which post a large number of news articles, we might require
a finer grained breakdown for thhe archives.  We could break down the archives
on a week-by-week basis.  This has not been written yet, but should be fairly
simple to do.
</p>

<h3>Comment System</h3>

<p>
One of the major components of the article features is the ability for visitors
to login and post comments.  Anybody is allowed to post comments, not just
regular visitors to the site.  The comment system uses a "threaded discussion
forum" format similar to Slashdot.
</p>

<p>
As of this writing, the moderation system for the comments has not been written
yet, but should be out shortly.  This will allow respected members to gain
moderation points for contributing to our community.  They can then use these
points to "score" other people's comments.  This way, the article comments
become a "self-policing" system and we can weed out all of the negative,
offensive and off-topic remarks.
</p>

<h3>Summary</h3>

<p>
The article system provides a flexible system for managing a large collection
of news stories.  It provides your readers with a simple way to discuss news
stories with other people who have the same interests.
</p>

<p>
With an article search feature, categories and the article archive, you can
organize your stories in an intuitive way that allows your readers to easily
browse your content.  The administration interface makes it easy to add,
edit and delete articles to the system.
</p>

<p>
<font class="tinytext">Last Updated: Oct 01, 2003</font>
</p>
<!-- #include file="../footer.asp" -->