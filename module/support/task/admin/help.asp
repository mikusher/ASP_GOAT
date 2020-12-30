<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../header_popup.asp" -->

<!-- SECTION_START:Overview -->
<h3>Tasks Help</h3>

<p>
The task management module is a very basic module for managing basic
tasks and simple projects that need to be completed.  Currently, it
is setup so that only administrators can create and modify tasks and
visitors to the site can view the tasks and add comments.
</p>

<p>
This task management is suited to updating visitors to the status of
open-source development for a software application such as ASP Nuke.
It can also be used for any other application where you wish to
publish the status of tasks within your Nuke application.
</p>

<p>
You can create a dynamic list of priorities and statuses which indicate
the current state of a task as well as how important it is to be
completed.  This is great for the programmer who wants to keep track of
all of the development tasks remaining and also for the administrator
who wants to be updated on the current status of various projects.
</p>

<!-- SECTION_END:Overview -->

<!-- SECTION_START:Task -->

<h3>Task Help</h3>

<p>
The central element for the task manager is, no surprise, tasks!  These
are the individual chores that need to be done for a project.  If this
were a full project management module, you would also have the ability to
create projects and then group the tasks under projects.  That may come
sometime in the future, but for now all of the tasks are lumped together.
</p>

<p>
You will notice that when you go to create a task priority, you can assign
an "HTML color code" with each priority.  For quick visualization of the tasks,
each task will have a color background to indicate its priority.  The tasks you
see in the task list will be sorted by priority with the most important items
appearing at the top.
</p>

<h4>Task Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Title</td>
	<td class="formd">This is the title of the task that appears in the task list and as the main task heading.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Priority</td>
	<td class="formd">Indicate the urgency or importance of the task.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Status</td>
	<td class="formd">Indicates the current state of progress on completing this task.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Comments</td>
	<td class="formd">Comments that should be attached to the task.  These are visible to the visitor on the public-side meaning they are NOT "admin-only".</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Percent Complete</td>
	<td class="formd">Indicate the percentage of work done to fully complete the task (reach a version 1.0 (stable release) for a software task.)</td>
</tr>
</table>
</p>

<!-- SECTION_END:Task -->

<!-- SECTION_START:Priority -->

<h3>Priority Help</h3>

<p>
The task priorities indicate the urgency or importance of completing a
task.  The priority administration allows you to define as many priority
statuses that you need to classify all of your current tasks.
</p>

<p>
Each priority can have an HTML color code assigned to it.  This will be
used as a background color in the task overview (Tasks tab) to give a
quick visual representation of the items.  If you need more information
on these codes, please check out:
<a href="http://hotwired.lycos.com/webmonkey/reference/color_codes/">Web Monkey HTML Color Codes</A>.
</p>

<p>
You will notice that you can order your priorities using the "up" and "down"
action links.  You will want to order your priorities from lowest importance
to highest.  When choosing a priority to assign to an individual task, they
will appear in this order.  When viewing ALL of the tasks on the task overview
page, they will appear in reverse order so that the tasks with the most
important priorities appear on the top.
</p>

<h4>Task Priority Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Priority Name</td>
	<td class="formd">The title that this priority should be known as.</td>
</tr><tr class="list1">
	<Td class="forml" valign="top" nowrap>English Text</td>
	<td class="formd">HTML color code for displaying tasks in task overview (eg: #F0E0F0).</td>
</tr><tr class="list0">
	<Td class="forml" valign="top" nowrap>Comments</td>
	<td class="formd">Additional clarification of the purpose for this task priority.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Priority -->

<!-- SECTION_START:Status -->

<h3>Status Help</h3>

<p>
The status attributes for tasks allow you to assign a status which is
independent of the priority.  The combination of priority and status
is a very powerful tool to classify your tasks.  This allows you to
indicate a task is urgent, but "delayed because of legal issues".
</p>

<p>
You can also use the statuses to denote which project each task is
assigned to.  Although this is more of a work-around until projects
are integrated into the task manager.  If you want to do a combination
of both the project and status, you can make entries such as "Framework - Delayed".
</p>

<p>
Because the statuses are dynamic, you can define as many or as little
statuses as you need to facilitate the classification of all of your
project tasks.  The statuses will be ordered alphabetically when you need
to choose a status from the list (for creating/modifying a task.)
</p>

<h4>Task Status Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Status Name</td>
	<td class="formd">Indicate the title this task status should be known as.</td>
</tr><tr class="list1">
	<Td class="forml" valign="top" nowrap>Comments</td>
	<td class="formd">Enter a long description of the meaning of this status to clarify its usage.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Status -->

<!-- #include file="../../../../footer_popup.asp" -->