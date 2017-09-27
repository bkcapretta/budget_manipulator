# Budget Manipulator
By: Bianca Capretta

# Purpose
I created an interface on Google Spreadsheets to handle a multi-faceted budget and 
coded its implementation using Google Apps' script editor. This 
implementation allows a user to manipulate the data (the budget, percent effort 
allocated per month, money allocated per month, money allocated semi-monthly, etc)
in one cell and have it immediately update the data in the rest of the spreadsheet.

A specific office on Tufts University's campus with many grants each summer asked
me to put together this tool to make their lives a bit easier (because doing 
this task on paper wasn't the most efficient method).

Link to Spreadsheet: https://docs.google.com/spreadsheets/d/1qVG3xqOvEDnbB6sZWrl_2uIxxsNrE8mV3prYlwbaaRs/edit?usp=sharing

# Example 
Let's say I have $30,000 budgeted for the summer and four grants during 
that block of time (Lego Project, Summer Workshops, Interlace Project, and STOMP).
I want to allocate 20%, 30%, 15%, and 35% of my resources respectively. First, 
this sheet will immediately calculate the corresponding monthly rate ($) and
semi-monthly rate. If I want to adjust the numbers a bit and decide that I'd like
to spend more than $4,500 on the Interlace Project, then I can edit that cell
to $5,000, and every other cell on the spreadsheet will change accordingly.
Less money will be left for the other projects so it will need to be taken
out from their individual budget's. When taking money away from other projects,
the algorithm favors the small projects by taking more from the bigger projects. 
When adding money to the other projects, the bigger grants are prioritized to 
get more money back.