Completed:

- Aspen integration (Aspen.js)
- Sheet + Manager class for managing Aspen Grades works (AspenGrades.js)
- Sheet + Manager class for managing Aspen Assignments work (AspenAssignments.js)
- Aspen set-up works (AspenSetupUI.js, AspenSetup.js)

TO DO:

AspenGradeSync.js:

- Create function to add all skill items to Aspen Assignments Tab via a menu item (skill or unit). When finished, it should foreground the assignments tab and pop up a toast saying the user needs to fill in dueDate in order to create assignments.
- Update AspenAssignments.js &possibly Aspen.js to support assignDate parameter in addition to dueDate which gets passed to OneRoster API.
- Update AspenAssignments.js to have methods to create new assignments and to not fail if we have assignments
  that aren't created -- the sheet will contain a list of both created and un-created assignments. We can only CREATE assignments that have a dueDate. We should add conditional formatting to maybe put rows in dark grey and italic if they haven't been created yet: we could have three levels of formatting: Aspen ID exists = solid existing look; due date exists = ready to roll look (maybe highlighted); due date doesn't exist yet, grey and italic (i.e. not real).
- Create new "magic" skill called "Unit Avg" which can be turned on as well -- add function to "add unit averages to assignments tab" which will add a "Unit Avg" for each Unit (this probably would be used INSTEAD of the skills+grades, not in addition).

Once all this is done, we're ready to automate the grading part... so we would have...

- Create new function which gets grades from Grades sheet -- we build a "comment" our of each skill levels feedback string (i.e. basic/intermediate/advanced checkmarks would become a comment with the skill level a colon and then the skill string), and then we pull the grade from the mastery grade column.
- For each grade/comment, we call the maybeSync function to "maybe Sync" that grade.
- Create new function to handle our magic "Unit Avg" skill
- For each assignment with type "Unit Avg." we are going to calculate the Unit Average and we'll build a comment which consists of scores on each sub-skill, so the comment would look like "Skill: Description - 3, Skill: Description - 4, etc." with newlines between each skill -- so basically we get a big block-o-text for our comment with full data. Then we just post the grade/average with maybeSync as per usual.

NOTE: All of this should be designed with the concept that we read the _whole_ grade sheet ONE TIME and then do all this magic -- we don't want to be in the business of getting bogged down in many spreadsheet reads.
