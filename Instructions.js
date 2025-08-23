function writePostSetupInstructions() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('Instructions');
  if (!sheet) {
    sheet = ss.insertSheet('Instructions');
  }
  setRichInstructions(
    sheet.getRange('A1'),
    `
      <h2>Post-Setup Instructions</h2>
      <li>First: make sure the <b>Levels</b> are set up the way you want.</li>
      <li>Next: populate the Students and Skills sheets</li>
      <li>Next: set up your Grade Book. You'll have to do some manual set-up to make it easy to assess either by student
      <b>or</b> by skill (see below).</li>
      <li>Finally, you can share a live view of each student's progress with students and/or families.</li>      
    `
  );
  setRichInstructions(
    sheet.getRange('A2'),
    `
    <h2>Setting up Levels</h2>
    <li>Define the different <b>Levels</b> you want to use.</li>
    <li>Map each <i>Skill</i> to the appropriate <u>Level</u>.</li>
    <li>Use the menu to adjust the <b>Level</b> settings as needed.</li>
    `
  );
  setRichInstructions(
    sheet.getRange('A3'),
    `<h2>Setting up Students and Skills</h2>
    <h4>Students</h4>
    <p>Make sure to include student emails if you're going to want to set up sharing.</p>
    <h5>Note:</h5>
    <p>You can add students at any time and run the "Add students" from the menu to keep adding more.</p>
    <h4>Skills</h4>
    <p>Define the different <b>skills</b> you want to grade. You can include a unit header and a skill number which
    will allow you to sort by unit or skill later when organizing your gradebook.</p>
    <p>You can add new skills at any time and run the "Add students and skills" from the menu to update your gradebook.</p>
    `
  );
  setRichInstructions(
    sheet.getRange('A4'),
    `<h2>Setting up the Grade Book</h2>
    <p>Note: when we set up the grade book, we set up one row per student per skill. So if you have 20 students and you assess
    20 different skills, that will be 400 rows of data!</p>
    <p>Obviously that's a lot to manage, but Spreadsheets make it easy!</p>
    <h3>Option 1: Google Sheets' new "Table" feature</h3>
    <p>Google doesn't yet let me write a script to set up a table for you automatically, but the table view will make
    it easy to toggle between different views. If you first convert your Gradebook to a table, you can then set up a
    "group by" view to group items by student. You can next set up a "group by" view to group items by skill. You can
    also set up "filter" views to view only one unit at a time.</p>
    <h3>Option 2: Filter Views</h3>
    <p>Filter views allow you to create different views of your data without changing the underlying data. You can create a filter view for each unit, for example, to focus on just the skills being assessed in that unit.</p>
    <p>To create a filter view, click on the "Data" menu, then "Filter views", and "Create new filter view". You can then set the filter criteria to only show the rows you want to see.</p>
    `
  );
  // set column width to 850
  sheet.setColumnWidth(1, 850);
}