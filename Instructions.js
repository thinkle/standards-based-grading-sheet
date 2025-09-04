/* eslint-disable no-unused-vars */
/* global getCellUrl, getSheetUrl, setRichInstructions, SpreadsheetApp,
  SHEET_SYMBOLS, SHEET_LEVELSETTINGS, MENU
*/
/* exported writePostSetupInstructions */

let counters = {
  default: 1,
}
function step(counter = 'default') {
  if (!counters[counter]) {
    counters[counter] = 1;
  }
  let stepString = `${counters[counter]}`;
  counters[counter]++;
  return stepString;
}

const blurbs = {
  // Top navigation lives in row 1 across multiple columns (frozen)
  'TOC_ROW': 1,
  // Intro/overview blurb merged across all TOC columns
  'Intro': 'A2',
  // Section anchors start below the intro; keep these stable for links
  'About': 'A3',
  'Setup': 'A4',
  'Setup-Symbols': 'A5',
  'Setup-LevelSettings': 'A6',
  'Setup-Build-Grade-Sheet': 'A7',
  'StudentsAndStandards': 'A8',
  'StudentsAndStandards-Students': 'A9',
  'StudentsAndStandards-Standards': 'A10',
  'Gradebook': 'A11',
  'Sharing': 'A12'
}

function writePostSetupInstructions() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('Instructions');
  if (!sheet) {
    sheet = ss.insertSheet('Instructions');
  }
  // Rebuild layout from scratch for predictable anchors
  sheet.clear();

  // Configure a frozen top navigation row with one link per column
  const tocItems = [
    { key: 'About', label: 'About' },
    { key: 'Setup', label: 'Initial Setup' },
    { key: 'StudentsAndStandards', label: 'Students &amp; Standards' },
    { key: 'Gradebook', label: 'Using the Gradebook' },
    { key: 'Sharing', label: 'Student Views' }
  ];
  // Ensure we have enough columns; set some friendly widths
  const tocCols = tocItems.length;
  for (let c = 1; c <= tocCols; c++) {
    sheet.setColumnWidth(c, 220);
  }
  sheet.setFrozenRows(1);
  // Write one link per column in row 1
  tocItems.forEach((item, idx) => {
    const col = idx + 1;
    const anchorRange = sheet.getRange(blurbs[item.key]);
    const url = getCellUrl(anchorRange);
    setRichInstructions(
      sheet.getRange(1, col),
      `<div style="text-align:center"><a href="${url}"><b>${item.label}</b></a></div>`
    );
  });
  // Merge row 2 across the TOC columns for a big intro header/blurb
  for (let row = 2; row < 20; row++) {
    sheet.getRange(row, 1, 1, tocCols).merge();
  }


  // Overview and Quick Start
  setRichInstructions(
    sheet.getRange(blurbs['Intro']),
    `<h1>Standards-Based Grading Sheet - Complete Guide</h1>
    <p>This guide walks you through four stages. Use the links above to jump directly to a section.</p>`);
  setRichInstructions(
    sheet.getRange(blurbs['About']),
    `<h2>About This Tool</h2>
    <p>
    <i>This standards-based grading system was inspired by <a href="https://sites.google.com/brzmath.com/btcrubric">the BTC Rubric from Tim Brzenzinski et al</a>.</i>
    <small>BTC stands for "Building Thinking Classrooms", a book by Peter Liljedahl.</small>
    </p>
    <p>If you are familiar with that tool, you will see this tool uses the same grading philosophy and approach, but makes
    the following improvements:</p>
    <li>1. Easier typing of grades (typable symbols instead of a drop-down menu of checks)</li>
    <li>2. No tab-based management: enter all grades in one view that you can sort by skill or by student rather than clicking through
    one tab per student.</li>
    <li>3. Real-time feedback: give students a "live" view of information rather than requiring reports to be generated.</li>
    <li>4. Flexibility: we don't hard-code things like the number of assessments you need, the number of tries you need for something
    to count as mastery, or how mastery indicators tally up into a final "score" used for a gradebook -- all of this is left configurable
    for teacher use.</li>    
    `
  );

  // Stage 1: Initial Setup
  setRichInstructions(
    sheet.getRange(blurbs['Setup']),
    `<h2>Stage ${step()}: Initial Setup</h2>
    <h3>${MENU.TITLE} -> ${MENU.SETUP}</h3>
    <p>When you run the <b>${MENU.SETUP}</b> item in the <i>${MENU.TITLE}</i> Menu, it will
    automatically set up some sheets that allow you to reconfigure the grading system to fit your needs,
    including items such as what you type to mark student work and how the levels of mastery you assess
    map to grades in your system.
    </p>
    `
  );
  setRichInstructions(
    sheet.getRange(blurbs['Setup-Symbols']),
    `<h3>Step ${step('setup')}: Define Grading Marks</h3>
    <p>The <b>${SHEET_SYMBOLS}</b> sheet uses a three-column system that enhances teacher data entry based on educator feedback:</p>
    <li><b>Character:</b> What teachers type for easy data entry (pick something you can type easily!)</li>
    <li><b>Mastery:</b> Whether this counts toward mastery (1 = Yes, 0 = No)</li>
    <li><b>Symbol:</b> What displays in reports and student views</li>
    <p><i>Remember: students will need to show a "streak" of mastery in order to demonstrate proficiency
        and earn credit for each level.</i></p>
    <p>You can customize the symbols on the <b>${ss.getSheetByName(SHEET_SYMBOLS) ? `<a href="${getSheetUrl(ss.getSheetByName(SHEET_SYMBOLS))}">Symbols sheet</a>` : 'Symbols sheet'}</b>.</p>
    <h4>Default Symbol System:</h4>
    <p>When you first run the set-up, we will give you the default symbol system described in
    the BTC Grading Rubric with easy to type shorthands.</p>
    <h5>Successful Attempts (Count Toward Mastery):</h5>
    <li><code>C</code> <b>✔</b> - KDI (Knowledge Demonstrated Individually)</li>
    <li><code>Co</code> <b>✔o</b> - KDI via teacher Observation</li>
    <li><code>Cc</code> <b>✔c</b> - KDI via Conversation with teacher</li>
    <li><code>Cs</code> <b>✔s</b> - KDI with a Silly mistake not related to the objective</li>
    <h5>Learning Attempts (Do Not Count Toward Mastery):</h5>
    <li><code>H</code> <b>H</b> - Knowledge Demonstrated with Help from a teacher or peer</li>
    <li><code>G</code> <b>G</b> - Knowledge demonstrated in a Group setting</li>
    <li><code>X</code> <b>✗</b> - Question attempted but answered incorrectly</li>
    <li><code>Xo</code> <b>✗o</b> - Teacher observes incorrect attempt</li>
    <li><code>Xc</code> <b>✗c</b> - Incorrect response in Conversation with teacher</li>
    <li><code>PC</code> <b>PC</b> - Partially correct</li>
    <li><code>N</code> <b>N</b> - Not Attempted (score of 0)</li>    
    <p><b>Customization:</b> You can define your own symbols and meanings. The three-column approach allows teachers to type simple characters (like "check" or "x") while displaying meaningful symbols (✓ or ✗) in student reports.</p>        
    `
  );
  setRichInstructions(
    sheet.getRange(blurbs['Setup-LevelSettings']),
    `<h3>Step ${step('setup')}: Configure Your Levels</h3>
    <p>Go to the <b>${SHEET_LEVELSETTINGS}</b> sheet to define your mastery levels. The default setup includes:</p>
    <li><b>Basic (B):</b> Entry-level understanding - Score: 2, Requires: 2 consecutive successes</li>
    <li><b>Intermediate (I):</b> Developing proficiency - Score: 3, Requires: 2 consecutive successes</li>
    <li><b>Mastery (M):</b> Full mastery - Score: 4, Requires: 2 consecutive successes</li>  
    <h4>Customizing Levels:</h4>
    <li><b>Level Name:</b> What you call this level (e.g., "Basic", "Proficient", "Advanced")</li>
    <li><b>Short Code:</b> Letter used in gradebook columns (e.g., "B", "P", "A")</li>
    <li><b>Default Attempts:</b> How many attempt columns to create for this level in your gradebook sheet
    (default: 5; you <i>can</i> insert more if you want later).</li>
    <li><b>Streak for Mastery:</b> How many consecutive successes needed to demonstrate mastery (default: 2).</li>
    <li><b>Score:</b> The numeric grade assigned when this level is mastered (default values are set up
    for a 0-4 grading scale, but this can easily be tweaked to work with 0-100 or whatever system you use).</li>    
    `);

  setRichInstructions(
    sheet.getRange(blurbs['Setup-Build-Grade-Sheet']),
    `<h3>Step ${step('setup')}: Build your Grade Sheet</h3>
    <p>From the <b>${MENU.TITLE}</b> menu, select <b>${MENU.SETUP_GRADE_SHEET}</b> to create the initial gradebook structure.</p>
    <p>You can check if the grading columns you want have been set up before moving forward to enter your students and standards.</p>
    `
  )


  // Stage 2: Students and Standards
  setRichInstructions(
    sheet.getRange(blurbs['StudentsAndStandards']),
    `<h2>Stage 2: Setting up Students and Standards</h2>`
  );
  setRichInstructions(
    sheet.getRange(blurbs['StudentsAndStandards-Students']),
    `<h3>Adding Students</h3>
    <p>Go to the <b>Students</b> sheet and enter:</p>
    <li><b>Student Names:</b> Full names as you want them to appear in grade reports</li>
    <li><b>Student Emails:</b> Required for sharing individual student views with families</li>
    <p><b>Run ${MENU.TITLE} → ${MENU.ADD_STUDENTS_AND_SKILLS}</b> to add the necessary rows in the gradebook
    for new students at any time.</p>
    <h4>Important Notes:</h4>
    <li>Include student emails if you plan to share live views with students or families</li>
    <li>You can add students at any time - just run "Add Students &amp; Skills" from the menu to update the gradebook</li>
    <li>Student order will be preserved in the gradebook and student views</li>
    `);
  setRichInstructions(
    sheet.getRange(blurbs['StudentsAndStandards-Standards']),
    `<h3>Defining Skills and Standards</h3>
    <p>Go to the <b>Skills</b> sheet and define what you'll be assessing:</p>
    <li><b>Unit:</b> Organizing header (e.g., "Algebra", "Chapter 1", "Quarter 1")</li>
    <li><b>Skill #:</b> Number or code for sorting (e.g., "1.1", "A", "01")</li>
    <li><b>Skill Description:</b> Clear description of what students must demonstrate</li>    
    <h4>Best Practices for Skills:</h4>
    <li>Use clear, specific descriptions that students and parents will understand</li>
    <li>Group related skills under the same Unit for easier organization</li>
    <li>Use consistent numbering schemes within units</li>
    <li>Consider breaking complex standards into smaller, assessable skills</li>            
    <p>Run <b>${MENU.TITLE} → ${MENU.ADD_STUDENTS_AND_SKILLS}</b> to add new skills to your gradebook
    at any time.</p>    
    `
  );

  // Stage 3: Using the Gradebook
  setRichInstructions(
    sheet.getRange(blurbs['Gradebook']),
    `<h2>Stage 3: Using as a Gradebook</h2>

    <h3>Understanding the Gradebook Layout</h3>
    <p>The <b>Grades</b> sheet contains one row per student per skill. For 20 students and 20 skills, that's 400 rows of data! Each row includes:</p>
    <li><b>Student Info:</b> Name and Email</li>
    <li><b>Skill Info:</b> Unit, Skill #, and Description</li>
    <li><b>Mastery Grade:</b> Automatically calculated overall score</li>
    <li><b>Attempt Columns:</b> Individual attempts at each level (B1, B2, I1, I2, M1, M2, etc.)</li>
    
    <h3>Recording Student Attempts</h3>
    <p>To grade student work:</p>
    <li>Find the student-skill row you want to assess</li>
    <li>Find the column for the task you are assessing: for example "B3" for the student's
    third attempt to demonstrate "basic" level mastery for the skill.</li>
    <li>Type the character you indicated for their "mark", such as "1o" for <b>✓o</b> (success observed by teacher)
    or <b>Xs</b> for <b>✗s</b> (unsuccessful attempt, silly mistake). <i>You use whatever symbols you established in the Symbols
    sheet during set up.</i>
    </li>
    <li>The system automatically generates the student-friendly view and the calculated grade based on what you type.</li>
    <li><b>Note: you can hide columns in the Gradebook view to focus on data entry as needed!</b></li>
    <h3>Organizing Your Gradebook View</h3>
    <p>With hundreds of rows, organization is key! Use these strategies:</p>
    
    <h4>Option 1: Google Sheets Tables (Recommended)</h4>
    <li>Select your data range and choose <b>Format → Convert to table</b></li>
    <li>Create <b>Group by Student</b> view to see all skills for one student at a time</li>
    <li>Create <b>Group by Skill</b> view to see all students' progress on one skill</li>
    <li>Create <b>Group by Unit</b> view to focus on one unit at a time</li>
    <li><b>Note: Under "Table Options" -> "View Options" choose to "Hide group by aggregation" to make the view more compact.</b>
    </li>
    
    <h4>Option 2: Filter Views</h4>
    <li>Go to <b>Data → Filter views → Create new filter view</b></li>
    <li>Create filters for individual students, skills, or units</li>
    <li>Save multiple filter views for quick switching between perspectives</li>
    
    <h3>Sorting Strategies</h3>
    <li><b>By Student:</b> Sort by Name column to group all skills for each student</li>
    <li><b>By Standard:</b> Sort by Unit, then Skill # to group by curricular sequence</li>
    <li><b>By Progress:</b> Sort by Mastery Grade to identify students needing support</li>
    
    <h3>Adding New Students or Skills</h3>
    <p>As your class evolves:</p>
    <li>Add new students to the <b>Students</b> sheet</li>
    <li>Add new skills to the <b>Skills</b> sheet</li>
    <li>Run <b>${MENU.TITLE} → Add Students &amp; Skills</b> to update the gradebook</li>
    <li>The system will automatically add missing student × skill combinations without duplicating existing data</li>
    
    <h3>Updating and Maintenance</h3>
    <li><b>Regular Updates:</b> Grade attempts as students demonstrate learning</li>
    <li><b>Progress Monitoring:</b> Use the Grade View sheet to see individual student progress</li>
    <li><b>Data Integrity:</b> The system prevents accidental overwrites of existing grades</li>
    `
  );

  // Stage 4: Sharing and Communication
  setRichInstructions(
    sheet.getRange(blurbs['Sharing']),
    `<h2>Stage 4: Sharing and Communication</h2>
    
    <h3>Understanding Student Views</h3>
    <p>The system creates clean, student-friendly views that show:</p>
    <li>Individual student's progress on all skills</li>
    <li>Current mastery grades with clear explanations</li>
    <li>Attempt history showing the path to mastery</li>
    <li>Level descriptions explaining what each score means</li>
    
    <h3>Generating Student Views</h3>
    <p>To create shareable student reports:</p>
    <li>Run <b>${MENU.TITLE} → ${MENU.GENERATE_STUDENT_VIEWS}</b></li>
    <li>This creates individual Google Sheets for each student</li>
    <li>Each sheet automatically updates when you change grades in the main gradebook</li>
    <li>A <b>Student Views</b> sheet appears with links to all individual reports</li>
    
    <h3>Sharing with Students and Families</h3>
    <p>Once views are generated:</p>
    <li>Run <b>${MENU.TITLE} → ${MENU.SHARE_STUDENT_VIEWS}</b> to share views with student emails.</li>
    <li>Students and families can receive <b>READ or COMMENT access only</b> - never EDIT access*</li>
    <li>Each family sees only their own student's progress</li>
    <li>Views update automatically as you add new grades</li>
    <small><i><b>*</b> If students had write access to one of the view sheets, they can change which
    data they see by editing the formula that imports data from the main gradebook.</i></small>
        
    <h4>Important Sharing Guidelines:</h4>
    <li><b>Security:</b> Never grant edit access - families could potentially see other students' grades</li>
    <li><b>Privacy:</b> Each family gets access only to their own student's view</li>
    <li><b>Updates:</b> Changes to the main gradebook automatically appear in shared views</li>
    <li><b>First Access:</b> Open each view once to authorize the IMPORTRANGE function</li>
    
    <h3>Communication Features</h3>
    <li><b>Live Updates:</b> Shared views reflect current grades in real-time</li>
    <li><b>Progress Tracking:</b> Families can see the complete journey toward mastery</li>
    <li><b>Clear Explanations:</b> Views include descriptions of what each level means</li>
    <li><b>Comment Access:</b> Families can add comments for two-way communication</li>
    
    <h3>Managing Shared Views</h3>
    <li>The <b>Student Views</b> sheet provides links to all individual reports</li>
    <li>You can regenerate views at any time if structure changes</li>
    <li>Views are organized in a dedicated folder for easy management</li>
    <li>Each student view is a separate Google Sheet that can be shared independently</li>
    
    <h3>Best Practices for Family Communication</h3>
    <li>Explain the grading system to families before sharing access</li>
    <li>Emphasize that mastery takes time and multiple attempts</li>
    <li>Encourage families to look for growth patterns, not just final scores</li>
    <li>Use the comment feature for ongoing dialogue about student progress</li>
    `
  );

  // Set column widths for readability (content spans the first N TOC columns)
  // Keep column 1 wide enough for merged paragraphs; the others are already set above.
  // Our width should be either 60px per heading OR 850px for the merged, whichever is
  // bigger...
  let colWidth = 850 / tocCols;
  if (colWidth < 60) {
    colWidth = 60;
  }
  for (let c = 1; c <= tocCols; c++) {
    sheet.setColumnWidth(c, colWidth);
  }
}