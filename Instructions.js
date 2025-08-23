function writePostSetupInstructions() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('Instructions');
  if (!sheet) {
    sheet = ss.insertSheet('Instructions');
  }
  
  // Overview and Quick Start
  setRichInstructions(
    sheet.getRange('A1'),
    `<h1>Standards-Based Grading Sheet - Complete Guide</h1>
    <p>This standards-based grading system was inspired by educational research and teacher feedback on effective assessment practices. This guide walks you through setting up and using your standards-based gradebook in four key stages:</p>
    <li><b>Stage 1:</b> Initial Setup - Configuring levels and mastery definitions</li>
    <li><b>Stage 2:</b> Students and Standards - Adding your roster and skills</li>
    <li><b>Stage 3:</b> Using the Gradebook - Daily grading and management</li>
    <li><b>Stage 4:</b> Sharing and Communication - Student and family views</li>
    `
  );

  // Stage 1: Initial Setup
  setRichInstructions(
    sheet.getRange('A2'),
    `<h2>Stage 1: Initial Setup</h2>
    <h3>Understanding the System</h3>
    <p>This standards-based grading system works by having students demonstrate mastery through multiple attempts at different skill levels. Students need to achieve a "streak" of successful attempts to show true mastery.</p>
    
    <h3>Step 1: Configure Your Levels</h3>
    <p>Go to the <b>LevelSettings</b> sheet to define your mastery levels. The default setup includes:</p>
    <li><b>Basic (B):</b> Entry-level understanding - Score: 2, Requires: 2 consecutive successes</li>
    <li><b>Intermediate (I):</b> Developing proficiency - Score: 3, Requires: 2 consecutive successes</li>
    <li><b>Mastery (M):</b> Full mastery - Score: 4, Requires: 2 consecutive successes</li>
    
    <h4>Customizing Levels:</h4>
    <li><b>Level Name:</b> What you call this level (e.g., "Basic", "Proficient", "Advanced")</li>
    <li><b>Short Code:</b> Letter used in gradebook columns (e.g., "B", "P", "A")</li>
    <li><b>Default Attempts:</b> How many attempt columns to create for this level</li>
    <li><b>Streak for Mastery:</b> How many consecutive successes needed to demonstrate mastery</li>
    <li><b>Score:</b> The numeric grade assigned when this level is mastered</li>
    
    <h3>Step 2: Define Mastery Criteria</h3>
    <p>The <b>Symbols</b> sheet uses a three-column system that enhances teacher data entry based on educator feedback:</p>
    <li><b>Character:</b> What teachers type for easy data entry</li>
    <li><b>Mastery:</b> Whether this counts toward mastery (1 = Yes, 0 = No)</li>
    <li><b>Symbol:</b> What displays in reports and student views</li>
    
    <h4>Comprehensive Symbol System:</h4>
    <p>The system supports a rich variety of assessment marks originally designed for standards-based grading:</p>
    
    <h5>Successful Attempts (Count Toward Mastery):</h5>
    <li><b>✔</b> - KDI (Knowledge Demonstrated Individually)</li>
    <li><b>✔o</b> - KDI via teacher Observation</li>
    <li><b>✔c</b> - KDI via Conversation with teacher</li>
    <li><b>✔s</b> - KDI with a Silly mistake not related to the objective</li>
    
    <h5>Learning Attempts (Do Not Count Toward Mastery):</h5>
    <li><b>H</b> - Knowledge Demonstrated with Help from a teacher or peer</li>
    <li><b>G</b> - Knowledge demonstrated in a Group setting</li>
    <li><b>✗</b> - Question attempted but answered incorrectly</li>
    <li><b>✗o</b> - Teacher observes incorrect attempt</li>
    <li><b>✗c</b> - Incorrect response in Conversation with teacher</li>
    <li><b>PC</b> - Partially correct</li>
    <li><b>N</b> - Not Attempted (score of 0)</li>
    
    <p><b>Customization:</b> You can define your own symbols and meanings. The three-column approach allows teachers to type simple characters (like "check" or "x") while displaying meaningful symbols (✓ or ✗) in student reports.</p>
    
    <h3>Step 3: Run Initial Setup</h3>
    <p>From the menu, select <b>Standards-Based Grading → Initial Setup</b> to create all necessary sheets and named ranges.</p>
    `
  );

  // Stage 2: Students and Standards
  setRichInstructions(
    sheet.getRange('A3'),
    `<h2>Stage 2: Setting up Students and Standards</h2>
    
    <h3>Adding Students</h3>
    <p>Go to the <b>Students</b> sheet and enter:</p>
    <li><b>Student Names:</b> Full names as you want them to appear in grade reports</li>
    <li><b>Student Emails:</b> Required for sharing individual student views with families</li>
    
    <h4>Important Notes:</h4>
    <li>Include student emails if you plan to share live views with students or families</li>
    <li>You can add students at any time - just run "Add Students & Skills" from the menu to update the gradebook</li>
    <li>Student order will be preserved in the gradebook and student views</li>
    
    <h3>Defining Skills and Standards</h3>
    <p>Go to the <b>Skills</b> sheet and define what you'll be assessing:</p>
    <li><b>Unit:</b> Organizing header (e.g., "Algebra", "Chapter 1", "Quarter 1")</li>
    <li><b>Skill #:</b> Number or code for sorting (e.g., "1.1", "A", "01")</li>
    <li><b>Skill Description:</b> Clear description of what students must demonstrate</li>
    
    <h4>Best Practices for Skills:</h4>
    <li>Use clear, specific descriptions that students and parents will understand</li>
    <li>Group related skills under the same Unit for easier organization</li>
    <li>Use consistent numbering schemes within units</li>
    <li>Consider breaking complex standards into smaller, assessable skills</li>
    
    <h3>Building the Gradebook Structure</h3>
    <p>Once students and skills are defined:</p>
    <li>Run <b>Standards-Based Grading → Setup Grade Sheet</b> to create the gradebook structure</li>
    <li>Run <b>Standards-Based Grading → Add Students & Skills</b> to populate all student × skill combinations</li>
    <p>This creates one row for each student-skill combination, allowing individual tracking of progress on every standard.</p>
    `
  );

  // Stage 3: Using the Gradebook
  setRichInstructions(
    sheet.getRange('A4'),
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
    <li>Choose the appropriate level column (Basic, Intermediate, or Mastery)</li>
    <li>Enter <b>✓</b> for successful attempts or <b>X</b> for unsuccessful attempts</li>
    <li>The system automatically calculates streaks and updates the overall Mastery Grade</li>
    
    <h3>Organizing Your Gradebook View</h3>
    <p>With hundreds of rows, organization is key! Use these strategies:</p>
    
    <h4>Option 1: Google Sheets Tables (Recommended)</h4>
    <li>Select your data range and choose <b>Format → Convert to table</b></li>
    <li>Create <b>Group by Student</b> view to see all skills for one student at a time</li>
    <li>Create <b>Group by Skill</b> view to see all students' progress on one skill</li>
    <li>Create <b>Group by Unit</b> view to focus on one unit at a time</li>
    
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
    <li>Run <b>Standards-Based Grading → Add Students & Skills</b> to update the gradebook</li>
    <li>The system will automatically add missing student × skill combinations without duplicating existing data</li>
    
    <h3>Updating and Maintenance</h3>
    <li><b>Regular Updates:</b> Grade attempts as students demonstrate learning</li>
    <li><b>Progress Monitoring:</b> Use the Grade View sheet to see individual student progress</li>
    <li><b>Data Integrity:</b> The system prevents accidental overwrites of existing grades</li>
    `
  );

  // Stage 4: Sharing and Communication
  setRichInstructions(
    sheet.getRange('A5'),
    `<h2>Stage 4: Sharing and Communication</h2>
    
    <h3>Understanding Student Views</h3>
    <p>The system creates clean, student-friendly views that show:</p>
    <li>Individual student's progress on all skills</li>
    <li>Current mastery grades with clear explanations</li>
    <li>Attempt history showing the path to mastery</li>
    <li>Level descriptions explaining what each score means</li>
    
    <h3>Generating Student Views</h3>
    <p>To create shareable student reports:</p>
    <li>Run <b>Standards-Based Grading → Generate student views</b></li>
    <li>This creates individual Google Sheets for each student</li>
    <li>Each sheet automatically updates when you change grades in the main gradebook</li>
    <li>A <b>Student Views</b> sheet appears with links to all individual reports</li>
    
    <h3>Sharing with Students and Families</h3>
    <p>Once views are generated:</p>
    <li>Run <b>Standards-Based Grading → Share student views</b> to set up sharing</li>
    <li>Students and families receive <b>READ or COMMENT access only</b> - never EDIT access</li>
    <li>Each family sees only their own student's progress</li>
    <li>Views update automatically as you add new grades</li>
    
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

  // Set column width for readability
  sheet.setColumnWidth(1, 850);
}