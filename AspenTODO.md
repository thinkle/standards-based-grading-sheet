## Completed

- [x] **Aspen integration** (`Aspen.js`)
- [x] **Sheet + Manager class for Aspen Grades** (`AspenGrades.js`)
- [x] **Sheet + Manager class for Aspen Assignments** (`AspenAssignments.js`)
- [x] **Aspen set-up** (`AspenSetupUI.js`, `AspenSetup.js`)

---

## To Do

### `AspenGradeSync.js`

- [x] **Add all skill items to Aspen Assignments Tab**
  - [x] Create a function (menu item) to add all skill or unit items.
  - [x] On completion: foreground the Assignments tab and show a toast prompting the user to fill in `dueDate` to create assignments.

---

#### Progress Notes (2025-09-13) — Assign Date

- [x] Started scaffolding for function to add all skill items to Aspen Assignments Tab in `AspenGradeSync.js`.
- [x] Implemented logic to add all unique (unit, skill) pairs from Grades to Aspen Assignments tab and foreground the sheet with a toast. No deviation from plan.

- [x] **Support `assignDate` parameter**
  - [x] Update `AspenAssignments.js` (and possibly `Aspen.js`) to support an `assignDate` parameter (in addition to `dueDate`) for the OneRoster API.

---

#### Progress Notes (2025-09-13) — Creation Logic

- [x] Added Assign Date column to Aspen Assignments sheet and updated all logic to support assignDate in `AspenAssignments.js`. No changes needed in `Aspen.js` as it already passes through assignDate in the API payload.

- [ ] **Improve assignment creation logic**

  - [x] Update `AspenAssignments.js` to:
    - [x] Add methods for creating new assignments from data in sheet (if row in sheet has all necessary fields EXCEPT AspenID), we create it, then update the sheet with the new ID
    - [x] Handle cases where assignments exist in the sheet but are not yet created (i.e., allow both created and uncreated assignments).
    - [x] Only create assignments with a `dueDate`.
    - [x] Add conditional formatting:
      - [x] **Aspen ID exists:** solid/existing look.
      - [x] **Due date exists:** "ready to roll" look (highlighted).
  - [x] NOTE: should only highlight if DUE DATE + categoryTitle EXIST!
  - [x] **No due date:** grey and italic (not real).

- [x] Add a data validation dropdown on the categoryTitle items which contains the category titles (found in Aspen Config -> Categories JSON -- each object in the JSON list has a title prop).
- [ ] Improve our title generation code in AspenAssignments.js / AspenIDGen to take _THREE_ arguments: Unit, Skill, Descriptor. We don't have to pass descriptors in -- we can look them up in the SKILLS sheet as needed (or make AspenAssignments just read the skills data once on load and then use that to look up the descriptor). Then, our title logic should do the shortening ONLY on the Unit + Skill and then should include the lengthy descriptor in the full title.
- [ ] Fix the Add Skills to Assignments menu item to use Unit + Skill# (NOT skill descriptor) as the main skill name

---

#### Progress Notes (2025-09-13)

- [x] Implemented `createMissingAssignmentsFromSheet(classId)` and manager method to scan rows and create only those with due dates; updates ID, title, created timestamp, and JSON. Added conditional formatting to visually distinguish states. Types updated in `functions.d.ts`.

- [ ] **"Unit Average" magic skill**
  - [ ] Create a new "magic" skill called **Unit Average**.
  - [x] Add a function to "add unit averages to assignments tab" (adds a "Unit Average" for each unit).
  - [ ] Intended for use instead of skills+grades, not in addition.

---

### Grading Automation

- [ ] **Automate grade syncing**

  - [ ] Create a function to get grades from the Grades sheet.
    - [ ] Build a comment from each skill level's feedback (e.g., checkmarks become "Skill Level: Description").
    - [ ] Pull the grade from the mastery grade column.
  - [ ] For each grade/comment, call `maybeSync` to sync the grade.

- [ ] **Handle "Unit Average" skill**
  - [ ] For assignments with type "Unit Average":
    - [ ] Calculate the unit average.
    - [ ] Build a comment listing scores for each sub-skill (e.g., `Skill: Description - 3`), separated by newlines.
    - [ ] Post the grade/average with `maybeSync` as usual.

---

> **Note:**  
> All logic should be designed to read the entire grade sheet **once** per operation to avoid excessive spreadsheet reads.
