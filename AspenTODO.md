## Completed

- [x] **Aspen integration** (`Aspen.js`)
- [x] **Sheet + Manager class for Aspen Grades** (`AspenGrades.js`)
- [x] **Sheet + Manager class for Aspen Assignments** (`AspenAssignments.js`)
- [x] **Aspen set-up** (`AspenSetupUI.js`, `AspenSetup.js`)

---

## To Do

### `AspenGradeSync.js`

- [ ] **Add all skill items to Aspen Assignments Tab**

  - [ ] Create a function (menu item) to add all skill or unit items.
  - [ ] On completion: foreground the Assignments tab and show a toast prompting the user to fill in `dueDate` to create assignments.

- [ ] **Support `assignDate` parameter**

  - [ ] Update `AspenAssignments.js` (and possibly `Aspen.js`) to support an `assignDate` parameter (in addition to `dueDate`) for the OneRoster API.

- [ ] **Improve assignment creation logic**

  - [ ] Update `AspenAssignments.js` to:
    - [ ] Add methods for creating new assignments.
    - [ ] Handle cases where assignments exist in the sheet but are not yet created (i.e., allow both created and uncreated assignments).
    - [ ] Only create assignments with a `dueDate`.
    - [ ] Add conditional formatting:
      - [ ] **Aspen ID exists:** solid/existing look.
      - [ ] **Due date exists:** "ready to roll" look (highlighted).
      - [ ] **No due date:** grey and italic (not real).

- [ ] **"Unit Avg" magic skill**
  - [ ] Create a new "magic" skill called **Unit Avg**.
  - [ ] Add a function to "add unit averages to assignments tab" (adds a "Unit Avg" for each unit).
  - [ ] Intended for use instead of skills+grades, not in addition.

---

### Grading Automation

- [ ] **Automate grade syncing**

  - [ ] Create a function to get grades from the Grades sheet.
    - [ ] Build a comment from each skill level's feedback (e.g., checkmarks become "Skill Level: Description").
    - [ ] Pull the grade from the mastery grade column.
  - [ ] For each grade/comment, call `maybeSync` to sync the grade.

- [ ] **Handle "Unit Avg" skill**
  - [ ] For assignments with type "Unit Avg":
    - [ ] Calculate the unit average.
    - [ ] Build a comment listing scores for each sub-skill (e.g., `Skill: Description - 3`), separated by newlines.
    - [ ] Post the grade/average with `maybeSync` as usual.

---

> **Note:**  
> All logic should be designed to read the entire grade sheet **once** per operation to avoid excessive spreadsheet reads.
