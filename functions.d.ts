/**
 * Type definitions for Google Apps Script functions used across files
 */

// AspenIdGen.js functions
declare function createAssignmentId(
  classId: string,
  unit: string,
  skill: string
): string;
declare function createAssignmentTitle(unit: string, skill: string): string;
declare function sanitizeForId(str: string, maxLength?: number): string;
declare function simpleHash(str: string): string;

// Aspen.js functions
declare function getAspenClassConfig(classId: string): any;
declare function createLineItem(id: string, lineItemData: any): any;
declare function postGrade(
  id: string,
  lineItem: string,
  student: string,
  score: number,
  comment?: string
): any;

// GradeSync.js functions
declare function getColumnIndex(cols: string[], header: string): number;
declare function initializeAspenIntegration(classId: string): void;
declare function getAvailableAspenCourses(): any[];

// Setup functions
declare function setupNamedRanges(): void;
declare function setupStudents(): void;
declare function setupSkills(): void;
declare function setupGradesSheet(): void;
declare function populateGrades(): void;
declare function writePostSetupInstructions(): void;
