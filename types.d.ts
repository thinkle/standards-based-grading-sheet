export interface User {
  sourcedId: string;
  status: "active" | "inactive"; // Adjust this if there are other possible statuses.
  dateLastModified: string;
  metadata: Record<string, unknown>; // Use an appropriate type if the structure of metadata is known.
  username: string;
  userIds: Array<unknown>; // Specify the type of userIds elements if known.
  enabledUser: boolean;
  givenName: string;
  familyName: string;
  middleName: string | null;
  role: "teacher" | "student" | "administrator"; // Adjust this based on the possible roles in the system.
  identifier: string | null;
  email: string;
  sms: string | null;
  phone: string | null;
  orgs: Org[];
  agents: Array<unknown>; // Specify the type of agents if known.
  grades: string[]; // Adjust if grades are represented differently.
  nameSuffix: string | null;
  birthDate: string | null;
}

export interface Org {
  href: string;
  sourcedId: string;
  type: "org"; // Adjust this if there are different types of organizations.
}

export interface Course {
  sourcedId: string;
  status: "active" | "tobedeleted";
  dateLastModified: string;
  title: string;
  courseCode: string;
  grades: string[];
  subjects: string[];
  schoolYear: string;
  org: Reference;
}

export interface Reference {
  href: string;
  sourcedId: string;
  type: string;
}

export interface LineItem {
  sourcedId: string;
  status: "active" | "tobedeleted";
  dateLastModified: string;
  title: string;
  description?: string;
  assignDate?: string;
  dueDate?: string;
  category?: Reference;
  gradingPeriod?: Reference;
  resultValueMin?: number;
  resultValueMax?: number;
}

export interface Reference {
  href: string;
  sourcedId: string;
  type: string;
}
// Interface for Category
export interface Category {
  href: string;
  sourcedId: string;
  type: string;
  title: string;
}

// Interface for GradingPeriod
export interface GradingPeriod {
  href: string;
  sourcedId: string;
  type: string;
}

export interface Level {
  id: string;
  title: string;
  description: string;
  points: number;
}

export interface Criterion {
  id: string;
  description: string;
  title: string;
  levels: Level[];
  // Populated after the map, so it holds a quick-reference by level ID
  levelsMap: Record<string, Level>;
}

export interface Rubric {
  id: string;
  courseId: string;
  criteria: Criterion[];
  // Populated after the map, so it holds a quick-reference by criterion ID
  criteriaMap: Record<string, Criterion>;
}
export interface RubricGrade {
  criterion: string;
  level: string;
  points: number;
  criterionId: string;
  levelId: string;
  description: string;
}
export interface Grade {
  studentEmail: string;
  studentName: string;
  assignedGrade: number | null;
  maximumGrade: number;
  submissionState: string;
  late: boolean;
  rubricGrades?: RubricGrade[];
}
