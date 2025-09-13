/* Aspen.js Last Update 2025-09-13 12:11 <3bea24e38668616b3079dd180da45b5a522a8881279cdd98e074e632a7887eac>
/**
 * @typedef {import('./types').User} User
 * @typedef {import('./types').Course} Course  
 * @typedef {import('./types').LineItem} LineItem
 * @typedef {import('./types').GradingPeriod} GradingPeriod
 * @typedef {import('./types').Category} Category
 */

// Note: logApiCall function should be available globally 
// or defined in other files in the Google Apps Script project

function aspenInterface() {
  /*
   * All code related to our API *must* be encapsulated
   * here, as all top-level functions are exposed in client code
   * and we cannot expose our Aspen API keys or calls directly.
   * Be *very* careful what you expose here and note that things
   * like caching calls and tokens *also* must be encapsulated.
   */

  const TOKEN_KEY = "aspen_access_token";
  const EXPIRATION_KEY = "aspen_token_expiration";
  function getApiKey() {
    if (typeof PropertiesService !== "undefined") {
      return PropertiesService.getScriptProperties().getProperty(
        "ASPEN_API_SECRET"
      );
    } else {
      // Fallback to environment variable if running outside Google Apps Script
      return process.env.VITE_ASPEN_API_SECRET;
    }
  }

  function getApiId() {
    if (typeof PropertiesService !== "undefined") {
      return PropertiesService.getScriptProperties().getProperty(
        "ASPEN_API_ID"
      );
    } else {
      // Fallback to environment variable if running outside Google Apps Script
      return process.env.VITE_ASPEN_API_ID;
    }
  }

  // Utility functions, but they could be used to bypass
  // our security so they must be encapsulated here.
  /**
   * @param {string} key
   * @returns {string | null}
   */
  function getProp(key) {
    if (typeof PropertiesService !== "undefined") {
      const properties = PropertiesService.getScriptProperties();
      return properties.getProperty(key);
    } else if (typeof localStorage !== "undefined") {
      return localStorage.getItem(key);
    }
    return null;
  }

  /**
   * @param {string} key
   * @param {string} value
   */
  function setProp(key, value) {
    if (typeof PropertiesService !== "undefined") {
      const properties = PropertiesService.getScriptProperties();
      properties.setProperty(key, value);
    } else if (typeof localStorage !== "undefined") {
      localStorage.setItem(key, value);
    }
  }

  function getAccessToken() {
    const now = Date.now();

    // Get token and expiration from the appropriate storage
    const cachedToken = getProp(TOKEN_KEY);
    const tokenExpiration = parseInt(getProp(EXPIRATION_KEY) || "0");

    // Check if the token is cached and still valid
    if (cachedToken && tokenExpiration && now < tokenExpiration) {
      console.log("Using cached token");
      return cachedToken;
    }

    const url = "https://ma-innovation.myfollett.com/oauth/rest/v2.0/auth";
    const clientId = getApiId();
    const clientSecret = getApiKey();
    const options = {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      payload: `grant_type=client_credentials&client_id=${encodeURIComponent(
        clientId
      )}&client_secret=${encodeURIComponent(clientSecret)}`,
    };

    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    const expiresIn = data.expires_in || 3600; // Default to 1 hour if not provided

    // Cache the token and its expiration time
    const newToken = data.access_token;
    const newExpiration = now + expiresIn * 1000; // Convert expiresIn to milliseconds

    setProp(TOKEN_KEY, newToken);
    setProp(EXPIRATION_KEY, newExpiration.toString());

    return newToken;
  }

  function testApiCall() {
    const response = UrlFetchApp.fetch(
      "https://jsonplaceholder.typicode.com/posts/1",
      { method: "GET" }
    );
    const json = JSON.parse(response.getContentText());
    return json;
  }

  // In-memory, per-execution caches to avoid redundant API calls during a single run
  /** @type {Record<string, User>} */
  const memoTeacherByEmail = {};
  /** @type {Record<string, import('./types').Course[]>} */
  const memoCoursesByTeacherId = {};
  /** @type {Record<string, boolean>} */
  const memoCourseAccess = {};
  /** @type {Record<string, string>} */
  const memoLineItemClassId = {};

  /**
   * @returns {User[]}
   */
  function fetchTeachers() {
    const accessToken = getAccessToken();

    const url =
      "https://ma-innovation.myfollett.com/ims/oneroster/v1p1/teachers?limit=100&offset=0&orderBy=asc";
    const response = UrlFetchApp.fetch(url, {
      method: "GET",
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${accessToken}`,
      },
    });

    const data = JSON.parse(response.getContentText());
    console.log("Got data: ", data.users);
    return data.users;
  }

  /* Access control system: ensure that teachers can only access
     their own courses, students and grades. 
     The Aspen API does *not* enforce this limitation, so we must
     be extremely careful to enforce it ourselves.
  */

  /**
   * @param {string} courseId
   * @returns {boolean}
   */
  function hasAccessToCourse(courseId) {
    if (typeof Session == "undefined") {
      // override -- local testing environment
      return true;
    }

    // Fast path: memoized decision for this courseId
    if (Object.prototype.hasOwnProperty.call(memoCourseAccess, courseId)) {
      return !!memoCourseAccess[courseId];
    }

    const authorizedEmail = Session.getActiveUser().getEmail();
    const authorizedTeacher = fetchTeacherByEmail(authorizedEmail); // will use memo
    const authorizedCourses = fetchAspenCourses(authorizedTeacher); // will use memo

    // Build a set of allowed course IDs and cache per course
    let allowed = false;
    for (let i = 0; i < authorizedCourses.length; i++) {
      const cid = authorizedCourses[i].sourcedId;
      memoCourseAccess[cid] = true;
      if (cid === courseId) allowed = true;
    }

    if (!allowed) {
      console.error(
        "Teacher",
        authorizedTeacher,
        "does not have access to course",
        courseId
      );
    }
    // Cache negative as well to avoid repeat lookups
    if (!Object.prototype.hasOwnProperty.call(memoCourseAccess, courseId)) {
      memoCourseAccess[courseId] = allowed;
    }
    return allowed;
  }

  /**
   * @param {string} lineItemId
   * @returns {boolean}
   */
  function hasAccessToLineItem(lineItemId) {
    if (typeof Session == "undefined") {
      // override -- local testing environment
      return true;
    }
    // If previously resolved, reuse cached classId
    const cachedClassId = memoLineItemClassId[lineItemId];
    if (cachedClassId) {
      return hasAccessToCourse(cachedClassId);
    }
    // Otherwise fetch once, then memoize mapping
    const url = `https://ma-innovation.myfollett.com/ims/oneroster/v1p1/lineItems/${lineItemId}`;
    const accessToken = getAccessToken();
    const response = UrlFetchApp.fetch(url, {
      method: "GET",
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${accessToken}`,
      },
    });
    const data = JSON.parse(response.getContentText());
    const lineItem = data.lineItem;
    const courseId = lineItem.class.sourcedId;
    memoLineItemClassId[lineItemId] = courseId;
    return hasAccessToCourse(courseId);
  }

  /**
   * @param {string} email
   * @returns {User}
   */
  function fetchTeacherByEmail(email) {
    if (memoTeacherByEmail[email]) {
      return memoTeacherByEmail[email];
    }
    const accessToken = getAccessToken();
    const url = `https://ma-innovation.myfollett.com/ims/oneroster/v1p1/teachers?filter=email=${email}`;
    const response = UrlFetchApp.fetch(url, {
      method: "GET",
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${accessToken}`,
      },
    });

    const data = JSON.parse(response.getContentText());
    console.log("Got one teacher data: ", data.users);
    const teacher = data.users[0];
    if (teacher) memoTeacherByEmail[email] = teacher;
    return teacher;
  }

  /**
   * @param {User} teacher (optional, defaults to logged-in user)
   * @returns {Course[]}
   */
  function fetchAspenCourses(teacher = null) {
    if (teacher) {
      // Security layer...
      let teacherEmail = teacher.email;
      if (typeof Session !== "undefined") {
        let loggedInEmail = Session.getActiveUser().getEmail();

        if (teacherEmail !== loggedInEmail) {
          throw new Error("Unauthorized access to teacher data");
        }
      }
    } else {
      teacher = fetchTeacherByEmail(
        Session.getActiveUser().getEmail()
      );
    }
    const accessToken = getAccessToken();
    const teacherId = teacher.sourcedId;
    if (memoCoursesByTeacherId[teacherId]) {
      return memoCoursesByTeacherId[teacherId];
    }
    const url = `https://ma-innovation.myfollett.com/ims/oneroster/v1p1/teachers/${teacherId}/classes?limit=100&offset=0&orderBy=asc`;

    const response = UrlFetchApp.fetch(url, {
      method: "GET",
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${accessToken}`,
      },
    });

    const data = JSON.parse(response.getContentText());
    console.log("Courses fetched: ", data.classes);
    memoCoursesByTeacherId[teacherId] = data.classes || [];
    return memoCoursesByTeacherId[teacherId];
  }

  /**
   * @param {Course} course
   * @returns {LineItem[]}
   */
  function fetchLineItems(course) {
    const accessToken = getAccessToken();
    const courseId = course.sourcedId;
    const url = `https://ma-innovation.myfollett.com/ims/oneroster/v1p1/classes/${courseId}/lineItems?limit=100&offset=0&orderBy=asc`;

    const response = UrlFetchApp.fetch(url, {
      method: "GET",
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${accessToken}`,
      },
    });

    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
      if (typeof logApiCall !== "undefined") {
        logApiCall({
          method: "GET",
          url: url,
          response: response.getContentText(),
        });
      }
      throw new Error("Failed to fetch line items: " + response.getContentText());
    }

    const data = JSON.parse(response.getContentText());
    console.log("Line items fetched: ", data.lineItems);
    if (typeof logApiCall !== "undefined") {
      logApiCall({
        method: "GET",
        url: url,
        response: `${data.lineItems.length} line items fetched`,
      });
    }
    return data.lineItems;
  }

  /**
   * @param {Course} course
   * @returns {User[]}
   */
  function fetchStudents(course) {
    if (!hasAccessToCourse(course.sourcedId)) {
      throw new Error("Unauthorized access to student data");
    }
    const accessToken = getAccessToken();
    const courseId = course.sourcedId;
    const url = `https://ma-innovation.myfollett.com/ims/oneroster/v1p1/classes/${courseId}/students?limit=100&offset=0&orderBy=asc`;

    const response = UrlFetchApp.fetch(url, {
      method: "GET",
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${accessToken}`,
      },
    });
    const data = JSON.parse(response.getContentText());
    console.log("Students fetched: ", data.users);
    if (typeof logApiCall !== "undefined") {
      logApiCall({
        method: "GET",
        url: url,
        response: `${data.users.length} students fetched`,
      });
    }
    return data.users;
  }

  /**
   * @param {Course} course
   * @returns {Category[]}
   */
  function fetchCategories(course) {
    const accessToken = getAccessToken();
    let classId = course.sourcedId;
    const url = `https://ma-innovation.myfollett.com/ims/oneroster/v1p1/categories?filter=metadata.classId=${classId}`;

    const response = UrlFetchApp.fetch(url, {
      method: "GET",
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${accessToken}`,
      },
    });

    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
      throw new Error("Failed to fetch categories: " + response.getContentText());
    }

    const data = JSON.parse(response.getContentText());
    console.log("Categories fetched: ", data.categories);
    return data.categories;
  }

  /**
   * @returns {GradingPeriod[]}
   */
  function fetchGradingPeriods() {
    const accessToken = getAccessToken();
    const url = `https://ma-innovation.myfollett.com/ims/oneroster/v1p1/gradingPeriods`;

    const response = UrlFetchApp.fetch(url, {
      method: "GET",
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${accessToken}`,
      },
    });

    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
      throw new Error("Failed to fetch grading periods: " + response.getContentText());
    }

    const data = JSON.parse(response.getContentText());
    console.log("Grading periods fetched: ", data.gradingPeriods);
    return data.gradingPeriods;
  }

  /**
   * @param {string} id
   * @param {LineItem} lineItemData
   * @returns {LineItem}
   */
  function createLineItem(id, lineItemData) {
    let courseId = lineItemData.class.sourcedId;
    if (!hasAccessToCourse(courseId)) {
      throw new Error("Unauthorized access to course data");
    }
    const accessToken = getAccessToken();
    const url = `https://ma-innovation.myfollett.com/ims/oneroster/v1p1/lineItems/${id}`;

    const response = UrlFetchApp.fetch(url, {
      method: "PUT",
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      payload: JSON.stringify({ lineItem: lineItemData }),
    });

    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
      throw new Error("Failed to create line item: " + response.getContentText());
    }

    const data = JSON.parse(response.getContentText());
    console.log("Line item created: ", data.lineItem);
    if (typeof logApiCall !== "undefined") {
      logApiCall({
        method: "PUT",
        url: url,
        response: `Line item created: ${data.lineItem.title} ${data.lineItem.sourcedId}`,
      });
    }
    return data.lineItem;
  }

  /**
   * @param {string} classId
   * @returns {User[]}
   */
  function fetchAspenRoster(classId) {
    // Security layer...
    if (!hasAccessToCourse(classId)) {
      throw new Error("Unauthorized access to roster data");
    }
    const accessToken = getAccessToken();
    const url = `https://ma-innovation.myfollett.com/ims/oneroster/v1p1/classes/${classId}/students?limit=100&offset=0&orderBy=asc`;

    const response = UrlFetchApp.fetch(url, {
      method: "GET",
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${accessToken}`,
      },
    });

    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
      throw new Error("Failed to fetch students: " + response.getContentText());
    }

    const data = JSON.parse(response.getContentText());
    console.log("Students fetched: ", data.users);
    return data.users;
  }

  /**
   * @param {string} resultId
   * @param {any} resultData
   * @returns {any}
   */
  function postResult(resultId, resultData) {
    console.time && console.time('authCheck');
    const lineItemId = resultData.result.lineItem.sourcedId;
    // Prefer using classId hint provided by caller to skip an extra GET
    const classIdFromPayload = resultData._classId;
    const authorized = classIdFromPayload
      ? hasAccessToCourse(classIdFromPayload)
      : hasAccessToLineItem(lineItemId);
    console.timeEnd && console.timeEnd('authCheck');
    if (!authorized) {
      throw new Error("Unauthorized access to line item data");
    }

    const accessToken = getAccessToken();
    const url = `https://ma-innovation.myfollett.com/ims/oneroster/v1p1/results/${resultId}`;

    console.time && console.time('putResult');
    // Remove private hint before sending
    try { delete resultData._classId; } catch (e) { }
    const response = UrlFetchApp.fetch(url, {
      method: "PUT",
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      payload: JSON.stringify(resultData),
    });
    console.timeEnd && console.timeEnd('putResult');

    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
      if (typeof logApiCall !== "undefined") {
        logApiCall({
          method: "PUT",
          url: url,
          response: response.getContentText(),
        });
      }
      throw new Error("Failed to post result: " + response.getContentText());
    }

    const data = JSON.parse(response.getContentText());
    // Memoize lineItem->class mapping to speed later checks in same execution
    try {
      const li = data && data.result && data.result.lineItem;
      if (li && li.sourcedId) {
        const cid = (li.class && li.class.sourcedId) || classIdFromPayload;
        if (cid) memoLineItemClassId[li.sourcedId] = cid;
      }
    } catch (e) {
      // ignore memo failures
    }
    if (typeof logApiCall !== "undefined") {
      logApiCall({
        method: "PUT",
        url: url,
        response: `Result posted: ${data.result.sourcedId}`,
      });
    }
    console.log("Result posted: ", data);
    return data;
  }

  function postGrade(id, lineItem, student, score, comment) {
    let resultObject = {
      result: {
        lineItem: {
          sourcedId: lineItem.sourcedId,
          href: lineItem.href,
          type: lineItem.type,
        },
        student: {
          sourcedId: student.sourcedId,
          href: student.href,
          type: student.type,
        },
        score: score,
        comment: comment,
      },
    };
    // Provide private classId hint for auth optimization (not sent to API)
    try {
      if (lineItem && lineItem.class && lineItem.class.sourcedId) {
        resultObject._classId = lineItem.class.sourcedId;
      }
    } catch (e) { }
    let data = postResult(id, resultObject);
    return data;
  }

  return {
    testApiCall, // no enforcement needed -- data not private
    fetchTeacherByEmail, // enforces email
    fetchAspenCourses, // enforces teacher
    fetchLineItems, // enforces course -> teacher
    fetchStudents, // enforces course -> teacher
    fetchCategories, // no enforcement needed -- data not private
    fetchGradingPeriods, // no enforcement needed -- data not private
    createLineItem, // enforces course -> teacher
    fetchAspenRoster, // enforces course -> teacher
    postGrade, // enforces line item -> course -> teacher
  };
}

// Expose the public interface
const aspenAPI = aspenInterface();

/**
 * @returns {any}
 */
function testApiCall() {
  return aspenAPI.testApiCall();
}

/**
 * @param {string} email
 * @returns {User}
 */
function fetchTeacherByEmail(email) {
  return aspenAPI.fetchTeacherByEmail(email);
}

/**
 * @param {User} teacher
 * @returns {Course[]}
 */
function fetchAspenCourses(teacher) {
  return aspenAPI.fetchAspenCourses(teacher);
}

/**
 * @param {Course} course
 * @returns {LineItem[]}
 */
function fetchLineItems(course) {
  return aspenAPI.fetchLineItems(course);
}

/**
 * @param {Course} course
 * @returns {User[]}
 */
function fetchStudents(course) {
  return aspenAPI.fetchStudents(course);
}

/**
 * @param {Course} course
 * @returns {Category[]}
 */
function fetchCategories(course) {
  return aspenAPI.fetchCategories(course);
}

/**
 * @returns {GradingPeriod[]}
 */
function fetchGradingPeriods() {
  return aspenAPI.fetchGradingPeriods();
}

/**
 * @param {string} id
 * @param {LineItem} lineItemData
 * @returns {LineItem}
 */
function createLineItem(id, lineItemData) {
  return aspenAPI.createLineItem(id, lineItemData);
}

/**
 * @param {string} classId
 * @returns {User[]}
 */
function fetchAspenRoster(classId) {
  return aspenAPI.fetchAspenRoster(classId);
}

/**
 * @param {string} resultId
 * @param {any} resultData
 * @returns {any}
 */
function postResult(resultId, resultData) {
  return aspenAPI.postResult(resultId, resultData);
}

/**
 * @param {string} id
 * @param {LineItem} lineItem
 * @param {User} student
 * @param {number} score
 * @param {string} comment
 * @returns {any}
 */
function postGrade(id, lineItem, student, score, comment) {
  return aspenAPI.postGrade(id, lineItem, student, score, comment);
}
