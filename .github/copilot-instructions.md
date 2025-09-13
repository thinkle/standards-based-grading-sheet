# GitHub Copilot Instructions

## Project Overview

This is a Google Apps Script (GAS) project for standards-based grading that integrates with Aspen SIS. The codebase manages grade sheets, student data, assignments, and synchronization with the Aspen gradebook system.

## Code Style & Architecture

### Functional over Object-Oriented

- **Prefer functions over classes** unless state management is truly needed
- **Use closures for API wrappers** (see `Aspen.js` pattern)
- **Avoid unnecessary object instantiation** - if it's just logic, make it a function
- **Classes only when managing state** (e.g., `AspenAssignmentManager` tracks assignments/config)

### JavaScript Patterns

- Use **modern JavaScript** (ES6+) features appropriately
- Prefer **const/let** over `var`
- Use **template literals** for string interpolation
- Use **destructuring** for cleaner parameter handling
- **Arrow functions** for simple callbacks, regular functions for main logic

### File Organization

- Each file has a specific purpose (assignments, grades, API calls, etc.)
- **External dependencies** declared with `/* global */` comments for linter compatibility
- **Type definitions** centralized in `functions.d.ts` for IntelliSense support
- **Helper functions** at top of files, main logic below

## Google Apps Script Specifics

### Sheet Management

- Use **structured headers** with constants (see `ASPEN_ASSIGNMENTS_HEADERS`)
- **Column arrays** for consistent ordering
- **Helper functions** like `getColumnIndex()` for maintainable sheet access
- **Create sheets on-demand** with proper initialization

### API Integration

- **Closure pattern** for API clients (OAuth, rate limiting, error handling)
- **Store raw API responses** in JSON columns for debugging
- **Human-readable keys** alongside system IDs (e.g., "Unit - Skill" format)
- **Graceful error handling** with meaningful messages

## Data Patterns

### Assignment & Grade Tracking

- **Assignment Key format**: `"Unit - Skill"` (human-readable)
- **Store both spec and API result** for assignments
- **Track sync state** with timestamps and change detection
- **Comments support** throughout grading workflow

### ID Generation

- **Functional approach** for ID generation (not classes)
- **Hash-based IDs** for predictable assignment identifiers
- **Sanitization** for Aspen compatibility (dots to "DOT", etc.)

## JSDoc Standards

- **Complete type definitions** with `@typedef` for complex objects
- **Parameter documentation** with types and descriptions
- **Return type specification**
- **Function signatures** declared in `functions.d.ts` for cross-file references
- **Optional parameters** marked with `[param=default]` in JSDoc or `param?:` in TypeScript

## External Function Dependencies

- **Use `/* global */` comments** at the top of files to declare external functions for ESLint
- **Define function signatures** in `functions.d.ts` for TypeScript IntelliSense
- **Example pattern**:
  ```javascript
  /* global createAssignmentId, createAssignmentTitle, getAspenClassConfig */
  // functions.d.ts handles the type definitions
  ```
- **Avoid JSDoc `@external` or `@function`** declarations - use the centralized TypeScript approach

## Error Handling

- **Meaningful error messages** that help users understand issues
- **Check prerequisites** (e.g., "Run initializeAspenIntegration first")
- **Graceful degradation** when possible
- **Console logging** for debugging during development

## Testing

- **Test functions** in dedicated files (e.g., `AspenTest.js`)
- **Use real class IDs** but comment them clearly
- **Test both success and failure cases**
- **Mock external dependencies** when appropriate

## Anti-Patterns to Avoid

- ❌ Creating classes for simple utility functions
- ❌ Java-style verbose object hierarchies
- ❌ Duplicating type definitions across files (use `functions.d.ts`)
- ❌ Hard-coding magic numbers or strings
- ❌ Ignoring Google Apps Script quotas and limitations
- ❌ Cryptic variable names or IDs without human-readable alternatives

## When Suggesting Code

- **Read existing patterns** in the codebase first
- **Follow the functional-first approach**
- **Use established naming conventions**
- **Include proper JSDoc documentation**
- **Consider Google Apps Script limitations**
- **Test suggestions against existing code style**
