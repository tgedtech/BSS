# Copilot Instructions for Behavior Support System (BSS)

This document provides guidance for AI coding agents working on the Behavior Support System (BSS) project. The goal is to ensure productive contributions by understanding the project's architecture, workflows, and conventions.

## Project Overview
The Behavior Support System (BSS) is a Google Apps Script project linked to a Google Sheet. The codebase is developed locally in TypeScript and transpiled into JavaScript for deployment in the Apps Script environment.

### Key Components
- **`src/`**: Contains TypeScript source files. The main entry point is `src/main.ts`.
- **`dist/Code.js`**: The transpiled output, compatible with Google Apps Script.
- **`appsscript.json`**: The manifest file for configuring the Apps Script project.

### Data Flow
- The system interacts with a Google Sheet, primarily the "Behavior Log" sheet.
- Functions in `src/main.ts` handle reading and writing data to the sheet.
- Developers can extend the `BehaviorRecord` interface and helper functions to support additional data fields.

## Developer Workflows
### Local Development
1. **Install dependencies**:
   ```sh
   npm install
   ```
2. **Build the project**:
   ```sh
   npm run build
   ```
   This generates the `dist/Code.js` file.
3. **Copy to Apps Script**:
   - Replace the code in the Apps Script editor with the contents of `dist/Code.js`.
   - Update the manifest file in the Apps Script editor with `appsscript.json` if needed.

### Debugging
- Use `npm run watch` to automatically rebuild while editing.
- Logs can be added using `console.log` statements, which will appear in the Apps Script execution logs.

### Spreadsheet Integration
- Update the `SPREADSHEET_ID` in the Apps Script project settings to link to the correct Google Sheet.
- Adjust `DEFAULT_SHEET_NAME` and `HEADER_VALUES` in `src/main.ts` to match the sheet layout.

## Project-Specific Conventions
- **TypeScript-first development**: All logic should be written in TypeScript under `src/`.
- **BehaviorRecord interface**: Extend this interface to define the structure of data stored in the sheet.
- **Menu actions**: Add custom menu actions in `src/main.ts` to trigger specific functions.

## External Dependencies
- **Node.js**: Required for building the project.
- **Google Apps Script**: The runtime environment for the transpiled code.
- **Optional**: [`@google/clasp`](https://github.com/google/clasp) for direct deployment.

## Examples
### Adding a New Function
1. Define the function in `src/main.ts`:
   ```typescript
   export function newFunction() {
       console.log("New function executed");
   }
   ```
2. Rebuild the project:
   ```sh
   npm run build
   ```
3. Copy the updated `dist/Code.js` to Apps Script.

### Extending BehaviorRecord
1. Update the interface in `src/main.ts`:
   ```typescript
   interface BehaviorRecord {
       timestamp: string;
       studentName: string;
       behavior: string;
       // Add new fields here
       notes?: string;
   }
   ```
2. Adjust helper functions to handle the new fields.

---

For further details, refer to the `README.md` file or the source code in `src/main.ts`. If any instructions are unclear or incomplete, please provide feedback to improve this document.