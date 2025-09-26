# Behavior Support System (BSS)

Local development workspace for the Behavior Support System Google Apps Script that powers the linked Google Sheet. Write and test your code in VS Code with TypeScript, then copy the compiled output into Apps Script.

## Prerequisites
- Node.js 18+
- Existing Google Sheet and Google Apps Script project bound to that sheet

## Getting started
1. Install dependencies:
   ```sh
   npm install
   ```
2. Update the spreadsheet ID:
   - In the Apps Script editor, open **Project Settings → Script properties**.
   - Create or update a `SPREADSHEET_ID` property with your sheet's ID.
   - The first local build writes the placeholder `PUT_SPREADSHEET_ID_HERE`; replace it in Apps Script or adjust `src/main.ts` before copying.
3. (Optional) Adjust the manifest template in `appsscript.json` to match your project settings.

## Local workflow
- Edit TypeScript sources under `src/`. The starter file `src/main.ts` contains helpers for reading and writing to the "Behavior Log" sheet along with sample menu actions.
- Build to generate GAS-compatible code:
  ```sh
  npm run build
  ```
  The transpiled script is written to `dist/Code.js`.
- Copy everything from `dist/Code.js` into your Apps Script editor, replacing the existing code.
- Copy `appsscript.json` into the Apps Script **Project Settings → Manifest file** editor if you need to update runtime configuration.

## Recommended edits
- Rename `DEFAULT_SHEET_NAME` or `HEADER_VALUES` in `src/main.ts` to match your sheet layout.
- Extend the `BehaviorRecord` interface and related helper functions for any additional data you need to store.
- Add more functions in TypeScript, then rebuild and copy the updated output to Apps Script.

## Helpful commands
- `npm run build` – one-off build
- `npm run watch` – rebuild automatically while editing

## Next steps
- Consider installing [`@google/clasp`](https://github.com/google/clasp) if you want to push code directly instead of copy/paste.
- Add tests or mocks using `@types/google-apps-script` to validate logic locally before deploying.
