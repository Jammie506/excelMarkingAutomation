# Word Processing Marking Sheet — 4N1123

An automated Excel workbook for generating individual learner marking sheets and assessment briefs for the **Word Processing 4N1123** (QQI Level 4) module. Designed to reduce manual effort when assessing a class of learners across the two assessment components.

## Assessment Structure

| Component | Weighting |
|---|---|
| Section 1 — Collection of Work | 70% (out of 70) |
| Section 2 — Examination | 30% (out of 30) |

**Section 1 — Collection of Work** covers five tasks:
- Task 1: File Management (MIMLO 2)
- Task 2: Document Creation — Text (MIMLO 3)
- Task 3: Document Creation — Graphics (MIMLO 3)
- Task 4: Document Creation — Tables (MIMLO 3)
- Task 5: Primary Functions & Processes (MIMLO 1)

**Section 2 — Examination** assesses: page setup, text insertion, character/paragraph formatting, object formatting, and use of review tools.

## What the Workbook Does

- Reads learner names from the `learnerList` sheet
- Generates individual marking sheets for each learner by copying and populating the `markingTemplate`
- Generates individual assessment briefs by copying and populating the `briefTemplate`
- Populates named cells with learner names and section scores automatically
- Logs each script run in the `RunLog` sheet
- Exports marking sheets and briefs as PDFs to configurable output folders

## Workbook Structure

| Sheet | Purpose |
|---|---|
| `config` | All configurable settings (paths, ranges, template names, cell references) |
| `learnerList` | Input list of learner names to be processed |
| `gradeSheet` | Overview of grades across all learners for both sections |
| `markingTemplate` | Template for individual learner marking sheets |
| `briefTemplate` | Template for the assessment brief (Collection of Work & Examination) |
| `RunLog` | Log of script runs and any errors encountered |

## Configuration

The `config` sheet controls how the script runs. Key settings include:

| Setting | Description |
|---|---|
| `learnerListRange` | Cell range where learner names are read from (`learnerList!A2:A200`) |
| `markingTemplateName` | Name of the marking sheet template tab |
| `briefTemplateName` | Name of the brief template tab |
| `markingFolder` | Relative path for exported marking sheet PDFs |
| `briefsFolder` | Relative path for exported brief PDFs |
| `splitMarkingSheets` | Toggle for generating separate marking sheets per learner |
| `markingPageRanges` | Page ranges used when splitting marking sheets |

> ⚠️ Check the `config` sheet carefully before running. Incorrect named ranges or template names will cause the script to fail.

## Requirements

- Microsoft 365 with Office Scripts enabled
- Excel for the web or a desktop version that supports Office Scripts
- Office Scripts must be enabled by your Microsoft 365 administrator

## Scripts & Macros

There are three automation scripts used with this workbook.

---

### `WordProcessingMarkingScript` — Office Script (TypeScript)

The main generation script. When run, it:

1. Reads all learner names from the `learnerList` sheet
2. Loads grade data from the `gradeSheet` for both Section 1 (Collection of Work) and Section 2 (Examination)
3. Copies the `markingTemplate` sheet once per learner, naming each tab `[Learner Name] MS`
4. Populates each marking sheet with the learner's name and their individual scores across all assessment criteria:
   - **Section 1:** File Management, Document Creation (Text), Document Creation (Graphics), Document Creation (Tables), Primary Functions & Processes
   - **Section 2:** Page Setup, Text Insertion, Character & Paragraph Formatting, Object Insertion & Formatting, Review Tools, File Management
5. Copies the `briefTemplate` sheet once per learner, naming each tab `[Learner Name] Brief`
6. If `splitMarkingSheets` is enabled in `config`, creates additional temporary split pages per learner for PDF export
7. Logs all activity and any errors to the `RunLog` sheet

---

### `Cleaner` — Office Script (TypeScript)

A cleanup script to reset the workbook after generation and export. When run, it deletes all generated sheets and keeps only the core workbook sheets:

- `config`
- `learnerList`
- `gradeSheet`
- `markingTemplate`
- `briefTemplate`
- `RunLog`

Any sheet not in that list — all generated marking sheets, briefs, and split pages — will be deleted. Run this after PDFs have been exported to return the workbook to its base state.

---

### `Export Sheets` — VBA Macro

The print/export macro embedded in the workbook. After `WordProcessingMarkingScript` has generated all learner sheets, running this macro exports each generated sheet as a PDF to a specified folder on your machine. Unlike the Office Scripts above, this macro **is** embedded in the `.xlsm` file and does not need to be added manually.

---

## Setup Instructions

Office Scripts do not transfer when a file is shared externally — they are tied to the Microsoft 365 account that created them. Follow these steps to add the script manually.

### Step 1 — Open the file

Open `wordProcMarking.xlsm` in Excel via Microsoft 365.

### Step 2 — Open the Office Scripts editor

In the ribbon, go to **Automate** → **New Script**.

### Step 3 — Add the scripts

You will need to add both Office Scripts separately. Repeat Steps 2–4 for each:

- Copy the contents of `WordProcessingMarkingScript.ts` from the [`/scripts`](./scripts) folder, paste it into the editor, and save it as `WordProcessingMarkingScript`
- Repeat for `Cleaner.ts`, saving it as `Cleaner`

### Step 4 — Save each script

Press `Ctrl + S` or click **Save** after pasting each one.

### Step 5 — Link the scripts to their buttons

There are two buttons to assign:

1. Right-click the **Generate** button → **Assign Script** → select `WordProcessingMarkingScript`
2. Right-click the **Clean** button → **Assign Script** → select `Cleaner`

The `Export Sheets` macro is already embedded in the workbook and does not need to be assigned manually.

## Folder Structure

The workbook relies on a specific folder structure to export PDFs correctly. **Do not rename or move any of these folders.**

```
Automated Sheet Generator/
├── wordProcMarking.xlsm
├── Briefs/
└── Marking Sheets/
```

- `Briefs/` — where exported brief PDFs will be saved
- `Marking Sheets/` — where exported marking sheet PDFs will be saved

The `Export Sheets` macro uses relative paths from the workbook's location, so as long as the two folders sit alongside the `.xlsm` file inside `Automated Sheet Generator`, exports will work correctly.

## Recommended Workflow

1. Populate `learnerList` and `gradeSheet` with learner data
2. Run `WordProcessingMarkingScript` to generate all marking sheets and briefs
3. Run the `Export Sheets` macro to export all generated sheets as PDFs
4. Run `Cleaner` to delete all generated sheets and reset the workbook

## Notes

- Ensure `learnerList` is fully populated before running
- Named ranges in the `config` sheet must match exactly — a mismatch in spelling will cause the script to fail silently or error
- Sheet names generated per learner are derived from learner names; avoid names that would exceed Excel's 31-character sheet name limit
- Check the `RunLog` sheet after each run to confirm all learners were processed successfully
- The `gradeSheet` provides an at-a-glance summary of Distinctions, Merits, Passes, and Unsuccessfuls across the class