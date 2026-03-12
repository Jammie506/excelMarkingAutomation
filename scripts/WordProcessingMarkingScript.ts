/**
 * Marking Sheet + Brief Generator
 * --------------------------------
 * This script:
 *  - Loads configuration from the config sheet
 *  - Loads learner names
 *  - Loads grade tables into dictionaries
 *  - Generates marking sheets
 *  - Generates briefs
 *  - Splits marking sheets into page ranges (for Power Automate PDF export)
 *  - Logs errors to RunLog
 *
 * NOTE:
 *  - This script DOES NOT export PDFs. Power Automate handles that.
 *  - This script DOES prepare temporary sheets for page splitting.
 */

async function main(workbook: ExcelScript.Workbook) {

    // ------------------------------------------------------------
    // LOGGING SETUP
    // ------------------------------------------------------------
    function log(message: string) {
        let logSheet = workbook.getWorksheet("RunLog");
        if (!logSheet) {
            logSheet = workbook.addWorksheet("RunLog");
            logSheet.getRange("A1").setValue("Timestamp");
            logSheet.getRange("B1").setValue("Message");
        }
        const lastRow = (logSheet.getUsedRange()?.getRowCount() ?? 1);
        logSheet.getCell(lastRow, 0).setValue(new Date().toISOString());
        logSheet.getCell(lastRow, 1).setValue(message);
    }

    log("Script started.");

    // ------------------------------------------------------------
    // CONFIG LOADING
    // ------------------------------------------------------------
    const configSheet = workbook.getWorksheet("config");
    if (!configSheet) throw new Error("Missing config sheet.");

    const configTable = configSheet.getTable("ConfigTable");
    if (!configTable) throw new Error("Missing ConfigTable.");

    function getConfig(key: string): string {
        const rows = configTable.getRangeBetweenHeaderAndTotal().getValues();
        for (let row of rows) {
            if (row[0] === key) return row[1] as string;
        }
        throw new Error(`Missing config key: ${key}`);
    }

    // Load config values
    const learnerListRange = getConfig("learnerListRange");
    const sectionOneRange = getConfig("sectionOneRange");
    const sectionTwoRange = getConfig("sectionTwoRange");

    const markingTemplateName = getConfig("markingTemplateName");
    const briefTemplateName = getConfig("briefTemplateName");

    const markingNameCell = getConfig("markingNameCell");
    const briefNameCell = getConfig("briefNameCell");

    const sectionOneP1 = getConfig("sectionOneP1");
    const sectionOneP2 = getConfig("sectionOneP2");
    const sectionOneP3 = getConfig("sectionOneP3");
    const sectionOneP4 = getConfig("sectionOneP4");
    const sectionOneP5 = getConfig("sectionOneP5");

    const sectionTwoP1 = getConfig("sectionTwoP1");
    const sectionTwoP2 = getConfig("sectionTwoP2");
    const sectionTwoP3 = getConfig("sectionTwoP3");
    const sectionTwoP4 = getConfig("sectionTwoP4");
    const sectionTwoP5 = getConfig("sectionTwoP5");
    const sectionTwoP6 = getConfig("sectionTwoP6");

    const splitMarkingSheets = getConfig("splitMarkingSheets") === "TRUE";
    const markingPageRanges = getConfig("markingPageRanges");

    // ------------------------------------------------------------
    // SAFE RANGE LOADER
    // ------------------------------------------------------------
    function getRangeFromAddress(address: string): ExcelScript.Range {
        const parts = address.split("!");
        if (parts.length !== 2) throw new Error(`Invalid range address: ${address}`);

        const sheetName = parts[0];
        const rangeAddress = parts[1];

        const sheet = workbook.getWorksheet(sheetName);
        if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);

        return sheet.getRange(rangeAddress);
    }

    // ------------------------------------------------------------
    // FLATTEN RANGE VALUES (no .flat() in Office Scripts)
    // ------------------------------------------------------------
    function flatten(values: (string | number | boolean)[][]): string[] {
        const result: string[] = [];
        for (let row of values) {
            for (let cell of row) {
                if (typeof cell === "string" && cell.trim() !== "") {
                    result.push(cell.trim());
                }
            }
        }
        return result;
    }

    // ------------------------------------------------------------
    // LOAD LEARNER LIST
    // ------------------------------------------------------------
    const learnerListRangeObj = getRangeFromAddress(learnerListRange);
    const learnerList: string[] = flatten(learnerListRangeObj.getValues());

    if (learnerList.length === 0) {
        log("No learners found. Ending script.");
        return;
    }

    // ------------------------------------------------------------
    // BUILD GRADE DICTIONARIES
    // ------------------------------------------------------------
    function buildGradeDict(address: string): Record<string, (string | number | boolean)[]> {
        const range = getRangeFromAddress(address);
        const values: (string | number | boolean)[][] = range.getValues();

        const dict: Record<string, (string | number | boolean)[]> = {};
        for (let row of values) {
            const name = row[0];
            if (typeof name === "string" && name.trim() !== "") {
                dict[name.trim()] = row;
            }
        }
        return dict;
    }

    const collectionDict = buildGradeDict(sectionOneRange);
    const examDict = buildGradeDict( sectionTwoRange);

    // ------------------------------------------------------------
    // PAGE RANGE PARSER
    // ------------------------------------------------------------
    function parsePageRanges(rangeString: string): { start: number, end: number }[] {
        const ranges = rangeString.split(";");
        const result: { start: number, end: number }[] = [];

        for (let r of ranges) {
            const parts = r.split("–");
            if (parts.length !== 2) {
                log(`Invalid page range: ${r}`);
                continue;
            }
            result.push({
                start: Number(parts[0]),
                end: Number(parts[1])
            });
        }
        return result;
    }

    const pageRanges = parsePageRanges(markingPageRanges);

    // ------------------------------------------------------------
    // MARKING SHEET FILLER
    // ------------------------------------------------------------
    function fillMarkingSheet(sheet: ExcelScript.Worksheet, learner: string) {

        // Collection of Work
        if (collectionDict[learner]) {
            const row = collectionDict[learner];
            //File Management
            sheet.getRange(sectionOneP1).setValue(Number(row[6])); 
            // Document Creation - Text
            sheet.getRange(sectionOneP2).setValue(Number(row[36])); 
            // Document Creation - Graphics
            sheet.getRange(sectionOneP3).setValue(Number(row[50])); 
            // Document Creation - Tables
            sheet.getRange(sectionOneP4).setValue(Number(row[66])); 
            // Primary Functions & Processes
            sheet.getRange(sectionOneP5).setValue(Number(row[73])); 
        } else {
            log(`Missing Collection of Work grade for ${learner}`);
        }

        // Examination
        if (examDict[learner]) {
            const row = examDict[learner];
            // Page Setup (Q4, 5, 25)
            sheet.getRange(sectionTwoP1).setValue(Number(row[7]) + Number(row[8]) + Number(row[28]));
            // Text Insertion (Q11, 13, 14, 19)
            sheet.getRange(sectionTwoP2).setValue(Number(row[14]) + Number(row[16]) + Number(row[17]) + Number(row[22]));
            //Character & Paragraph Formatting (Q6, 9, 10, 7, 8, 16, 18, 12)
            sheet.getRange(sectionTwoP3).setValue(Number(row[9]) + Number(row[12]) + Number(row[13]) + Number(row[10]) + Number(row[11]) + Number(row[19]) + Number(row[21]) + Number(row[15]));
            //Object Insertion & Formatting (Q17, 20, 21, 22)
            sheet.getRange(sectionTwoP4).setValue(Number(row[20]) + Number(row[23]) + Number(row[24]) + Number(row[25]));
            //Review Tools (Q3, 15, 23, 24)
            sheet.getRange(sectionTwoP5).setValue(Number(row[6]) + Number(row[18]) + Number(row[26]) + Number(row[27]));
            //File Management (Q1, 2, 26, 27, 0)
            sheet.getRange(sectionTwoP6).setValue(Number(row[4]) + Number(row[5]) + Number(row[29]) + Number(row[30]) + Number(row[3]));
        } else {
            log(`Missing Examination grade for ${learner}`);
        }
    }

    // ------------------------------------------------------------
    // BRIEF SHEET FILLER
    // ------------------------------------------------------------
    function fillBriefSheet(sheet: ExcelScript.Worksheet, learner: string) {
        const ms = workbook.getWorksheet(`${learner} MS`);
        if (!ms) {
            log(`Missing marking sheet for brief: ${learner}`);
            return;
        }
    }

    // ------------------------------------------------------------
    // PAGE SPLITTING (for Power Automate PDF export)
    // ------------------------------------------------------------
    function createSplitPages(sheet: ExcelScript.Worksheet, learner: string) {
        if (!splitMarkingSheets) return;

        for (let i = 0; i < pageRanges.length; i++) {
            const { start, end } = pageRanges[i];

            const temp = sheet.copy(ExcelScript.WorksheetPositionType.after, sheet);
            temp.setName(`${learner} MS Page ${i + 1}`);

            const used = temp.getUsedRange();
            const rowCount = used.getRowCount();

            // Hide all rows outside the page range
            for (let r = rowCount - 1; r >= 0; r--) {
                const rowIndex = r + 1; // 1-based index
                const keep = rowIndex >= start && rowIndex <= end;

                if (!keep) {
                    temp.getRange(`${rowIndex}:${rowIndex}`).delete(ExcelScript.DeleteShiftDirection.up);
                }
            }
        }
    }

    // ------------------------------------------------------------
    // GENERIC SHEET GENERATOR
    // ------------------------------------------------------------
    function generateSheets(templateName: string, nameCell: string, suffix: string, fillCallback: (sheet: ExcelScript.Worksheet, learner: string) => void) {
        const template = workbook.getWorksheet(templateName);
        if (!template) throw new Error(`Missing template: ${templateName}`);

        for (const learner of learnerList) {
            const newSheet = template.copy(ExcelScript.WorksheetPositionType.after, template);
            newSheet.setName(`${learner} ${suffix}`);

            if (newSheet.getName().includes("MS")) {
              newSheet.getRange(nameCell).setValue("Learner Name - " + learner);
            } else {
              newSheet.getRange(nameCell).setValue(learner);
            }

            

            fillCallback(newSheet, learner);

            if (suffix === "MS") {
                createSplitPages(newSheet, learner);
            }
        }
    }

    // ------------------------------------------------------------
    // RUN GENERATION
    // ------------------------------------------------------------
    generateSheets(markingTemplateName, markingNameCell, "MS", fillMarkingSheet);
    generateSheets(briefTemplateName, briefNameCell, "Brief", fillBriefSheet);

    log("Script completed successfully.");
}
