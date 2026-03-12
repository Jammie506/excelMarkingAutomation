/**
 * Cleaner Script
 * --------------
 * Keeps only the original sheets and deletes all generated ones.
 */

function main(workbook: ExcelScript.Workbook) {

    // List the sheets you want to KEEP.
    // Everything else will be deleted.
    const keepSheets = [
        "config",
        "learnerList",
        "gradeSheet",
        "markingTemplate",
        "briefTemplate",
        "RunLog" // optional — remove if you want this deleted too
    ];

    const allSheets = workbook.getWorksheets();

    for (let sheet of allSheets) {
        const name = sheet.getName();

        // If the sheet is NOT in the keep list, delete it
        if (!keepSheets.includes(name)) {
            sheet.delete();
        }
    }
}
