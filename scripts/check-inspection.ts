/**
 * Check if an inspection file exists in SharePoint
 *
 * @description
 * Verifies whether an inspection PDF has been uploaded to SharePoint.
 * Useful for checking if the automated worker succeeded or if manual upload is needed.
 *
 * @usage
 * ```bash
 * bun scripts/check-inspection.ts "<contractor>" "<project>" [date]
 * ```
 *
 * @example
 * ```bash
 * # Check today's inspection
 * bun scripts/check-inspection.ts "ARCO" "KTEC PHX"
 *
 * # Check specific date
 * bun scripts/check-inspection.ts "ARCO" "KTEC PHX" "01.29.26"
 * ```
 *
 * @module scripts/check-inspection
 */

import { buildInspectionPath, fileExists } from "./lib";

/**
 * Prints usage instructions and exits
 */
function printUsage(): never {
  console.log(`
Usage: bun scripts/check-inspection.ts "<contractor>" "<project>" [date]

Arguments:
  contractor  Contractor name (e.g., "ARCO", "BPR COMPANIES")
  project     Project name (e.g., "KTEC PHX", "PV LOT C3")
  date        Optional. Date in MM.DD.YY format (defaults to today)

Examples:
  bun scripts/check-inspection.ts "ARCO" "KTEC PHX"
  bun scripts/check-inspection.ts "BPR COMPANIES" "PV LOT C3" "01.26.26"
`);
  process.exit(1);
}

/**
 * Main entry point
 */
async function main(): Promise<void> {
  const contractor = process.argv[2];
  const project = process.argv[3];
  const dateArg = process.argv[4];

  if (!(contractor && project)) {
    printUsage();
  }

  const path = buildInspectionPath(contractor, project, dateArg);

  console.log(`\nChecking: ${contractor} - ${project}`);
  console.log(`Path: ${path.fullPath}`);

  const exists = await fileExists(path.folderPath, path.fileName);

  if (exists) {
    console.log("\n✅ File exists in SharePoint");
    process.exit(0);
  } else {
    console.log("\n❌ NOT FOUND - needs manual upload");
    console.log("\nTo upload, run:");
    console.log(
      `bun scripts/manual-upload.ts "<report-url>" "${contractor}" "${project}"${dateArg ? ` "${dateArg}"` : ""}`
    );
    process.exit(1);
  }
}

main().catch((err) => {
  console.error("Error:", err.message);
  process.exit(1);
});
