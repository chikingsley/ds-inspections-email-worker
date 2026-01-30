/**
 * Manual upload script for failed inspections
 *
 * @description
 * Generates a PDF from a ComplianceGo report URL and uploads it to SharePoint.
 * Use this when the automated worker fails to upload an inspection.
 *
 * @usage
 * ```bash
 * bun scripts/manual-upload.ts "<report-url>" "<contractor>" "<project>" [date]
 * ```
 *
 * @example
 * ```bash
 * # Upload with today's date
 * bun scripts/manual-upload.ts "https://cdn3.compliancego.com/..." "ARCO" "KTEC PHX"
 *
 * # Upload with specific date
 * bun scripts/manual-upload.ts "https://cdn3.compliancego.com/..." "ARCO" "KTEC PHX" "01.29.26"
 * ```
 *
 * @module scripts/manual-upload
 */

import puppeteer from "puppeteer";
import { buildInspectionPath, fileExists, getSharePointClient } from "./lib";

/** Chrome executable path on macOS */
const CHROME_PATH =
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";

/**
 * Prints usage instructions and exits
 */
function printUsage(): never {
  console.log(`
Usage: bun scripts/manual-upload.ts "<report-url>" "<contractor>" "<project>" [date]

Arguments:
  report-url  ComplianceGo report URL (https://cdn3.compliancego.com/...)
  contractor  Contractor name (e.g., "ARCO", "BPR COMPANIES")
  project     Project name (e.g., "KTEC PHX", "PV LOT C3")
  date        Optional. Date in MM.DD.YY format (defaults to today)

Examples:
  bun scripts/manual-upload.ts "https://cdn3.compliancego.com/..." "ARCO" "KTEC PHX"
  bun scripts/manual-upload.ts "https://cdn3.compliancego.com/..." "BPR COMPANIES" "PV LOT C3" "01.26.26"
`);
  process.exit(1);
}

/**
 * Generates a PDF from a ComplianceGo report URL using Puppeteer.
 *
 * @param url - ComplianceGo report URL
 * @returns PDF content as Buffer
 */
async function generatePdf(url: string): Promise<Buffer> {
  console.log("Launching browser...");
  const browser = await puppeteer.launch({
    executablePath: CHROME_PATH,
    headless: true,
  });

  try {
    const page = await browser.newPage();
    console.log(`Navigating to ${url}...`);
    await page.goto(url, { waitUntil: "networkidle2", timeout: 60_000 });

    console.log("Generating PDF...");
    const pdf = await page.pdf({
      format: "letter",
      printBackground: true,
      margin: { top: "0.5in", bottom: "0.5in", left: "0.5in", right: "0.5in" },
    });

    return Buffer.from(pdf);
  } finally {
    await browser.close();
  }
}

/**
 * Main entry point
 */
async function main(): Promise<void> {
  const reportUrl = process.argv[2];
  const contractor = process.argv[3];
  const project = process.argv[4];
  const dateArg = process.argv[5];

  if (!(reportUrl && contractor && project)) {
    printUsage();
  }

  // Build the SharePoint path
  const path = buildInspectionPath(contractor, project, dateArg);

  console.log(`\nUploading to: ${path.fullPath}`);

  // Check if file already exists
  const exists = await fileExists(path.folderPath, path.fileName);
  if (exists) {
    console.log("\n✅ File already exists in SharePoint!");
    process.exit(0);
  }

  // Generate PDF
  console.log("\nGenerating PDF...");
  const pdfContent = await generatePdf(reportUrl);
  console.log(`PDF generated: ${pdfContent.length} bytes`);

  // Upload to SharePoint
  console.log("\nUploading to SharePoint...");
  const client = await getSharePointClient();
  const result = await client.upload(
    path.folderPath,
    path.fileName,
    pdfContent
  );

  console.log("\n✅ Uploaded successfully!");
  console.log(`URL: ${result.webUrl}`);
}

main().catch((err) => {
  console.error("Error:", err.message || err);
  process.exit(1);
});
