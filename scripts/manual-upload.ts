/**
 * Manual upload script for failed inspections
 * Uses the existing SharePointClient with credentials from .env
 *
 * Usage: bun scripts/manual-upload.ts <report-url> <contractor> <project>
 * Example: bun scripts/manual-upload.ts "https://cdn3.compliancego.com/..." "BPR COMPANIES" "PV LOT C3"
 */

import puppeteer from "puppeteer";
import { SharePointClient } from "../sharepoint-inspections-folders-sync/client";

// Load env from the sync folder
const envPath = "./sharepoint-inspections-folders-sync/.env";
const envFile = Bun.file(envPath);
if (await envFile.exists()) {
  const envContent = await envFile.text();
  for (const line of envContent.split("\n")) {
    const trimmed = line.trim();
    if (trimmed && !trimmed.startsWith("#")) {
      const eqIndex = trimmed.indexOf("=");
      if (eqIndex > 0) {
        const key = trimmed.slice(0, eqIndex);
        const value = trimmed.slice(eqIndex + 1);
        process.env[key] = value;
      }
    }
  }
}

function getProjectsFolder(contractor: string): string {
  const firstChar = contractor.charAt(0).toUpperCase();
  const isNumberOrAtoM =
    (firstChar >= "0" && firstChar <= "9") ||
    (firstChar >= "A" && firstChar <= "M");
  return isNumberOrAtoM ? "PROJECTS A-M" : "PROJECTS N-Z";
}

async function generatePdf(url: string): Promise<Buffer> {
  console.log("Launching browser...");
  const browser = await puppeteer.launch({
    executablePath:
      "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
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

async function main() {
  const reportUrl = process.argv[2];
  const contractor = process.argv[3];
  const project = process.argv[4];

  if (!(reportUrl && contractor && project)) {
    console.log(
      "Usage: bun scripts/manual-upload.ts <report-url> <contractor> <project>"
    );
    console.log(
      'Example: bun scripts/manual-upload.ts "https://cdn3.compliancego.com/..." "BPR COMPANIES" "PV LOT C3"'
    );
    process.exit(1);
  }

  // Validate env vars
  const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET } = process.env;
  if (!(AZURE_TENANT_ID && AZURE_CLIENT_ID && AZURE_CLIENT_SECRET)) {
    console.error("Missing Azure credentials in environment");
    process.exit(1);
  }

  const year = new Date().getFullYear();
  const date = new Date();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const yy = String(date.getFullYear()).slice(-2);
  const fileName = `${mm}.${dd}.${yy}.pdf`;

  const folder = getProjectsFolder(contractor);
  const folderPath = `SWPPP/INSPECTIONS/PROJECTS/${folder}/${contractor}/${project}/${year}`;

  console.log(`\nUploading to: ${folderPath}/${fileName}`);

  // Initialize SharePoint client
  const client = new SharePointClient({
    azureTenantId: AZURE_TENANT_ID,
    azureClientId: AZURE_CLIENT_ID,
    azureClientSecret: AZURE_CLIENT_SECRET,
  });

  // Check if file already exists
  try {
    const existingFiles = await client.listFiles(folderPath);
    const exists = existingFiles.some((f) => f.name === fileName);
    if (exists) {
      console.log("\n✅ File already exists in SharePoint!");
      process.exit(0);
    }
  } catch {
    // Folder might not exist yet, continue with upload
  }

  console.log("\nGenerating PDF...");
  const pdfContent = await generatePdf(reportUrl);
  console.log(`PDF generated: ${pdfContent.length} bytes`);

  console.log("\nUploading to SharePoint...");
  const result = await client.upload(folderPath, fileName, pdfContent);

  console.log("\n✅ Uploaded successfully!");
  console.log(`URL: ${result.webUrl}`);
}

main().catch((err) => {
  console.error("Error:", err);
  process.exit(1);
});
