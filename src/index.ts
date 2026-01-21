/**
 * Desert Services Inspection Report Email Worker
 *
 * Receives ComplianceGo inspection emails, generates PDFs using Browser Rendering,
 * and uploads to SharePoint - all within the worker.
 *
 * Email: inspections@desertservices.app
 * Worker: inspection-router.cheez2012.workers.dev
 */

import puppeteer from "@cloudflare/puppeteer";

// =============================================================================
// Top-level regex patterns (for performance)
// =============================================================================

const SITE_NAME_HTML_REGEX =
  /Site\/Location Name:<\/span>\s*<span[^>]*>([^<]+)<\/span>/i;
const SITE_NAME_TEXT_REGEX = /Site\/Location Name:\s*([^\r\n]+)/i;
const SITE_ADDRESS_REGEX =
  /Site\/Location Address:<\/span>\s*<a[^>]*>(?:<span[^>]*>)?([^<]+)/i;
const REPORT_URL_REGEX =
  /href="(https:\/\/cdn3\.compliancego\.com\/[^"]+\.html)"/i;
const DATE_FROM_URL_REGEX = /(\d{1,2})([A-Za-z]{3})(\d{2})-/;

// =============================================================================
// Types
// =============================================================================

export interface Env {
  // Browser Rendering binding
  BROWSER: Fetcher;

  // Azure/SharePoint credentials
  AZURE_TENANT_ID: string;
  AZURE_CLIENT_ID: string;
  AZURE_CLIENT_SECRET: string;
}

interface InspectionData {
  siteName: string;
  siteAddress: string;
  reportUrl: string;
  inspectionDate: Date;
}

interface SharePointResult {
  success: boolean;
  webUrl?: string;
  error?: string;
}

// =============================================================================
// Worker Entry Point
// =============================================================================

export default {
  async email(
    message: ForwardableEmailMessage,
    env: Env,
    ctx: ExecutionContext
  ) {
    const from = message.from;
    const to = message.to;
    const subject = message.headers.get("subject") ?? "";

    console.log(`[Worker] Email from: ${from}`);
    console.log(`[Worker] Subject: ${subject}`);

    // Only process ComplianceGo inspection emails
    if (!isComplianceGoEmail(from, subject)) {
      console.log("[Worker] Not a ComplianceGo inspection, forwarding");
      await message.forward(to);
      return;
    }

    console.log("[Worker] ComplianceGo inspection detected!");

    // Parse email in try block, always forward
    try {
      const rawEmail = await streamToString(message.raw);
      const inspection = parseComplianceGoEmail(rawEmail);

      if (!inspection) {
        console.error("[Worker] Failed to parse inspection data");
        await message.forward(to);
        return;
      }

      console.log(`[Worker] Site: ${inspection.siteName}`);
      console.log(`[Worker] Date: ${inspection.inspectionDate.toISOString()}`);
      console.log(`[Worker] Report URL: ${inspection.reportUrl}`);

      // Process in background, forward immediately
      ctx.waitUntil(
        processInspection(env, inspection).catch((err) =>
          console.error(`[Worker] Processing failed: ${err}`)
        )
      );

      await message.forward(to);
      console.log(`[Worker] Email forwarded to: ${to}`);
    } catch (error) {
      console.error(`[Worker] Error: ${error}`);
      await message.forward(to);
    }
  },
};

// =============================================================================
// Main Processing
// =============================================================================

async function processInspection(
  env: Env,
  inspection: InspectionData
): Promise<void> {
  // 1. Generate PDF from ComplianceGo report
  console.log("[Worker] Generating PDF...");
  const pdfBuffer = await generatePdf(env, inspection.reportUrl);
  console.log(`[Worker] PDF generated: ${pdfBuffer.byteLength} bytes`);

  // 2. Upload to SharePoint
  console.log("[Worker] Uploading to SharePoint...");
  const folderPath = siteNameToSharePointPath(inspection.siteName);
  const fileName = formatFilename(inspection.inspectionDate);

  if (!folderPath) {
    console.error(
      `[Worker] Could not determine folder for: ${inspection.siteName}`
    );
    return;
  }

  const result = await uploadToSharePoint(env, folderPath, fileName, pdfBuffer);

  if (result.success) {
    console.log(`[Worker] Uploaded: ${result.webUrl}`);
  } else {
    console.error(`[Worker] Upload failed: ${result.error}`);
  }
}

// =============================================================================
// PDF Generation (Browser Rendering)
// =============================================================================

async function generatePdf(env: Env, url: string): Promise<Uint8Array> {
  const browser = await puppeteer.launch(env.BROWSER);

  try {
    const page = await browser.newPage();
    await page.goto(url, { waitUntil: "networkidle0" });

    const pdf = await page.pdf({
      format: "letter",
      printBackground: true,
      margin: { top: "0.5in", bottom: "0.5in", left: "0.5in", right: "0.5in" },
    });

    return pdf;
  } finally {
    await browser.close();
  }
}

// =============================================================================
// SharePoint Upload (Microsoft Graph API)
// =============================================================================

async function uploadToSharePoint(
  env: Env,
  folderPath: string,
  fileName: string,
  content: Uint8Array
): Promise<SharePointResult> {
  try {
    const token = await getGraphToken(env);

    const driveEndpoint =
      "https://graph.microsoft.com/v1.0/sites/desertservices.sharepoint.com:/sites/DataDrive:/drives";

    const drivesRes = await fetch(driveEndpoint, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!drivesRes.ok) {
      return {
        success: false,
        error: `Failed to get drives: ${drivesRes.status}`,
      };
    }

    const drivesData = (await drivesRes.json()) as {
      value: Array<{ id: string; name: string }>;
    };
    const docDrive = drivesData.value.find(
      (d) => d.name === "Documents" || d.name === "Shared Documents"
    );

    if (!docDrive) {
      return { success: false, error: "Could not find Documents drive" };
    }

    const uploadPath = `${folderPath}/${fileName}`;
    const uploadUrl = `https://graph.microsoft.com/v1.0/drives/${docDrive.id}/root:/${uploadPath}:/content`;

    const uploadRes = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/pdf",
      },
      body: content,
    });

    if (!uploadRes.ok) {
      const errorText = await uploadRes.text();
      return {
        success: false,
        error: `Upload failed: ${uploadRes.status} - ${errorText}`,
      };
    }

    const uploadData = (await uploadRes.json()) as { webUrl: string };
    return { success: true, webUrl: uploadData.webUrl };
  } catch (error) {
    return { success: false, error: String(error) };
  }
}

async function getGraphToken(env: Env): Promise<string> {
  const tokenUrl = `https://login.microsoftonline.com/${env.AZURE_TENANT_ID}/oauth2/v2.0/token`;

  const body = new URLSearchParams({
    client_id: env.AZURE_CLIENT_ID,
    client_secret: env.AZURE_CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });

  if (!res.ok) {
    throw new Error(`Token request failed: ${res.status}`);
  }

  const data = (await res.json()) as { access_token: string };
  return data.access_token;
}

// =============================================================================
// Email Parsing
// =============================================================================

function isComplianceGoEmail(from: string, subject: string): boolean {
  return (
    from.toLowerCase().includes("compliancego.com") &&
    subject.toLowerCase().includes("inspection report")
  );
}

function parseComplianceGoEmail(content: string): InspectionData | null {
  // Extract site name
  const siteMatch = content.match(SITE_NAME_HTML_REGEX);
  let siteName = siteMatch?.[1]?.trim();

  if (!siteName) {
    const altMatch = content.match(SITE_NAME_TEXT_REGEX);
    siteName = altMatch?.[1]?.trim();
  }

  if (!siteName) {
    return null;
  }

  // Extract address
  const addrMatch = content.match(SITE_ADDRESS_REGEX);
  const siteAddress = addrMatch?.[1]?.trim() ?? "";

  // Extract report URL
  const urlMatch = content.match(REPORT_URL_REGEX);
  const reportUrl = urlMatch?.[1];

  if (!reportUrl) {
    return null;
  }

  // Parse date from URL: IR_Arco-KtecPhx_21Jan26-11:36AM_uuid.html
  const dateMatch = reportUrl.match(DATE_FROM_URL_REGEX);
  let inspectionDate = new Date();

  if (dateMatch) {
    const day = Number.parseInt(dateMatch[1], 10);
    const monthStr = dateMatch[2];
    const year = 2000 + Number.parseInt(dateMatch[3], 10);
    const monthMap: Record<string, number> = {
      Jan: 0,
      Feb: 1,
      Mar: 2,
      Apr: 3,
      May: 4,
      Jun: 5,
      Jul: 6,
      Aug: 7,
      Sep: 8,
      Oct: 9,
      Nov: 10,
      Dec: 11,
    };
    inspectionDate = new Date(year, monthMap[monthStr] ?? 0, day, 12, 0, 0);
  }

  return { siteName, siteAddress, reportUrl, inspectionDate };
}

// =============================================================================
// SharePoint Path Helpers
// =============================================================================

function siteNameToSharePointPath(siteName: string): string | null {
  // "ARCO - KTEC PHX" -> "SWPPP/INSPECTIONS/PROJECTS/PROJECTS A-M/ARCO/KTEC PHX"
  const parts = siteName.split(" - ");
  if (parts.length < 2) {
    return null;
  }

  const contractor = parts[0].trim().toUpperCase();
  const project = parts.slice(1).join(" - ").trim();
  const firstLetter = contractor.charAt(0);
  const folder =
    firstLetter >= "A" && firstLetter <= "M" ? "PROJECTS A-M" : "PROJECTS N-Z";

  return `SWPPP/INSPECTIONS/PROJECTS/${folder}/${contractor}/${project}`;
}

function formatFilename(date: Date): string {
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const yy = String(date.getFullYear()).slice(-2);
  return `${mm}.${dd}.${yy}.pdf`;
}

// =============================================================================
// Utilities
// =============================================================================

async function streamToString(
  stream: ReadableStream<Uint8Array>
): Promise<string> {
  const reader = stream.getReader();
  const chunks: Uint8Array[] = [];

  let result = await reader.read();
  while (!result.done) {
    chunks.push(result.value);
    result = await reader.read();
  }

  let length = 0;
  for (const c of chunks) {
    length += c.length;
  }

  const combined = new Uint8Array(length);
  let offset = 0;
  for (const c of chunks) {
    combined.set(c, offset);
    offset += c.length;
  }

  return new TextDecoder().decode(combined);
}
