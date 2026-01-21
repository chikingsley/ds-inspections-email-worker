/**
 * Desert Services Inspection Report Email Worker
 *
 * Receives ComplianceGo inspection emails, generates PDFs using Browser Rendering,
 * and uploads to SharePoint - all within the worker.
 *
 * Email: inspections@desertservices.app
 * Worker: inspection-router.cheez2012.workers.dev
 */

import { EmailMessage } from "cloudflare:email";
import puppeteer from "@cloudflare/puppeteer";
import {
  formatFilename,
  type InspectionData,
  parseComplianceGoEmail,
  parseRawEmail,
  parseSiteName,
  shouldProcessEmail,
  siteNameToSharePointPath,
} from "./parser";

// =============================================================================
// Types
// =============================================================================

export interface Env {
  // Browser Rendering binding
  BROWSER: Fetcher;

  // Send email binding for notifications
  SEND_EMAIL: SendEmail;

  // Azure/SharePoint credentials
  AZURE_TENANT_ID: string;
  AZURE_CLIENT_ID: string;
  AZURE_CLIENT_SECRET: string;
}

interface NotificationData {
  success: boolean;
  siteName: string;
  contractor: string;
  project: string;
  fileName: string;
  folderPath: string;
  sharePointUrl?: string;
  error?: string;
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
    const subject = message.headers.get("subject") ?? "";

    // Forward destination (verified in Cloudflare)
    const forwardTo = "chi@desertservices.net";

    console.log(`[Worker] Email from: ${from}`);
    console.log(`[Worker] Subject: ${subject}`);

    try {
      // Check if sender is allowed to trigger inspection processing
      if (!shouldProcessEmail(from)) {
        console.log("[Worker] Not from allowed sender, forwarding");
        await message.forward(forwardTo);
        return;
      }

      // Read and decode raw MIME email
      const rawEmailBuffer = await streamToArrayBuffer(message.raw);
      const decoded = await parseRawEmail(rawEmailBuffer);

      console.log(
        `[Worker] Decoded email - HTML: ${decoded.html.length} chars, Text: ${decoded.text.length} chars`
      );

      const inspection = parseComplianceGoEmail(decoded.html, decoded.text);

      if (!inspection) {
        console.log("[Worker] No ComplianceGo report URL found, forwarding");
        await message.forward(forwardTo);
        return;
      }

      console.log("[Worker] ComplianceGo inspection detected!");

      console.log(`[Worker] Site: ${inspection.siteName}`);
      console.log(`[Worker] Date: ${inspection.inspectionDate.toISOString()}`);
      console.log(`[Worker] Report URL: ${inspection.reportUrl}`);

      // Process in background, forward immediately
      ctx.waitUntil(
        processInspection(env, inspection).catch((err) =>
          console.error(`[Worker] Processing failed: ${err}`)
        )
      );

      await message.forward(forwardTo);
      console.log(`[Worker] Email forwarded to: ${forwardTo}`);
    } catch (error) {
      console.error(`[Worker] Error: ${error}`);
      // Try to forward even on error
      try {
        await message.forward(forwardTo);
      } catch (forwardError) {
        console.error(`[Worker] Forward also failed: ${forwardError}`);
      }
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
  const year = inspection.inspectionDate.getFullYear();
  const folderPath = siteNameToSharePointPath(inspection.siteName, year);
  const fileName = formatFilename(inspection.inspectionDate);
  const parsed = parseSiteName(inspection.siteName);
  const contractor = parsed?.contractor ?? "UNKNOWN";
  const project = parsed?.project ?? "UNKNOWN";

  if (!folderPath) {
    console.error(
      `[Worker] Could not determine folder for: ${inspection.siteName}`
    );
    await sendNotification(env, {
      success: false,
      siteName: inspection.siteName,
      contractor,
      project,
      fileName,
      folderPath: "N/A",
      error: "Could not determine SharePoint folder path",
    });
    return;
  }

  try {
    // 1. Generate PDF from ComplianceGo report
    console.log("[Worker] Generating PDF...");
    const pdfBuffer = await generatePdf(env, inspection.reportUrl);
    console.log(`[Worker] PDF generated: ${pdfBuffer.byteLength} bytes`);

    // 2. Upload to SharePoint
    console.log("[Worker] Uploading to SharePoint...");
    const result = await uploadToSharePoint(
      env,
      folderPath,
      fileName,
      pdfBuffer
    );

    if (result.success) {
      console.log(`[Worker] Uploaded: ${result.webUrl}`);
      await sendNotification(env, {
        success: true,
        siteName: inspection.siteName,
        contractor,
        project,
        fileName,
        folderPath,
        sharePointUrl: result.webUrl,
      });
    } else {
      console.error(`[Worker] Upload failed: ${result.error}`);
      await sendNotification(env, {
        success: false,
        siteName: inspection.siteName,
        contractor,
        project,
        fileName,
        folderPath,
        error: result.error,
      });
    }
  } catch (error) {
    console.error(`[Worker] Processing error: ${error}`);
    await sendNotification(env, {
      success: false,
      siteName: inspection.siteName,
      contractor,
      project,
      fileName,
      folderPath,
      error: String(error),
    });
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
// Email Notifications
// =============================================================================

async function sendNotification(
  env: Env,
  data: NotificationData
): Promise<void> {
  const status = data.success ? "SUCCESS" : "FAILED";
  const subject = `[Inspection ${status}] ${data.contractor} - ${data.project}`;

  const body = data.success
    ? `Inspection report processed successfully!

Contractor: ${data.contractor}
Project: ${data.project}
File: ${data.fileName}
Path: ${data.folderPath}

SharePoint URL: ${data.sharePointUrl}
`
    : `Inspection report processing failed.

Contractor: ${data.contractor}
Project: ${data.project}
Site Name: ${data.siteName}
File: ${data.fileName}
Path: ${data.folderPath}

Error: ${data.error}
`;

  try {
    const messageId = `${Date.now()}.${crypto.randomUUID()}@desertservices.app`;
    const date = new Date().toUTCString();

    // RFC 5322 formatted email
    const rawEmail = [
      `From: "Inspection Router" <inspections@desertservices.app>`,
      "To: chi@desertservices.net",
      `Subject: ${subject}`,
      `Date: ${date}`,
      `Message-ID: <${messageId}>`,
      "MIME-Version: 1.0",
      "Content-Type: text/plain; charset=UTF-8",
      "",
      body,
    ].join("\r\n");

    const message = new EmailMessage(
      "inspections@desertservices.app",
      "chi@desertservices.net",
      rawEmail
    );
    await env.SEND_EMAIL.send(message);
    console.log(`[Worker] Notification sent: ${subject}`);
  } catch (error) {
    console.error(`[Worker] Failed to send notification: ${error}`);
  }
}

// =============================================================================
// Utilities
// =============================================================================

async function streamToArrayBuffer(
  stream: ReadableStream<Uint8Array>
): Promise<ArrayBuffer> {
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

  return combined.buffer;
}
