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
  // HTTP handler for testing
  async fetch(request: Request, env: Env): Promise<Response> {
    const url = new URL(request.url);

    // Test endpoint: /test-email?type=success or /test-email?type=failure
    if (url.pathname === "/test-email") {
      const type = url.searchParams.get("type") ?? "success";

      const testData: NotificationData = {
        success: type === "success",
        siteName: "ARCO - KTEC PHX",
        contractor: "ARCO",
        project: "KTEC PHX",
        fileName: "01.22.26.pdf",
        folderPath:
          "SWPPP/INSPECTIONS/PROJECTS/PROJECTS A-M/ARCO/KTEC PHX/2026",
        sharePointUrl:
          "https://desertservices.sharepoint.com/sites/DataDrive/Shared%20Documents/SWPPP/INSPECTIONS/PROJECTS/PROJECTS%20A-M/ARCO/KTEC%20PHX/2026",
        error:
          type === "failure"
            ? "Test error: Could not connect to SharePoint"
            : undefined,
      };

      await sendNotification(env, testData);

      return new Response(`Test ${type} email sent to chi@desertservices.net`, {
        headers: { "Content-Type": "text/plain" },
      });
    }

    // Preview endpoint: /preview?type=success or /preview?type=failure
    if (url.pathname === "/preview") {
      const type = url.searchParams.get("type") ?? "success";

      const testData: NotificationData = {
        success: type === "success",
        siteName: "ARCO - KTEC PHX",
        contractor: "ARCO",
        project: "KTEC PHX",
        fileName: "01.22.26.pdf",
        folderPath:
          "SWPPP/INSPECTIONS/PROJECTS/PROJECTS A-M/ARCO/KTEC PHX/2026",
        sharePointUrl:
          "https://desertservices.sharepoint.com/sites/DataDrive/Shared%20Documents/SWPPP/INSPECTIONS/PROJECTS/PROJECTS%20A-M/ARCO/KTEC%20PHX/2026",
        error:
          type === "failure"
            ? "Test error: Could not connect to SharePoint"
            : undefined,
      };

      return new Response(buildNotificationHtml(testData), {
        headers: { "Content-Type": "text/html" },
      });
    }

    return new Response("Inspection Router Worker", {
      headers: { "Content-Type": "text/plain" },
    });
  },

  // Email handler
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

const MAX_PDF_RETRIES = 5;

/**
 * Check if error is a rate limit or browser availability error
 */
function isRateLimitError(error: Error): boolean {
  const msg = error.message.toLowerCase();
  return (
    msg.includes("429") ||
    msg.includes("503") ||
    msg.includes("rate limit") ||
    msg.includes("no browser available")
  );
}

/**
 * Calculate exponential backoff delay
 * Base: 3s, max: 60s, with jitter to prevent thundering herd
 */
function getBackoffDelay(attempt: number, isRateLimit: boolean): number {
  // Longer base delay for rate limit errors
  const baseDelay = isRateLimit ? 10_000 : 3_000;
  const maxDelay = 60_000;

  // Exponential backoff: base * 2^(attempt-1)
  const exponentialDelay = baseDelay * Math.pow(2, attempt - 1);

  // Add jitter (±20%) to prevent all retries hitting at same time
  const jitter = exponentialDelay * 0.2 * (Math.random() * 2 - 1);

  return Math.min(exponentialDelay + jitter, maxDelay);
}

async function generatePdf(env: Env, url: string): Promise<Uint8Array> {
  let lastError: Error | null = null;
  let browser: Awaited<ReturnType<typeof puppeteer.launch>> | null = null;

  for (let attempt = 1; attempt <= MAX_PDF_RETRIES; attempt++) {
    try {
      // Browser launch is now inside try-catch to handle 503/429 errors
      console.log(`[Worker] PDF attempt ${attempt}/${MAX_PDF_RETRIES} - launching browser...`);
      browser = await puppeteer.launch(env.BROWSER);

      const page = await browser.newPage();

      // Set longer timeout for slow pages
      page.setDefaultTimeout(60_000);
      page.setDefaultNavigationTimeout(60_000);

      // Use networkidle2 (allows 2 in-flight requests) - more forgiving than networkidle0
      await page.goto(url, {
        waitUntil: "networkidle2",
        timeout: 60_000,
      });

      // Small delay to ensure page is fully rendered
      await new Promise((resolve) => setTimeout(resolve, 1000));

      const pdf = await page.pdf({
        format: "letter",
        printBackground: true,
        margin: {
          top: "0.5in",
          bottom: "0.5in",
          left: "0.5in",
          right: "0.5in",
        },
      });

      console.log(`[Worker] PDF generated successfully on attempt ${attempt}`);
      return pdf;
    } catch (error) {
      lastError = error instanceof Error ? error : new Error(String(error));
      const isRateLimit = isRateLimitError(lastError);

      console.error(
        `[Worker] PDF attempt ${attempt}/${MAX_PDF_RETRIES} failed: ${lastError.message}` +
          (isRateLimit ? " (rate limit/availability error - will use longer backoff)" : "")
      );

      if (attempt < MAX_PDF_RETRIES) {
        const delay = getBackoffDelay(attempt, isRateLimit);
        console.log(`[Worker] Waiting ${Math.round(delay / 1000)}s before retry...`);
        await new Promise((resolve) => setTimeout(resolve, delay));
      }
    } finally {
      if (browser) {
        await browser.close().catch(() => {}); // Ignore close errors
        browser = null;
      }
    }
  }

  throw lastError ?? new Error("PDF generation failed after retries");
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
    // Return folder URL (strip filename) so user can verify file is there
    const folderUrl = uploadData.webUrl.substring(
      0,
      uploadData.webUrl.lastIndexOf("/")
    );
    return { success: true, webUrl: folderUrl };
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

function buildNotificationHtml(data: NotificationData): string {
  if (data.success) {
    return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: #10b981; color: white; padding: 20px; border-radius: 8px 8px 0 0; }
    .content { background: #f9fafb; padding: 20px; border: 1px solid #e5e7eb; border-top: none; border-radius: 0 0 8px 8px; }
    .field { margin-bottom: 12px; }
    .label { font-weight: 600; color: #6b7280; font-size: 12px; text-transform: uppercase; }
    .value { font-size: 16px; margin-top: 2px; }
    .btn { display: inline-block; background: #2563eb; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; margin-top: 16px; }
    .btn:hover { background: #1d4ed8; }
  </style>
</head>
<body>
  <div class="header">
    <h2 style="margin:0;">Inspection Uploaded Successfully</h2>
  </div>
  <div class="content">
    <div class="field">
      <div class="label">Contractor</div>
      <div class="value">${data.contractor}</div>
    </div>
    <div class="field">
      <div class="label">Project</div>
      <div class="value">${data.project}</div>
    </div>
    <div class="field">
      <div class="label">File</div>
      <div class="value">${data.fileName}</div>
    </div>
    <div class="field">
      <div class="label">Folder</div>
      <div class="value">${data.folderPath}</div>
    </div>
    <a href="${data.sharePointUrl}" class="btn">View Folder in SharePoint</a>
  </div>
</body>
</html>`;
  }

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: #ef4444; color: white; padding: 20px; border-radius: 8px 8px 0 0; }
    .content { background: #f9fafb; padding: 20px; border: 1px solid #e5e7eb; border-top: none; border-radius: 0 0 8px 8px; }
    .field { margin-bottom: 12px; }
    .label { font-weight: 600; color: #6b7280; font-size: 12px; text-transform: uppercase; }
    .value { font-size: 16px; margin-top: 2px; }
    .error { background: #fef2f2; border: 1px solid #fecaca; padding: 12px; border-radius: 6px; color: #991b1b; margin-top: 16px; }
  </style>
</head>
<body>
  <div class="header">
    <h2 style="margin:0;">Inspection Upload Failed</h2>
  </div>
  <div class="content">
    <div class="field">
      <div class="label">Contractor</div>
      <div class="value">${data.contractor}</div>
    </div>
    <div class="field">
      <div class="label">Project</div>
      <div class="value">${data.project}</div>
    </div>
    <div class="field">
      <div class="label">Site Name</div>
      <div class="value">${data.siteName}</div>
    </div>
    <div class="field">
      <div class="label">File</div>
      <div class="value">${data.fileName}</div>
    </div>
    <div class="field">
      <div class="label">Folder</div>
      <div class="value">${data.folderPath}</div>
    </div>
    <div class="error">
      <strong>Error:</strong> ${data.error}
    </div>
  </div>
</body>
</html>`;
}

async function sendNotification(
  env: Env,
  data: NotificationData
): Promise<void> {
  const status = data.success ? "SUCCESS" : "FAILED";
  const subject = `[Inspection ${status}] ${data.contractor} - ${data.project}`;
  const htmlBody = buildNotificationHtml(data);

  try {
    const messageId = `${Date.now()}.${crypto.randomUUID()}@desertservices.app`;
    const date = new Date().toUTCString();

    // RFC 5322 formatted email with HTML content
    const rawEmail = [
      `From: "Inspection Router" <inspections@desertservices.app>`,
      "To: chi@desertservices.net",
      `Subject: ${subject}`,
      `Date: ${date}`,
      `Message-ID: <${messageId}>`,
      "MIME-Version: 1.0",
      "Content-Type: text/html; charset=UTF-8",
      "",
      htmlBody,
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
