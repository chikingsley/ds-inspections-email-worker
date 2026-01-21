/**
 * Email parsing utilities for ComplianceGo inspection emails
 */

import PostalMime from "postal-mime";

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

export interface InspectionData {
  siteName: string;
  siteAddress: string;
  reportUrl: string;
  inspectionDate: Date;
}

export interface ParsedEmail {
  from: string;
  subject: string;
  html: string;
  text: string;
}

// =============================================================================
// MIME Parsing
// =============================================================================

/**
 * Parse raw MIME email and extract decoded content.
 * Handles base64, quoted-printable, and multipart emails.
 */
export async function parseRawEmail(
  rawEmail: ArrayBuffer | string
): Promise<ParsedEmail> {
  const parser = new PostalMime();
  const email = await parser.parse(rawEmail);

  return {
    from: email.from?.address ?? "",
    subject: email.subject ?? "",
    html: email.html ?? "",
    text: email.text ?? "",
  };
}

// =============================================================================
// Email Validation
// =============================================================================

// Allowed senders who can forward inspection emails
const ALLOWED_SENDERS = ["chi@desertservices.net", "compliancego.com"];

/**
 * Check if email should be processed.
 * Accepts: direct from ComplianceGo OR from allowed senders.
 * Actual validation happens in parseComplianceGoEmail (checks for valid report URL).
 */
export function shouldProcessEmail(from: string): boolean {
  const fromLower = from.toLowerCase();

  // Direct from ComplianceGo
  if (fromLower.includes("compliancego.com")) {
    return true;
  }

  // From allowed sender - let parseComplianceGoEmail determine if it's valid
  return ALLOWED_SENDERS.some((sender) =>
    fromLower.includes(sender.toLowerCase())
  );
}

// =============================================================================
// ComplianceGo Email Parsing
// =============================================================================

/**
 * Parse ComplianceGo inspection data from decoded email content.
 * Searches both HTML and text content for inspection details.
 */
export function parseComplianceGoEmail(
  html: string,
  text?: string
): InspectionData | null {
  // Combine html and text for searching (html takes priority)
  const content = html || text || "";
  const allContent = `${html}\n${text ?? ""}`;

  // Extract site name from HTML first, then text
  const siteMatch = content.match(SITE_NAME_HTML_REGEX);
  let siteName = siteMatch?.[1]?.trim();

  if (!siteName) {
    const altMatch = allContent.match(SITE_NAME_TEXT_REGEX);
    siteName = altMatch?.[1]?.trim();
  }

  if (!siteName) {
    return null;
  }

  // Extract address
  const addrMatch = content.match(SITE_ADDRESS_REGEX);
  const siteAddress = addrMatch?.[1]?.trim() ?? "";

  // Extract report URL - search in all content
  const urlMatch = allContent.match(REPORT_URL_REGEX);
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

/**
 * Parses a site name into contractor and project parts.
 * Handles both " - " and "-" separators.
 */
export function parseSiteName(
  siteName: string
): { contractor: string; project: string } | null {
  // Try " - " first (e.g., "ARCO - KTEC PHX")
  let parts = siteName.split(" - ");
  if (parts.length >= 2) {
    return {
      contractor: parts[0].trim().toUpperCase(),
      project: parts.slice(1).join(" - ").trim(),
    };
  }

  // Fall back to "-" (e.g., "3411 BUILDERS-ATLAS KEIRLAND")
  parts = siteName.split("-");
  if (parts.length >= 2) {
    return {
      contractor: parts[0].trim().toUpperCase(),
      project: parts.slice(1).join("-").trim(),
    };
  }

  return null;
}

/**
 * Determines which folder (A-M or N-Z) based on first character.
 * Numbers and A-M go to PROJECTS A-M, N-Z go to PROJECTS N-Z.
 */
export function getProjectsFolder(contractor: string): string {
  const firstChar = contractor.charAt(0).toUpperCase();
  const isNumberOrAtoM =
    (firstChar >= "0" && firstChar <= "9") ||
    (firstChar >= "A" && firstChar <= "M");
  return isNumberOrAtoM ? "PROJECTS A-M" : "PROJECTS N-Z";
}

/**
 * Converts site name and date to full SharePoint path including year subfolder.
 * Example: "3411 BUILDERS-ATLAS KEIRLAND" with 2026 date ->
 *   "SWPPP/INSPECTIONS/PROJECTS/PROJECTS A-M/3411 BUILDERS/ATLAS KEIRLAND/2026"
 */
export function siteNameToSharePointPath(
  siteName: string,
  year?: number
): string | null {
  const parsed = parseSiteName(siteName);
  if (!parsed) {
    return null;
  }

  const { contractor, project } = parsed;
  const folder = getProjectsFolder(contractor);
  const yearFolder = year ?? new Date().getFullYear();

  return `SWPPP/INSPECTIONS/PROJECTS/${folder}/${contractor}/${project}/${yearFolder}`;
}

export function formatFilename(date: Date): string {
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const yy = String(date.getFullYear()).slice(-2);
  return `${mm}.${dd}.${yy}.pdf`;
}
