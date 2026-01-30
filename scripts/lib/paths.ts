/**
 * SharePoint path utilities for inspection files
 * @module scripts/lib/paths
 */

/** Inspection file location info */
export interface InspectionPath {
  /** Full SharePoint folder path */
  folderPath: string;
  /** Filename in MM.DD.YY.pdf format */
  fileName: string;
  /** Full path including filename */
  fullPath: string;
}

/** Base path for all inspection files */
const INSPECTIONS_BASE = "SWPPP/INSPECTIONS/PROJECTS";

/** Date format regex (MM.DD.YY) */
const DATE_FORMAT_REGEX = /^\d{2}\.\d{2}\.\d{2}$/;

/**
 * Determines the projects folder (A-M or N-Z) based on contractor name.
 * Numbers and A-M go to "PROJECTS A-M", N-Z go to "PROJECTS N-Z".
 *
 * @param contractor - The contractor name
 * @returns "PROJECTS A-M" or "PROJECTS N-Z"
 * @example
 * ```ts
 * getProjectsFolder("ARCO") // "PROJECTS A-M"
 * getProjectsFolder("SD CRANE") // "PROJECTS N-Z"
 * getProjectsFolder("3411 BUILDERS") // "PROJECTS A-M"
 * ```
 */
export function getProjectsFolder(contractor: string): string {
  const firstChar = contractor.charAt(0).toUpperCase();
  const isNumberOrAtoM =
    (firstChar >= "0" && firstChar <= "9") ||
    (firstChar >= "A" && firstChar <= "M");
  return isNumberOrAtoM ? "PROJECTS A-M" : "PROJECTS N-Z";
}

/**
 * Formats a date as MM.DD.YY for use in filenames.
 *
 * @param date - Date to format (defaults to today)
 * @returns Formatted date string
 * @example
 * ```ts
 * formatDateForFilename(new Date(2026, 0, 29)) // "01.29.26"
 * ```
 */
export function formatDateForFilename(date: Date = new Date()): string {
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const yy = String(date.getFullYear()).slice(-2);
  return `${mm}.${dd}.${yy}`;
}

/**
 * Parses a date string in MM.DD.YY format.
 *
 * @param dateStr - Date string in MM.DD.YY format
 * @returns Parsed Date object, or null if invalid
 * @example
 * ```ts
 * parseDateString("01.29.26") // Date for Jan 29, 2026
 * parseDateString("invalid") // null
 * ```
 */
export function parseDateString(dateStr: string): Date | null {
  if (!DATE_FORMAT_REGEX.test(dateStr)) {
    return null;
  }

  const [mm, dd, yy] = dateStr.split(".").map(Number);
  return new Date(2000 + yy, mm - 1, dd);
}

/**
 * Builds the full SharePoint path for an inspection file.
 *
 * @param contractor - Contractor name (e.g., "ARCO")
 * @param project - Project name (e.g., "KTEC PHX")
 * @param dateStr - Date in MM.DD.YY format, or Date object (defaults to today)
 * @returns Inspection path info
 * @example
 * ```ts
 * const path = buildInspectionPath("ARCO", "KTEC PHX", "01.29.26");
 * // {
 * //   folderPath: "SWPPP/INSPECTIONS/PROJECTS/PROJECTS A-M/ARCO/KTEC PHX/2026",
 * //   fileName: "01.29.26.pdf",
 * //   fullPath: "SWPPP/INSPECTIONS/PROJECTS/PROJECTS A-M/ARCO/KTEC PHX/2026/01.29.26.pdf"
 * // }
 * ```
 */
export function buildInspectionPath(
  contractor: string,
  project: string,
  dateStr?: string | Date
): InspectionPath {
  let year: number;
  let formattedDate: string;

  if (dateStr instanceof Date) {
    year = dateStr.getFullYear();
    formattedDate = formatDateForFilename(dateStr);
  } else if (dateStr) {
    const parsed = parseDateString(dateStr);
    if (!parsed) {
      throw new Error(`Invalid date format: ${dateStr}. Expected MM.DD.YY`);
    }
    year = parsed.getFullYear();
    formattedDate = dateStr;
  } else {
    const now = new Date();
    year = now.getFullYear();
    formattedDate = formatDateForFilename(now);
  }

  const folder = getProjectsFolder(contractor);
  const folderPath = `${INSPECTIONS_BASE}/${folder}/${contractor}/${project}/${year}`;
  const fileName = `${formattedDate}.pdf`;

  return {
    folderPath,
    fileName,
    fullPath: `${folderPath}/${fileName}`,
  };
}
