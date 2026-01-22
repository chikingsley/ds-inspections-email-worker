/**
 * SharePoint folder paths - centralized constants
 *
 * This file defines all SharePoint paths used across the codebase.
 * Update this file when folder structure changes.
 *
 * @see /docs/sharepoint-structure.md - Current structure documentation
 * @see /docs/planning/sharepoint-project-structure-v2.md - Proposed new structure
 *
 * Last Updated: 2025-12-19
 */

// =============================================================================
// NEW STRUCTURE (Proposed - use for new projects)
// =============================================================================

/**
 * New project folder structure paths
 * @see /docs/planning/sharepoint-project-structure-v2.md
 */
export const PROJECT_PATHS = {
  /** Root folder for all projects */
  ROOT: "Projects",

  /** Projects with estimates submitted, waiting on award */
  BIDDING: "Projects/00-Bidding",

  /** Contract signed, work in progress */
  ACTIVE: "Projects/01-Active",

  /** Work done, closeout complete */
  COMPLETED: "Projects/02-Completed",

  /** Didn't win bid, kept for reference */
  LOST: "Projects/03-Lost",
} as const;

/**
 * Subfolders within each project folder
 * Use with buildProjectPath() helper
 */
export const PROJECT_SUBFOLDERS = {
  ESTIMATES: "01-Estimates",
  CONTRACTS: "02-Contracts",
  PERMITS: "03-Permits",
  SWPPP: "04-SWPPP",
  INSPECTIONS: "05-Inspections",
  INSPECTION_PHOTOS: "05-Inspections/Photos",
  BILLING: "06-Billing",
  CLOSEOUT: "07-Closeout",
} as const;

/**
 * Project status values (for metadata and folder placement)
 */
export type ProjectStatus = "Bidding" | "Active" | "Completed" | "Lost";

const STATUS_PATH_MAP: Record<ProjectStatus, string> = {
  Bidding: PROJECT_PATHS.BIDDING,
  Active: PROJECT_PATHS.ACTIVE,
  Completed: PROJECT_PATHS.COMPLETED,
  Lost: PROJECT_PATHS.LOST,
} as const;

/**
 * Build a full path to a project folder
 *
 * @example
 * buildProjectPath('Active', 'Kiwanis Playground - Caliente Construction')
 * // => 'Projects/01-Active/Kiwanis Playground - Caliente Construction'
 *
 * buildProjectPath('Active', 'Kiwanis Playground - Caliente Construction', 'INSPECTIONS')
 * // => 'Projects/01-Active/Kiwanis Playground - Caliente Construction/05-Inspections'
 */
export function buildProjectPath(
  status: ProjectStatus,
  projectFolderName: string,
  subfolder?: keyof typeof PROJECT_SUBFOLDERS
): string {
  const basePath = `${STATUS_PATH_MAP[status]}/${projectFolderName}`;

  return subfolder ? `${basePath}/${PROJECT_SUBFOLDERS[subfolder]}` : basePath;
}

/**
 * Build a project folder name from project and contractor
 *
 * @example
 * buildProjectFolderName('Kiwanis Playground', 'Caliente Construction')
 * // => 'Kiwanis Playground - Caliente Construction'
 */
export function buildProjectFolderName(
  projectName: string,
  contractorName: string
): string {
  return `${projectName} - ${contractorName}`;
}

// =============================================================================
// LEGACY STRUCTURE (Current - for reading existing data)
// =============================================================================

/**
 * Legacy SharePoint paths - existing folder structure
 * Use these for reading/querying existing data
 *
 * @deprecated For new projects, use PROJECT_PATHS instead
 * @see /docs/sharepoint-structure.md
 */
export const LEGACY_PATHS = {
  // -------------------------------------------------------------------------
  // SWPPP Inspections (where projects are listed for inspections)
  // -------------------------------------------------------------------------

  /** Active inspections - contractors A-M */
  INSPECTIONS_ACTIVE_AM: "SWPPP/INSPECTIONS/PROJECTS/PROJECTS A-M",

  /** Active inspections - contractors N-Z */
  INSPECTIONS_ACTIVE_NZ: "SWPPP/INSPECTIONS/PROJECTS/PROJECTS N-Z",

  /** Completed/archived inspections */
  INSPECTIONS_COMPLETED:
    "SWPPP/INSPECTIONS/PROJECTS/-ZZZZ COMPLETED INSPECTIONS",

  /** Inspection templates and master docs */
  INSPECTIONS_MASTER: "SWPPP/INSPECTIONS/PROJECTS/1 1 THE SWPPP MASTER",

  // -------------------------------------------------------------------------
  // SWPPP Books (where SWPPP documentation is stored)
  // -------------------------------------------------------------------------

  /** Completed SWPPP books (2000+ projects) */
  SWPPP_BOOKS_COMPLETED: "SWPPP/SWPPP Book/SWPPP Books Completed",

  /** SWPPP books pending NOI submission */
  SWPPP_BOOKS_PENDING_NOI: "SWPPP/SWPPP Book/SWPPP Books Waiting for the NOI",

  /** Cancelled SWPPP books */
  SWPPP_BOOKS_CANCELLED: "SWPPP/SWPPP Book/Books Cancelled",

  /** SWPPP reference materials */
  SWPPP_BOOK_INFO: "SWPPP/SWPPP Book/SWPPP Book Info",

  /** SWPPP narrative templates */
  SWPPP_NARRATIVE_TEMPLATES:
    "SWPPP/SWPPP Book/SWPPP Narrative Examples and Templates",

  // -------------------------------------------------------------------------
  // Other SWPPP folders
  // -------------------------------------------------------------------------

  /** Daily inspection schedules (by date) */
  SCHEDULES: "SWPPP/Schedules",

  /** Loose inspection photos (legacy - unorganized) */
  PICTURES: "SWPPP/Pictures",

  /** Best Management Practices docs */
  BMPS: "SWPPP/BMP's",

  /** NOI documents */
  NOIS: "SWPPP/NOI's",

  // -------------------------------------------------------------------------
  // Billing / AIA
  // -------------------------------------------------------------------------

  /** AIA billing documentation by contractor */
  AIA_JOBS: "AIA JOBS",

  // -------------------------------------------------------------------------
  // Other top-level folders
  // -------------------------------------------------------------------------

  /** Project plans (PDFs) */
  PLANS: "Plans",

  /** Permit documents */
  PERMITS: "Permits",

  /** Customer-specific project folders */
  CUSTOMER_PROJECTS: "Customer Projects",

  /** Accounting records */
  ACCOUNTING: "Accounting",

  /** Insurance certificates */
  INSURANCE_CERTS: "Insurance Certs",

  /** Prequalification documents */
  PREQUALS: "Prequals",

  /** Price books */
  PRICE_BOOKS: "Price Books",
} as const;

/**
 * Get the inspection folder path for a contractor
 * (Legacy structure uses A-M and N-Z split)
 *
 * @example
 * getLegacyInspectionPath('AR MAYS')
 * // => 'SWPPP/INSPECTIONS/PROJECTS/PROJECTS A-M/AR MAYS'
 */
export function getLegacyInspectionPath(contractorName: string): string {
  const firstLetter = contractorName.charAt(0).toUpperCase();
  const isInFirstHalf = firstLetter >= "A" && firstLetter <= "M";
  const basePath = isInFirstHalf
    ? LEGACY_PATHS.INSPECTIONS_ACTIVE_AM
    : LEGACY_PATHS.INSPECTIONS_ACTIVE_NZ;

  return `${basePath}/${contractorName}`;
}

type SwpppBookStatus = "completed" | "pending_noi" | "cancelled";

const SWPPP_BOOK_PATH_MAP: Record<SwpppBookStatus, string> = {
  completed: LEGACY_PATHS.SWPPP_BOOKS_COMPLETED,
  pending_noi: LEGACY_PATHS.SWPPP_BOOKS_PENDING_NOI,
  cancelled: LEGACY_PATHS.SWPPP_BOOKS_CANCELLED,
} as const;

/**
 * Get the SWPPP book folder path (legacy naming: "Contractor (Project)")
 *
 * @example
 * getLegacySwpppBookPath('AR Mays Construction', 'Kiwanis Playground', 'completed')
 * // => 'SWPPP/SWPPP Book/SWPPP Books Completed/AR Mays Construction (Kiwanis Playground)'
 */
export function getLegacySwpppBookPath(
  contractorName: string,
  projectName: string,
  status: SwpppBookStatus = "completed"
): string {
  const folderName = `${contractorName} (${projectName})`;
  return `${SWPPP_BOOK_PATH_MAP[status]}/${folderName}`;
}

// =============================================================================
// FILE NAMING HELPERS
// =============================================================================

/**
 * Document types for file naming
 */
export const DOCUMENT_TYPES = {
  ESTIMATE: "Estimate",
  PLANS: "Plans",
  CONTRACT: "Contract",
  PO: "PO",
  INSURANCE: "Insurance-COI",
  SCHEDULE_OF_VALUES: "ScheduleOfValues",
  NOI: "NOI",
  NDC: "NDC",
  DUST_PERMIT: "DustPermit",
  DUST_APPLICATION: "DustApplication",
  SWPPP_PLAN: "SWPPP-Plan",
  NARRATIVE: "Narrative",
  INVOICE: "Invoice",
  CHANGE_ORDER: "ChangeOrder",
  LIEN_WAIVER: "Lien-Waiver",
} as const;

/**
 * Format a date for file naming (YYYY-MM-DD)
 */
export function formatDateForFilename(date: Date = new Date()): string {
  return date.toISOString().split("T")[0] ?? "";
}

export interface FilenameOptions {
  type: keyof typeof DOCUMENT_TYPES;
  date: string;
  identifier?: string;
  modifier?: string;
  extension?: string;
}

/**
 * Build a standardized filename
 *
 * @example
 * buildFilename({ type: 'ESTIMATE', date: '2025-12-15', identifier: 'SWPPP', modifier: 'v2' })
 * // => 'Estimate-SWPPP-2025-12-15-v2.pdf'
 *
 * buildFilename({ type: 'INVOICE', date: '2025-12-08', identifier: 'IV086336' })
 * // => 'Invoice-IV086336-2025-12-08.pdf'
 */
export function buildFilename(options: FilenameOptions): string {
  const { type, date, identifier, modifier, extension = "pdf" } = options;
  const parts = [DOCUMENT_TYPES[type], identifier, date, modifier].filter(
    Boolean
  );

  return `${parts.join("-")}.${extension}`;
}

/**
 * Build an inspection report filename
 *
 * @example
 * buildInspectionFilename(new Date('2025-07-22'))
 * // => '2025-07-22.pdf'
 *
 * buildInspectionFilename(new Date('2025-09-28'), true)
 * // => '2025-09-28-Rain.pdf'
 */
export function buildInspectionFilename(
  date: Date,
  isRainEvent = false
): string {
  const dateStr = formatDateForFilename(date);
  return isRainEvent ? `${dateStr}-Rain.pdf` : `${dateStr}.pdf`;
}
