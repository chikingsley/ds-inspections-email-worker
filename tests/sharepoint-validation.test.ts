/**
 * SharePoint Path Validation Tests
 *
 * Validates the worker's path generation logic against the actual
 * SharePoint folder structure synced to SQLite.
 *
 * This is a "dry run" test that ensures the worker can correctly
 * route inspection reports to any existing project folder.
 */

import { Database } from "bun:sqlite";
import { describe, expect, it } from "bun:test";
import {
  getProjectsFolder,
  parseSiteName,
  siteNameToSharePointPath,
} from "../src/parser";

const DB_PATH = "./sharepoint-inspections-folders-sync/inspections.db";

interface ProjectRow {
  contractor: string;
  project: string;
  folder_path: string;
}

/**
 * Build a ComplianceGo-style site name from contractor and project
 * Format: "CONTRACTOR - PROJECT"
 */
function buildSiteName(contractor: string, project: string): string {
  return `${contractor} - ${project}`;
}

/**
 * Extract expected base path (without year suffix) from siteNameToSharePointPath result
 */
function getBasePathFromGenerated(generatedPath: string | null): string | null {
  if (!generatedPath) return null;
  // Remove trailing year folder (e.g., /2026)
  return generatedPath.replace(/\/\d{4}$/, "");
}

describe("SharePoint Path Validation", () => {
  const db = new Database(DB_PATH, { readonly: true });

  // Get all active projects (excluding completed/archive folders)
  const activeProjects = db
    .query<ProjectRow, []>(
      `
      SELECT c.name as contractor, p.name as project, p.folder_path
      FROM contractors c
      JOIN projects p ON c.id = p.contractor_id
      WHERE p.name NOT LIKE 'ZZZ%'
        AND p.name NOT LIKE 'ZZZZ%'
      ORDER BY c.name, p.name
    `
    )
    .all();

  const allProjects = db
    .query<ProjectRow, []>(
      `
      SELECT c.name as contractor, p.name as project, p.folder_path
      FROM contractors c
      JOIN projects p ON c.id = p.contractor_id
      ORDER BY c.name, p.name
    `
    )
    .all();

  it(`should have projects to validate (found ${activeProjects.length} active)`, () => {
    expect(activeProjects.length).toBeGreaterThan(0);
  });

  describe("A-M / N-Z Folder Split", () => {
    const amProjects = allProjects.filter((p) =>
      p.folder_path.includes("PROJECTS A-M")
    );
    const nzProjects = allProjects.filter((p) =>
      p.folder_path.includes("PROJECTS N-Z")
    );

    it("should correctly identify A-M contractors", () => {
      const mismatches: string[] = [];

      for (const project of amProjects) {
        const folder = getProjectsFolder(project.contractor);
        if (folder !== "PROJECTS A-M") {
          mismatches.push(`${project.contractor}: expected A-M, got ${folder}`);
        }
      }

      if (mismatches.length > 0) {
        console.log("A-M mismatches:", mismatches.slice(0, 10));
      }
      expect(mismatches).toEqual([]);
    });

    it("should correctly identify N-Z contractors", () => {
      const mismatches: string[] = [];

      for (const project of nzProjects) {
        const folder = getProjectsFolder(project.contractor);
        if (folder !== "PROJECTS N-Z") {
          mismatches.push(`${project.contractor}: expected N-Z, got ${folder}`);
        }
      }

      if (mismatches.length > 0) {
        console.log("N-Z mismatches:", mismatches.slice(0, 10));
      }
      expect(mismatches).toEqual([]);
    });
  });

  describe("Site Name Parsing", () => {
    it("should parse all active project site names correctly", () => {
      const failures: string[] = [];

      for (const project of activeProjects) {
        const siteName = buildSiteName(project.contractor, project.project);
        const parsed = parseSiteName(siteName);

        if (!parsed) {
          failures.push(`Failed to parse: "${siteName}"`);
          continue;
        }

        // The contractor should match (case-insensitive since we uppercase)
        if (
          parsed.contractor.toUpperCase() !== project.contractor.toUpperCase()
        ) {
          failures.push(
            `Contractor mismatch for "${siteName}": got "${parsed.contractor}", expected "${project.contractor}"`
          );
        }

        // The project should match exactly
        if (parsed.project !== project.project) {
          failures.push(
            `Project mismatch for "${siteName}": got "${parsed.project}", expected "${project.project}"`
          );
        }
      }

      if (failures.length > 0) {
        console.log(`\nParsing failures (${failures.length}):`);
        for (const f of failures.slice(0, 20)) {
          console.log(`  - ${f}`);
        }
      }

      expect(failures.length).toBe(0);
    });
  });

  describe("Full Path Generation", () => {
    it("should generate correct paths for all active projects", () => {
      const mismatches: Array<{
        siteName: string;
        expected: string;
        generated: string | null;
      }> = [];

      for (const project of activeProjects) {
        const siteName = buildSiteName(project.contractor, project.project);
        const generatedPath = siteNameToSharePointPath(siteName, 2026);
        const basePath = getBasePathFromGenerated(generatedPath);

        if (basePath !== project.folder_path) {
          mismatches.push({
            siteName,
            expected: project.folder_path,
            generated: basePath,
          });
        }
      }

      if (mismatches.length > 0) {
        console.log(`\nPath mismatches (${mismatches.length}):`);
        for (const m of mismatches.slice(0, 20)) {
          console.log(`  Site: "${m.siteName}"`);
          console.log(`    Expected:  ${m.expected}`);
          console.log(`    Generated: ${m.generated}`);
        }
      }

      // Report success rate
      const successRate = (
        ((activeProjects.length - mismatches.length) / activeProjects.length) *
        100
      ).toFixed(1);
      console.log(
        `\nPath generation success: ${activeProjects.length - mismatches.length}/${activeProjects.length} (${successRate}%)`
      );

      expect(mismatches.length).toBe(0);
    });
  });

  describe("Edge Cases from Real Data", () => {
    // Find projects with special characters
    const specialCharProjects = activeProjects.filter(
      (p) =>
        p.contractor.includes("&") ||
        p.contractor.includes("(") ||
        p.contractor.includes("-") ||
        p.project.includes("&") ||
        p.project.includes("(") ||
        p.project.includes("-")
    );

    it(`should handle ${specialCharProjects.length} projects with special characters`, () => {
      const failures: string[] = [];

      for (const project of specialCharProjects) {
        const siteName = buildSiteName(project.contractor, project.project);
        const generatedPath = siteNameToSharePointPath(siteName, 2026);
        const basePath = getBasePathFromGenerated(generatedPath);

        if (basePath !== project.folder_path) {
          failures.push(
            `"${siteName}" => ${basePath} (expected: ${project.folder_path})`
          );
        }
      }

      if (failures.length > 0) {
        console.log(`\nSpecial character failures (${failures.length}):`);
        for (const f of failures.slice(0, 15)) {
          console.log(`  ${f}`);
        }
      }

      expect(failures.length).toBe(0);
    });

    // Find contractors starting with numbers
    const numericContractors = activeProjects.filter((p) =>
      /^\d/.test(p.contractor)
    );

    it(`should handle ${numericContractors.length} projects with numeric contractor names`, () => {
      const failures: string[] = [];

      for (const project of numericContractors) {
        const siteName = buildSiteName(project.contractor, project.project);
        const generatedPath = siteNameToSharePointPath(siteName, 2026);
        const basePath = getBasePathFromGenerated(generatedPath);

        if (basePath !== project.folder_path) {
          failures.push(
            `"${siteName}" => ${basePath} (expected: ${project.folder_path})`
          );
        }
      }

      if (failures.length > 0) {
        console.log("\nNumeric contractor failures:");
        for (const f of failures) {
          console.log(`  ${f}`);
        }
      }

      expect(failures.length).toBe(0);
    });
  });

  db.close();
});
