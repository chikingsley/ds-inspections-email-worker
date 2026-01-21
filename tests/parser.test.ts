import { describe, expect, it } from "bun:test";
import {
  formatFilename,
  getProjectsFolder,
  parseComplianceGoEmail,
  parseSiteName,
  shouldProcessEmail,
  siteNameToSharePointPath,
} from "../src/parser";

// Real forwarded email content from chi@desertservices.net
const REAL_FORWARDED_EMAIL = `
Best,
--

Chi Ejimofor
Project Coordinator
E: chi@desertservices.net
M: (304) 216-8700

---------- Forwarded message ----------
From: Chi Ejimofor <chi@desertservices.net>
Date: January 21, 2026 at 1:45 PM
To: inspections@desertservices.app
Subject: Fwd: Fw: ARCO - KTEC PHX - Inspection Report

---------- Forwarded message ----------
From: Logan <Logan@desertservices.net>
Date: January 21, 2026 at 12:58 PM
To: Dust Permits <dustpermits@desertservices.net>
Subject: Fw: ARCO - KTEC PHX - Inspection Report

Get Outlook for iOS
From: support@compliancego.com <support@compliancego.com> Sent: Wednesday, January 21, 2026 11:36:56 AM To: Logan <logan@desertservices.net>; jvanderbeck@arco1.com <jvanderbeck@arco1.com>; dgrumich@arco1.com <dgrumich@arco1.com> Subject: ARCO - KTEC PHX - Inspection Report

Desert Services, LLC Logo

Inspection Completion Report

Site Information

Company: Desert Services, LLC
Division:
Site Contact:
Site/Location Name: ARCO - KTEC PHX
Site/Location Address: 16741 W Northern Ave, Waddell, AZ 85355 USA
Site/Location Address: 33.5516089, -112.4195679

The New AZPDES Construction Stormwater Inspection Report inspection for this site has been completed and it can be accessed by clicking the link below.

href="https://cdn3.compliancego.com/reports/IR_Arco-KtecPhx_21Jan26-11:36AM_abc123.html"

You are receiving this email because you were designated as an inspection report recipient for this site by the site administrator.
support@compliancego.com
www.compliancego.com
© 2007 - 2026. ComplianceGo. All Rights Reserved.
`;

describe("shouldProcessEmail", () => {
  it("accepts direct email from ComplianceGo", () => {
    expect(shouldProcessEmail("support@compliancego.com")).toBe(true);
  });

  it("accepts email from chi@desertservices.net (allowed sender)", () => {
    expect(shouldProcessEmail("chi@desertservices.net")).toBe(true);
  });

  it("rejects email from random sender", () => {
    expect(shouldProcessEmail("random@gmail.com")).toBe(false);
  });

  it("rejects email from unknown desertservices address", () => {
    expect(shouldProcessEmail("logan@desertservices.net")).toBe(false);
  });
});

describe("parseComplianceGoEmail", () => {
  it("parses real forwarded email correctly", () => {
    const result = parseComplianceGoEmail(REAL_FORWARDED_EMAIL);

    expect(result).not.toBeNull();
    expect(result?.siteName).toBe("ARCO - KTEC PHX");
    expect(result?.reportUrl).toContain("cdn3.compliancego.com");
  });

  it("parses HTML format email correctly", () => {
    const htmlContent = `
      <html>
        <body>
          Site/Location Name:</span> <span class="value">ARCO - KTEC PHX</span>
          Site/Location Address:</span> <a href="#"><span>123 Main St, Phoenix AZ</span></a>
          <a href="https://cdn3.compliancego.com/reports/IR_Arco-KtecPhx_21Jan26-11:36AM_abc123.html">View Report</a>
        </body>
      </html>
    `;

    const result = parseComplianceGoEmail(htmlContent);

    expect(result).not.toBeNull();
    expect(result?.siteName).toBe("ARCO - KTEC PHX");
    expect(result?.siteAddress).toBe("123 Main St, Phoenix AZ");
    expect(result?.reportUrl).toBe(
      "https://cdn3.compliancego.com/reports/IR_Arco-KtecPhx_21Jan26-11:36AM_abc123.html"
    );
    expect(result?.inspectionDate.getDate()).toBe(21);
    expect(result?.inspectionDate.getMonth()).toBe(0); // January
    expect(result?.inspectionDate.getFullYear()).toBe(2026);
  });

  it("parses plain text format email via text parameter", () => {
    const textContent = `
      Site/Location Name: ARCO - KTEC PHX
      href="https://cdn3.compliancego.com/reports/IR_Test_15Feb26-09:00AM_xyz.html"
    `;

    const result = parseComplianceGoEmail("", textContent);

    expect(result).not.toBeNull();
    expect(result?.siteName).toBe("ARCO - KTEC PHX");
    expect(result?.inspectionDate.getMonth()).toBe(1); // February
  });

  it("finds URL in text when not in HTML", () => {
    const htmlContent =
      "<html><body>Site/Location Name:</span> <span>Test Site</span></body></html>";
    const textContent = `href="https://cdn3.compliancego.com/reports/IR_Test_15Mar26-09:00AM_xyz.html"`;

    const result = parseComplianceGoEmail(htmlContent, textContent);

    expect(result).not.toBeNull();
    expect(result?.siteName).toBe("Test Site");
    expect(result?.reportUrl).toContain("cdn3.compliancego.com");
  });

  it("returns null when site name is missing", () => {
    const htmlContent = `href="https://cdn3.compliancego.com/reports/IR_Test_15Feb26-09:00AM_xyz.html"`;

    const result = parseComplianceGoEmail(htmlContent);
    expect(result).toBeNull();
  });

  it("returns null when report URL is missing", () => {
    const htmlContent = "Site/Location Name: ARCO - KTEC PHX";

    const result = parseComplianceGoEmail(htmlContent);
    expect(result).toBeNull();
  });

  it("handles all month abbreviations correctly", () => {
    const months = [
      { abbr: "Jan", expected: 0 },
      { abbr: "Feb", expected: 1 },
      { abbr: "Mar", expected: 2 },
      { abbr: "Apr", expected: 3 },
      { abbr: "May", expected: 4 },
      { abbr: "Jun", expected: 5 },
      { abbr: "Jul", expected: 6 },
      { abbr: "Aug", expected: 7 },
      { abbr: "Sep", expected: 8 },
      { abbr: "Oct", expected: 9 },
      { abbr: "Nov", expected: 10 },
      { abbr: "Dec", expected: 11 },
    ];

    for (const { abbr, expected } of months) {
      const htmlContent = `
        Site/Location Name: Test Site
        href="https://cdn3.compliancego.com/reports/IR_Test_15${abbr}26-09:00AM_xyz.html"
      `;

      const result = parseComplianceGoEmail(htmlContent);
      expect(result?.inspectionDate.getMonth()).toBe(expected);
    }
  });
});

describe("parseSiteName", () => {
  it("parses site name with space-dash-space separator", () => {
    const result = parseSiteName("ARCO - KTEC PHX");
    expect(result).not.toBeNull();
    expect(result?.contractor).toBe("ARCO");
    expect(result?.project).toBe("KTEC PHX");
  });

  it("parses site name with dash-only separator", () => {
    const result = parseSiteName("3411 BUILDERS-ATLAS KEIRLAND");
    expect(result).not.toBeNull();
    expect(result?.contractor).toBe("3411 BUILDERS");
    expect(result?.project).toBe("ATLAS KEIRLAND");
  });

  it("prefers space-dash-space over dash-only", () => {
    const result = parseSiteName("ARCO-WEST - KTEC-PHX PROJECT");
    expect(result).not.toBeNull();
    expect(result?.contractor).toBe("ARCO-WEST");
    expect(result?.project).toBe("KTEC-PHX PROJECT");
  });

  it("returns null for invalid site name without separator", () => {
    const result = parseSiteName("InvalidSiteName");
    expect(result).toBeNull();
  });

  it("handles multi-part project names with space-dash-space", () => {
    const result = parseSiteName("ARCO - Project A - Phase 1");
    expect(result).not.toBeNull();
    expect(result?.contractor).toBe("ARCO");
    expect(result?.project).toBe("Project A - Phase 1");
  });

  it("handles multi-part project names with dash-only", () => {
    const result = parseSiteName("3411 BUILDERS-ATLAS-KEIRLAND-PHASE2");
    expect(result).not.toBeNull();
    expect(result?.contractor).toBe("3411 BUILDERS");
    expect(result?.project).toBe("ATLAS-KEIRLAND-PHASE2");
  });

  it("uppercases the contractor name", () => {
    const result = parseSiteName("arco - Some Project");
    expect(result?.contractor).toBe("ARCO");
  });
});

describe("getProjectsFolder", () => {
  it("returns A-M for contractors starting with A-M", () => {
    expect(getProjectsFolder("ARCO")).toBe("PROJECTS A-M");
    expect(getProjectsFolder("BUILDERS")).toBe("PROJECTS A-M");
    expect(getProjectsFolder("MARTIN")).toBe("PROJECTS A-M");
  });

  it("returns N-Z for contractors starting with N-Z", () => {
    expect(getProjectsFolder("NATIONAL")).toBe("PROJECTS N-Z");
    expect(getProjectsFolder("SUNBELT")).toBe("PROJECTS N-Z");
    expect(getProjectsFolder("PHOENIX")).toBe("PROJECTS N-Z");
    expect(getProjectsFolder("ZENITH")).toBe("PROJECTS N-Z");
  });

  it("returns A-M for contractors starting with numbers", () => {
    expect(getProjectsFolder("3411 BUILDERS")).toBe("PROJECTS A-M");
    expect(getProjectsFolder("123 COMPANY")).toBe("PROJECTS A-M");
  });
});

describe("siteNameToSharePointPath", () => {
  it("generates correct path for A-M contractor with year", () => {
    const path = siteNameToSharePointPath("ARCO - KTEC PHX", 2026);
    expect(path).toBe(
      "SWPPP/INSPECTIONS/PROJECTS/PROJECTS A-M/ARCO/KTEC PHX/2026"
    );
  });

  it("generates correct path for numeric contractor with year", () => {
    const path = siteNameToSharePointPath("3411 BUILDERS-ATLAS KEIRLAND", 2026);
    expect(path).toBe(
      "SWPPP/INSPECTIONS/PROJECTS/PROJECTS A-M/3411 BUILDERS/ATLAS KEIRLAND/2026"
    );
  });

  it("returns null for invalid site name", () => {
    const path = siteNameToSharePointPath("InvalidSiteName", 2026);
    expect(path).toBeNull();
  });

  it("generates correct path for N-Z contractor", () => {
    const path = siteNameToSharePointPath("PHOENIX - Desert Project", 2026);
    expect(path).toBe(
      "SWPPP/INSPECTIONS/PROJECTS/PROJECTS N-Z/PHOENIX/Desert Project/2026"
    );
  });

  it("handles multi-part project names", () => {
    const path = siteNameToSharePointPath("ARCO - Project A - Phase 1", 2025);
    expect(path).toBe(
      "SWPPP/INSPECTIONS/PROJECTS/PROJECTS A-M/ARCO/Project A - Phase 1/2025"
    );
  });

  it("uses current year when year not provided", () => {
    const path = siteNameToSharePointPath("ARCO - KTEC PHX");
    const currentYear = new Date().getFullYear();
    expect(path).toBe(
      `SWPPP/INSPECTIONS/PROJECTS/PROJECTS A-M/ARCO/KTEC PHX/${currentYear}`
    );
  });
});

describe("formatFilename", () => {
  it("formats date correctly as MM.DD.YY.pdf", () => {
    const date = new Date(2026, 0, 21); // January 21, 2026
    expect(formatFilename(date)).toBe("01.21.26.pdf");
  });

  it("pads single digit months and days", () => {
    const date = new Date(2026, 2, 5); // March 5, 2026
    expect(formatFilename(date)).toBe("03.05.26.pdf");
  });

  it("handles end of year dates", () => {
    const date = new Date(2026, 11, 31); // December 31, 2026
    expect(formatFilename(date)).toBe("12.31.26.pdf");
  });
});
