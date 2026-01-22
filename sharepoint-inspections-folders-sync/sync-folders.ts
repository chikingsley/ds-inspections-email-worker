/**
 * Sync SharePoint inspection folder structure to SQLite
 *
 * Creates/updates a local database with:
 * - Contractors
 * - Projects (under each contractor)
 * - Inspection files (with parsed dates)
 */
import { Database } from "bun:sqlite";
import { SharePointClient } from "./client";

const DB_PATH = "inspections.db";

// Initialize database schema
function initDb(db: Database): void {
  db.run(`
    CREATE TABLE IF NOT EXISTS contractors (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL UNIQUE,
      folder_path TEXT NOT NULL,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS projects (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      contractor_id INTEGER NOT NULL,
      name TEXT NOT NULL,
      folder_path TEXT NOT NULL,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP,
      FOREIGN KEY (contractor_id) REFERENCES contractors(id),
      UNIQUE(contractor_id, name)
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS inspections (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      project_id INTEGER NOT NULL,
      filename TEXT NOT NULL,
      inspection_date TEXT,
      year INTEGER,
      is_rain_event INTEGER DEFAULT 0,
      file_size INTEGER,
      sharepoint_id TEXT,
      web_url TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP,
      FOREIGN KEY (project_id) REFERENCES projects(id),
      UNIQUE(project_id, filename)
    )
  `);

  // Indexes for common queries
  db.run(
    "CREATE INDEX IF NOT EXISTS idx_inspections_date ON inspections(inspection_date)"
  );
  db.run(
    "CREATE INDEX IF NOT EXISTS idx_inspections_project ON inspections(project_id)"
  );
  db.run(
    "CREATE INDEX IF NOT EXISTS idx_projects_contractor ON projects(contractor_id)"
  );
}

// Parse inspection filename to extract date and rain event status
// Formats: "01.07.26.pdf", "11.20.25 Rain.pdf", "04.02.24 - Rain & Reg.pdf"
function parseInspectionFilename(filename: string): {
  date: string | null;
  year: number | null;
  isRain: boolean;
} {
  const isRain = /rain/i.test(filename);

  // Match date pattern: MM.DD.YY
  const dateMatch = filename.match(/(\d{2})\.(\d{2})\.(\d{2})/);
  if (!dateMatch) {
    return { date: null, year: null, isRain };
  }

  const [, month, day, shortYear] = dateMatch;
  const year =
    Number(shortYear) > 50
      ? 1900 + Number(shortYear)
      : 2000 + Number(shortYear);
  const date = `${year}-${month}-${day}`;

  return { date, year, isRain };
}

async function syncFolders(): Promise<void> {
  const db = new Database(DB_PATH);
  initDb(db);

  const client = new SharePointClient({
    azureTenantId: process.env.AZURE_TENANT_ID!,
    azureClientId: process.env.AZURE_CLIENT_ID!,
    azureClientSecret: process.env.AZURE_CLIENT_SECRET!,
  });

  const AM_PATH = "SWPPP/INSPECTIONS/PROJECTS/PROJECTS A-M";
  const NZ_PATH = "SWPPP/INSPECTIONS/PROJECTS/PROJECTS N-Z";

  // Get all contractor folders
  console.log("Fetching contractor folders...");
  const [amContractors, nzContractors] = await Promise.all([
    client.listFiles(AM_PATH),
    client.listFiles(NZ_PATH),
  ]);

  const allContractors = [
    ...amContractors
      .filter((i) => i.folder)
      .map((i) => ({ name: i.name, path: `${AM_PATH}/${i.name}` })),
    ...nzContractors
      .filter((i) => i.folder)
      .map((i) => ({ name: i.name, path: `${NZ_PATH}/${i.name}` })),
  ];

  console.log(`Found ${allContractors.length} contractors`);

  // Prepared statements
  const insertContractor = db.prepare(`
    INSERT OR IGNORE INTO contractors (name, folder_path) VALUES (?, ?)
  `);
  const getContractorId = db.prepare(
    "SELECT id FROM contractors WHERE name = ?"
  );
  const insertProject = db.prepare(`
    INSERT OR IGNORE INTO projects (contractor_id, name, folder_path) VALUES (?, ?, ?)
  `);
  const getProjectId = db.prepare(
    "SELECT id FROM projects WHERE contractor_id = ? AND name = ?"
  );
  const insertInspection = db.prepare(`
    INSERT OR REPLACE INTO inspections
    (project_id, filename, inspection_date, year, is_rain_event, file_size, sharepoint_id, web_url)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
  `);

  let totalProjects = 0;
  let totalInspections = 0;

  // Process contractors in batches
  const BATCH_SIZE = 5;
  for (let i = 0; i < allContractors.length; i += BATCH_SIZE) {
    const batch = allContractors.slice(i, i + BATCH_SIZE);

    await Promise.all(
      batch.map(async (contractor) => {
        // Insert contractor
        insertContractor.run(contractor.name, contractor.path);
        const contractorRow = getContractorId.get(contractor.name) as {
          id: number;
        };
        const contractorId = contractorRow.id;

        // Get projects for this contractor
        try {
          const items = await client.listFiles(contractor.path);
          const projects = items.filter((item) => item.folder);

          for (const project of projects) {
            const projectPath = `${contractor.path}/${project.name}`;
            insertProject.run(contractorId, project.name, projectPath);
            const projectRow = getProjectId.get(contractorId, project.name) as {
              id: number;
            };
            const projectId = projectRow.id;
            totalProjects++;

            // Get inspection files (check year subfolders and root)
            try {
              const projectItems = await client.listFiles(projectPath);

              // Process files at project level and in year subfolders
              const yearFolders = projectItems.filter(
                (item) => item.folder && /^\d{4}$/.test(item.name)
              );
              const rootFiles = projectItems.filter(
                (item) => item.file && item.name.endsWith(".pdf")
              );

              // Process root-level PDFs
              for (const file of rootFiles) {
                const parsed = parseInspectionFilename(file.name);
                insertInspection.run(
                  projectId,
                  file.name,
                  parsed.date,
                  parsed.year,
                  parsed.isRain ? 1 : 0,
                  file.size ?? null,
                  file.id,
                  file.webUrl
                );
                totalInspections++;
              }

              // Process year subfolders
              for (const yearFolder of yearFolders) {
                const yearPath = `${projectPath}/${yearFolder.name}`;
                const yearFiles = await client.listFiles(yearPath);
                const pdfs = yearFiles.filter(
                  (item) => item.file && item.name.endsWith(".pdf")
                );

                for (const file of pdfs) {
                  const parsed = parseInspectionFilename(file.name);
                  insertInspection.run(
                    projectId,
                    file.name,
                    parsed.date,
                    parsed.year ?? Number(yearFolder.name),
                    parsed.isRain ? 1 : 0,
                    file.size ?? null,
                    file.id,
                    file.webUrl
                  );
                  totalInspections++;
                }
              }
            } catch (err) {
              console.error(
                `  Error getting inspections for ${project.name}:`,
                err
              );
            }
          }
        } catch (err) {
          console.error(`Error getting projects for ${contractor.name}:`, err);
        }
      })
    );

    console.log(
      `Processed ${Math.min(i + BATCH_SIZE, allContractors.length)}/${allContractors.length} contractors...`
    );
  }

  db.close();

  console.log("\n✓ Sync complete!");
  console.log(`  Contractors: ${allContractors.length}`);
  console.log(`  Projects: ${totalProjects}`);
  console.log(`  Inspections: ${totalInspections}`);
  console.log(`  Database: ${DB_PATH}`);
}

// Run
syncFolders().catch(console.error);
