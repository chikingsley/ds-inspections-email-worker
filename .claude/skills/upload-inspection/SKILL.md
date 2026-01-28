# Upload Inspection

Manually upload a failed inspection PDF to SharePoint.

## When to use

Use this when:

- The automated worker failed to upload an inspection
- User says "upload the inspection", "manually upload", or similar
- User provides a failed inspection notification with contractor/project details

## Required information

You need these 3 things from the user:

1. **Report URL** - The ComplianceGo report URL (starts with `https://cdn3.compliancego.com/...`)
2. **Contractor** - The contractor name (e.g., "BPR COMPANIES")
3. **Project** - The project name (e.g., "PV LOT C3")

If the user provides a failure notification email, extract these from it.

## How to run

Run the script from the project root:

```bash
bun scripts/manual-upload.ts "<report-url>" "<contractor>" "<project>"
```

Example:

```bash
bun scripts/manual-upload.ts "https://cdn3.compliancego.com/s_xxx/insp/rpt/IR_BprCompanies-PvLotC3_26Jan26.html" "BPR COMPANIES" "PV LOT C3"
```

## What it does

1. Loads Azure credentials from `sharepoint-inspections-folders-sync/.env`
2. Checks if file already exists in SharePoint
3. Generates PDF from the ComplianceGo report URL using Puppeteer
4. Uploads to the correct SharePoint folder based on contractor/project
5. Returns the SharePoint URL on success

## Notes

- The script uses today's date for the filename (MM.DD.YY.pdf)
- Credentials are read from env vars, never hardcoded
- If the file already exists, it exits without re-uploading
