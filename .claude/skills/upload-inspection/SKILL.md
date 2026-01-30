# Inspection Upload Management

Check and manually upload inspection PDFs to SharePoint.

## When to use

Use this skill when:

- User asks to check if an inspection was uploaded
- User says "check inspection", "was this uploaded", "verify upload"
- User provides a ComplianceGo URL and asks to upload it
- User mentions a failed inspection notification
- User provides contractor/project details for verification

## Workflow

### 1. Check if Inspection Exists

First, verify if the inspection is already in SharePoint:

```bash
bun scripts/check-inspection.ts "<contractor>" "<project>" [date]
```

**Examples:**

```bash
# Check today's inspection
bun scripts/check-inspection.ts "ARCO" "KTEC PHX"

# Check specific date
bun scripts/check-inspection.ts "BPR COMPANIES" "PV LOT C3" "01.26.26"
```

**Exit codes:**
- `0` = File exists
- `1` = File not found (needs upload)

### 2. Manual Upload (if needed)

If the file doesn't exist, upload it:

```bash
bun scripts/manual-upload.ts "<report-url>" "<contractor>" "<project>" [date]
```

**Examples:**

```bash
# Upload with today's date
bun scripts/manual-upload.ts "https://cdn3.compliancego.com/..." "ARCO" "KTEC PHX"

# Upload with specific date (for catch-up uploads)
bun scripts/manual-upload.ts "https://cdn3.compliancego.com/..." "ARCO" "KTEC PHX" "01.29.26"
```

## Extracting Information

### From ComplianceGo Report URL

The report page shows:
- **Client Company** → Use as `contractor`
- **Site Name** → Usually `CONTRACTOR - PROJECT` format, use the project part

### From Failed Upload Notification

The error email contains:
- **CONTRACTOR** → Use directly
- **PROJECT** → Use directly
- **FILE** → Contains the date (e.g., `01.29.26.pdf`)

### Site Name Format

Sites should follow `CONTRACTOR - PROJECT` naming (space-dash-space). If a site doesn't parse correctly, it's likely a data entry issue in ComplianceGo.

## SharePoint Path Structure

Files are organized as:
```text
SWPPP/INSPECTIONS/PROJECTS/
├── PROJECTS A-M/       # Contractors starting with 0-9 or A-M
│   └── {CONTRACTOR}/
│       └── {PROJECT}/
│           └── {YEAR}/
│               └── MM.DD.YY.pdf
└── PROJECTS N-Z/       # Contractors starting with N-Z
    └── ...
```

## Troubleshooting

### "Could not determine SharePoint folder path"
The site name in ComplianceGo doesn't have a `-` separator. Ask the user to fix it in ComplianceGo to: `CONTRACTOR - PROJECT`

### Wrong date on filename
Use the optional date parameter: `"01.29.26"` (MM.DD.YY format)

### File already exists
The script exits cleanly - no action needed.
