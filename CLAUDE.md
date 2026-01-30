# DS Inspections Email Worker

Cloudflare Worker that receives ComplianceGo inspection emails, generates PDFs, and uploads to SharePoint.

## Project Overview

- **Email**: `inspections@desertservices.app`
- **Worker URL**: `inspection-router.cheez2012.workers.dev`
- **SharePoint**: DataDrive site → Shared Documents → SWPPP/INSPECTIONS/PROJECTS/

## Skills

### Inspection Upload Management
**Path**: `.claude/skills/upload-inspection/SKILL.md`

Use when:
- Checking if an inspection was uploaded
- Manually uploading failed inspections
- User provides ComplianceGo URLs or contractor/project details

## Quick Commands

```bash
# Check if inspection exists
bun scripts/check-inspection.ts "<contractor>" "<project>" [date]

# Manual upload
bun scripts/manual-upload.ts "<report-url>" "<contractor>" "<project>" [date]

# Deploy worker
bun run deploy

# View worker logs
bun run tail
```

## Architecture

```text
src/
├── index.ts          # Worker entry point (email handler, HTTP endpoints)
└── parser.ts         # Email parsing, site name mapping

scripts/
├── lib/              # Shared utilities
│   ├── env.ts        # Environment loading
│   ├── paths.ts      # SharePoint path building
│   └── sharepoint.ts # SharePoint client helpers
├── check-inspection.ts   # Verify upload exists
└── manual-upload.ts      # Manual PDF upload

sharepoint-inspections-folders-sync/
├── client.ts         # SharePointClient class (Graph API)
└── .env              # Azure credentials (not in git)
```

## Development

### Stack
- **Runtime**: Bun (scripts), Cloudflare Workers (production)
- **PDF Generation**: Puppeteer (local), @cloudflare/puppeteer (worker)
- **SharePoint**: Microsoft Graph API via Azure AD app

### Path Aliases

```typescript
import { parseComplianceGoEmail } from "@/parser";
import { buildInspectionPath } from "@scripts/lib";
import { SharePointClient } from "@sharepoint/client";
```

### Environment

Azure credentials in `sharepoint-inspections-folders-sync/.env`:
```text
AZURE_TENANT_ID=...
AZURE_CLIENT_ID=...
AZURE_CLIENT_SECRET=...
```

Worker secrets configured via `wrangler secret put`.

## Common Issues

### Upload fails with "JSON Content-Type" error
Path contains spaces that aren't URL-encoded. Already fixed in worker and client.

### "Could not determine SharePoint folder path"
Site name in ComplianceGo missing separator. Should be: `CONTRACTOR - PROJECT`

### Wrong date on manual upload
Use 4th parameter: `bun scripts/manual-upload.ts "..." "CONTRACTOR" "PROJECT" "01.29.26"`

## Bun Usage

- Use `bun <file>` instead of `node` or `ts-node`
- Use `bun test` for tests
- Use `bun install` for dependencies
- Use `Bun.file()` for file operations
- Bun auto-loads `.env` files
