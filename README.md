# Desert Services Inspection Email Worker

Cloudflare Email Worker that receives ComplianceGo inspection emails, generates PDFs using Browser Rendering, and uploads to SharePoint - all serverless.

**Email:** `inspections@desertservices.app`
**Worker:** `inspection-router.cheez2012.workers.dev`

## Setup

### 1. Install dependencies

```bash
npm install
```

### 2. Configure secrets in Cloudflare Dashboard

Go to Workers & Pages → inspection-router → Settings → Variables and Secrets:

| Secret | Value |
|--------|-------|
| `AZURE_TENANT_ID` | Your Azure tenant ID |
| `AZURE_CLIENT_ID` | Your Azure app client ID |
| `AZURE_CLIENT_SECRET` | Your Azure app client secret |

### 3. Deploy

Deploys automatically when you push to `main`.

Or manually:
```bash
npm run deploy
```

### 4. Configure Email Routing

In Cloudflare Dashboard → Email → Email Routing:
- Route emails from `support@compliancego.com` with subject "Inspection Report" → Worker `inspection-router`

## How It Works

1. ComplianceGo sends inspection report email
2. Email routing forwards to this worker
3. Worker parses email → extracts site name, report URL, date
4. **Browser Rendering** generates PDF from the report URL
5. **Microsoft Graph API** uploads PDF to SharePoint
6. Email forwarded to original recipient

## SharePoint Path Mapping

Site name → SharePoint folder:

- `ARCO - KTEC PHX` → `SWPPP/INSPECTIONS/PROJECTS/PROJECTS A-M/ARCO/KTEC PHX`
- `Sundt - Mesa` → `SWPPP/INSPECTIONS/PROJECTS/PROJECTS N-Z/SUNDT/Mesa`

Filename: `MM.DD.YY.pdf` (e.g., `01.21.26.pdf`)

## Development

```bash
# Run with remote browser (required for Browser Rendering)
npm run dev

# View logs
npm run tail
```

## Observability

Logs are available in Cloudflare Dashboard → Workers & Pages → inspection-router → Logs.
