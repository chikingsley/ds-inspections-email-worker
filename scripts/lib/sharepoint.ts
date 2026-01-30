/**
 * SharePoint client initialization and helpers
 * @module scripts/lib/sharepoint
 */

import { SharePointClient } from "@sharepoint/client";
import { getAzureCredentials, loadEnv } from "./env";

/** Cached client instance */
let clientInstance: SharePointClient | null = null;

/**
 * Gets or creates a SharePoint client instance.
 * Automatically loads environment variables on first call.
 *
 * @returns Initialized SharePoint client
 * @example
 * ```ts
 * const client = await getSharePointClient();
 * const files = await client.listFiles("SWPPP/INSPECTIONS");
 * ```
 */
export async function getSharePointClient(): Promise<SharePointClient> {
  if (clientInstance) {
    return clientInstance;
  }

  await loadEnv();
  const creds = getAzureCredentials();

  clientInstance = new SharePointClient({
    azureTenantId: creds.tenantId,
    azureClientId: creds.clientId,
    azureClientSecret: creds.clientSecret,
  });

  return clientInstance;
}

/**
 * Checks if a file exists in SharePoint.
 *
 * @param folderPath - SharePoint folder path
 * @param fileName - Name of file to check
 * @returns True if file exists, false otherwise
 * @example
 * ```ts
 * const exists = await fileExists(
 *   "SWPPP/INSPECTIONS/PROJECTS/PROJECTS A-M/ARCO/KTEC PHX/2026",
 *   "01.29.26.pdf"
 * );
 * ```
 */
export async function fileExists(
  folderPath: string,
  fileName: string
): Promise<boolean> {
  const client = await getSharePointClient();

  try {
    const files = await client.listFiles(folderPath);
    return files.some((f) => f.name === fileName);
  } catch {
    // Folder doesn't exist
    return false;
  }
}
