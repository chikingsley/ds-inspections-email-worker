/**
 * Environment loading utilities for scripts
 * @module scripts/lib/env
 */

/** Azure credentials required for SharePoint access */
export interface AzureCredentials {
  tenantId: string;
  clientId: string;
  clientSecret: string;
}

/** Path to the .env file containing Azure credentials */
const ENV_PATH = "./sharepoint-inspections-folders-sync/.env";

/**
 * Loads environment variables from the SharePoint sync .env file.
 * Bun auto-loads root .env, but we need the nested one for Azure creds.
 *
 * @returns Promise that resolves when env vars are loaded
 * @example
 * ```ts
 * await loadEnv();
 * const creds = getAzureCredentials();
 * ```
 */
export async function loadEnv(): Promise<void> {
  const envFile = Bun.file(ENV_PATH);

  if (!(await envFile.exists())) {
    throw new Error(`Environment file not found: ${ENV_PATH}`);
  }

  const content = await envFile.text();

  for (const line of content.split("\n")) {
    const trimmed = line.trim();
    if (trimmed && !trimmed.startsWith("#")) {
      const eqIndex = trimmed.indexOf("=");
      if (eqIndex > 0) {
        const key = trimmed.slice(0, eqIndex);
        const value = trimmed.slice(eqIndex + 1);
        process.env[key] = value;
      }
    }
  }
}

/**
 * Gets Azure credentials from environment variables.
 * Call loadEnv() first to ensure variables are loaded.
 *
 * @returns Azure credentials object
 * @throws Error if any required credential is missing
 * @example
 * ```ts
 * await loadEnv();
 * const { tenantId, clientId, clientSecret } = getAzureCredentials();
 * ```
 */
export function getAzureCredentials(): AzureCredentials {
  const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET } = process.env;

  if (!(AZURE_TENANT_ID && AZURE_CLIENT_ID && AZURE_CLIENT_SECRET)) {
    throw new Error(
      "Missing Azure credentials. Ensure AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET are set."
    );
  }

  return {
    tenantId: AZURE_TENANT_ID,
    clientId: AZURE_CLIENT_ID,
    clientSecret: AZURE_CLIENT_SECRET,
  };
}
