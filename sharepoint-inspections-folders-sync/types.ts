/**
 * Type definitions for SharePoint client
 */

export interface SharePointConfig {
  azureTenantId: string;
  azureClientId: string;
  azureClientSecret: string;
}

export interface SharePointSite {
  id: string;
  name: string;
  displayName: string;
  webUrl: string;
  description?: string;
}

export interface SharePointDrive {
  id: string;
  name: string;
  driveType: string;
  webUrl: string;
}

export interface SharePointFolder {
  childCount: number;
}

export interface SharePointFile {
  mimeType: string;
}

export interface SharePointItem {
  id: string;
  name: string;
  webUrl: string;
  size?: number;
  createdDateTime: string;
  lastModifiedDateTime: string;
  folder?: SharePointFolder;
  file?: SharePointFile;
}

export interface SharePointList {
  id: string;
  name: string;
  displayName: string;
  webUrl: string;
  description?: string;
}

export interface SharePointListItem {
  id: string;
  fields: Record<string, unknown>;
  createdDateTime: string;
  lastModifiedDateTime: string;
}
