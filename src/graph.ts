import fetch from "node-fetch";
import { getAccessToken } from "./auth.js";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

async function graphRequest<T>(
  method: string,
  endpoint: string,
  body?: unknown
): Promise<T> {
  const token = await getAccessToken();
  const res = await fetch(`${GRAPH_BASE}${endpoint}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: body !== undefined ? JSON.stringify(body) : undefined,
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph API ${method} ${endpoint} → ${res.status}: ${text}`);
  }

  // 204 No Content
  if (res.status === 204) return {} as T;
  return res.json() as Promise<T>;
}

// ── OneDrive / Files ──────────────────────────────────────────────────────────

export interface DriveItem {
  id: string;
  name: string;
  size?: number;
  lastModifiedDateTime?: string;
  webUrl?: string;
  folder?: object;
  file?: object;
}

/** List files/folders in the root of the user's OneDrive */
export async function listOneDriveItems(
  folderId?: string
): Promise<DriveItem[]> {
  const path = folderId
    ? `/me/drive/items/${folderId}/children`
    : "/me/drive/root/children";
  const data = await graphRequest<{ value: DriveItem[] }>("GET", path);
  return data.value;
}

/** Download the text content of a OneDrive file by item ID */
export async function downloadOneDriveFile(itemId: string): Promise<Buffer> {
  const token = await getAccessToken();
  const res = await fetch(
    `${GRAPH_BASE}/me/drive/items/${itemId}/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`Download failed: ${res.status}`);
  return Buffer.from(await res.arrayBuffer());
}

/** Upload / replace a file on OneDrive by path (e.g. "Documents/report.xlsx") */
export async function uploadOneDriveFile(
  remotePath: string,
  content: Buffer,
  mimeType = "application/octet-stream"
): Promise<DriveItem> {
  const token = await getAccessToken();
  const encoded = encodeURIComponent(remotePath);
  const res = await fetch(
    `${GRAPH_BASE}/me/drive/root:/${encoded}:/content`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": mimeType,
      },
      body: content,
    }
  );
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Upload failed: ${res.status}: ${text}`);
  }
  return res.json() as Promise<DriveItem>;
}

// ── Excel (Graph Workbook API) ────────────────────────────────────────────────

export interface WorksheetRow {
  [column: string]: unknown;
}

/** List worksheets in an Excel file stored on OneDrive */
export async function listWorksheets(itemId: string): Promise<string[]> {
  const data = await graphRequest<{ value: Array<{ name: string }> }>(
    "GET",
    `/me/drive/items/${itemId}/workbook/worksheets`
  );
  return data.value.map((ws) => ws.name);
}

/** Read cells from a named range or address (e.g. "Sheet1!A1:D10") */
export async function readExcelRange365(
  itemId: string,
  address: string
): Promise<unknown[][]> {
  const data = await graphRequest<{ values: unknown[][] }>(
    "GET",
    `/me/drive/items/${itemId}/workbook/worksheets('${encodeURIComponent(address.split("!")[0])}')/range(address='${encodeURIComponent(address)}')`
  );
  return data.values;
}

/** Write values to a range (values must match the range dimensions) */
export async function updateExcelRange365(
  itemId: string,
  sheetName: string,
  address: string,
  values: unknown[][]
): Promise<void> {
  await graphRequest(
    "PATCH",
    `/me/drive/items/${itemId}/workbook/worksheets('${encodeURIComponent(sheetName)}')/range(address='${encodeURIComponent(address)}')`,
    { values }
  );
}

// ── Word / Documents ──────────────────────────────────────────────────────────

/** Read the plain-text content of a Word document via Graph (converts to text) */
export async function readWordDoc365(itemId: string): Promise<string> {
  // Download raw bytes and extract text via mammoth
  const buffer = await downloadOneDriveFile(itemId);
  const mammoth = await import("mammoth");
  const result = await mammoth.extractRawText({ buffer });
  return result.value;
}

// ── SharePoint / Sites ────────────────────────────────────────────────────────

export interface Site {
  id: string;
  displayName: string;
  webUrl: string;
}

/** Search for SharePoint sites by keyword */
export async function searchSites(query: string): Promise<Site[]> {
  const data = await graphRequest<{ value: Site[] }>(
    "GET",
    `/sites?search=${encodeURIComponent(query)}`
  );
  return data.value;
}

/** List document libraries (drives) in a site */
export async function listSiteDrives(
  siteId: string
): Promise<Array<{ id: string; name: string }>> {
  const data = await graphRequest<{
    value: Array<{ id: string; name: string }>;
  }>("GET", `/sites/${siteId}/drives`);
  return data.value;
}

/** List files in a site drive */
export async function listSiteDriveFiles(
  siteId: string,
  driveId: string,
  folderId?: string
): Promise<DriveItem[]> {
  const endpoint = folderId
    ? `/sites/${siteId}/drives/${driveId}/items/${folderId}/children`
    : `/sites/${siteId}/drives/${driveId}/root/children`;
  const data = await graphRequest<{ value: DriveItem[] }>("GET", endpoint);
  return data.value;
}
