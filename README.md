# office-mcp-server

A [Model Context Protocol (MCP)](https://modelcontextprotocol.io) server that lets AI assistants (like [IBM Bob](https://www.ibm.com/bob)) read and write Microsoft Office files — both **locally on disk** and in **Microsoft 365 / OneDrive / SharePoint** via the Microsoft Graph API.

---

## Features

| Format | Local files | Microsoft 365 |
|---|---|---|
| **Word** (.docx) | Read (text + HTML), create, append | Read via OneDrive item ID |
| **Excel** (.xlsx) | Read all sheets, read range, write sheet, update cell | List sheets, read range, write range |
| **PowerPoint** (.pptx) | Read slides, create presentation | — |
| **OneDrive** | — | List, upload, download files |
| **SharePoint** | — | Search sites, list drives, list files |

---

## Prerequisites

- [Node.js](https://nodejs.org) v18 or later
- An MCP-compatible AI assistant (e.g. IBM Bob)
- For Microsoft 365 features: an Azure app registration (see [Azure Setup](#azure-setup))

---

## Installation

```bash
git clone https://github.com/Shedwar/office-mcp-server.git
cd office-mcp-server
npm install
npm run build
```

The compiled server will be at `build/index.js`.

---

## Registering with Bob (mcp.json)

Add the following entry to your Bob MCP config. Use the **global** config (`~/.bob/settings/mcp.json`) to make it available in every workspace, or a **workspace** config (`.bob/mcp.json`) to scope it to one project.

```json
{
  "mcpServers": {
    "office-mcp-server": {
      "command": "node",
      "args": ["/absolute/path/to/office-mcp-server/build/index.js"],
      "env": {
        "AZURE_CLIENT_ID": "${env:AZURE_CLIENT_ID}",
        "AZURE_TENANT_ID": "${env:AZURE_TENANT_ID}",
        "AZURE_CLIENT_SECRET": "${env:AZURE_CLIENT_SECRET}"
      }
    }
  }
}
```

> Replace `/absolute/path/to/office-mcp-server` with the actual path on your machine.
> The `${env:...}` references are expanded from the environment Bob is running in — see [Environment Variables](#environment-variables).

Bob hot-reloads the config on save. After saving, the server should appear as connected in Bob's MCP panel.

---

## Azure Setup

Microsoft 365 tools require an Azure app registration with delegated Microsoft Graph permissions.

### 1. Create an App Registration

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Set a name (e.g. `office-mcp-server`)
3. Set **Supported account types** to your organisation (or multi-tenant if needed)
4. Add a **Redirect URI**: `http://localhost:3456/callback` (type: Web)
5. Click **Register**

### 2. Add API Permissions

Under **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**, add:

- `Files.ReadWrite.All`
- `Sites.ReadWrite.All`
- `offline_access`

Click **Grant admin consent** if your tenant requires it.

### 3. Create a Client Secret

Under **Certificates & secrets** → **New client secret**, create a secret and copy the **Value** (shown once).

### 4. Note Your IDs

You need three values from the app registration **Overview** page:

| Variable | Where to find it |
|---|---|
| `AZURE_CLIENT_ID` | Application (client) ID |
| `AZURE_TENANT_ID` | Directory (tenant) ID |
| `AZURE_CLIENT_SECRET` | The secret value you copied above |

---

## Environment Variables

Set these in the shell environment where Bob is launched, then **restart Bob**:

```bash
export AZURE_CLIENT_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
export AZURE_TENANT_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
export AZURE_CLIENT_SECRET="your-client-secret-value"
```

> Variables set in a separate terminal after Bob is already running will **not** be picked up. You must restart Bob from a shell where these are exported.

---

## First-time Microsoft 365 Authentication

The first time any Microsoft 365 tool is called, the server will print an authentication URL to its log output:

```
=== Microsoft 365 Authentication Required ===
Open this URL in your browser:

https://login.microsoftonline.com/...
```

1. Open that URL in a browser
2. Sign in with your Microsoft 365 account
3. The server catches the OAuth callback on `localhost:3456` and stores the token at `~/.office-mcp-token-cache.json`

Subsequent calls use the cached refresh token silently. If the token expires, the interactive flow will trigger again.

---

## Available Tools

### Local — Word

#### `word_read_local`
Read the text (or HTML) content of a local `.docx` file.

| Parameter | Type | Description |
|---|---|---|
| `file_path` | string | Absolute or relative path to the `.docx` file |
| `format` | `"text"` \| `"html"` | Output format (default: `"text"`) |

**Example prompt:**
> "Read the contents of `~/Documents/report.docx`"

---

#### `word_create_local`
Create a new `.docx` file from an array of paragraphs.

| Parameter | Type | Description |
|---|---|---|
| `file_path` | string | Path where the file will be saved |
| `paragraphs` | array | List of paragraph objects (see below) |

Each paragraph object:

| Field | Type | Description |
|---|---|---|
| `text` | string | Paragraph text |
| `heading` | `"HEADING_1"` \| `"HEADING_2"` \| `"HEADING_3"` | Optional heading level |
| `bold` | boolean | Optional bold formatting |
| `italic` | boolean | Optional italic formatting |

**Example prompt:**
> "Create a Word document at `~/Documents/summary.docx` with a HEADING_1 'Executive Summary' and two body paragraphs"

---

#### `word_append_local`
Append paragraphs to an existing `.docx` file (creates the file if it doesn't exist).

| Parameter | Type | Description |
|---|---|---|
| `file_path` | string | Path to the `.docx` file |
| `paragraphs` | array | Paragraphs to append (same shape as `word_create_local`) |

---

### Local — Excel

#### `excel_read_local`
Read all sheets from a local `.xlsx` file as JSON.

| Parameter | Type | Description |
|---|---|---|
| `file_path` | string | Path to the `.xlsx` file |

Returns an object keyed by sheet name, where each value is an array of row objects.

---

#### `excel_read_range_local`
Read a specific cell range from a sheet.

| Parameter | Type | Description |
|---|---|---|
| `file_path` | string | Path to the `.xlsx` file |
| `sheet_name` | string | Name of the worksheet |
| `range` | string | Cell range e.g. `A1:D10` |

---

#### `excel_list_sheets_local`
List all sheet names in a local `.xlsx` file.

| Parameter | Type | Description |
|---|---|---|
| `file_path` | string | Path to the `.xlsx` file |

---

#### `excel_write_sheet_local`
Write rows to a named sheet (creates the file and/or sheet if missing).

| Parameter | Type | Description |
|---|---|---|
| `file_path` | string | Path to the `.xlsx` file |
| `sheet_name` | string | Target sheet name |
| `rows` | array | Array of objects where keys are column headers |

**Example prompt:**
> "Write this table to Sheet1 of `data.xlsx`: Name, Age, City …"

---

#### `excel_update_cell_local`
Update a single cell value.

| Parameter | Type | Description |
|---|---|---|
| `file_path` | string | Path to the `.xlsx` file |
| `sheet_name` | string | Target sheet name |
| `cell` | string | Cell address e.g. `B2` |
| `value` | string \| number \| boolean | New cell value |

---

### Local — PowerPoint

#### `pptx_read_local`
Extract slide titles and body text from a local `.pptx` file.

| Parameter | Type | Description |
|---|---|---|
| `file_path` | string | Path to the `.pptx` file |

Returns a JSON array of `{ title, body }` objects — one per slide.

---

#### `pptx_create_local`
Create a new `.pptx` file from an array of slides.

| Parameter | Type | Description |
|---|---|---|
| `file_path` | string | Path where the file will be saved |
| `slides` | array | Array of `{ title: string, body: string }` objects |

**Example prompt:**
> "Create a PowerPoint at `~/pitch.pptx` with 3 slides: Introduction, Problem, Solution"

---

### Microsoft 365 — OneDrive

#### `onedrive_list_files`
List files and folders in the signed-in user's OneDrive.

| Parameter | Type | Description |
|---|---|---|
| `folder_id` | string (optional) | Item ID of a subfolder; omit for root |

---

#### `onedrive_upload_file`
Upload a local file to OneDrive at a specified path.

| Parameter | Type | Description |
|---|---|---|
| `local_path` | string | Path to the local file |
| `remote_path` | string | Destination path in OneDrive e.g. `Documents/report.xlsx` |

---

#### `onedrive_download_file`
Download a OneDrive file by item ID and save it locally.

| Parameter | Type | Description |
|---|---|---|
| `item_id` | string | OneDrive file item ID (from `onedrive_list_files`) |
| `local_path` | string | Local path to save the file |

---

### Microsoft 365 — Excel

#### `excel365_list_sheets`
List worksheets in an Excel file stored on OneDrive.

| Parameter | Type | Description |
|---|---|---|
| `item_id` | string | OneDrive item ID of the `.xlsx` file |

---

#### `excel365_read_range`
Read a cell range from an Excel file on OneDrive.

| Parameter | Type | Description |
|---|---|---|
| `item_id` | string | OneDrive item ID of the `.xlsx` file |
| `address` | string | Range including sheet name e.g. `Sheet1!A1:D10` |

---

#### `excel365_update_range`
Write a 2D array of values to a range in a OneDrive Excel file.

| Parameter | Type | Description |
|---|---|---|
| `item_id` | string | OneDrive item ID of the `.xlsx` file |
| `sheet_name` | string | Target worksheet name |
| `address` | string | Range address e.g. `A1:C3` |
| `values` | 2D array | Values to write — must match range dimensions |

---

### Microsoft 365 — Word

#### `word365_read`
Read the plain-text content of a Word document stored on OneDrive.

| Parameter | Type | Description |
|---|---|---|
| `item_id` | string | OneDrive item ID of the `.docx` file |

---

### Microsoft 365 — SharePoint

#### `sharepoint_search_sites`
Search for SharePoint sites by keyword.

| Parameter | Type | Description |
|---|---|---|
| `query` | string | Search keyword e.g. `marketing` |

---

#### `sharepoint_list_drives`
List document libraries (drives) in a SharePoint site.

| Parameter | Type | Description |
|---|---|---|
| `site_id` | string | SharePoint site ID (from `sharepoint_search_sites`) |

---

#### `sharepoint_list_files`
List files in a SharePoint document library.

| Parameter | Type | Description |
|---|---|---|
| `site_id` | string | SharePoint site ID |
| `drive_id` | string | Drive ID (from `sharepoint_list_drives`) |
| `folder_id` | string (optional) | Item ID of a subfolder; omit for root |

---

## Development

```bash
# Watch mode — recompiles on save
npm run dev

# One-off build
npm run build
```

Source files are in [`src/`](./src/). Each module is focused:

| File | Responsibility |
|---|---|
| [`src/index.ts`](./src/index.ts) | MCP server entry point, all tool registrations |
| [`src/auth.ts`](./src/auth.ts) | Azure OAuth + silent token refresh |
| [`src/graph.ts`](./src/graph.ts) | Microsoft Graph API calls |
| [`src/local-word.ts`](./src/local-word.ts) | Local `.docx` read/write |
| [`src/local-excel.ts`](./src/local-excel.ts) | Local `.xlsx` read/write |
| [`src/local-powerpoint.ts`](./src/local-powerpoint.ts) | Local `.pptx` read/write |

---

## Token Cache

The OAuth token cache is stored at `~/.office-mcp-token-cache.json`. Delete this file to force re-authentication.

---

## License

ISC
