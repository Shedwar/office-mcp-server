#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";

// Local tools
import { readWordFile, readWordFileAsHtml, createWordFile, appendToWordFile } from "./local-word.js";
import {
  readExcelFile,
  readExcelRange,
  writeExcelSheet,
  updateExcelCell,
  listExcelSheets,
} from "./local-excel.js";
import { readPowerPointFile, createPowerPointFile } from "./local-powerpoint.js";

// Graph (365) tools
import {
  listOneDriveItems,
  downloadOneDriveFile,
  uploadOneDriveFile,
  listWorksheets,
  readExcelRange365,
  updateExcelRange365,
  readWordDoc365,
  searchSites,
  listSiteDrives,
  listSiteDriveFiles,
} from "./graph.js";

const server = new McpServer({
  name: "office-mcp-server",
  version: "1.0.0",
});

// ── LOCAL WORD ────────────────────────────────────────────────────────────────

server.registerTool(
  "word_read_local",
  {
    description: "Read the text content of a local .docx file",
    inputSchema: z.object({
      file_path: z.string().describe("Absolute or relative path to the .docx file"),
      format: z.enum(["text", "html"]).default("text").describe("Output format: plain text or HTML"),
    }),
  },
  async ({ file_path, format }) => {
    try {
      const content =
        format === "html"
          ? await readWordFileAsHtml(file_path)
          : await readWordFile(file_path);
      return { content: [{ type: "text", text: content }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

server.registerTool(
  "word_create_local",
  {
    description: "Create a new .docx file with the specified paragraphs",
    inputSchema: z.object({
      file_path: z.string().describe("Path where the .docx should be saved"),
      paragraphs: z.array(
        z.object({
          text: z.string(),
          heading: z.enum(["HEADING_1", "HEADING_2", "HEADING_3"]).optional(),
          bold: z.boolean().optional(),
          italic: z.boolean().optional(),
        })
      ).describe("Array of paragraph objects to write into the document"),
    }),
  },
  async ({ file_path, paragraphs }) => {
    try {
      await createWordFile(file_path, paragraphs);
      return { content: [{ type: "text", text: `Created ${file_path}` }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

server.registerTool(
  "word_append_local",
  {
    description: "Append paragraphs to an existing (or new) local .docx file",
    inputSchema: z.object({
      file_path: z.string(),
      paragraphs: z.array(
        z.object({
          text: z.string(),
          bold: z.boolean().optional(),
          italic: z.boolean().optional(),
        })
      ),
    }),
  },
  async ({ file_path, paragraphs }) => {
    try {
      await appendToWordFile(file_path, paragraphs);
      return { content: [{ type: "text", text: `Updated ${file_path}` }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

// ── LOCAL EXCEL ───────────────────────────────────────────────────────────────

server.registerTool(
  "excel_read_local",
  {
    description: "Read all sheets from a local .xlsx file as JSON",
    inputSchema: z.object({
      file_path: z.string().describe("Path to the .xlsx file"),
    }),
  },
  async ({ file_path }) => {
    try {
      const data = readExcelFile(file_path);
      return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

server.registerTool(
  "excel_read_range_local",
  {
    description: "Read a specific cell range from a local .xlsx sheet",
    inputSchema: z.object({
      file_path: z.string(),
      sheet_name: z.string().describe("Name of the worksheet"),
      range: z.string().describe("Cell range e.g. A1:D10"),
    }),
  },
  async ({ file_path, sheet_name, range }) => {
    try {
      const data = readExcelRange(file_path, sheet_name, range);
      return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

server.registerTool(
  "excel_list_sheets_local",
  {
    description: "List sheet names in a local .xlsx file",
    inputSchema: z.object({ file_path: z.string() }),
  },
  async ({ file_path }) => {
    try {
      const sheets = listExcelSheets(file_path);
      return { content: [{ type: "text", text: sheets.join(", ") }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

server.registerTool(
  "excel_write_sheet_local",
  {
    description: "Write rows to a sheet in a local .xlsx file (creates file if missing)",
    inputSchema: z.object({
      file_path: z.string(),
      sheet_name: z.string(),
      rows: z.array(z.record(z.unknown())).describe("Array of row objects (keys = column headers)"),
    }),
  },
  async ({ file_path, sheet_name, rows }) => {
    try {
      writeExcelSheet(file_path, sheet_name, rows);
      return { content: [{ type: "text", text: `Written ${rows.length} rows to ${sheet_name} in ${file_path}` }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

server.registerTool(
  "excel_update_cell_local",
  {
    description: "Update a single cell in a local .xlsx file",
    inputSchema: z.object({
      file_path: z.string(),
      sheet_name: z.string(),
      cell: z.string().describe("Cell address e.g. B2"),
      value: z.union([z.string(), z.number(), z.boolean()]),
    }),
  },
  async ({ file_path, sheet_name, cell, value }) => {
    try {
      updateExcelCell(file_path, sheet_name, cell, value);
      return { content: [{ type: "text", text: `Cell ${cell} updated in ${sheet_name}` }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

// ── LOCAL POWERPOINT ──────────────────────────────────────────────────────────

server.registerTool(
  "pptx_read_local",
  {
    description: "Extract slides (title + body text) from a local .pptx file",
    inputSchema: z.object({ file_path: z.string() }),
  },
  async ({ file_path }) => {
    try {
      const slides = await readPowerPointFile(file_path);
      return { content: [{ type: "text", text: JSON.stringify(slides, null, 2) }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

server.registerTool(
  "pptx_create_local",
  {
    description: "Create a new .pptx file with the given slides",
    inputSchema: z.object({
      file_path: z.string(),
      slides: z.array(
        z.object({
          title: z.string(),
          body: z.string(),
        })
      ).describe("Array of slides, each with a title and body text"),
    }),
  },
  async ({ file_path, slides }) => {
    try {
      await createPowerPointFile(file_path, slides);
      return { content: [{ type: "text", text: `Created ${file_path} with ${slides.length} slide(s)` }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

// ── MICROSOFT 365 / ONEDRIVE ──────────────────────────────────────────────────

server.registerTool(
  "onedrive_list_files",
  {
    description: "List files and folders in the user's OneDrive root (or a specific folder)",
    inputSchema: z.object({
      folder_id: z.string().optional().describe("OneDrive item ID of the folder (omit for root)"),
    }),
  },
  async ({ folder_id }) => {
    try {
      const items = await listOneDriveItems(folder_id);
      return { content: [{ type: "text", text: JSON.stringify(items, null, 2) }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

server.registerTool(
  "onedrive_upload_file",
  {
    description: "Upload a local file to OneDrive at a specified remote path",
    inputSchema: z.object({
      local_path: z.string().describe("Path to the local file to upload"),
      remote_path: z.string().describe("Destination path in OneDrive e.g. Documents/report.xlsx"),
    }),
  },
  async ({ local_path, remote_path }) => {
    try {
      const fs = await import("fs");
      const content = fs.readFileSync(local_path);
      const item = await uploadOneDriveFile(remote_path, content);
      return { content: [{ type: "text", text: `Uploaded: ${item.name} — ${item.webUrl}` }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

server.registerTool(
  "onedrive_download_file",
  {
    description: "Download a OneDrive file by item ID and save it locally",
    inputSchema: z.object({
      item_id: z.string().describe("OneDrive file item ID"),
      local_path: z.string().describe("Local path to save the file"),
    }),
  },
  async ({ item_id, local_path }) => {
    try {
      const buffer = await downloadOneDriveFile(item_id);
      const fs = await import("fs");
      fs.writeFileSync(local_path, buffer);
      return { content: [{ type: "text", text: `Saved to ${local_path}` }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

// ── MICROSOFT 365 / EXCEL ─────────────────────────────────────────────────────

server.registerTool(
  "excel365_list_sheets",
  {
    description: "List worksheets in an Excel file stored on OneDrive",
    inputSchema: z.object({
      item_id: z.string().describe("OneDrive item ID of the .xlsx file"),
    }),
  },
  async ({ item_id }) => {
    try {
      const sheets = await listWorksheets(item_id);
      return { content: [{ type: "text", text: sheets.join(", ") }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

server.registerTool(
  "excel365_read_range",
  {
    description: "Read a cell range from an Excel file on OneDrive (e.g. Sheet1!A1:D10)",
    inputSchema: z.object({
      item_id: z.string(),
      address: z.string().describe("Range address including sheet name e.g. Sheet1!A1:D10"),
    }),
  },
  async ({ item_id, address }) => {
    try {
      const values = await readExcelRange365(item_id, address);
      return { content: [{ type: "text", text: JSON.stringify(values, null, 2) }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

server.registerTool(
  "excel365_update_range",
  {
    description: "Write values to a cell range in an Excel file on OneDrive",
    inputSchema: z.object({
      item_id: z.string(),
      sheet_name: z.string(),
      address: z.string().describe("Range address e.g. A1:C3"),
      values: z.array(z.array(z.union([z.string(), z.number(), z.null()]))).describe("2D array of values matching the range dimensions"),
    }),
  },
  async ({ item_id, sheet_name, address, values }) => {
    try {
      await updateExcelRange365(item_id, sheet_name, address, values);
      return { content: [{ type: "text", text: `Range ${address} updated in ${sheet_name}` }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

// ── MICROSOFT 365 / WORD ──────────────────────────────────────────────────────

server.registerTool(
  "word365_read",
  {
    description: "Read the text content of a Word document stored on OneDrive",
    inputSchema: z.object({
      item_id: z.string().describe("OneDrive item ID of the .docx file"),
    }),
  },
  async ({ item_id }) => {
    try {
      const text = await readWordDoc365(item_id);
      return { content: [{ type: "text", text }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

// ── SHAREPOINT ────────────────────────────────────────────────────────────────

server.registerTool(
  "sharepoint_search_sites",
  {
    description: "Search for SharePoint sites by keyword",
    inputSchema: z.object({
      query: z.string().describe("Search keyword e.g. 'marketing'"),
    }),
  },
  async ({ query }) => {
    try {
      const sites = await searchSites(query);
      return { content: [{ type: "text", text: JSON.stringify(sites, null, 2) }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

server.registerTool(
  "sharepoint_list_drives",
  {
    description: "List document libraries (drives) in a SharePoint site",
    inputSchema: z.object({
      site_id: z.string().describe("SharePoint site ID"),
    }),
  },
  async ({ site_id }) => {
    try {
      const drives = await listSiteDrives(site_id);
      return { content: [{ type: "text", text: JSON.stringify(drives, null, 2) }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

server.registerTool(
  "sharepoint_list_files",
  {
    description: "List files in a SharePoint document library",
    inputSchema: z.object({
      site_id: z.string(),
      drive_id: z.string(),
      folder_id: z.string().optional().describe("Item ID of a subfolder (omit for root)"),
    }),
  },
  async ({ site_id, drive_id, folder_id }) => {
    try {
      const files = await listSiteDriveFiles(site_id, drive_id, folder_id);
      return { content: [{ type: "text", text: JSON.stringify(files, null, 2) }] };
    } catch (e) {
      return { content: [{ type: "text", text: String(e) }], isError: true };
    }
  }
);

// ── Start ─────────────────────────────────────────────────────────────────────

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("office-mcp-server running on stdio");
}

main().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
