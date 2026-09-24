import fs from "fs";
import path from "path";
import mammoth from "mammoth";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
} from "docx";

/** Read all text from a .docx file */
export async function readWordFile(filePath: string): Promise<string> {
  const abs = path.resolve(filePath);
  if (!fs.existsSync(abs)) throw new Error(`File not found: ${abs}`);
  const result = await mammoth.extractRawText({ path: abs });
  return result.value;
}

/** Read a .docx file and return HTML */
export async function readWordFileAsHtml(filePath: string): Promise<string> {
  const abs = path.resolve(filePath);
  if (!fs.existsSync(abs)) throw new Error(`File not found: ${abs}`);
  const result = await mammoth.convertToHtml({ path: abs });
  return result.value;
}

export interface WordParagraph {
  text: string;
  heading?: "HEADING_1" | "HEADING_2" | "HEADING_3";
  bold?: boolean;
  italic?: boolean;
}

/** Create a new .docx file from an array of paragraphs */
export async function createWordFile(
  filePath: string,
  paragraphs: WordParagraph[]
): Promise<void> {
  const abs = path.resolve(filePath);

  const docParagraphs = paragraphs.map((p) => {
    const headingMap: Record<string, (typeof HeadingLevel)[keyof typeof HeadingLevel]> = {
      HEADING_1: HeadingLevel.HEADING_1,
      HEADING_2: HeadingLevel.HEADING_2,
      HEADING_3: HeadingLevel.HEADING_3,
    };
    return new Paragraph({
      heading: p.heading ? headingMap[p.heading] : undefined,
      children: [
        new TextRun({
          text: p.text,
          bold: p.bold ?? false,
          italics: p.italic ?? false,
        }),
      ],
    });
  });

  const doc = new Document({ sections: [{ children: docParagraphs }] });
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(abs, buffer);
}

/** Append paragraphs to an existing .docx by re-reading text and appending */
export async function appendToWordFile(
  filePath: string,
  paragraphs: WordParagraph[]
): Promise<void> {
  const abs = path.resolve(filePath);
  let existing: WordParagraph[] = [];

  if (fs.existsSync(abs)) {
    const extracted = await mammoth.extractRawText({ path: abs });
    // Re-wrap existing text as plain paragraphs
    existing = extracted.value
      .split("\n")
      .filter((l) => l.trim().length > 0)
      .map((l) => ({ text: l }));
  }

  await createWordFile(abs, [...existing, ...paragraphs]);
}
