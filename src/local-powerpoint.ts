import fs from "fs";
import path from "path";
// @ts-ignore — officegen ships CJS without type defs
import officegen from "officegen";

export interface Slide {
  title: string;
  body: string;
}

/** Create a new .pptx file from an array of slides */
export function createPowerPointFile(
  filePath: string,
  slides: Slide[]
): Promise<void> {
  return new Promise((resolve, reject) => {
    const abs = path.resolve(filePath);
    const pptx = officegen("pptx");

    pptx.on("error", (err: Error) => reject(err));

    for (const slide of slides) {
      const sl = pptx.makeNewSlide();
      sl.addText(slide.title, {
        x: "5%",
        y: "5%",
        cx: "90%",
        cy: "15%",
        font_size: 32,
        bold: true,
        color: "000000",
      });
      sl.addText(slide.body, {
        x: "5%",
        y: "22%",
        cx: "90%",
        cy: "70%",
        font_size: 20,
        color: "333333",
      });
    }

    const out = fs.createWriteStream(abs);
    out.on("error", reject);
    out.on("close", resolve);
    pptx.generate(out);
  });
}

/** Read slide titles and body text from a .pptx (basic extraction via xml) */
export async function readPowerPointFile(
  filePath: string
): Promise<Slide[]> {
  const abs = path.resolve(filePath);
  if (!fs.existsSync(abs)) throw new Error(`File not found: ${abs}`);

  // pptx is a zip archive; use the built-in unzip to extract text
  const { execSync } = await import("child_process");
  const tmpDir = fs.mkdtempSync("/tmp/pptx-");

  try {
    execSync(`unzip -o "${abs}" "ppt/slides/*.xml" -d "${tmpDir}"`, {
      stdio: "pipe",
    });

    const slidesDir = path.join(tmpDir, "ppt", "slides");
    if (!fs.existsSync(slidesDir)) return [];

    const slideFiles = fs
      .readdirSync(slidesDir)
      .filter((f) => f.endsWith(".xml"))
      .sort();

    const slides: Slide[] = [];

    for (const file of slideFiles) {
      const xml = fs.readFileSync(path.join(slidesDir, file), "utf-8");

      // Extract all <a:t>text</a:t> nodes
      const texts: string[] = [];
      const matches = xml.matchAll(/<a:t[^>]*>([^<]*)<\/a:t>/g);
      for (const m of matches) {
        if (m[1].trim()) texts.push(m[1].trim());
      }

      const [title = "", ...rest] = texts;
      slides.push({ title, body: rest.join("\n") });
    }

    return slides;
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
}
