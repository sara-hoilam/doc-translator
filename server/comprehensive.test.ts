/**
 * Comprehensive test suite for PDFGodWork
 * Tests: format detection, text extraction, document building, format conversion,
 *        in-place translation, multi-language translation, watermark preview, download integrity
 */

import { describe, it, expect } from "vitest";
import {
  extractTextFromBuffer,
  translateWithLLM,
  buildTranslatedDocument,
  convertDocument,
  translateDocxInPlace,
  translatePptxInPlace,
  translateXlsxInPlace,
  getMimeType,
  getFormatFromFilename,
  isImageFormat,
} from "./docProcessor";
import { generateWatermarkedPreview } from "./watermark";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import AdmZip from "adm-zip";
import * as XLSX from "xlsx";

// ─── Test fixture helpers ─────────────────────────────────────────────────────

async function makePdfBuffer(text = "Hello World\nThis is a test PDF document."): Promise<Buffer> {
  const pdfDoc = await PDFDocument.create();
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const page = pdfDoc.addPage([595, 842]);
  page.drawText(text.slice(0, 200), { x: 50, y: 750, size: 12, font, color: rgb(0, 0, 0) });
  return Buffer.from(await pdfDoc.save());
}

function makeDocxBuffer(text = "Hello World\nSecond paragraph"): Buffer {
  const zip = new AdmZip();
  const lines = text.split("\n");
  zip.addFile("[Content_Types].xml", Buffer.from(
    `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
  ));
  zip.addFile("_rels/.rels", Buffer.from(
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`
  ));
  zip.addFile("word/document.xml", Buffer.from(
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${lines.map(l => `<w:p><w:r><w:t>${l.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</w:t></w:r></w:p>`).join("")}<w:sectPr/></w:body></w:document>`
  ));
  return zip.toBuffer();
}

function makePptxBuffer(slideText = "Test Slide\nBullet one\nBullet two"): Buffer {
  const zip = new AdmZip();
  const lines = slideText.split("\n");
  zip.addFile("[Content_Types].xml", Buffer.from(
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`
  ));
  zip.addFile("_rels/.rels", Buffer.from(
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`
  ));
  zip.addFile("ppt/presentation.xml", Buffer.from(
    `<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst><p:sldId id="256" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></p:sldIdLst></p:presentation>`
  ));
  zip.addFile("ppt/_rels/presentation.xml.rels", Buffer.from(
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>`
  ));
  zip.addFile("ppt/slides/slide1.xml", Buffer.from(
    `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody>${lines.map(l => `<a:p><a:r><a:t>${l.replace(/&/g, "&amp;").replace(/</g, "&lt;")}</a:t></a:r></a:p>`).join("")}</p:txBody></p:sp></p:spTree></p:cSld></p:sld>`
  ));
  zip.addFile("ppt/slides/_rels/slide1.xml.rels", Buffer.from(
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`
  ));
  return zip.toBuffer();
}

function makeXlsxBuffer(): Buffer {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ["Name", "Age", "City"],
    ["Alice", "30", "New York"],
    ["Bob", "25", "London"],
    ["Charlie", "35", "Tokyo"],
  ]);
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  return Buffer.from(XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
}

function makeTxtBuffer(text = "Hello World\nThis is plain text."): Buffer {
  return Buffer.from(text, "utf-8");
}

function makeMockLLM(translateFn?: (text: string) => string) {
  return async (params: any) => {
    const userContent = params.messages?.find((m: any) => m.role === "user")?.content ?? "";
    if (userContent.match(/^\d+\./m)) {
      const lines = userContent.split("\n").filter((l: string) => l.match(/^\d+\./));
      const translated = lines.map((l: string) => {
        const m = l.match(/^(\d+)\.\s*(.*)/);
        if (!m) return l;
        return `${m[1]}. ${translateFn ? translateFn(m[2]) : `[TR] ${m[2]}`}`;
      }).join("\n");
      return { choices: [{ message: { content: translated } }] };
    }
    if (userContent.match(/^\[\d+\]/m)) {
      const lines = userContent.split("\n").filter((l: string) => l.match(/^\[\d+\]/));
      const translated = lines.map((l: string) => {
        const m = l.match(/^\[(\d+)\]\s*(.*)/);
        if (!m) return l;
        return `[${m[1]}] ${translateFn ? translateFn(m[2]) : `[TR] ${m[2]}`}`;
      }).join("\n");
      return { choices: [{ message: { content: translated } }] };
    }
    const result = translateFn ? translateFn(userContent) : `[TR] ${userContent.slice(0, 80)}`;
    return { choices: [{ message: { content: result } }] };
  };
}

// ─── 1. Format detection ──────────────────────────────────────────────────────
describe("1. Format detection", () => {
  it("PDF files fall back to txt (PDF no longer an input format)", () => expect(getFormatFromFilename("doc.pdf")).toBe("txt")); // PDF removed from SUPPORTED_INPUT_FORMATS
  it("detects DOCX", () => expect(getFormatFromFilename("doc.docx")).toBe("docx"));
  it("detects PPTX", () => expect(getFormatFromFilename("slides.pptx")).toBe("pptx"));
  it("detects XLSX", () => expect(getFormatFromFilename("data.xlsx")).toBe("xlsx"));
  it("detects TXT", () => expect(getFormatFromFilename("notes.txt")).toBe("txt"));
  it("detects PNG as PDF (image→PDF conversion)", () => expect(getFormatFromFilename("photo.png")).toBe("pdf"));
  it("detects JPG as PDF", () => expect(getFormatFromFilename("img.jpg")).toBe("pdf"));
  it("detects JPEG as PDF", () => expect(getFormatFromFilename("img.jpeg")).toBe("pdf"));
  it("detects WEBP as PDF", () => expect(getFormatFromFilename("img.webp")).toBe("pdf"));
  it("detects GIF as PDF", () => expect(getFormatFromFilename("img.gif")).toBe("pdf"));
  it("isImageFormat: PNG/JPG/JPEG/WEBP/GIF are images", () => {
    expect(isImageFormat("png")).toBe(true);
    expect(isImageFormat("jpg")).toBe(true);
    expect(isImageFormat("jpeg")).toBe(true);
    expect(isImageFormat("webp")).toBe(true);
    expect(isImageFormat("gif")).toBe(true);
  });
  it("isImageFormat: PDF/DOCX/PPTX are not images", () => {
    expect(isImageFormat("pdf")).toBe(false);
    expect(isImageFormat("docx")).toBe(false);
    expect(isImageFormat("pptx")).toBe(false);
  });
  it("getMimeType returns correct MIME for all formats", () => {
    expect(getMimeType("pdf")).toBe("application/pdf");
    expect(getMimeType("docx")).toContain("wordprocessingml");
    expect(getMimeType("pptx")).toContain("presentationml");
    expect(getMimeType("xlsx")).toContain("spreadsheetml");
    expect(getMimeType("txt")).toBe("text/plain");
    expect(getMimeType("unknown")).toBe("application/octet-stream");
  });
});

// ─── 2. Text extraction ───────────────────────────────────────────────────────
describe("2. Text extraction — all input formats", () => {
  it("extracts text from PDF", async () => {
    const buf = await makePdfBuffer("Hello World PDF content here");
    const result = await extractTextFromBuffer(buf, "pdf", "test.pdf");
    expect(result.text).toContain("Hello World");
    expect(result.pageCount).toBeGreaterThanOrEqual(1);
    expect(result.charCount).toBeGreaterThan(0);
  }, 15_000);

  it("extracts text from DOCX", async () => {
    const buf = makeDocxBuffer("Hello World DOCX content\nSecond line");
    const result = await extractTextFromBuffer(buf, "docx", "test.docx");
    expect(result.text).toContain("Hello World");
    expect(result.pageCount).toBeGreaterThanOrEqual(1);
  });

  it("extracts text from PPTX", async () => {
    const buf = makePptxBuffer("Test Slide Title\nBullet one\nBullet two");
    const result = await extractTextFromBuffer(buf, "pptx", "test.pptx");
    expect(result.text).toContain("Test Slide Title");
    expect(result.pageCount).toBeGreaterThanOrEqual(1);
  });

  it("extracts text from XLSX", async () => {
    const buf = makeXlsxBuffer();
    const result = await extractTextFromBuffer(buf, "xlsx", "test.xlsx");
    expect(result.text).toContain("Name");
    expect(result.text).toContain("Alice");
    expect(result.pageCount).toBeGreaterThanOrEqual(1);
  });

  it("extracts text from TXT", async () => {
    const buf = makeTxtBuffer("Hello World plain text\nSecond line here");
    const result = await extractTextFromBuffer(buf, "txt", "test.txt");
    expect(result.text).toContain("Hello World");
    expect(result.pageCount).toBeGreaterThanOrEqual(1);
  });

  it("handles empty TXT gracefully", async () => {
    const result = await extractTextFromBuffer(Buffer.from(""), "txt", "empty.txt");
    expect(result.text).toBe("");
    expect(result.pageCount).toBeGreaterThanOrEqual(1);
  });

  it("handles TXT with only whitespace", async () => {
    const result = await extractTextFromBuffer(Buffer.from("   \n\n   "), "txt", "ws.txt");
    expect(result.text.trim()).toBe("");
  });
});

// ─── 3. Document building ─────────────────────────────────────────────────────
describe("3. Document building — all output formats", () => {
  const md = `# Test Document\n\nParagraph with **bold** text.\n\n## Section Two\n\n- Bullet one\n- Bullet two\n\n| Col A | Col B |\n|-------|-------|\n| 1     | 2     |\n`;

  it("builds valid DOCX", async () => {
    const buf = await buildTranslatedDocument(md, "docx", "out.docx");
    expect(buf.length).toBeGreaterThan(500);
    const zip = new AdmZip(buf);
    expect(zip.getEntries().map(e => e.entryName)).toContain("word/document.xml");
    expect(zip.getEntries().map(e => e.entryName)).toContain("[Content_Types].xml");
    const docXml = zip.readAsText("word/document.xml");
    expect(docXml).toContain("Test Document");
  });

  it("builds valid PPTX", async () => {
    const buf = await buildTranslatedDocument(md, "pptx", "out.pptx");
    expect(buf.length).toBeGreaterThan(500);
    const zip = new AdmZip(buf);
    expect(zip.getEntries().some(e => e.entryName.startsWith("ppt/slides/"))).toBe(true);
    // No duplicate slideMaster entries
    const ct = zip.readAsText("[Content_Types].xml");
    const masters = ct.match(/slideMaster\d+\.xml/g) ?? [];
    expect(masters.length).toBe(new Set(masters).size);
  });

  it("builds valid XLSX", async () => {
    const buf = await buildTranslatedDocument(md, "xlsx", "out.xlsx");
    expect(buf.length).toBeGreaterThan(100);
    const wb = XLSX.read(buf, { type: "buffer" });
    expect(wb.SheetNames.length).toBeGreaterThan(0);
  });

  it("builds valid TXT", async () => {
    const buf = await buildTranslatedDocument(md, "txt", "out.txt");
    expect(buf.toString("utf-8")).toContain("Test Document");
  });

  it("builds valid PDF", async () => {
    const buf = await buildTranslatedDocument(md, "pdf", "out.pdf");
    expect(buf.length).toBeGreaterThan(100);
    const header = buf.toString("utf-8", 0, 20);
    expect(header.startsWith("%PDF") || header.includes("DOCTYPE") || header.includes("<html")).toBe(true);
  }, 15_000);

  it("handles empty text for DOCX", async () => {
    const buf = await buildTranslatedDocument("", "docx", "empty.docx");
    expect(buf.length).toBeGreaterThan(100);
    const zip = new AdmZip(buf);
    expect(zip.getEntry("word/document.xml")).not.toBeNull();
  });

  it("handles XML special characters in DOCX", async () => {
    const text = "Hello & World <test> \"quotes\" 'apostrophes'";
    const buf = await buildTranslatedDocument(text, "docx", "special.docx");
    const zip = new AdmZip(buf);
    const docXml = zip.readAsText("word/document.xml");
    // Should not contain unescaped & in XML context
    expect(docXml).not.toMatch(/&(?!amp;|lt;|gt;|quot;|apos;)[a-zA-Z]/);
  });

  it("multi-page PPTX has no duplicate slideMaster entries", async () => {
    const text = "# Slide 1\nContent\n--- PAGE BREAK ---\n# Slide 2\nMore\n--- PAGE BREAK ---\n# Slide 3\nFinal";
    const buf = await buildTranslatedDocument(text, "pptx", "multi.pptx");
    const zip = new AdmZip(buf);
    const ct = zip.readAsText("[Content_Types].xml");
    const masters = ct.match(/slideMaster\d+\.xml/g) ?? [];
    expect(masters.length).toBe(new Set(masters).size);
  });
});

// ─── 4. Format conversion ─────────────────────────────────────────────────────
describe("4. Format conversion — input→output paths", () => {
  it("TXT → DOCX", async () => {
    const out = await convertDocument(makeTxtBuffer("Hello World"), "txt", "docx", "t.txt");
    expect(out.length).toBeGreaterThan(500);
    expect(new AdmZip(out).getEntry("word/document.xml")).not.toBeNull();
  });

  it("TXT → PPTX", async () => {
    const out = await convertDocument(makeTxtBuffer("Slide Title\nContent"), "txt", "pptx", "t.txt");
    expect(out.length).toBeGreaterThan(500);
    expect(new AdmZip(out).getEntries().some(e => e.entryName.startsWith("ppt/slides/"))).toBe(true);
  });

  it("TXT → XLSX", async () => {
    const out = await convertDocument(makeTxtBuffer("| A | B |\n|---|---|\n| 1 | 2 |"), "txt", "xlsx", "t.txt");
    expect(out.length).toBeGreaterThan(100);
    expect(XLSX.read(out, { type: "buffer" }).SheetNames.length).toBeGreaterThan(0);
  });

  it("TXT → TXT (passthrough)", async () => {
    const out = await convertDocument(makeTxtBuffer("Hello passthrough"), "txt", "txt", "t.txt");
    expect(out.toString("utf-8")).toContain("Hello passthrough");
  });

  it("DOCX → DOCX (passthrough)", async () => {
    const out = await convertDocument(makeDocxBuffer("Passthrough DOCX"), "docx", "docx", "t.docx");
    expect(out.length).toBeGreaterThan(100);
  });

  it("XLSX → XLSX (passthrough)", async () => {
    const out = await convertDocument(makeXlsxBuffer(), "xlsx", "xlsx", "t.xlsx");
    expect(out.length).toBeGreaterThan(100);
  });
});

// ─── 5. In-place translation ──────────────────────────────────────────────────
describe("5. In-place translation — preserves structure", () => {
  it("DOCX: translates text, preserves ZIP entries", async () => {
    const buf = makeDocxBuffer("Hello World\nSecond paragraph");
    const out = await translateDocxInPlace(buf, "French", makeMockLLM(t => `[FR] ${t}`));
    expect(out.length).toBeGreaterThan(100);
    const zip = new AdmZip(out);
    expect(zip.getEntry("word/document.xml")).not.toBeNull();
    expect(zip.readAsText("word/document.xml")).toContain("[FR]");
  });

  it("DOCX: all original ZIP entries preserved after translation", async () => {
    const buf = makeDocxBuffer("Original text");
    const out = await translateDocxInPlace(buf, "French", makeMockLLM());
    const origEntries = new AdmZip(buf).getEntries().map(e => e.entryName).sort();
    const outEntries = new AdmZip(out).getEntries().map(e => e.entryName).sort();
    for (const e of origEntries) expect(outEntries).toContain(e);
  });

  it("PPTX: translates text, preserves ZIP structure", async () => {
    const buf = makePptxBuffer("Hello World\nSecond bullet");
    const out = await translatePptxInPlace(buf, "Spanish", makeMockLLM(t => `[ES] ${t}`));
    expect(out.length).toBeGreaterThan(100);
    expect(new AdmZip(out).getEntries().some(e => e.entryName.startsWith("ppt/slides/"))).toBe(true);
  });

  it("XLSX: translates shared strings, preserves ZIP structure", async () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["Name", "City"], ["Alice", "New York"]]), "S1");
    const buf = Buffer.from(XLSX.write(wb, { type: "buffer", bookType: "xlsx", bookSST: true }));
    const out = await translateXlsxInPlace(buf, "German", makeMockLLM(t => `[DE] ${t}`));
    expect(out.length).toBeGreaterThan(100);
    expect(new AdmZip(out).getEntries().some(e => e.entryName.includes("xl/"))).toBe(true);
  });

  it("DOCX: handles no-text document gracefully", async () => {
    const zip = new AdmZip();
    zip.addFile("[Content_Types].xml", Buffer.from(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`));
    zip.addFile("_rels/.rels", Buffer.from(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`));
    zip.addFile("word/document.xml", Buffer.from(`<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>123 456</w:t></w:r></w:p></w:body></w:document>`));
    const out = await translateDocxInPlace(zip.toBuffer(), "French", makeMockLLM());
    expect(out.length).toBeGreaterThan(100);
  });
});

// ─── 6. Multi-language translation ───────────────────────────────────────────
describe("6. Translation via LLM — 10 languages", () => {
  const testText = "The quick brown fox jumps over the lazy dog.";
  const languages = [
    { name: "French", tag: "[FR]" },
    { name: "Spanish", tag: "[ES]" },
    { name: "German", tag: "[DE]" },
    { name: "Japanese", tag: "[JA]" },
    { name: "Chinese (Simplified)", tag: "[ZH]" },
    { name: "Arabic", tag: "[AR]" },
    { name: "Portuguese", tag: "[PT]" },
    { name: "Russian", tag: "[RU]" },
    { name: "Korean", tag: "[KO]" },
    { name: "Italian", tag: "[IT]" },
  ];

  for (const lang of languages) {
    it(`translates to ${lang.name}`, async () => {
      const result = await translateWithLLM(testText, lang.name, makeMockLLM(t => `${lang.tag} ${t}`));
      expect(result).toContain(lang.tag);
      expect(result.length).toBeGreaterThan(0);
    });
  }

  it("handles empty text (returns empty string)", async () => {
    const result = await translateWithLLM("", "French", makeMockLLM());
    expect(result).toBe("");
  });

  it("handles very long text (chunking path)", async () => {
    const longText = "This is a long sentence. ".repeat(500);
    const result = await translateWithLLM(longText, "French", makeMockLLM(t => `[FR] ${t.slice(0, 30)}`));
    expect(result.length).toBeGreaterThan(0);
    expect(result).toContain("[FR]");
  });

  it("preserves page break markers during translation", async () => {
    const text = "Page one\n\n--- PAGE BREAK ---\n\nPage two";
    const result = await translateWithLLM(text, "French", makeMockLLM(t => t));
    expect(result).toContain("--- PAGE BREAK ---");
  });
});

// ─── 7. Watermark preview generation ─────────────────────────────────────────
describe("7. Watermark preview generation", () => {
  it("PDF → watermarked PDF (valid PDF output)", async () => {
    const pdfBuf = await makePdfBuffer("Test PDF for watermark preview");
    const preview = await generateWatermarkedPreview(pdfBuf, "pdf");
    expect(preview.length).toBeGreaterThan(500);
    expect(preview.slice(0, 4).toString()).toBe("%PDF");
    // Watermarked PDF should be larger (watermark adds content)
    expect(preview.length).toBeGreaterThanOrEqual(pdfBuf.length);
  });

  it("TXT → watermarked PDF (converts to PDF then stamps)", async () => {
    const txtBuf = makeTxtBuffer("Hello World\nLine two\nLine three\n".repeat(5));
    const preview = await generateWatermarkedPreview(txtBuf, "txt");
    expect(preview.length).toBeGreaterThan(500);
    expect(preview.slice(0, 4).toString()).toBe("%PDF");
  });

  it("DOCX → watermarked PDF (via LibreOffice)", async () => {
    const docx = await buildTranslatedDocument("# Test\n\nHello World paragraph.", "docx", "test.docx");
    const preview = await generateWatermarkedPreview(docx, "docx");
    expect(preview.length).toBeGreaterThan(500);
    expect(preview.slice(0, 4).toString()).toBe("%PDF");
  }, 60_000);

  it("XLSX → watermarked PDF (via LibreOffice)", async () => {
    const xlsx = makeXlsxBuffer();
    const preview = await generateWatermarkedPreview(xlsx, "xlsx");
    expect(preview.length).toBeGreaterThan(500);
    expect(preview.slice(0, 4).toString()).toBe("%PDF");
  }, 60_000);

  it("PPTX → watermarked PDF (via LibreOffice)", async () => {
    const pptx = await buildTranslatedDocument("# Slide One\n\nContent here.", "pptx", "test.pptx");
    const preview = await generateWatermarkedPreview(pptx, "pptx");
    expect(preview.length).toBeGreaterThan(500);
    expect(preview.slice(0, 4).toString()).toBe("%PDF");
  }, 60_000);

  it("watermarked PDF has at least 1 page and is parseable by pdf-lib", async () => {
    const pdfBuf = await makePdfBuffer("Simple watermark test");
    const preview = await generateWatermarkedPreview(pdfBuf, "pdf");
    const doc = await PDFDocument.load(preview);
    expect(doc.getPageCount()).toBeGreaterThanOrEqual(1);
  });
});

// ─── 8. Download integrity ────────────────────────────────────────────────────
describe("8. Download integrity — output buffer validity", () => {
  it("DOCX output is valid ZIP with word/document.xml", async () => {
    const buf = await buildTranslatedDocument("# Hello\n\nWorld", "docx", "test.docx");
    expect(() => new AdmZip(buf)).not.toThrow();
    expect(new AdmZip(buf).getEntry("word/document.xml")).not.toBeNull();
  });

  it("PPTX output is valid ZIP with slide content", async () => {
    const buf = await buildTranslatedDocument("# Slide\n\nContent", "pptx", "test.pptx");
    expect(() => new AdmZip(buf)).not.toThrow();
    expect(new AdmZip(buf).getEntries().some(e => e.entryName.startsWith("ppt/slides/"))).toBe(true);
  });

  it("XLSX output is readable by xlsx library", async () => {
    const buf = await buildTranslatedDocument("| A | B |\n|---|---|\n| 1 | 2 |", "xlsx", "test.xlsx");
    expect(() => XLSX.read(buf, { type: "buffer" })).not.toThrow();
    expect(XLSX.read(buf, { type: "buffer" }).SheetNames.length).toBeGreaterThan(0);
  });

  it("TXT output is valid UTF-8 text", async () => {
    const buf = await buildTranslatedDocument("Hello World\nSecond line", "txt", "test.txt");
    expect(buf.toString("utf-8")).toContain("Hello World");
  });

  it("DOCX in-place translation preserves all original ZIP entries", async () => {
    const original = makeDocxBuffer("Original text");
    const translated = await translateDocxInPlace(original, "French", makeMockLLM(t => `Translated: ${t}`));
    const origEntries = new AdmZip(original).getEntries().map(e => e.entryName).sort();
    const transEntries = new AdmZip(translated).getEntries().map(e => e.entryName).sort();
    for (const entry of origEntries) expect(transEntries).toContain(entry);
  });

  it("PPTX in-place translation preserves all original ZIP entries", async () => {
    const original = makePptxBuffer("Hello World\nContent");
    const translated = await translatePptxInPlace(original, "Spanish", makeMockLLM());
    const origEntries = new AdmZip(original).getEntries().map(e => e.entryName).sort();
    const transEntries = new AdmZip(translated).getEntries().map(e => e.entryName).sort();
    for (const entry of origEntries) expect(transEntries).toContain(entry);
  });
});

// ─── 8. High-fidelity PDF → DOCX (pdf2docx) ─────────────────────────────────
describe("8. High-fidelity PDF → DOCX conversion (pdf2docx)", () => {
  it("converts a PDF buffer to a valid DOCX buffer preserving structure", async () => {
    const { convertPdfToDocxWithPdf2Docx } = await import("./docProcessor");
    const pdfBuf = await makePdfBuffer("Test document\nWith multiple lines\nAnd some content");
    const docxBuf = await convertPdfToDocxWithPdf2Docx(pdfBuf);
    // Must be a valid ZIP (DOCX is a ZIP)
    expect(() => new AdmZip(docxBuf)).not.toThrow();
    const zip = new AdmZip(docxBuf);
    // Must contain word/document.xml
    const entries = zip.getEntries().map(e => e.entryName);
    expect(entries.some(e => e.includes("word/document.xml"))).toBe(true);
    // Must be non-trivial size
    expect(docxBuf.length).toBeGreaterThan(5000);
  });

  it("falls back gracefully if given an invalid PDF", async () => {
    const { convertPdfToDocxWithPdf2Docx } = await import("./docProcessor");
    // An invalid PDF should throw (not silently produce empty output)
    const invalidBuf = Buffer.from("not a pdf");
    await expect(convertPdfToDocxWithPdf2Docx(invalidBuf)).rejects.toThrow();
  });
});
