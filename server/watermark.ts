/**
 * watermark.ts
 * Generates a native-format preview from any supported output format,
 * truncated to the first PREVIEW_PAGE_LIMIT pages/slides.
 *
 * - PDF   → pdf-lib: crop to N pages + diagonal watermark overlay
 * - DOCX  → JSZip: trim to half-page content, watermark header, native DOCX
 * - PPTX  → AdmZip: keep first N slides + add watermark shape to each
 * - XLSX  → AdmZip: add watermark row in each sheet
 * - TXT   → pdf-lib: render as text PDF (fallback)
 */

import { PDFDocument, rgb, StandardFonts, PDFPage, degrees } from "pdf-lib";

export const PREVIEW_PAGE_LIMIT = 3;
const WATERMARK_TEXT = "Done by Docu Translantes";

// ─── PDF helpers ──────────────────────────────────────────────────────────────

function stampWatermarkOnPdf(previewDoc: PDFDocument, watermarkFont: Awaited<ReturnType<PDFDocument["embedFont"]>>) {
  const watermarkSize = 28;
  for (const pg of previewDoc.getPages()) {
    const { width: pw, height: ph } = pg.getSize();
    const positions = [
      { x: pw * 0.08, y: ph * 0.72 },
      { x: pw * 0.22, y: ph * 0.46 },
      { x: pw * 0.36, y: ph * 0.20 },
    ];
    for (const pos of positions) {
      pg.drawText(WATERMARK_TEXT, {
        x: pos.x,
        y: pos.y,
        size: watermarkSize,
        font: watermarkFont,
        color: rgb(0.75, 0.75, 0.75),
        opacity: 0.35,
        rotate: degrees(-45),
      });
    }
  }
}

async function truncatePdf(
  pdfBuffer: Buffer,
  maxPages: number,
): Promise<{ truncated: Buffer; totalPages: number; previewPages: number }> {
  const srcDoc = await PDFDocument.load(pdfBuffer, { ignoreEncryption: true });
  const totalPages = srcDoc.getPageCount();

  const previewDoc = await PDFDocument.create();
  const keepCount = Math.min(maxPages, totalPages);
  const pageIndices = Array.from({ length: keepCount }, (_, i) => i);
  const copiedPages = await previewDoc.copyPages(srcDoc, pageIndices);
  for (const pg of copiedPages) previewDoc.addPage(pg);

  // Watermark each content page
  const watermarkFont = await previewDoc.embedFont(StandardFonts.HelveticaBold);
  stampWatermarkOnPdf(previewDoc, watermarkFont);

  // "Preview ends here" gate page (only if we actually truncated)
  if (totalPages > maxPages) {
    const { width, height } = copiedPages[copiedPages.length - 1].getSize();
    const gatePage = previewDoc.addPage([width, height]);
    const font = await previewDoc.embedFont(StandardFonts.HelveticaBold);
    const bodyFont = await previewDoc.embedFont(StandardFonts.Helvetica);

    gatePage.drawRectangle({ x: 0, y: 0, width, height, color: rgb(0.97, 0.97, 0.97) });

    const titleText = "Preview ends here";
    const titleSize = 22;
    gatePage.drawText(titleText, {
      x: (width - font.widthOfTextAtSize(titleText, titleSize)) / 2,
      y: height / 2 + 60,
      size: titleSize, font, color: rgb(0.15, 0.15, 0.15),
    });

    const subText = `You've seen ${maxPages} of ${totalPages} pages`;
    const subSize = 14;
    gatePage.drawText(subText, {
      x: (width - bodyFont.widthOfTextAtSize(subText, subSize)) / 2,
      y: height / 2 + 28,
      size: subSize, font: bodyFont, color: rgb(0.4, 0.4, 0.4),
    });

    const ctaText = "Purchase to download the full document";
    const ctaSize = 13;
    gatePage.drawText(ctaText, {
      x: (width - bodyFont.widthOfTextAtSize(ctaText, ctaSize)) / 2,
      y: height / 2 - 4,
      size: ctaSize, font: bodyFont, color: rgb(0.85, 0.45, 0.1),
    });
  }

  const bytes = await previewDoc.save();
  return { truncated: Buffer.from(bytes), totalPages, previewPages: keepCount };
}

// ─── TXT fallback ─────────────────────────────────────────────────────────────

function sanitizeForWinAnsi(raw: string): { text: string; hadNonLatin: boolean } {
  let hadNonLatin = false;
  const text = raw.split("").map(ch => {
    const cp = ch.charCodeAt(0);
    if ((cp >= 32 && cp <= 126) || cp === 10 || cp === 13 || cp === 9) return ch;
    if (cp >= 160 && cp <= 255) return ch;
    hadNonLatin = true;
    return "";
  }).join("");
  return { text, hadNonLatin };
}

async function buildTextPdf(textBuffer: Buffer): Promise<Buffer> {
  const { text: rawText, hadNonLatin } = sanitizeForWinAnsi(textBuffer.toString("utf-8"));
  const header = hadNonLatin
    ? "[Preview: translated document — download the file to view full formatting]\n\n"
    : "";
  const text = header + rawText;

  const pdfDoc = await PDFDocument.create();
  const font = await pdfDoc.embedFont(StandardFonts.Courier);
  const pageWidth = 595, pageHeight = 842, margin = 50, lineHeight = 14, fontSize = 11;
  const maxWidth = pageWidth - margin * 2;
  const maxLinesPerPage = Math.floor((pageHeight - margin * 2) / lineHeight);

  const rawLines = text.split(/\r?\n/);
  const wrappedLines: string[] = [];
  for (const rawLine of rawLines) {
    if (rawLine.length === 0) { wrappedLines.push(""); continue; }
    let remaining = rawLine;
    while (remaining.length > 0) {
      let end = remaining.length;
      try {
        while (end > 0 && font.widthOfTextAtSize(remaining.slice(0, end), fontSize) > maxWidth) end--;
      } catch { end = 1; }
      if (end === 0) end = 1;
      wrappedLines.push(remaining.slice(0, end));
      remaining = remaining.slice(end);
    }
  }

  let page: PDFPage | null = null;
  let lineY = 0, lineCount = 0;
  for (const line of wrappedLines) {
    if (!page || lineCount >= maxLinesPerPage) {
      page = pdfDoc.addPage([pageWidth, pageHeight]);
      lineY = pageHeight - margin - lineHeight;
      lineCount = 0;
    }
    if (line.length > 0) {
      try {
        page.drawText(line, { x: margin, y: lineY, size: fontSize, font, color: rgb(0.1, 0.1, 0.1) });
      } catch { /* skip unencodable lines */ }
    }
    lineY -= lineHeight;
    lineCount++;
  }
  if (pdfDoc.getPageCount() === 0) pdfDoc.addPage([pageWidth, pageHeight]);
  return Buffer.from(await pdfDoc.save());
}

// ─── DOCX preview ────────────────────────────────────────────────────────────
// Produces a native DOCX preview by modifying ONLY the document body:
//  1. Prepends a watermark paragraph at the top
//  2. Keeps only the top half of content (up to 3 pages max)
//  3. Appends a "Download for the full document" message
// NO header/footer/relationship changes — Office Online rejects those.

const PARAS_PER_PAGE = 20;
const MAX_PREVIEW_PAGES = 3;

const WATERMARK_PARAGRAPH = `<w:p><w:pPr><w:jc w:val="center"/><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="4" w:color="D0D0D0"/></w:pBdr><w:spacing w:after="120"/></w:pPr><w:r><w:rPr><w:color w:val="B0B0B0"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:i/><w:iCs/></w:rPr><w:t xml:space="preserve">PREVIEW — ${WATERMARK_TEXT}</w:t></w:r></w:p>`;

const DOWNLOAD_MSG_PARAGRAPH = `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="1200" w:after="200"/></w:pPr><w:r><w:rPr><w:color w:val="999999"/><w:sz w:val="28"/><w:szCs w:val="28"/><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">— Download for the full document —</w:t></w:r></w:p>`;

async function previewDocx(buffer: Buffer): Promise<Buffer> {
  const JSZip = (await import("jszip")).default;
  const zip = await JSZip.loadAsync(buffer);
  const docFile = zip.file("word/document.xml");
  if (!docFile) return buffer;

  let docXml = await docFile.async("text");

  const bodyOpenMatch = docXml.match(/<w:body(?:\s[^>]*)?>/);
  const bodyCloseIdx = docXml.lastIndexOf("</w:body>");
  if (!bodyOpenMatch || bodyOpenMatch.index === undefined || bodyCloseIdx === -1) return buffer;

  const bodyOpenEnd = bodyOpenMatch.index + bodyOpenMatch[0].length;
  const bodyContent = docXml.slice(bodyOpenEnd, bodyCloseIdx);

  // Count total paragraphs
  let totalParas = 0;
  const paraCountRe = /<w:p[\s>]/g;
  while (paraCountRe.exec(bodyContent) !== null) totalParas++;

  // Estimate pages and determine how many paragraphs to keep (top half)
  const estimatedPages = Math.max(1, Math.ceil(totalParas / PARAS_PER_PAGE));
  const pagesToPreview = Math.min(estimatedPages, MAX_PREVIEW_PAGES);
  const parasForPages = pagesToPreview * PARAS_PER_PAGE;
  const parasToKeep = Math.max(1, Math.floor(parasForPages / 2));

  // Preserve the sectPr (page layout settings) — always needed at end of body
  const sectPrMatch = bodyContent.match(/<w:sectPr(?:\s[^>]*)?>[\s\S]*?<\/w:sectPr>/);
  const sectPr = sectPrMatch ? sectPrMatch[0] : "";

  let trimmedBody: string;
  if (totalParas > parasToKeep) {
    const paraRe = /<w:p[\s>]/g;
    let count = 0;
    let cutIdx = -1;
    let m: RegExpExecArray | null;
    while ((m = paraRe.exec(bodyContent)) !== null) {
      count++;
      if (count > parasToKeep) { cutIdx = m.index; break; }
    }
    trimmedBody = cutIdx !== -1 ? bodyContent.slice(0, cutIdx) : bodyContent;
  } else {
    // Short doc — keep all content but remove sectPr (we'll re-append it)
    trimmedBody = sectPr ? bodyContent.replace(sectPr, "") : bodyContent;
  }

  // Reassemble: watermark + trimmed content + download message + sectPr
  docXml =
    docXml.slice(0, bodyOpenEnd) +
    WATERMARK_PARAGRAPH +
    trimmedBody +
    DOWNLOAD_MSG_PARAGRAPH +
    sectPr +
    "</w:body>" +
    docXml.slice(bodyCloseIdx + "</w:body>".length);

  zip.file("word/document.xml", docXml);

  return Buffer.from(await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } }));
}

// ─── PPTX preview (first N slides + watermark per slide) ─────────────────────
// Uses JSZip instead of AdmZip to handle all PPTX ZIP variants correctly.

async function previewPptx(buffer: Buffer, maxSlides: number): Promise<{ preview: Buffer; slideCount: number }> {
  const JSZip = (await import("jszip")).default;
  const zip = await JSZip.loadAsync(buffer);

  const slidePattern = /^ppt\/slides\/slide(\d+)\.xml$/;
  const slideNames = Object.keys(zip.files)
    .filter(n => slidePattern.test(n))
    .sort((a, b) => {
      const na = parseInt(a.match(/\d+/)![0]);
      const nb = parseInt(b.match(/\d+/)![0]);
      return na - nb;
    });

  const totalSlides = slideNames.length;
  const keepCount = Math.min(maxSlides, totalSlides);

  // Watermark shape XML — diagonal text overlaid on each kept slide
  const wmShape = `<p:sp><p:nvSpPr><p:cNvPr id="9999" name="DocuTranslantesWatermark"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm rot="-2700000"><a:off x="914400" y="2286000"/><a:ext cx="7086000" cy="1524000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="2800" b="1" dirty="0"><a:solidFill><a:srgbClr val="BFBFBF"><a:alpha val="40000"/></a:srgbClr></a:solidFill></a:rPr><a:t>${WATERMARK_TEXT}</a:t></a:r></a:p></p:txBody></p:sp>`;

  // Stamp watermark on kept slides
  for (let i = 0; i < keepCount; i++) {
    const file = zip.file(slideNames[i]);
    if (!file) continue;
    let xml = await file.async("text");
    xml = xml.replace(/<\/p:spTree>/, `${wmShape}</p:spTree>`);
    zip.file(slideNames[i], xml);
  }

  // Remove slides beyond keepCount
  if (totalSlides > keepCount) {
    const presRelsFile = zip.file("ppt/_rels/presentation.xml.rels");
    let ridsToRemove: string[] = [];

    if (presRelsFile) {
      let presRelsXml = await presRelsFile.async("text");
      for (let i = keepCount; i < slideNames.length; i++) {
        const target = slideNames[i].replace("ppt/", "");
        const escapedTarget = target.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        const m = presRelsXml.match(new RegExp(`Id="([^"]+)"[^>]+Target="${escapedTarget}"`));
        if (m) ridsToRemove.push(m[1]);
        presRelsXml = presRelsXml.replace(
          new RegExp(`<Relationship[^>]+Target="${escapedTarget}"[^>]*/>`), "",
        );
      }
      zip.file("ppt/_rels/presentation.xml.rels", presRelsXml);
    }

    // Remove sldId from presentation.xml
    const presFile = zip.file("ppt/presentation.xml");
    if (presFile && ridsToRemove.length > 0) {
      let presXml = await presFile.async("text");
      for (const rid of ridsToRemove) {
        presXml = presXml.replace(new RegExp(`<p:sldId[^>]+r:id="${rid}"[^>]*/>`), "");
      }
      zip.file("ppt/presentation.xml", presXml);
    }

    // Delete slide files, their rels, and clean Content_Types
    for (let i = keepCount; i < slideNames.length; i++) {
      const name = slideNames[i];
      zip.remove(name);
      zip.remove(name.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels");

      const ctFile = zip.file("[Content_Types].xml");
      if (ctFile) {
        let ct = await ctFile.async("text");
        const escapedName = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        ct = ct.replace(new RegExp(`<Override PartName="/${escapedName}"[^/]*/>`), "");
        zip.file("[Content_Types].xml", ct);
      }
    }
  }

  const preview = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } });
  return { preview, slideCount: totalSlides };
}

// ─── XLSX preview (first 3 sheets, first 100 rows per sheet) ─────────────────
// Watermark is achieved by prefixing each kept sheet name with "PREVIEW:"
// in workbook.xml — zero cell manipulation means zero risk of XML corruption.

/** Stateful row trimmer — safer than a global greedy regex on large files. */
function trimSheetRows(wsXml: string, maxRows: number): string {
  const sdStart = wsXml.indexOf("<sheetData");
  if (sdStart === -1) return wsXml;

  // Self-closing <sheetData/> — nothing to trim
  const sdTagEnd = wsXml.indexOf(">", sdStart) + 1;
  if (wsXml.slice(sdStart, sdTagEnd).trimEnd().endsWith("/>")) return wsXml;

  const sdClose = wsXml.lastIndexOf("</sheetData>");
  if (sdClose === -1) return wsXml;

  const before = wsXml.slice(0, sdTagEnd);
  const content = wsXml.slice(sdTagEnd, sdClose);
  const after = wsXml.slice(sdClose);

  let result = "";
  let pos = 0;

  while (pos < content.length) {
    const rowStart = content.indexOf("<row", pos);
    if (rowStart === -1) { result += content.slice(pos); break; }

    // Preserve whitespace/text before this row tag
    result += content.slice(pos, rowStart);

    const headerEnd = content.indexOf(">", rowStart);
    if (headerEnd === -1) { result += content.slice(rowStart); break; }

    const rowHeader = content.slice(rowStart, headerEnd + 1);
    const rMatch = rowHeader.match(/\br="(\d+)"/);
    const rowNum = rMatch ? parseInt(rMatch[1], 10) : 0;

    // Self-closing row (<row ... />)
    if (rowHeader.trimEnd().endsWith("/>")) {
      if (rowNum <= maxRows) result += rowHeader;
      pos = rowStart + rowHeader.length;
      continue;
    }

    const rowClose = content.indexOf("</row>", headerEnd);
    if (rowClose === -1) { result += content.slice(rowStart); break; }

    const fullRow = content.slice(rowStart, rowClose + 6); // 6 = "</row>".length
    if (rowNum <= maxRows) result += fullRow;
    pos = rowClose + 6;
  }

  return before + result + after;
}

async function previewXlsx(buffer: Buffer, maxSheets = 3, maxRows = 100): Promise<Buffer> {
  const JSZip = (await import("jszip")).default;
  const zip = await JSZip.loadAsync(buffer);

  // ── 1. Determine which sheets to keep via workbook.xml ───────────────────
  const wbFile = zip.file("xl/workbook.xml");
  if (!wbFile) return buffer;

  let wbXml = await wbFile.async("text");

  const allSheetEls: string[] = [];
  const sheetElRe = /<sheet\s[^>]*\/>/g;
  let sm: RegExpExecArray | null;
  while ((sm = sheetElRe.exec(wbXml)) !== null) allSheetEls.push(sm[0]);

  const wbRelsFile = zip.file("xl/_rels/workbook.xml.rels");
  let wbRelsXml = wbRelsFile ? await wbRelsFile.async("text") : "";
  const rIdToTarget = new Map<string, string>();
  const relRe = /Id="([^"]+)"[^>]+Target="([^"]+)"/g;
  let rm: RegExpExecArray | null;
  while ((rm = relRe.exec(wbRelsXml)) !== null) rIdToTarget.set(rm[1], rm[2]);

  const keptRIds: Set<string> = new Set();
  for (let i = 0; i < Math.min(maxSheets, allSheetEls.length); i++) {
    const rIdMatch = allSheetEls[i].match(/r:id="([^"]+)"/);
    if (rIdMatch) keptRIds.add(rIdMatch[1]);
    wbXml = wbXml.replace(allSheetEls[i], allSheetEls[i].replace(/name="([^"]*)"/, `name="PREVIEW: $1"`));
  }

  for (let i = maxSheets; i < allSheetEls.length; i++) {
    const rIdMatch = allSheetEls[i].match(/r:id="([^"]+)"/);
    const rId = rIdMatch?.[1];
    wbXml = wbXml.replace(allSheetEls[i], "");

    if (rId) {
      const target = rIdToTarget.get(rId);
      if (target) {
        const fullPath = `xl/${target}`;
        zip.remove(fullPath);
        zip.remove(fullPath.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels");

        const ctFile = zip.file("[Content_Types].xml");
        if (ctFile) {
          const escapedPath = fullPath.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
          let ct = await ctFile.async("text");
          ct = ct.replace(new RegExp(`<Override[^>]+PartName="/${escapedPath}"[^>]*/>`), "");
          zip.file("[Content_Types].xml", ct);
        }
      }
      wbRelsXml = wbRelsXml.replace(
        new RegExp(`<Relationship[^>]+Id="${rId}"[^>]*/>`), "",
      );
    }
  }

  zip.file("xl/workbook.xml", wbXml);
  if (wbRelsFile) zip.file("xl/_rels/workbook.xml.rels", wbRelsXml);

  // ── 2. Trim rows in each kept sheet ──────────────────────────────────────
  const wsPattern = /^xl\/worksheets\/sheet\d+\.xml$/;
  for (const name of Object.keys(zip.files)) {
    if (!wsPattern.test(name)) continue;
    const file = zip.file(name);
    if (!file) continue;
    const trimmed = trimSheetRows(await file.async("text"), maxRows);
    zip.file(name, trimmed);
  }

  return Buffer.from(await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } }));
}

// ─── Main export ──────────────────────────────────────────────────────────────

/**
 * Given the final translated output buffer and its format, return a preview
 * in the SAME format (not always PDF), truncated to PREVIEW_PAGE_LIMIT pages
 * and stamped with the "Done by Docu Translantes" watermark.
 */
export async function generatePreview(
  outputBuffer: Buffer,
  outputFormat: string,
): Promise<{ previewBuffer: Buffer; previewPages: number; previewFormat: string; previewMimeType: string }> {
  const fmt = outputFormat.toLowerCase();

  if (fmt === "pdf") {
    const { truncated, previewPages } = await truncatePdf(outputBuffer, PREVIEW_PAGE_LIMIT);
    return { previewBuffer: truncated, previewPages, previewFormat: "pdf", previewMimeType: "application/pdf" };
  }

  if (fmt === "docx") {
    const preview = await previewDocx(outputBuffer);
    return { previewBuffer: preview, previewPages: PREVIEW_PAGE_LIMIT, previewFormat: "docx", previewMimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" };
  }

  if (fmt === "pptx") {
    const { preview, slideCount } = await previewPptx(outputBuffer, PREVIEW_PAGE_LIMIT);
    return { previewBuffer: preview, previewPages: Math.min(slideCount, PREVIEW_PAGE_LIMIT), previewFormat: "pptx", previewMimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" };
  }

  if (fmt === "xlsx") {
    const preview = await previewXlsx(outputBuffer);
    return { previewBuffer: preview, previewPages: 1, previewFormat: "xlsx", previewMimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" };
  }

  if (fmt === "txt" || fmt === "csv") {
    const pdfBuf = await buildTextPdf(outputBuffer);
    const { truncated, previewPages } = await truncatePdf(pdfBuf, PREVIEW_PAGE_LIMIT);
    return { previewBuffer: truncated, previewPages, previewFormat: "pdf", previewMimeType: "application/pdf" };
  }

  // Unknown format — return raw buffer as-is
  return { previewBuffer: outputBuffer, previewPages: 1, previewFormat: fmt, previewMimeType: "application/octet-stream" };
}

// ─── Legacy compat ────────────────────────────────────────────────────────────

export async function convertToPdfBuffer(inputBuffer: Buffer, _extension: string): Promise<Buffer> {
  return inputBuffer;
}

export async function generateWatermarkedPreview(outputBuffer: Buffer, outputFormat: string): Promise<Buffer> {
  const { previewBuffer } = await generatePreview(outputBuffer, outputFormat);
  return previewBuffer;
}
