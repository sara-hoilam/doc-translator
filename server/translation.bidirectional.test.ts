/**
 * Bidirectional translation test suite
 *
 * Tests all 9 target languages × all 9 source languages (72 pairs) to verify:
 * 1. pptxShouldSkip() correctly decides whether to translate or skip text
 *    based on the source script AND the target language direction
 * 2. translatePptxInPlace() actually sends CJK/Arabic source text to the LLM
 *    when the target is a Latin-script language (the bug that was fixed)
 * 3. translateWithLLM() and translateBatchWithLLM() correctly call the LLM
 *    for all language pairs
 *
 * Uses a mock LLM (no real API calls) to keep tests fast and deterministic.
 */

import { describe, expect, it, vi } from "vitest";
import AdmZip from "adm-zip";
import {
  translatePptxInPlace,
  translateWithLLM,
  translateBatchWithLLM,
} from "./docProcessor";

// ─── Language definitions ─────────────────────────────────────────────────────

const LANGUAGES = [
  {
    name: "Simplified Chinese",
    sampleText: "市场分析和商业策略概述",          // Market analysis and business strategy overview
    isCJKOrRTL: true,
    script: "cjk",
  },
  {
    name: "Traditional Chinese",
    sampleText: "市場分析和商業策略概述",           // Traditional Chinese characters
    isCJKOrRTL: true,
    script: "cjk",
  },
  {
    name: "English",
    sampleText: "Market analysis and business strategy overview",
    isCJKOrRTL: false,
    script: "latin",
  },
  {
    name: "Spanish",
    sampleText: "Análisis de mercado y descripción general de la estrategia empresarial",
    isCJKOrRTL: false,
    script: "latin",
  },
  {
    name: "Latin",
    sampleText: "Analysis mercatus et conspectus consilii negotii",
    isCJKOrRTL: false,
    script: "latin",
  },
  {
    name: "Korean",
    sampleText: "시장 분석 및 비즈니스 전략 개요",   // Market analysis and business strategy overview
    isCJKOrRTL: true,
    script: "cjk",
  },
  {
    name: "Japanese",
    sampleText: "市場分析とビジネス戦略の概要",       // Market analysis and business strategy overview
    isCJKOrRTL: true,
    script: "cjk",
  },
  {
    name: "French",
    sampleText: "Analyse de marché et aperçu de la stratégie commerciale",
    isCJKOrRTL: false,
    script: "latin",
  },
  {
    name: "Arabic",
    sampleText: "تحليل السوق ونظرة عامة على استراتيجية الأعمال",
    isCJKOrRTL: true,
    script: "rtl",
  },
] as const;

// ─── Helper: build a minimal valid PPTX buffer with one slide ────────────────

function buildMinimalPptx(slideText: string): Buffer {
  const zip = new AdmZip();

  // [Content_Types].xml
  zip.addFile(
    "[Content_Types].xml",
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`
    )
  );

  // _rels/.rels
  zip.addFile(
    "_rels/.rels",
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`
    )
  );

  // ppt/presentation.xml
  zip.addFile(
    "ppt/presentation.xml",
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst/>
  <p:sldSz cx="9144000" cy="6858000"/>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
</p:presentation>`
    )
  );

  // ppt/_rels/presentation.xml.rels
  zip.addFile(
    "ppt/_rels/presentation.xml.rels",
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`
    )
  );

  // Escape XML special chars in the slide text
  const escaped = slideText
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");

  // ppt/slides/slide1.xml
  zip.addFile(
    "ppt/slides/slide1.xml",
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="2400" b="1"/>
              <a:t>${escaped}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`
    )
  );

  // ppt/slides/_rels/slide1.xml.rels
  zip.addFile(
    "ppt/slides/_rels/slide1.xml.rels",
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`
    )
  );

  return zip.toBuffer();
}

// ─── Helper: extract text from a PPTX buffer ─────────────────────────────────

function extractTextFromPptx(buffer: Buffer): string {
  const zip = new AdmZip(buffer);
  const entry = zip.getEntry("ppt/slides/slide1.xml");
  if (!entry) return "";
  const xml = entry.getData().toString("utf-8");
  const matches = xml.match(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g) ?? [];
  return matches.map(m => m.replace(/<[^>]+>/g, "")).join(" ").trim();
}

// ─── Mock LLM factory ────────────────────────────────────────────────────────
// Returns a mock LLM that:
//   - Records every call (so we can assert it was/wasn't called)
//   - Returns a predictable "TRANSLATED: <original>" response for numbered lists
//   - Returns a predictable "TRANSLATED: <original>" response for plain text

function makeMockLLM() {
  const calls: { messages: any[] }[] = [];

  const mockLLM = async (params: { messages: any[] }) => {
    calls.push(params);
    const userContent = params.messages[params.messages.length - 1]?.content ?? "";
    const text = typeof userContent === "string" ? userContent : JSON.stringify(userContent);

    // Detect [N] format used by translateBatchWithLLM (DOCX/XLSX path)
    const bracketLines = text.match(/^\[\d+\]\s+.+/gm) ?? [];
    if (bracketLines.length > 0) {
      const translated = bracketLines
        .map((line: string) => {
          const m = line.match(/^(\[\d+\])\s+([\s\S]*)/);
          return m ? `${m[1]} TRANSLATED: ${m[2].trim()}` : line;
        })
        .join("\n");
      return { choices: [{ message: { content: translated } }] };
    }

    // Detect N. format used by translatePptxInPlace (PPTX path)
    const numberedLines = text.match(/^\d+\.\s+.+/gm) ?? [];
    if (numberedLines.length > 0) {
      const translated = numberedLines
        .map((line: string) => {
          const m = line.match(/^(\d+)\.\s+([\s\S]*)/);
          return m ? `${m[1]}. TRANSLATED: ${m[2].trim()}` : line;
        })
        .join("\n");
      return { choices: [{ message: { content: translated } }] };
    }

    // Plain text (from translateWithLLM)
    return { choices: [{ message: { content: `TRANSLATED: ${text.slice(0, 100)}` } }] };
  };

  return { mockLLM, calls };
}

// ─────────────────────────────────────────────────────────────────────────────
// TEST SUITE 1: pptxShouldSkip direction-awareness
// Tests that CJK/Arabic source text is correctly included or excluded
// based on the target language direction.
// ─────────────────────────────────────────────────────────────────────────────

describe("PPTX skip logic — direction-aware CJK/RTL handling", () => {
  // For each source language, test translation to every target language
  for (const source of LANGUAGES) {
    for (const target of LANGUAGES) {
      if (source.name === target.name) continue; // skip same-language pairs

      it(`${source.name} → ${target.name}: text should ${source.isCJKOrRTL && target.isCJKOrRTL ? "be SKIPPED (same script family)" : "be TRANSLATED"}`, async () => {
        const { mockLLM, calls } = makeMockLLM();
        const pptxBuffer = buildMinimalPptx(source.sampleText);

        await translatePptxInPlace(pptxBuffer, target.name, mockLLM);

        const shouldTranslate = !(source.isCJKOrRTL && target.isCJKOrRTL);

        if (shouldTranslate) {
          // LLM must have been called — text was sent for translation
          expect(calls.length, `Expected LLM to be called for ${source.name}→${target.name} but it wasn't`).toBeGreaterThan(0);
        } else {
          // LLM must NOT have been called — text was correctly skipped
          expect(calls.length, `Expected LLM to be skipped for ${source.name}→${target.name} but it was called`).toBe(0);
        }
      });
    }
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// TEST SUITE 2: translatePptxInPlace — output verification
// Verifies that the translated text is actually written back into the PPTX
// ─────────────────────────────────────────────────────────────────────────────

describe("translatePptxInPlace — output written back to PPTX", () => {
  const latinToNonLatin = [
    { source: "English", target: "Simplified Chinese", text: "Market analysis overview" },
    { source: "English", target: "Japanese", text: "Business strategy summary" },
    { source: "English", target: "Arabic", text: "Quarterly revenue report" },
    { source: "English", target: "Korean", text: "Product roadmap highlights" },
  ];

  const nonLatinToLatin = [
    { source: "Simplified Chinese", target: "English", text: "市场分析和商业策略概述" },
    { source: "Traditional Chinese", target: "Spanish", text: "市場分析和商業策略概述" },
    { source: "Japanese", target: "French", text: "市場分析とビジネス戦略の概要" },
    { source: "Korean", target: "English", text: "시장 분석 및 비즈니스 전략 개요" },
    { source: "Arabic", target: "English", text: "تحليل السوق ونظرة عامة على استراتيجية الأعمال" },
  ];

  for (const { source, target, text } of [...latinToNonLatin, ...nonLatinToLatin]) {
    it(`${source} → ${target}: translated text is written into output PPTX`, async () => {
      const { mockLLM } = makeMockLLM();
      const pptxBuffer = buildMinimalPptx(text);

      const outputBuffer = await translatePptxInPlace(pptxBuffer, target, mockLLM);

      const outputText = extractTextFromPptx(outputBuffer);
      // The mock LLM prepends "TRANSLATED: " — verify it appears in the output
      expect(outputText, `Output PPTX for ${source}→${target} should contain translated text`).toContain("TRANSLATED:");
    });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// TEST SUITE 3: translateWithLLM — all language pairs (plain text)
// ─────────────────────────────────────────────────────────────────────────────

describe("translateWithLLM — all 9 target languages", () => {
  for (const target of LANGUAGES) {
    it(`translates plain text to ${target.name}`, async () => {
      const { mockLLM, calls } = makeMockLLM();
      const result = await translateWithLLM(
        "This is a test sentence for translation.",
        target.name,
        mockLLM
      );
      expect(calls.length).toBeGreaterThan(0);
      expect(result).toContain("TRANSLATED:");
    });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// TEST SUITE 4: translateBatchWithLLM — all language pairs (batch)
// ─────────────────────────────────────────────────────────────────────────────

describe("translateBatchWithLLM — all 9 target languages", () => {
  for (const target of LANGUAGES) {
    it(`batch-translates multiple strings to ${target.name}`, async () => {
      const { mockLLM, calls } = makeMockLLM();
      const inputs = [
        "Market analysis",
        "Business strategy",
        "Quarterly revenue",
        "Product roadmap",
      ];
      const result = await translateBatchWithLLM(inputs, target.name, mockLLM);
      expect(calls.length).toBeGreaterThan(0);
      expect(result).toHaveLength(inputs.length);
      // At least one item should be translated (non-empty)
      expect(result.some(r => r.length > 0)).toBe(true);
    });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// TEST SUITE 5: Edge cases
// ─────────────────────────────────────────────────────────────────────────────

describe("Translation edge cases", () => {
  it("empty string is not sent to LLM", async () => {
    const { mockLLM, calls } = makeMockLLM();
    const result = await translateWithLLM("", "English", mockLLM);
    expect(calls.length).toBe(0);
    expect(result).toBe("");
  });

  it("pure numbers are skipped in PPTX translation", async () => {
    const { mockLLM, calls } = makeMockLLM();
    const pptxBuffer = buildMinimalPptx("42");
    await translatePptxInPlace(pptxBuffer, "English", mockLLM);
    expect(calls.length).toBe(0);
  });

  it("URLs are skipped in PPTX translation", async () => {
    const { mockLLM, calls } = makeMockLLM();
    const pptxBuffer = buildMinimalPptx("https://example.com");
    await translatePptxInPlace(pptxBuffer, "English", mockLLM);
    expect(calls.length).toBe(0);
  });

  it("mixed CJK+Latin text (e.g. 'KPI 关键指标') is translated when target is Latin", async () => {
    const { mockLLM, calls } = makeMockLLM();
    const pptxBuffer = buildMinimalPptx("KPI 关键指标");
    await translatePptxInPlace(pptxBuffer, "English", mockLLM);
    // Mixed text contains CJK — should be translated to English
    expect(calls.length).toBeGreaterThan(0);
  });

  it("Arabic text is translated when target is English", async () => {
    const { mockLLM, calls } = makeMockLLM();
    const pptxBuffer = buildMinimalPptx("تحليل السوق");
    await translatePptxInPlace(pptxBuffer, "English", mockLLM);
    expect(calls.length).toBeGreaterThan(0);
  });

  it("Arabic text is skipped when target is Arabic", async () => {
    const { mockLLM, calls } = makeMockLLM();
    const pptxBuffer = buildMinimalPptx("تحليل السوق");
    await translatePptxInPlace(pptxBuffer, "Arabic", mockLLM);
    expect(calls.length).toBe(0);
  });

  it("Korean text is translated when target is Spanish", async () => {
    const { mockLLM, calls } = makeMockLLM();
    const pptxBuffer = buildMinimalPptx("시장 분석");
    await translatePptxInPlace(pptxBuffer, "Spanish", mockLLM);
    expect(calls.length).toBeGreaterThan(0);
  });

  it("Korean text is skipped when target is Japanese", async () => {
    const { mockLLM, calls } = makeMockLLM();
    const pptxBuffer = buildMinimalPptx("시장 분석");
    await translatePptxInPlace(pptxBuffer, "Japanese", mockLLM);
    expect(calls.length).toBe(0);
  });

  it("Traditional Chinese is translated when target is English", async () => {
    const { mockLLM, calls } = makeMockLLM();
    const pptxBuffer = buildMinimalPptx("市場分析和商業策略概述");
    await translatePptxInPlace(pptxBuffer, "English", mockLLM);
    expect(calls.length).toBeGreaterThan(0);
  });

  it("Traditional Chinese is skipped when target is Simplified Chinese", async () => {
    const { mockLLM, calls } = makeMockLLM();
    const pptxBuffer = buildMinimalPptx("市場分析和商業策略概述");
    await translatePptxInPlace(pptxBuffer, "Simplified Chinese", mockLLM);
    expect(calls.length).toBe(0);
  });

  it("batch translation preserves array length for all 9 languages", async () => {
    const inputs = ["Hello", "World", "Test"];
    for (const target of LANGUAGES) {
      const { mockLLM } = makeMockLLM();
      const result = await translateBatchWithLLM(inputs, target.name, mockLLM);
      expect(result).toHaveLength(inputs.length);
    }
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// TEST SUITE 6: translateDocxInPlace — bidirectional for all 9 languages
//
// DOCX uses translateBatchWithLLM which has NO CJK skip rule, so all source
// scripts should always reach the LLM regardless of direction.
// ─────────────────────────────────────────────────────────────────────────────

import AdmZip2 from "adm-zip";
import {
  translateDocxInPlace,
  translateXlsxInPlace,
} from "./docProcessor";

/** Build a minimal valid DOCX buffer with one paragraph of text */
function buildMinimalDocx(text: string): Buffer {
  const zip = new AdmZip2();

  const escaped = text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");

  zip.addFile(
    "[Content_Types].xml",
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
    )
  );

  zip.addFile(
    "_rels/.rels",
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
    )
  );

  zip.addFile(
    "word/document.xml",
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>${escaped}</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`
    )
  );

  return zip.toBuffer();
}

/** Extract all <w:t> text from a DOCX buffer */
function extractTextFromDocx(buffer: Buffer): string {
  const zip = new AdmZip2(buffer);
  const entry = zip.getEntry("word/document.xml");
  if (!entry) return "";
  const xml = entry.getData().toString("utf-8");
  const matches = xml.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g) ?? [];
  return matches.map(m => m.replace(/<[^>]+>/g, "")).join(" ").trim();
}

/** Build a minimal valid XLSX buffer with one shared string */
function buildMinimalXlsx(text: string): Buffer {
  const zip = new AdmZip2();

  const escaped = text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");

  zip.addFile(
    "[Content_Types].xml",
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`
    )
  );

  zip.addFile(
    "xl/sharedStrings.xml",
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>${escaped}</t></si>
</sst>`
    )
  );

  return zip.toBuffer();
}

/** Extract all <t> text from an XLSX sharedStrings buffer */
function extractTextFromXlsx(buffer: Buffer): string {
  const zip = new AdmZip2(buffer);
  const entry = zip.getEntry("xl/sharedStrings.xml");
  if (!entry) return "";
  const xml = entry.getData().toString("utf-8");
  const matches = xml.match(/<t[^>]*>([\s\S]*?)<\/t>/g) ?? [];
  return matches.map(m => m.replace(/<[^>]+>/g, "")).join(" ").trim();
}

describe("translateDocxInPlace — all 9 source languages × all 9 target languages", () => {
  for (const source of LANGUAGES) {
    for (const target of LANGUAGES) {
      if (source.name === target.name) continue;

      it(`DOCX: ${source.name} → ${target.name}: LLM is always called (no CJK skip in DOCX path)`, async () => {
        const { mockLLM, calls } = makeMockLLM();
        const docxBuffer = buildMinimalDocx(source.sampleText);

        await translateDocxInPlace(docxBuffer, target.name, mockLLM);

        // DOCX has no CJK skip rule — LLM must always be called
        expect(calls.length, `Expected LLM to be called for DOCX ${source.name}→${target.name}`).toBeGreaterThan(0);
      });
    }
  }
});

describe("translateDocxInPlace — translated text written back to DOCX", () => {
  const pairs = [
    { source: "Simplified Chinese", target: "English",  text: "市场分析和商业策略概述" },
    { source: "Traditional Chinese", target: "Spanish", text: "市場分析和商業策略概述" },
    { source: "Japanese",            target: "French",  text: "市場分析とビジネス戦略の概要" },
    { source: "Korean",              target: "English", text: "시장 분석 및 비즈니스 전략 개요" },
    { source: "Arabic",              target: "English", text: "تحليل السوق ونظرة عامة على استراتيجية الأعمال" },
    { source: "English",             target: "Simplified Chinese", text: "Market analysis and business strategy" },
    { source: "French",              target: "Arabic",  text: "Analyse de marché et stratégie" },
    { source: "Spanish",             target: "Japanese", text: "Análisis de mercado y estrategia" },
    { source: "Latin",               target: "Korean",  text: "Analysis mercatus et consilii" },
  ];

  for (const { source, target, text } of pairs) {
    it(`DOCX: ${source} → ${target}: translated text written into output`, async () => {
      const { mockLLM } = makeMockLLM();
      const outputBuffer = await translateDocxInPlace(buildMinimalDocx(text), target, mockLLM);
      const outputText = extractTextFromDocx(outputBuffer);
      expect(outputText).toContain("TRANSLATED:");
    });
  }
});

describe("translateXlsxInPlace — all 9 source languages × all 9 target languages", () => {
  for (const source of LANGUAGES) {
    for (const target of LANGUAGES) {
      if (source.name === target.name) continue;

      it(`XLSX: ${source.name} → ${target.name}: LLM is always called (no CJK skip in XLSX path)`, async () => {
        const { mockLLM, calls } = makeMockLLM();
        const xlsxBuffer = buildMinimalXlsx(source.sampleText);

        await translateXlsxInPlace(xlsxBuffer, target.name, mockLLM);

        expect(calls.length, `Expected LLM to be called for XLSX ${source.name}→${target.name}`).toBeGreaterThan(0);
      });
    }
  }
});

describe("translateXlsxInPlace — translated text written back to XLSX", () => {
  const pairs = [
    { source: "Simplified Chinese", target: "English",  text: "市场分析" },
    { source: "Arabic",             target: "French",   text: "تحليل السوق" },
    { source: "Korean",             target: "Spanish",  text: "시장 분석" },
    { source: "English",            target: "Japanese", text: "Market analysis" },
  ];

  for (const { source, target, text } of pairs) {
    it(`XLSX: ${source} → ${target}: translated text written into output`, async () => {
      const { mockLLM } = makeMockLLM();
      const outputBuffer = await translateXlsxInPlace(buildMinimalXlsx(text), target, mockLLM);
      const outputText = extractTextFromXlsx(outputBuffer);
      expect(outputText).toContain("TRANSLATED:");
    });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// TEST SUITE 7: XML hallucination regression tests
//
// Regression tests for the bug where the LLM returns XML tags as literal text
// in its translation response, causing raw XML to appear on slides.
// ─────────────────────────────────────────────────────────────────────────────

import { translatePptxInPlace } from "./docProcessor";

describe("translatePptxInPlace — XML hallucination sanitization", () => {
  it("strips XML tags from LLM response before writing to slide", async () => {
    // Simulate an LLM that returns XML tags in its translation response
    const xmlHallucinatingLLM = async (params: { messages: any[] }) => {
      const userContent = params.messages[params.messages.length - 1]?.content ?? "";
      const text = typeof userContent === "string" ? userContent : "";
      // Return response with XML tags hallucinated (like the real bug)
      const numberedLines = text.match(/^\d+\.\s+.+/gm) ?? [];
      if (numberedLines.length > 0) {
        const translated = numberedLines
          .map((line: string) => {
            const m = line.match(/^(\d+)\.\s+([\s\S]*)/);
            // Simulate LLM hallucinating XML tags in the response
            return m ? `${m[1]}. <a:rPr lang="en-GB" sz="1400"/><a:t>Translated: ${m[2].trim()}</a:t>` : line;
          })
          .join("\n");
        return { choices: [{ message: { content: translated } }] };
      }
      return { choices: [{ message: { content: `<a:rPr/><a:t>Translated text</a:t>` } }] };
    };

    const pptxBuffer = buildMinimalPptx("市场分析和商业策略");
    const output = await translatePptxInPlace(pptxBuffer, "English", xmlHallucinatingLLM);
    const outputText = extractTextFromPptx(output);

    // The output should NOT contain any XML tags
    expect(outputText).not.toMatch(/<[a-z:]+/i);
    // The output should contain the translated content (without XML tags)
    expect(outputText).toContain("Translated:");
  });

  it("handles HTML-entity-escaped content in <a:t> nodes correctly", async () => {
    // Build a PPTX where the <a:t> contains HTML-escaped text (e.g. &lt;a:tab&gt;)
    const AdmZip = (await import("adm-zip")).default;
    const zip = new AdmZip();

    zip.addFile("[Content_Types].xml", Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`
    ));
    zip.addFile("_rels/.rels", Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`
    ));
    zip.addFile("ppt/presentation.xml", Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst/>
  <p:sldSz cx="9144000" cy="6858000"/>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
</p:presentation>`
    ));
    zip.addFile("ppt/_rels/presentation.xml.rels", Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`
    ));
    // Slide with HTML-escaped content in <a:t> (simulates the &lt;a:tab&gt; bug)
    zip.addFile("ppt/slides/slide1.xml", Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="1400" b="1"/>
              <a:t>Next Steps &amp; Goals</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`
    ));

    const pptxBuffer = zip.toBuffer();
    const { mockLLM } = makeMockLLM();
    const output = await translatePptxInPlace(pptxBuffer, "Simplified Chinese", mockLLM);
    const outputText = extractTextFromPptx(output);

    // The & entity should be unescaped for LLM lookup, and the output should contain translated text
    expect(outputText).toContain("TRANSLATED:");
    // No raw XML tags in output
    expect(outputText).not.toMatch(/<[a-z:]+/i);
  });
});
