import { invokeLLM } from "./_core/llm";
import { storagePut } from "./storage";
import { generatePreview } from "./watermark";
import {
  createJob,
  getJobById,
  updateJob,
  getAnonUserId,
  createJobStep,
  updateJobStep,
  appendJobLog,
  isJobCancelled,
} from "./db";
import {
  extractTextFromBuffer,
  translateWithLLM,
  translatePptxInPlace,
  translateDocxInPlace,
  translateXlsxInPlace,
  convertDocument,
  buildTranslatedDocument,
  convertPdfToDocxWithPdf2Docx,
  makeModelInvoker,
  getFormatFromFilename,
  getMimeType,
  isImageFormat,
  COST_PER_PAGE_EXTRACTION,
  COST_PER_PAGE_TRANSLATION,
  MAX_FILE_SIZE_BYTES,
  JOB_TIMEOUT_MS,
  type SupportedFormat,
} from "./docProcessor";

function nanoid(len = 12) {
  return Math.random().toString(36).slice(2, 2 + len);
}

export async function convertImageToPdf(imageBuffer: Buffer, mimeType: string): Promise<Buffer> {
  const sharp = (await import("sharp")).default;
  const { PDFDocument } = await import("pdf-lib");

  const meta = await sharp(imageBuffer).metadata();
  const width = meta.width ?? 800;
  const height = meta.height ?? 600;

  const pngBuf = await sharp(imageBuffer).png().toBuffer();
  const pngBytes = new Uint8Array(pngBuf.buffer, pngBuf.byteOffset, pngBuf.byteLength);

  const pdfDoc = await PDFDocument.create();
  const pngImage = await pdfDoc.embedPng(pngBytes);
  const page = pdfDoc.addPage([width, height]);
  page.drawImage(pngImage, { x: 0, y: 0, width, height });

  const pdfBytes = await pdfDoc.save();
  const pdfBuf = Buffer.allocUnsafe(pdfBytes.byteLength);
  pdfBuf.set(pdfBytes);
  return pdfBuf;
}

export async function convertMultipleImagesToPdf(imageBuffers: Buffer[], mimeTypes: string[]): Promise<Buffer> {
  const sharp = (await import("sharp")).default;
  const { PDFDocument } = await import("pdf-lib");

  const pdfDoc = await PDFDocument.create();

  for (let i = 0; i < imageBuffers.length; i++) {
    const imgBuf = imageBuffers[i];
    const meta = await sharp(imgBuf).metadata();
    const width = meta.width ?? 800;
    const height = meta.height ?? 600;

    const pngBuf = await sharp(imgBuf).png().toBuffer();
    const pngBytes = new Uint8Array(pngBuf.buffer, pngBuf.byteOffset, pngBuf.byteLength);

    const pngImage = await pdfDoc.embedPng(pngBytes);
    const page = pdfDoc.addPage([width, height]);
    page.drawImage(pngImage, { x: 0, y: 0, width, height });
  }

  const pdfBytes = await pdfDoc.save();
  const pdfBuf = Buffer.allocUnsafe(pdfBytes.byteLength);
  pdfBuf.set(pdfBytes);
  return pdfBuf;
}

export interface StartDocumentJobInput {
  filename: string;
  mimeType: string;
  buffer: Buffer;
  outputFormat?: string;
  targetLanguage?: string;
  targetLanguageName?: string;
  modelId?: string;
}

/** Prepare buffer (images→PDF), create DB job, and start background processing. */
export async function startDocumentJob(input: StartDocumentJobInput): Promise<{ jobId: number }> {
  if (input.buffer.length > MAX_FILE_SIZE_BYTES) {
    throw new Error(
      `File too large: ${(input.buffer.length / 1024 / 1024).toFixed(1)} MB. Maximum allowed is ${MAX_FILE_SIZE_BYTES / 1024 / 1024} MB.`,
    );
  }

  let processBuffer = input.buffer;
  let processFilename = input.filename;
  let processMimeType = input.mimeType;
  let isImageInput = false;

  const rawExt = input.filename.split(".").pop()?.toLowerCase() ?? "";
  if (isImageFormat(rawExt)) {
    processBuffer = await convertImageToPdf(input.buffer, input.mimeType);
    processFilename = input.filename.replace(/\.[^.]+$/, ".pdf");
    processMimeType = "application/pdf";
    isImageInput = true;
  }

  const format = getFormatFromFilename(processFilename);
  const fileKey = `uploads/anon/${nanoid()}_${processFilename}`;
  const anonUserId = await getAnonUserId();

  const jobId = await createJob({
    userId: anonUserId,
    originalFileName: input.filename,
    originalFileKey: fileKey,
    originalFileUrl: "",
    originalFormat: format,
    outputFormat: input.outputFormat ?? format,
    targetLanguage: input.targetLanguage,
    targetLanguageName: input.targetLanguageName,
    status: "pending",
  });

  await Promise.all([
    createJobStep(jobId, "upload"),
    createJobStep(jobId, "extract"),
    ...(input.targetLanguage ? [createJobStep(jobId, "translate")] : []),
    ...(input.outputFormat && input.outputFormat !== format ? [createJobStep(jobId, "convert")] : []),
  ]);

  await Promise.all([
    updateJobStep(jobId, "upload", { status: "done", completedAt: new Date() }),
    appendJobLog(jobId, "success", `File uploaded: ${input.filename} (${(processBuffer.length / 1024).toFixed(0)} KB)`),
  ]);

  storagePut(fileKey, processBuffer, processMimeType)
    .then(({ url }) => updateJob(jobId, { originalFileUrl: url }))
    .catch(err => console.error("[GCS] Background original-file upload failed:", err));

  scheduleJobProcessing({
    jobId,
    processBuffer,
    format,
    outputFormat: input.outputFormat as SupportedFormat | undefined,
    targetLanguage: input.targetLanguage,
    modelId: input.modelId,
    isImageInput,
  });

  return { jobId };
}

export function scheduleJobProcessing(opts: {
  jobId: number;
  processBuffer: Buffer;
  format: SupportedFormat;
  outputFormat?: SupportedFormat;
  targetLanguage?: string;
  originalFileUrl?: string;
  modelId?: string;
  isImageInput?: boolean;
}): void {
  let timeoutHandle: ReturnType<typeof setTimeout>;
  const timeoutPromise = new Promise<never>((_, reject) => {
    timeoutHandle = setTimeout(() => reject(new Error("__TIMEOUT__")), JOB_TIMEOUT_MS);
  });

  Promise.race([
    processJobAsync(
      opts.jobId,
      opts.processBuffer,
      opts.format,
      opts.outputFormat,
      opts.targetLanguage,
      opts.originalFileUrl ?? "",
      opts.modelId,
      opts.isImageInput,
    ).then(() => clearTimeout(timeoutHandle)),
    timeoutPromise,
  ]).catch(async err => {
    const msg: string = err.message ?? "Unknown error";
    if (msg === "__TIMEOUT__") {
      const currentJob = await getJobById(opts.jobId).catch(() => null);
      if (currentJob && ["done", "cancelled", "error"].includes(currentJob.status)) return;
      console.warn(`[Job ${opts.jobId}] Timed out — pausing for user decision`);
      await appendJobLog(opts.jobId, "warning", `Job is taking longer than ${JOB_TIMEOUT_MS / 60000} minutes. Paused — waiting for your decision.`).catch(() => {});
      await updateJob(opts.jobId, { status: "paused", errorMessage: `Paused after ${JOB_TIMEOUT_MS / 60000} minutes` }).catch(() => {});
    } else {
      console.error(`[Job ${opts.jobId}] Fatal error:`, err);
      await appendJobLog(opts.jobId, "error", msg).catch(() => {});
      await updateJob(opts.jobId, { status: "error", errorMessage: msg }).catch(() => {});
    }
  });
}

export async function waitForJobDone(
  jobId: number,
  opts?: { pollMs?: number; maxMs?: number },
): Promise<Awaited<ReturnType<typeof getJobById>>> {
  const pollMs = opts?.pollMs ?? 2500;
  const maxMs = opts?.maxMs ?? JOB_TIMEOUT_MS + 60_000;
  const deadline = Date.now() + maxMs;

  while (Date.now() < deadline) {
    const job = await getJobById(jobId);
    if (!job) throw new Error("Job not found");
    if (["done", "error", "cancelled", "paused"].includes(job.status)) return job;
    await new Promise(r => setTimeout(r, pollMs));
  }
  throw new Error("Processing timed out. Please try again later.");
}

export async function processJobAsync(
  jobId: number,
  buffer: Buffer,
  inputFormat: SupportedFormat,
  outputFormat: SupportedFormat | undefined,
  targetLanguage: string | undefined,
  originalFileUrl: string,
  modelId?: string,
  isImageInput?: boolean,
) {
  const llm = makeModelInvoker(invokeLLM, modelId);
  const startTime = Date.now();

  const log = (msg: string, level: "info" | "progress" | "success" | "warning" | "error" = "info") =>
    appendJobLog(jobId, level, msg).catch(() => {});

  const logTiming = (stepName: string, startMs: number) => {
    const duration = Date.now() - startMs;
    console.log(`[Timing] ${stepName}: ${duration}ms`);
    return duration;
  };

  const checkCancelled = async () => {
    const cancelled = await isJobCancelled(jobId);
    if (cancelled) throw new Error("Job cancelled by user");
  };

  try {
    const finalFormat = outputFormat ?? inputFormat;

    await updateJob(jobId, { status: "extracting" });
    await updateJobStep(jobId, "extract", { status: "running", startedAt: new Date() });

    let extractedText = "";
    let pageCount = 1;

    if (isImageInput) {
      await log("Image converted to PDF — no text extraction needed.", "info");
      await updateJob(jobId, { extractedText: "", pageCount: 1, charCount: 0, estimatedCost: 0 });
      await updateJobStep(jobId, "extract", { status: "done", completedAt: new Date(), message: "Image converted to PDF (1 page)" });
      await log("Image embedded as PDF page ✓", "success");

      const job = await getJobById(jobId);
      const outKey = `outputs/${job?.userId}/${nanoid()}_converted.pdf`;
      const { url: outputUrl } = await storagePut(outKey, buffer, getMimeType("pdf"));

      await log("Generating preview...", "info");
      let imgPreviewUrl: string | undefined;
      let imgPreviewKey: string | undefined;
      let imgPreviewPageCount = 0;
      try {
        const { previewBuffer, previewPages } = await generatePreview(buffer, "pdf");
        const pKey = `previews/${job?.userId}/${nanoid()}_preview.pdf`;
        const { url: pUrl } = await storagePut(pKey, previewBuffer, "application/pdf");
        imgPreviewUrl = pUrl;
        imgPreviewKey = pKey;
        imgPreviewPageCount = previewPages;
      } catch (e) {
        console.error("[Preview] Preview generation failed (non-fatal):", e);
      }

      await updateJob(jobId, {
        status: "done",
        outputFileKey: outKey,
        outputFileUrl: outputUrl,
        outputFormat: "pdf",
        previewFileKey: imgPreviewKey,
        previewFileUrl: imgPreviewUrl,
        previewPageCount: imgPreviewPageCount,
        conversionCostUsd: 0,
        downloadPriceUsd: 0,
        paid: true,
      });
      await log("Done! Your file is ready to download.", "success");
      logTiming("Total conversion time", startTime);
      return;
    }

    await log(`Extracting text from ${inputFormat.toUpperCase()} document...`, "info");
    await log(`Parsing document structure and identifying text blocks...`, "info");

    try {
      await log(`Extracting text locally (fast)...`, "info");
      await log(`Scanning ${inputFormat.toUpperCase()} for text content, tables, and metadata...`, "info");
      const extractStart = Date.now();
      const result = await extractTextFromBuffer(buffer, inputFormat, `doc.${inputFormat}`);
      logTiming("Text extraction", extractStart);
      extractedText = result.text;
      pageCount = result.pageCount;
    } catch {
      await log("Extraction failed — falling back to direct text extraction", "warning");
      const extractStart = Date.now();
      const result = await extractTextFromBuffer(buffer, inputFormat, `doc.${inputFormat}`);
      logTiming("Text extraction (fallback)", extractStart);
      extractedText = result.text;
      pageCount = result.pageCount;
    }

    const charCount = extractedText.length;
    const estimatedCost =
      pageCount * COST_PER_PAGE_EXTRACTION + (targetLanguage ? pageCount * COST_PER_PAGE_TRANSLATION : 0);

    await updateJob(jobId, { extractedText, pageCount, charCount, estimatedCost });
    await updateJobStep(jobId, "extract", {
      status: "done",
      completedAt: new Date(),
      message: `Extracted ${pageCount} pages, ${charCount} characters`,
    });
    await log(`Extracted ${pageCount} page${pageCount !== 1 ? "s" : ""} · ${charCount.toLocaleString()} characters`, "success");

    await checkCancelled();

    let processedText = extractedText;
    let inPlaceBuffer: Buffer | null = null;

    if (targetLanguage) {
      await updateJob(jobId, { status: "translating" });
      await updateJobStep(jobId, "translate", { status: "running", startedAt: new Date() });

      try {
        const job = await getJobById(jobId);
        const langName = job?.targetLanguageName ?? targetLanguage;
        const translateStart = Date.now();

        if (inputFormat === "pdf" && finalFormat === "docx") {
          await log(`Converting PDF layout to DOCX (preserving text, images, columns)...`, "info");
          const convertStart = Date.now();
          const convertedDocx = await convertPdfToDocxWithPdf2Docx(buffer, async (msg: string) => {
            await log(msg, "info");
          });
          logTiming("PDF to DOCX conversion", convertStart);
          await log(`PDF converted to DOCX — now translating to ${langName}...`, "info");
          const translateDocxStart = Date.now();
          inPlaceBuffer = await translateDocxInPlace(convertedDocx, langName, llm);
          logTiming("DOCX translation", translateDocxStart);
          logTiming("PDF→DOCX translation total", translateStart);
          await updateJobStep(jobId, "translate", {
            status: "done",
            completedAt: new Date(),
            message: `Converted PDF→DOCX and translated to ${langName}`,
          });
          await log(`PDF translated to ${langName} DOCX — layout and images preserved`, "success");
        } else if (inputFormat === "pptx" && (finalFormat === "pptx" || !outputFormat)) {
          await log(`Translating PPTX slides to ${langName} (in-place — preserving layout)...`, "info");
          inPlaceBuffer = await translatePptxInPlace(buffer, langName, llm, async (msg: string) => {
            await log(msg, "info");
          });
          logTiming("PPTX translation", translateStart);
          await updateJobStep(jobId, "translate", {
            status: "done",
            completedAt: new Date(),
            message: `Translated PPTX in-place to ${langName} (layout preserved)`,
          });
          await log(`PPTX translated to ${langName} — all images and formatting preserved`, "success");
        } else if (inputFormat === "docx" && (finalFormat === "docx" || !outputFormat)) {
          await log(`Translating DOCX to ${langName} (in-place — preserving layout)...`, "info");
          inPlaceBuffer = await translateDocxInPlace(buffer, langName, llm);
          logTiming("DOCX translation", translateStart);
          await updateJobStep(jobId, "translate", {
            status: "done",
            completedAt: new Date(),
            message: `Translated DOCX in-place to ${langName} (layout preserved)`,
          });
          await log(`DOCX translated to ${langName} — all images and formatting preserved`, "success");
        } else if (inputFormat === "xlsx" && (finalFormat === "xlsx" || !outputFormat)) {
          await log(`Translating XLSX to ${langName} (in-place — preserving layout)...`, "info");
          inPlaceBuffer = await translateXlsxInPlace(buffer, langName, llm);
          logTiming("XLSX translation", translateStart);
          await updateJobStep(jobId, "translate", {
            status: "done",
            completedAt: new Date(),
            message: `Translated XLSX in-place to ${langName} (layout preserved)`,
          });
          await log(`XLSX translated to ${langName} — all formatting preserved`, "success");
        } else {
          await log(`Translating document content to ${langName}...`, "info");
          processedText = await translateWithLLM(extractedText, langName, llm);
          logTiming("Text translation", translateStart);
          await updateJobStep(jobId, "translate", {
            status: "done",
            completedAt: new Date(),
            message: `Translated to ${langName} via ${modelId ?? "GPT-4o-mini"}`,
          });
          await log(`Translation to ${langName} complete`, "success");
        }
      } catch (e: any) {
        await updateJobStep(jobId, "translate", { status: "error", completedAt: new Date(), message: e.message });
        await log(`Translation failed: ${e.message}`, "error");
      }
    }

    await checkCancelled();

    let outputBuffer: Buffer;

    if (outputFormat && outputFormat !== inputFormat) {
      await updateJob(jobId, { status: "converting" });
      await updateJobStep(jobId, "convert", { status: "running", startedAt: new Date() });
      await log(`Converting to ${finalFormat.toUpperCase()}...`, "info");
    }

    if (inPlaceBuffer) {
      outputBuffer = inPlaceBuffer;
    } else if (targetLanguage) {
      await log(`Building ${finalFormat.toUpperCase()} output document...`, "info");
      outputBuffer = await buildTranslatedDocument(processedText, finalFormat, `translated.${finalFormat}`);
    } else if (outputFormat && outputFormat !== inputFormat) {
      outputBuffer = await convertDocument(buffer, inputFormat, finalFormat, `doc.${inputFormat}`);
    } else {
      outputBuffer = buffer;
    }

    if (outputFormat && outputFormat !== inputFormat) {
      await updateJobStep(jobId, "convert", {
        status: "done",
        completedAt: new Date(),
        message: `Converted to ${finalFormat.toUpperCase()}`,
      });
      await log(`Converted to ${finalFormat.toUpperCase()} successfully`, "success");
    }

    await log("Finalizing output file...", "info");
    const job = await getJobById(jobId);
    const outKey = `outputs/${job?.userId}/${nanoid()}_translated.${finalFormat}`;
    const { url: outputUrl } = await storagePut(outKey, outputBuffer, getMimeType(finalFormat));

    const finalJob = await getJobById(jobId);
    const finalCostUsd = finalJob?.estimatedCost ?? 0;
    await updateJob(jobId, {
      status: "done",
      outputFileKey: outKey,
      outputFileUrl: outputUrl,
      outputFormat: finalFormat,
      conversionCostUsd: +finalCostUsd.toFixed(4),
      downloadPriceUsd: 0,
      paid: true,
    });
    await log("Done! Your document is ready to download.", "success");
    logTiming("Total conversion time", startTime);

    (async () => {
      const PREVIEW_TIMEOUT = 60_000;
      const timeout = setTimeout(() => {}, PREVIEW_TIMEOUT);
      try {
        const { previewBuffer, previewPages, previewFormat, previewMimeType } = await generatePreview(
          outputBuffer,
          finalFormat,
        );
        const pKey = `previews/${job?.userId}/${nanoid()}_preview.${previewFormat}`;
        const { url: pUrl } = await storagePut(pKey, previewBuffer, previewMimeType);
        await updateJob(jobId, {
          previewFileKey: pKey,
          previewFileUrl: pUrl,
          previewPageCount: previewPages,
        });
      } catch (e: any) {
        console.error("[Preview] Background preview failed:", e?.message ?? e);
      } finally {
        clearTimeout(timeout);
      }
    })().catch(() => {});
  } catch (err: any) {
    console.error(`[Job ${jobId}] Processing failed:`, err);
    const msg = err.message ?? "Unknown error";
    if (msg === "Job cancelled by user") {
      await updateJob(jobId, { status: "cancelled" });
      await log("Processing stopped.", "warning");
    } else {
      await updateJob(jobId, { status: "error", errorMessage: msg });
      await log(`Error: ${msg}`, "error");
    }
  }
}
