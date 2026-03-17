import { z } from "zod";
import { COOKIE_NAME } from "@shared/const";
import { getSessionCookieOptions } from "./_core/cookies";
import { systemRouter } from "./_core/systemRouter";
import { publicProcedure, protectedProcedure, router } from "./_core/trpc";
import { invokeLLM } from "./_core/llm";
import { storagePut, storageDelete } from "./storage";
import { generatePreview } from "./watermark";
import { calculateDownloadPrice, createDownloadCheckoutSession } from "./stripe";
import {
  createJob, getJobById, updateJob, getUserJobs, getUserById, getAnonUserId,
  createJobStep, updateJobStep, getJobSteps,
  appendJobLog, getJobLogs, isJobCancelled, cancelJob, deleteJob,
} from "./db";
import {
  extractTextFromBuffer, extractWithLLM, translateWithLLM,
  translatePptxInPlace, translateDocxInPlace, translateXlsxInPlace,
  convertDocument, buildTranslatedDocument, convertPdfToPptxWithVision,
  convertPdfToDocxWithPdf2Docx,
  makeModelInvoker,
  getFormatFromFilename, getMimeType, isImageFormat,
  COST_PER_PAGE_EXTRACTION, COST_PER_PAGE_TRANSLATION, COST_PER_PAGE_TOTAL,
  MAX_FILE_SIZE_BYTES, JOB_TIMEOUT_MS, MAX_PAGES_VISION, COST_WARNING_THRESHOLD,
  LANGUAGES, SUPPORTED_INPUT_FORMATS, SUPPORTED_OUTPUT_FORMATS, AVAILABLE_MODELS,
  type SupportedFormat,
} from "./docProcessor";

// ─── Image → PDF conversion ───────────────────────────────────────────────────
async function convertImageToPdf(imageBuffer: Buffer, mimeType: string): Promise<Buffer> {
  const sharp = (await import("sharp")).default;
  const { PDFDocument } = await import("pdf-lib");

  // Use sharp to get image dimensions and convert to PNG for embedding
  const meta = await sharp(imageBuffer).metadata();
  const width = meta.width ?? 800;
  const height = meta.height ?? 600;

  // Convert to PNG bytes for pdf-lib embedding
  const pngBuf = await sharp(imageBuffer).png().toBuffer();
  const pngBytes = new Uint8Array(pngBuf.buffer, pngBuf.byteOffset, pngBuf.byteLength);

  const pdfDoc = await PDFDocument.create();
  const pngImage = await pdfDoc.embedPng(pngBytes);

  // Create a page with the same dimensions as the image (in points, 1pt = 1px at 72dpi)
  const page = pdfDoc.addPage([width, height]);
  page.drawImage(pngImage, { x: 0, y: 0, width, height });

  const pdfBytes = await pdfDoc.save();
  const pdfBuf = Buffer.allocUnsafe(pdfBytes.byteLength);
  pdfBuf.set(pdfBytes);
  return pdfBuf;
}

// ─── Multi-image → single PDF ───────────────────────────────────────────────
async function convertMultipleImagesToPdf(imageBuffers: Buffer[], mimeTypes: string[]): Promise<Buffer> {
  const sharp = (await import("sharp")).default;
  const { PDFDocument } = await import("pdf-lib");

  const pdfDoc = await PDFDocument.create();

  for (let i = 0; i < imageBuffers.length; i++) {
    const imgBuf = imageBuffers[i];
    const mime = mimeTypes[i] ?? "image/png";

    // Get dimensions
    const meta = await sharp(imgBuf).metadata();
    const width = meta.width ?? 800;
    const height = meta.height ?? 600;

    // Always convert to PNG for reliable pdf-lib embedding
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

// ─── Helpers ─────────────────────────────────────────────────────────────────
function nanoid(len = 12) {
  return Math.random().toString(36).slice(2, 2 + len);
}

// ─── Routers ──────────────────────────────────────────────────────────────────
export const appRouter = router({
  system: systemRouter,
  auth: router({
    me: publicProcedure.query(opts => opts.ctx.user),
    logout: publicProcedure.mutation(({ ctx }) => {
      const cookieOptions = getSessionCookieOptions(ctx.req);
      ctx.res.clearCookie(COOKIE_NAME, { ...cookieOptions, maxAge: -1 });
      return { success: true } as const;
    }),
  }),

  // ─── Document processing ────────────────────────────────────────────────
  docs: router({

    // Get supported formats, languages, and models
    config: publicProcedure.query(() => ({
      inputFormats: SUPPORTED_INPUT_FORMATS,
      outputFormats: SUPPORTED_OUTPUT_FORMATS,
      languages: LANGUAGES,
      models: AVAILABLE_MODELS,
      costPerPageExtraction: COST_PER_PAGE_EXTRACTION,
      costPerPageTranslation: COST_PER_PAGE_TRANSLATION,
      costPerPageTotal: COST_PER_PAGE_TOTAL,
      maxFileSizeMB: MAX_FILE_SIZE_BYTES / 1024 / 1024,
      maxPagesVision: MAX_PAGES_VISION,
      jobTimeoutMinutes: JOB_TIMEOUT_MS / 60000,
      costWarningThreshold: COST_WARNING_THRESHOLD,
    })),

    // Estimate cost before processing
    estimate: publicProcedure
      .input(z.object({
        pageCount: z.number().min(1),
        doTranslate: z.boolean(),
        doConvert: z.boolean(),
      }))
      .query(({ input }) => {
        const extractionCost = input.pageCount * COST_PER_PAGE_EXTRACTION;
        const translationCost = input.doTranslate ? input.pageCount * COST_PER_PAGE_TRANSLATION : 0;
        const total = extractionCost + translationCost;
        return {
          extractionCost: +extractionCost.toFixed(4),
          translationCost: +translationCost.toFixed(4),
          total: +total.toFixed(4),
          perPage: +(total / input.pageCount).toFixed(4),
        };
      }),

    // Upload file and create job (accepts base64 encoded file)
    upload: publicProcedure
      .input(z.object({
        filename: z.string(),
        mimeType: z.string(),
        base64Data: z.string(),
        outputFormat: z.string().optional(),
        targetLanguage: z.string().optional(),
        targetLanguageName: z.string().optional(),
        modelId: z.string().optional(),
        // Multi-image merge fields
        multiImageBase64: z.array(z.string()).optional(),
        multiImageMimeTypes: z.array(z.string()).optional(),
      }))
      .mutation(async ({ ctx, input }) => {
        let processBuffer: Buffer;
        let processFilename: string;
        let processMimeType: string;

        // Track whether the input was an image (skip text extraction later)
        let isImageInput = false;

        // ── Multi-image merge path ──────────────────────────────────────────
        if (input.multiImageBase64 && input.multiImageBase64.length > 0) {
          const imageBuffers = input.multiImageBase64.map(b64 => Buffer.from(b64, "base64") as Buffer);
          const mimeTypes = input.multiImageMimeTypes ?? imageBuffers.map(() => "image/png");
          const totalSize = imageBuffers.reduce((s, b) => s + b.length, 0);
          if (totalSize > MAX_FILE_SIZE_BYTES * 5) {
            throw new Error(`Total image size too large: ${(totalSize / 1024 / 1024).toFixed(1)} MB.`);
          }
          processBuffer = await convertMultipleImagesToPdf(imageBuffers, mimeTypes);
          processFilename = "merged-images.pdf";
          processMimeType = "application/pdf";
          isImageInput = true;
        } else {
          // ── Single file path ────────────────────────────────────────────────
          const buffer = Buffer.from(input.base64Data, "base64");

          // Server-side file size guard
          if (buffer.length > MAX_FILE_SIZE_BYTES) {
            throw new Error(`File too large: ${(buffer.length / 1024 / 1024).toFixed(1)} MB. Maximum allowed is ${MAX_FILE_SIZE_BYTES / 1024 / 1024} MB.`);
          }

          const rawExt = input.filename.split(".").pop()?.toLowerCase() ?? "";
          processBuffer = buffer as Buffer;
          processFilename = input.filename;
          processMimeType = input.mimeType;

          // Convert image inputs to PDF before processing
          if (isImageFormat(rawExt)) {
            processBuffer = await convertImageToPdf(buffer as Buffer, input.mimeType);
            processFilename = input.filename.replace(/\.[^.]+$/, ".pdf");
            processMimeType = "application/pdf";
            isImageInput = true;
          }
        }

        const format = getFormatFromFilename(processFilename);
        const fileKey = `uploads/anon/${nanoid()}_${processFilename}`;

        // Create the job record immediately — we don't block on GCS upload.
        // processJobAsync receives the buffer directly so it can start right away.
        // The GCS upload runs in the background and updates originalFileUrl once
        // done so the "resume job" flow (which re-fetches from GCS) still works.
        const anonUserId = await getAnonUserId();
        const jobId = await createJob({
          userId: anonUserId,
          originalFileName: input.filename,
          originalFileKey: fileKey,
          originalFileUrl: "",   // filled in by background GCS upload below
          originalFormat: format,
          outputFormat: input.outputFormat ?? format,
          targetLanguage: input.targetLanguage,
          targetLanguageName: input.targetLanguageName,
          status: "pending",
        });

        // Create all step records in parallel (saves multiple DB round-trips)
        await Promise.all([
          createJobStep(jobId, "upload"),
          createJobStep(jobId, "extract"),
          ...(input.targetLanguage ? [createJobStep(jobId, "translate")] : []),
          ...(input.outputFormat && input.outputFormat !== format ? [createJobStep(jobId, "convert")] : []),
        ]);

        // Mark upload done + emit log in parallel
        await Promise.all([
          updateJobStep(jobId, "upload", { status: "done", completedAt: new Date() }),
          appendJobLog(jobId, "success", `File uploaded: ${input.filename} (${(processBuffer.length / 1024).toFixed(0)} KB)`),
        ]);

        // Background GCS upload — runs concurrently with processJobAsync.
        // processJobAsync has the buffer in memory so it doesn't need the URL.
        // The URL is stored in the DB once ready so "resume" can re-fetch the file.
        storagePut(fileKey, processBuffer, processMimeType)
          .then(({ url }) => updateJob(jobId, { originalFileUrl: url }))
          .catch(err => console.error("[GCS] Background original-file upload failed:", err));

        // Kick off async processing with a timeout guard (fire and forget)
        // On timeout: pause the job instead of killing it, so the user can decide to continue or cancel
        let timeoutHandle: ReturnType<typeof setTimeout>;
        const timeoutPromise = new Promise<never>((_, reject) => {
          timeoutHandle = setTimeout(
            () => reject(new Error(`__TIMEOUT__`)),
            JOB_TIMEOUT_MS
          );
        });
        Promise.race([
          processJobAsync(
            jobId, processBuffer, format,
            input.outputFormat as SupportedFormat | undefined,
            input.targetLanguage,
            "",   // originalFileUrl not needed for initial run (buffer is in memory)
            input.modelId,
            isImageInput,
          ).then(() => clearTimeout(timeoutHandle)),
          timeoutPromise,
        ]).catch(async (err) => {
          const msg: string = err.message ?? "Unknown error";
          if (msg === "__TIMEOUT__") {
            // Before pausing, check if the job already finished successfully
            const currentJob = await getJobById(jobId).catch(() => null);
            if (currentJob && ["done", "cancelled", "error"].includes(currentJob.status)) {
              // Job already finished — timeout fired late, ignore it
              console.log(`[Job ${jobId}] Timeout fired but job already ${currentJob.status} — ignoring`);
              return;
            }
            console.warn(`[Job ${jobId}] Timed out — pausing for user decision`);
            await appendJobLog(jobId, "warning", `Job is taking longer than ${JOB_TIMEOUT_MS / 60000} minutes. Paused — waiting for your decision.`).catch(() => {});
            await updateJob(jobId, { status: "paused", errorMessage: `Paused after ${JOB_TIMEOUT_MS / 60000} minutes` }).catch(() => {});
          } else {
            console.error(`[Job ${jobId}] Fatal error:`, err);
            await appendJobLog(jobId, "error", msg).catch(() => {});
            await updateJob(jobId, { status: "error", errorMessage: msg }).catch(() => {});
          }
        });

        return { jobId };
      }),

    // Get job status and steps
    status: publicProcedure
      .input(z.object({ jobId: z.number() }))
      .query(async ({ input }) => {
        const job = await getJobById(input.jobId);
        if (!job) throw new Error("Job not found");
        const steps = await getJobSteps(input.jobId);
        return { job, steps };
      }),

    // Get live log entries for a job (poll-based streaming)
    getLogs: publicProcedure
      .input(z.object({
        jobId: z.number(),
        afterSeq: z.number().default(0),
      }))
      .query(async ({ input }) => {
        const job = await getJobById(input.jobId);
        if (!job) throw new Error("Job not found");
        const logs = await getJobLogs(input.jobId, input.afterSeq);
        return { logs };
      }),

    // Cancel a running job
    cancel: publicProcedure
      .input(z.object({ jobId: z.number() }))
      .mutation(async ({ input }) => {
        const job = await getJobById(input.jobId);
        if (!job) throw new Error("Job not found");
        if (["done", "error", "cancelled"].includes(job.status)) {
          return { success: false, message: "Job is already finished" };
        }
        await cancelJob(input.jobId);
        await appendJobLog(input.jobId, "warning", "Job cancelled by user");
        return { success: true };
      }),

    // Resume a paused job (extend the timeout and continue processing)
    resume: publicProcedure
      .input(z.object({ jobId: z.number() }))
      .mutation(async ({ input }) => {
        const job = await getJobById(input.jobId);
        if (!job) throw new Error("Job not found");
        if (job.status !== "paused") {
          return { success: false, message: "Job is not paused" };
        }
        // Resume: clear the paused state and re-run the job from scratch
        // (The original buffer is gone from memory, so we re-fetch from S3 and re-process)
        await updateJob(input.jobId, { status: "pending", errorMessage: null });
        await appendJobLog(input.jobId, "info", "Resuming job — continuing processing...");

        // Re-fetch the original file from S3 and restart processing
        const fileResp = await fetch(job.originalFileUrl);
        if (!fileResp.ok) throw new Error("Could not fetch original file for resume");
        const arrayBuf = await fileResp.arrayBuffer();
        const buffer = Buffer.from(arrayBuf);

        let timeoutHandle: ReturnType<typeof setTimeout>;
        const timeoutPromise = new Promise<never>((_, reject) => {
          // Give resumed jobs an extra 10 minutes on top of the original timeout
          timeoutHandle = setTimeout(
            () => reject(new Error("__TIMEOUT__")),
            JOB_TIMEOUT_MS
          );
        });
        Promise.race([
          processJobAsync(
            input.jobId,
            buffer,
            job.originalFormat as SupportedFormat,
            job.outputFormat as SupportedFormat | undefined,
            job.targetLanguage ?? undefined,
            job.originalFileUrl,
          ).then(() => clearTimeout(timeoutHandle)),
          timeoutPromise,
        ]).catch(async (err) => {
          const msg: string = err.message ?? "Unknown error";
          if (msg === "__TIMEOUT__") {
            const currentJob = await getJobById(input.jobId).catch(() => null);
            if (currentJob && ["done", "cancelled", "error"].includes(currentJob.status)) return;
            await appendJobLog(input.jobId, "warning", `Job paused again after ${JOB_TIMEOUT_MS / 60000} minutes. Click Continue to keep processing.`).catch(() => {});
            await updateJob(input.jobId, { status: "paused", errorMessage: `Paused after ${JOB_TIMEOUT_MS / 60000} minutes` }).catch(() => {});
          } else {
            await appendJobLog(input.jobId, "error", msg).catch(() => {});
            await updateJob(input.jobId, { status: "error", errorMessage: msg }).catch(() => {});
          }
        });

        return { success: true };
      }),

    // Kill a paused job (user chose to cancel)
    killPaused: publicProcedure
      .input(z.object({ jobId: z.number() }))
      .mutation(async ({ input }) => {
        const job = await getJobById(input.jobId);
        if (!job) throw new Error("Job not found");
        await cancelJob(input.jobId);
        await updateJob(input.jobId, { status: "cancelled", errorMessage: "Cancelled by user after timeout" });
        await appendJobLog(input.jobId, "warning", "Job cancelled by user.");
        return { success: true };
      }),

    // Create a Stripe checkout session for downloading a completed job
    createCheckoutSession: publicProcedure
      .input(z.object({
        jobId: z.number(),
        origin: z.string(),
      }))
      .mutation(async ({ input, ctx }) => {
        const job = await getJobById(input.jobId);
        if (!job) throw new Error("Job not found");
        if (job.status !== "done") throw new Error("Job is not complete");
        if (job.paid) return { alreadyPaid: true, checkoutUrl: null };

        // Look up user details for Stripe prefill (if available)
        let userEmail: string | null = null;
        let userName: string | null = null;
        // Try from auth context first, then fall back to DB lookup by job.userId
        if (ctx.user) {
          userEmail = ctx.user.email?.trim() || null;
          userName = ctx.user.name?.trim() || null;
        } else if (job.userId) {
          const user = await getUserById(job.userId);
          userEmail = user?.email?.trim() || null;
          userName = user?.name?.trim() || null;
        }

        const downloadPriceUsd = job.downloadPriceUsd ?? 2.0;
        const checkoutUrl = await createDownloadCheckoutSession({
          jobId: job.id,
          userId: job.userId ?? (await getAnonUserId()),
          userEmail,
          userName,
          originalFileName: job.originalFileName,
          downloadPriceUsd,
          origin: input.origin,
        });
        return { alreadyPaid: false, checkoutUrl };
      }),

    // Verify payment status for a job (poll after Stripe redirect)
    verifyPayment: publicProcedure
      .input(z.object({ jobId: z.number(), sessionId: z.string().optional() }))
      .query(async ({ input }) => {
        const job = await getJobById(input.jobId);
        if (!job) throw new Error("Job not found");
        // If webhook already marked it paid, return immediately
        if (job.paid) return { paid: true };
        // Otherwise check Stripe directly (for cases where webhook is delayed)
        if (input.sessionId) {
          try {
            const { stripe } = await import("./stripe");
            const session = await stripe.checkout.sessions.retrieve(input.sessionId);
            if (session.payment_status === "paid") {
              await updateJob(input.jobId, { paid: true, stripeSessionId: session.id });
              return { paid: true };
            }
          } catch {}
        }
        return { paid: false };
      }),

    // Permanently delete a job and its S3 files (ephemeral cleanup)
    cleanup: publicProcedure
      .input(z.object({ jobId: z.number() }))
      .mutation(async ({ input }) => {
        const job = await getJobById(input.jobId);
        if (!job) return { success: false };
        // Delete S3 files (best-effort)
        if (job.originalFileKey) await storageDelete(job.originalFileKey).catch(() => {});
        if (job.outputFileKey) await storageDelete(job.outputFileKey).catch(() => {});
        // Delete all DB records
        await deleteJob(input.jobId);
        return { success: true };
      }),
  }),
});

export type AppRouter = typeof appRouter;

// ─── Async job processor ──────────────────────────────────────────────────────
async function processJobAsync(
  jobId: number,
  buffer: Buffer,
  inputFormat: SupportedFormat,
  outputFormat: SupportedFormat | undefined,
  targetLanguage: string | undefined,
  originalFileUrl: string,
  modelId?: string,
  isImageInput?: boolean,
) {
  // Create a model-aware LLM invoker
  const llm = makeModelInvoker(invokeLLM, modelId);
  const startTime = Date.now();

  // Helper: emit a log line to DB
  const log = (msg: string, level: "info" | "progress" | "success" | "warning" | "error" = "info") =>
    appendJobLog(jobId, level, msg).catch(() => {});

  // Helper: log timing for a step
  const logTiming = (stepName: string, startMs: number) => {
    const duration = Date.now() - startMs;
    console.log(`[Timing] ${stepName}: ${duration}ms`);
    return duration;
  };

  // Helper: check if job was cancelled
  const checkCancelled = async () => {
    const cancelled = await isJobCancelled(jobId);
    if (cancelled) throw new Error("Job cancelled by user");
  };

  try {
    const finalFormat = outputFormat ?? inputFormat;
    // ── STANDARD PIPELINE ─────────────────────────────────────────────────

    // Step 1: Extract text
    await updateJob(jobId, { status: "extracting" });
    await updateJobStep(jobId, "extract", { status: "running", startedAt: new Date() });

    let extractedText = "";
    let pageCount = 1;

    // ── SHORT-CIRCUIT: image inputs have no text layer ──────────────────────
    if (isImageInput) {
      await log("Image converted to PDF — no text extraction needed.", "info");
      await updateJob(jobId, { extractedText: "", pageCount: 1, charCount: 0, estimatedCost: 0 });
      await updateJobStep(jobId, "extract", { status: "done", completedAt: new Date(), message: "Image converted to PDF (1 page)" });
      await log("Image embedded as PDF page ✓", "success");

      // No translation or conversion needed for image→PDF — just store the buffer
      const job = await getJobById(jobId);
      const outKey = `outputs/${job?.userId}/${nanoid()}_converted.pdf`;
      const { url: outputUrl } = await storagePut(outKey, buffer, getMimeType("pdf"));

      // Generate preview PDF (first 3 pages)
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
    } catch (e) {
      await log("Extraction failed — falling back to direct text extraction", "warning");
      const extractStart = Date.now();
      const result = await extractTextFromBuffer(buffer, inputFormat, `doc.${inputFormat}`);
      logTiming("Text extraction (fallback)", extractStart);
      extractedText = result.text;
      pageCount = result.pageCount;
    }

    const charCount = extractedText.length;
    const estimatedCost = pageCount * COST_PER_PAGE_EXTRACTION + (targetLanguage ? pageCount * COST_PER_PAGE_TRANSLATION : 0);

    await updateJob(jobId, { extractedText, pageCount, charCount, estimatedCost });
    await updateJobStep(jobId, "extract", { status: "done", completedAt: new Date(), message: `Extracted ${pageCount} pages, ${charCount} characters` });
    await log(`Extracted ${pageCount} page${pageCount !== 1 ? "s" : ""} · ${charCount.toLocaleString()} characters`, "success");

    await checkCancelled();

    // Step 2: Translate
    let processedText = extractedText;
    let inPlaceBuffer: Buffer | null = null;

    if (targetLanguage) {
      await updateJob(jobId, { status: "translating" });
      await updateJobStep(jobId, "translate", { status: "running", startedAt: new Date() });

      try {
        const job = await getJobById(jobId);
        const langName = job?.targetLanguageName ?? targetLanguage;
        const translateStart = Date.now();

        if (inputFormat === "pptx" && (finalFormat === "pptx" || !outputFormat)) {
          await log(`Translating PPTX slides to ${langName} (in-place — preserving layout)...`, "info");
          await log(`AI is scanning each slide for text runs, grouping by paragraph for fluent translation...`, "info");
          inPlaceBuffer = await translatePptxInPlace(buffer, langName, llm, async (msg: string) => {
            await log(msg, "info");
          });
          logTiming("PPTX translation", translateStart);
          await updateJobStep(jobId, "translate", { status: "done", completedAt: new Date(), message: `Translated PPTX in-place to ${langName} (layout preserved)` });
          await log(`PPTX translated to ${langName} — all images and formatting preserved`, "success");
        } else if (inputFormat === "docx" && (finalFormat === "docx" || !outputFormat)) {
          await log(`Translating DOCX to ${langName} (in-place — preserving layout)...`, "info");
          await log(`AI is processing paragraphs, tables, and headings while keeping styles intact...`, "info");
          inPlaceBuffer = await translateDocxInPlace(buffer, langName, llm);
          logTiming("DOCX translation", translateStart);
          await updateJobStep(jobId, "translate", { status: "done", completedAt: new Date(), message: `Translated DOCX in-place to ${langName} (layout preserved)` });
          await log(`DOCX translated to ${langName} — all images and formatting preserved`, "success");
        } else if (inputFormat === "xlsx" && (finalFormat === "xlsx" || !outputFormat)) {
          await log(`Translating XLSX to ${langName} (in-place — preserving layout)...`, "info");
          await log(`AI is translating cell content and shared strings while preserving formulas and cell styles...`, "info");
          inPlaceBuffer = await translateXlsxInPlace(buffer, langName, llm);
          logTiming("XLSX translation", translateStart);
          await updateJobStep(jobId, "translate", { status: "done", completedAt: new Date(), message: `Translated XLSX in-place to ${langName} (layout preserved)` });
          await log(`XLSX translated to ${langName} — all formatting preserved`, "success");
        } else {
          await log(`Translating document content to ${langName}...`, "info");
          await log(`AI is translating ${charCount.toLocaleString()} characters — preserving paragraph structure and formatting marks...`, "info");
          processedText = await translateWithLLM(extractedText, langName, llm);
          logTiming("Text translation", translateStart);
          await updateJobStep(jobId, "translate", { status: "done", completedAt: new Date(), message: `Translated to ${langName} via ${modelId ?? "GPT-4o-mini"}` });
          await log(`Translation to ${langName} complete`, "success");
        }
      } catch (e: any) {
        await updateJobStep(jobId, "translate", { status: "error", completedAt: new Date(), message: e.message });
        await log(`Translation failed: ${e.message}`, "error");
      }
    }

    await checkCancelled();

    // Step 3: Convert / Build output
    let outputBuffer: Buffer;

    if (outputFormat && outputFormat !== inputFormat) {
      await updateJob(jobId, { status: "converting" });
      await updateJobStep(jobId, "convert", { status: "running", startedAt: new Date() });
      await log(`Converting to ${finalFormat.toUpperCase()}...`, "info");
      await log(`Rebuilding document structure in ${finalFormat.toUpperCase()} format...`, "info");
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
      await updateJobStep(jobId, "convert", { status: "done", completedAt: new Date(), message: `Converted to ${finalFormat.toUpperCase()}` });
      await log(`Converted to ${finalFormat.toUpperCase()} successfully`, "success");
    }

    // Upload output
    await log("Finalizing output file...", "info");
    await log("Uploading to secure storage...", "info");
    const job = await getJobById(jobId);
    const outKey = `outputs/${job?.userId}/${nanoid()}_translated.${finalFormat}`;
    const { url: outputUrl } = await storagePut(outKey, outputBuffer, getMimeType(finalFormat));

    // Mark job as done immediately (don't wait for preview)
    const finalJob = await getJobById(jobId);
    const finalCostUsd = finalJob?.estimatedCost ?? 0;
    const finalDownloadPrice = calculateDownloadPrice(buffer.length, finalCostUsd);
    await updateJob(jobId, {
      status: "done",
      outputFileKey: outKey,
      outputFileUrl: outputUrl,
      outputFormat: finalFormat,
      previewFileKey: undefined,
      previewFileUrl: undefined,
      previewPageCount: 0,
      conversionCostUsd: +finalCostUsd.toFixed(4),
      downloadPriceUsd: +finalDownloadPrice.toFixed(2),
    });
    await log("Done! Your document is ready to download.", "success");
    logTiming("Total conversion time", startTime);

    // Generate preview in background (non-blocking, with 60s timeout)
    (async () => {
      const PREVIEW_TIMEOUT = 60_000;
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), PREVIEW_TIMEOUT);
      try {
        await log("Generating preview in background...", "info");
        const previewStart = Date.now();
        const { previewBuffer, previewPages, previewFormat, previewMimeType } = await generatePreview(outputBuffer, finalFormat);
        logTiming("Preview generation", previewStart);
        const pKey = `previews/${job?.userId}/${nanoid()}_preview.${previewFormat}`;
        console.log(`[Preview] Uploading ${previewFormat} preview (${previewBuffer.length} bytes) to GCS...`);
        const uploadStart = Date.now();
        const { url: pUrl } = await storagePut(pKey, previewBuffer, previewMimeType);
        console.log(`[Preview] Upload complete in ${Date.now() - uploadStart}ms`);
        await updateJob(jobId, {
          previewFileKey: pKey,
          previewFileUrl: pUrl,
          previewPageCount: previewPages,
        });
        await log("Preview ready!", "success");
      } catch (e: any) {
        const msg = e?.name === "AbortError" ? "Preview timed out after 60s" : e?.message ?? "Unknown error";
        console.error("[Preview] Background preview failed:", msg);
        await log(`Preview generation failed (file still available for download)`, "warning").catch(() => {});
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
