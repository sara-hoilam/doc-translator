import { z } from "zod";
import { COOKIE_NAME } from "@shared/const";
import { getSessionCookieOptions } from "./_core/cookies";
import { systemRouter } from "./_core/systemRouter";
import { publicProcedure, protectedProcedure, router } from "./_core/trpc";
import { storageDelete } from "./storage";
import { createDownloadCheckoutSession } from "./stripe";
import {
  getJobById, updateJob, getUserById, getAnonUserId,
  appendJobLog, getJobLogs, cancelJob, deleteJob, getJobSteps,
} from "./db";
import {
  COST_PER_PAGE_EXTRACTION, COST_PER_PAGE_TRANSLATION, COST_PER_PAGE_TOTAL,
  MAX_FILE_SIZE_BYTES, JOB_TIMEOUT_MS, MAX_PAGES_VISION, COST_WARNING_THRESHOLD,
  LANGUAGES, SUPPORTED_INPUT_FORMATS, SUPPORTED_OUTPUT_FORMATS, AVAILABLE_MODELS,
  type SupportedFormat,
} from "./docProcessor";
import {
  startDocumentJob,
  convertMultipleImagesToPdf,
  scheduleJobProcessing,
} from "./documentJob";

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
      .mutation(async ({ input }) => {
        if (input.multiImageBase64 && input.multiImageBase64.length > 0) {
          const imageBuffers = input.multiImageBase64.map(b64 => Buffer.from(b64, "base64") as Buffer);
          const mimeTypes = input.multiImageMimeTypes ?? imageBuffers.map(() => "image/png");
          const totalSize = imageBuffers.reduce((s, b) => s + b.length, 0);
          if (totalSize > MAX_FILE_SIZE_BYTES * 5) {
            throw new Error(`Total image size too large: ${(totalSize / 1024 / 1024).toFixed(1)} MB.`);
          }
          const processBuffer = await convertMultipleImagesToPdf(imageBuffers, mimeTypes);
          const { jobId } = await startDocumentJob({
            filename: input.filename || "merged-images.pdf",
            mimeType: "application/pdf",
            buffer: processBuffer,
            outputFormat: input.outputFormat ?? "pdf",
            targetLanguage: input.targetLanguage,
            targetLanguageName: input.targetLanguageName,
            modelId: input.modelId,
          });
          return { jobId };
        }

        const buffer = Buffer.from(input.base64Data, "base64");
        const { jobId } = await startDocumentJob({
          filename: input.filename,
          mimeType: input.mimeType,
          buffer,
          outputFormat: input.outputFormat,
          targetLanguage: input.targetLanguage,
          targetLanguageName: input.targetLanguageName,
          modelId: input.modelId,
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

        scheduleJobProcessing({
          jobId: input.jobId,
          processBuffer: buffer,
          format: job.originalFormat as SupportedFormat,
          outputFormat: job.outputFormat as SupportedFormat | undefined,
          targetLanguage: job.targetLanguage ?? undefined,
          originalFileUrl: job.originalFileUrl,
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
