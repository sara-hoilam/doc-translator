import "dotenv/config";
import express, { type Request, type Response } from "express";
import { createServer } from "http";
import net from "net";
import { createExpressMiddleware } from "@trpc/server/adapters/express";
import { registerOAuthRoutes } from "./oauth";
import { appRouter } from "../routers";
import { createContext } from "./context";
import { serveStatic, setupVite } from "./vite";
import { registerStripeWebhook } from "../stripeWebhook";
import { getJobById, deleteJob } from "../db";
import { storageDelete } from "../storage";

function isPortAvailable(port: number): Promise<boolean> {
  return new Promise(resolve => {
    const server = net.createServer();
    server.listen(port, () => {
      server.close(() => resolve(true));
    });
    server.on("error", () => resolve(false));
  });
}

async function findAvailablePort(startPort: number = 3000): Promise<number> {
  for (let port = startPort; port < startPort + 20; port++) {
    if (await isPortAvailable(port)) {
      return port;
    }
  }
  throw new Error(`No available port found starting from ${startPort}`);
}

async function startServer() {
  const app = express();
  const server = createServer(app);
  // Stripe webhook MUST be registered before express.json() to get raw body
  registerStripeWebhook(app);
  // Configure body parser with larger size limit for file uploads
  app.use(express.json({ limit: "50mb" }));
  app.use(express.urlencoded({ limit: "50mb", extended: true }));
  // OAuth callback under /api/oauth/callback
  registerOAuthRoutes(app);

  // ── Secure download proxy ────────────────────────────────────────────────
  // Fetches the output file from GCS server-side and streams it to the browser.
  // This avoids CORS issues and ensures cleanup only happens AFTER delivery.
  app.get("/api/download/:jobId", async (req: Request, res: Response) => {
    const jobId = parseInt(req.params.jobId, 10);
    if (isNaN(jobId)) { res.status(400).json({ error: "Invalid job ID" }); return; }

    const job = await getJobById(jobId).catch(() => null);
    if (!job || !job.outputFileUrl || !job.paid) {
      res.status(404).json({ error: "File not found or not paid" });
      return;
    }

    try {
      const upstream = await fetch(job.outputFileUrl);
      if (!upstream.ok) {
        res.status(502).json({ error: "Could not fetch file from storage" });
        return;
      }

      const ext = job.outputFormat ?? job.originalFormat ?? "bin";
      const rawName = job.originalFileName
        ? job.originalFileName.replace(/\.[^.]+$/, "") + `_translated.${ext}`
        : `download.${ext}`;
      const safeName = encodeURIComponent(rawName);

      res.setHeader("Content-Type", upstream.headers.get("content-type") ?? "application/octet-stream");
      res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${safeName}`);
      if (upstream.headers.get("content-length")) {
        res.setHeader("Content-Length", upstream.headers.get("content-length")!);
      }

      // Stream the file body to the client
      const reader = upstream.body!.getReader();
      const pump = async () => {
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          res.write(value);
        }
        res.end();
      };
      await pump();

      // Cleanup GCS files after successful delivery (best-effort)
      setImmediate(async () => {
        try {
          if (job.originalFileKey) await storageDelete(job.originalFileKey);
          if (job.outputFileKey)   await storageDelete(job.outputFileKey);
          await deleteJob(jobId);
        } catch {}
      });
    } catch (err) {
      console.error("[Download] Failed:", err);
      if (!res.headersSent) res.status(500).json({ error: "Download failed" });
    }
  });
  // tRPC API
  app.use(
    "/api/trpc",
    createExpressMiddleware({
      router: appRouter,
      createContext,
    })
  );
  // development mode uses Vite, production mode uses static files
  if (process.env.NODE_ENV === "development") {
    await setupVite(app, server);
  } else {
    serveStatic(app);
  }

  const preferredPort = parseInt(process.env.PORT || "3000");
  const port = await findAvailablePort(preferredPort);

  if (port !== preferredPort) {
    console.log(`Port ${preferredPort} is busy, using port ${port} instead`);
  }

  server.listen(port, () => {
    console.log(`Server running on http://localhost:${port}/`);
  });
}

startServer().catch(console.error);
