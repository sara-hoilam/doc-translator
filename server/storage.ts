/**
 * storage.ts — Google Cloud Storage backend
 *
 * Replaces the Manus Forge storage proxy.
 *
 * Required env vars:
 *   GCS_BUCKET_NAME   — the GCS bucket name (e.g. "pdfgodwork-files")
 *   GCS_PROJECT_ID    — your GCP project ID
 *   GCS_KEY_FILE_PATH — (local dev only) path to service-account JSON key
 *                       In Cloud Run / production, leave empty and use ADC.
 */

import { Storage } from "@google-cloud/storage";
import path from "path";
import { ENV } from "./_core/env";

function getStorageClient(): Storage {
  if (ENV.gcsKeyFilePath) {
    return new Storage({
      projectId: ENV.gcsProjectId || undefined,
      keyFilename: ENV.gcsKeyFilePath,
    });
  }
  // Application Default Credentials (ADC) — works on Cloud Run, GCE, etc.
  return new Storage({ projectId: ENV.gcsProjectId || undefined });
}

function getBucket() {
  if (!ENV.gcsBucketName) {
    throw new Error(
      "GCS_BUCKET_NAME is not configured. Set it in your .env file."
    );
  }
  return getStorageClient().bucket(ENV.gcsBucketName);
}

/** Signed URL expiry: 7 days (max allowed for service-account signed URLs) */
const SIGNED_URL_EXPIRY_MS = 7 * 24 * 60 * 60 * 1000;

function normalizeKey(relKey: string): string {
  return relKey.replace(/^\/+/, "");
}

/**
 * Upload a buffer to GCS and return a time-limited signed download URL.
 */
export async function storagePut(
  relKey: string,
  data: Buffer | Uint8Array | string,
  contentType = "application/octet-stream"
): Promise<{ key: string; url: string }> {
  const key = normalizeKey(relKey);
  const bucket = getBucket();
  const file = bucket.file(key);

  const buffer =
    typeof data === "string"
      ? Buffer.from(data, "utf-8")
      : Buffer.from(data as Uint8Array);

  await file.save(buffer, {
    contentType,
    resumable: false, // resumable uploads need extra config; fine to skip for <50 MB
  });

  const [signedUrl] = await file.getSignedUrl({
    action: "read",
    expires: Date.now() + SIGNED_URL_EXPIRY_MS,
  });

  return { key, url: signedUrl };
}

/**
 * Return a fresh signed download URL for an existing GCS object.
 */
export async function storageGet(
  relKey: string
): Promise<{ key: string; url: string }> {
  const key = normalizeKey(relKey);
  const bucket = getBucket();
  const file = bucket.file(key);

  const [signedUrl] = await file.getSignedUrl({
    action: "read",
    expires: Date.now() + SIGNED_URL_EXPIRY_MS,
  });

  return { key, url: signedUrl };
}

/**
 * Delete an object from GCS. Errors are swallowed (best-effort).
 */
export async function storageDelete(relKey: string): Promise<void> {
  const key = normalizeKey(relKey);
  const bucket = getBucket();
  try {
    await bucket.file(key).delete();
  } catch {
    // Ignore — file may already be gone
  }
}
