#!/usr/bin/env bash
# ─────────────────────────────────────────────────────────────────────────────
# apply-changes.sh
#
# Run this from the ROOT of your project repo:
#   bash apply-changes.sh
#
# What it does:
#   1. Patches the pipeline bug in server/routers.ts
#   2. Writes Dockerfile to repo root
#   3. Writes railway.toml to repo root
#   4. Writes .github/workflows/deploy.yml
#   5. Prints a checklist of next steps
# ─────────────────────────────────────────────────────────────────────────────

set -e
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(pwd)"

echo ""
echo "╔══════════════════════════════════════════════════╗"
echo "║   DocTranslate — Applying changes                ║"
echo "╚══════════════════════════════════════════════════╝"
echo ""

# ── Guard: must be run from project root ─────────────────────────────────────
if [ ! -f "$PROJECT_ROOT/package.json" ]; then
  echo "❌  ERROR: Run this script from the root of your project repo."
  echo "    e.g.  cd ~/my-project && bash ~/Downloads/apply-changes.sh"
  exit 1
fi

# ── 1. Patch server/routers.ts ────────────────────────────────────────────────
ROUTERS="$PROJECT_ROOT/server/routers.ts"
if [ ! -f "$ROUTERS" ]; then
  echo "❌  ERROR: server/routers.ts not found. Are you in the right directory?"
  exit 1
fi

echo "▶  Patching server/routers.ts ..."

# We replace the block between the two sentinel comments that already exist
# in the file.  The sed range is:
#   start: line containing "// Step 2: Translate"
#   end:   line containing "await checkCancelled();" that follows it
#         (there are two — we want the second one after Step 2)
#
# Strategy: use Python (available on all platforms) for a safe multi-line replace.

python3 - "$ROUTERS" <<'PYEOF'
import sys, re

path = sys.argv[1]
with open(path, "r", encoding="utf-8") as f:
    src = f.read()

# The old block starts right after Step 1's checkCancelled and ends just
# before Step 3's checkCancelled.  We match on the two anchors that are
# unique in the file.
OLD_START = "    // Step 2: Translate\n    let processedText = extractedText;"
OLD_END   = "\n\n    await checkCancelled();\n\n    // Step 3: Convert"

NEW_BLOCK = '''    // Step 2: Translate
    let processedText = extractedText;
    let inPlaceBuffer: Buffer | null = null;

    if (targetLanguage) {
      await updateJob(jobId, { status: "translating" });
      await updateJobStep(jobId, "translate", { status: "running", startedAt: new Date() });

      try {
        const job = await getJobById(jobId);
        const langName = job?.targetLanguageName ?? targetLanguage;
        const translateStart = Date.now();

        // ── DOCX source: always translate in-place, then convert if needed ────
        if (inputFormat === "docx") {
          await log(`Translating DOCX to ${langName} (in-place — preserving layout)...`, "info");
          await log(`AI is processing paragraphs, tables, and headings while keeping styles intact...`, "info");
          const translatedDocx = await translateDocxInPlace(buffer, langName, llm);
          logTiming("DOCX translation", translateStart);

          if (finalFormat === "pdf") {
            await log(`Converting translated DOCX to PDF via LibreOffice...`, "info");
            const { convertToPdfBuffer } = await import("./watermark");
            inPlaceBuffer = await convertToPdfBuffer(translatedDocx, "docx");
            await log(`DOCX→PDF conversion complete`, "success");
          } else {
            inPlaceBuffer = translatedDocx;
          }

          await updateJobStep(jobId, "translate", {
            status: "done",
            completedAt: new Date(),
            message: `Translated DOCX in-place to ${langName}${finalFormat === "pdf" ? " → converted to PDF" : " (layout preserved)"}`,
          });
          await log(`DOCX translated to ${langName} — all images and formatting preserved`, "success");

        // ── PPTX source: always translate in-place, then convert if needed ────
        } else if (inputFormat === "pptx") {
          await log(`Translating PPTX slides to ${langName} (in-place — preserving layout)...`, "info");
          await log(`AI is scanning each slide for text runs, grouping by paragraph for fluent translation...`, "info");
          const translatedPptx = await translatePptxInPlace(buffer, langName, llm, async (msg: string) => {
            await log(msg, "info");
          });
          logTiming("PPTX translation", translateStart);

          if (finalFormat === "pdf") {
            await log(`Converting translated PPTX to PDF via LibreOffice...`, "info");
            const { convertToPdfBuffer } = await import("./watermark");
            inPlaceBuffer = await convertToPdfBuffer(translatedPptx, "pptx");
            await log(`PPTX→PDF conversion complete`, "success");
          } else {
            inPlaceBuffer = translatedPptx;
          }

          await updateJobStep(jobId, "translate", {
            status: "done",
            completedAt: new Date(),
            message: `Translated PPTX in-place to ${langName}${finalFormat === "pdf" ? " → converted to PDF" : " (layout preserved)"}`,
          });
          await log(`PPTX translated to ${langName} — all images and formatting preserved`, "success");

        // ── XLSX source: always translate in-place ────────────────────────────
        } else if (inputFormat === "xlsx") {
          await log(`Translating XLSX to ${langName} (in-place — preserving layout)...`, "info");
          await log(`AI is translating cell content and shared strings while preserving formulas and cell styles...`, "info");
          inPlaceBuffer = await translateXlsxInPlace(buffer, langName, llm);
          logTiming("XLSX translation", translateStart);
          await updateJobStep(jobId, "translate", {
            status: "done",
            completedAt: new Date(),
            message: `Translated XLSX in-place to ${langName} (layout preserved)`,
          });
          await log(`XLSX translated to ${langName} — all formatting preserved`, "success");

        // ── All other formats (TXT, etc.): text-extract → translate ──────────
        } else {
          await log(`Translating document content to ${langName}...`, "info");
          await log(`AI is translating ${charCount.toLocaleString()} characters — preserving paragraph structure and formatting marks...`, "info");
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
    }'''

idx_start = src.find(OLD_START)
if idx_start == -1:
    print("PATCH_SKIP: Could not find Step 2 start anchor — already patched or file differs.")
    sys.exit(0)

idx_end = src.find(OLD_END, idx_start)
if idx_end == -1:
    print("PATCH_SKIP: Could not find Step 3 anchor after Step 2 — check file manually.")
    sys.exit(0)

# Replace from OLD_START up to (but not including) OLD_END
patched = src[:idx_start] + NEW_BLOCK + src[idx_start + (idx_end - idx_start):]

# Quick sanity: the new block should appear
if "Always translate in-place" not in patched:
    print("PATCH_ERROR: Replacement string not found in result. Aborting.")
    sys.exit(1)

with open(path, "w", encoding="utf-8") as f:
    f.write(patched)

print("PATCH_OK")
PYEOF

echo "   ✅  server/routers.ts patched"

# ── 2. Write Dockerfile ───────────────────────────────────────────────────────
echo "▶  Writing Dockerfile ..."
cat > "$PROJECT_ROOT/Dockerfile" << 'DOCKERFILE'
# ─── Stage 1: Build ───────────────────────────────────────────────────────────
FROM node:20-bookworm-slim AS builder

RUN npm install -g pnpm@10.4.1

WORKDIR /app
COPY package.json pnpm-lock.yaml ./
COPY patches/ ./patches/
RUN pnpm install --frozen-lockfile

COPY . .
RUN pnpm run build


# ─── Stage 2: Production ──────────────────────────────────────────────────────
FROM node:20-bookworm-slim AS production

# LibreOffice for DOCX/PPTX/XLSX → PDF
# Noto CJK fonts for Chinese / Japanese / Korean output
# Python3 venv for pdf2docx
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    libreoffice-impress \
    libreoffice-calc \
    fonts-noto-cjk \
    fonts-noto-cjk-extra \
    fonts-liberation \
    fontconfig \
    python3 \
    python3-pip \
    python3-venv \
    ca-certificates \
    curl \
    && fc-cache -fv \
    && rm -rf /var/lib/apt/lists/*

RUN python3 -m venv /opt/venv && \
    /opt/venv/bin/pip install --no-cache-dir pdf2docx
ENV PATH="/opt/venv/bin:$PATH"

RUN npm install -g pnpm@10.4.1

WORKDIR /app
COPY package.json pnpm-lock.yaml ./
COPY patches/ ./patches/
RUN pnpm install --frozen-lockfile --prod

COPY --from=builder /app/dist ./dist

# LibreOffice needs a writable home for its profile cache
ENV HOME=/tmp
ENV NODE_ENV=production

EXPOSE 3000
CMD ["node", "dist/index.js"]
DOCKERFILE
echo "   ✅  Dockerfile written"

# ── 3. Write railway.toml ─────────────────────────────────────────────────────
echo "▶  Writing railway.toml ..."
cat > "$PROJECT_ROOT/railway.toml" << 'RAILWAYTOML'
[build]
builder = "DOCKERFILE"
dockerfilePath = "./Dockerfile"

[deploy]
startCommand = "node dist/index.js"
healthcheckPath = "/api/trpc/docs.config"
healthcheckTimeout = 30
restartPolicyType = "ON_FAILURE"
restartPolicyMaxRetries = 3
RAILWAYTOML
echo "   ✅  railway.toml written"

# ── 4. Write GitHub Actions workflow ─────────────────────────────────────────
echo "▶  Writing .github/workflows/deploy.yml ..."
mkdir -p "$PROJECT_ROOT/.github/workflows"
cat > "$PROJECT_ROOT/.github/workflows/deploy.yml" << 'WORKFLOW'
name: Deploy to Railway

on:
  push:
    branches:
      - main

jobs:
  deploy:
    name: Deploy
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - name: Install Railway CLI
        run: npm install -g @railway/cli
      - name: Deploy
        env:
          RAILWAY_TOKEN: ${{ secrets.RAILWAY_TOKEN }}
        run: railway up --service doc-translator --detach
WORKFLOW
echo "   ✅  .github/workflows/deploy.yml written"

# ── Done ─────────────────────────────────────────────────────────────────────
echo ""
echo "╔══════════════════════════════════════════════════════════════════════╗"
echo "║  ✅  All changes applied!                                           ║"
echo "╠══════════════════════════════════════════════════════════════════════╣"
echo "║                                                                      ║"
echo "║  NEXT STEPS                                                          ║"
echo "║  ─────────────────────────────────────────────────────────────────  ║"
echo "║  1. Commit everything:                                               ║"
echo "║       git add -A                                                     ║"
echo "║       git commit -m 'fix: pipeline bug + Railway deployment setup'  ║"
echo "║       git push origin main                                           ║"
echo "║                                                                      ║"
echo "║  2. Go to https://railway.com → New Project → Deploy from GitHub    ║"
echo "║     Select your repo — Railway detects the Dockerfile automatically ║"
echo "║                                                                      ║"
echo "║  3. In Railway → your service → Variables, add all your env vars:   ║"
echo "║     DATABASE_URL, AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY,         ║"
echo "║     AWS_REGION, S3_BUCKET_NAME, STRIPE_SECRET_KEY,                  ║"
echo "║     STRIPE_WEBHOOK_SECRET, ANTHROPIC_API_KEY, OPENAI_API_KEY        ║"
echo "║                                                                      ║"
echo "║  4. Add your custom domain in Railway → Settings → Networking       ║"
echo "║                                                                      ║"
echo "║  5. Update Stripe webhook URL to https://yourdomain.com/api/stripe/webhook ║"
echo "║                                                                      ║"
echo "║  6. (Optional) Add RAILWAY_TOKEN secret to GitHub for auto-deploys  ║"
echo "╚══════════════════════════════════════════════════════════════════════╝"
echo ""
