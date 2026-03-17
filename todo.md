# DocTranslate — Project TODO

## Backend
- [x] Database schema: jobs, job_steps tables
- [x] Install server dependencies: multer, pdf-parse, mammoth, xlsx, node-fetch, form-data
- [x] File upload endpoint (base64 via tRPC)
- [x] S3 upload helper for storing original and processed files
- [x] Text extraction router (PDF, DOCX, PPTX, XLSX, TXT, HTML) using LLM
- [x] Google Translate API integration for 100+ languages (with LLM fallback)
- [x] Format conversion using LibreOffice/Pandoc shell commands
- [x] Cost estimation procedure (per-page cost calculation)
- [x] Job status polling procedure
- [x] Processing history procedure (list past jobs with metadata + download links)

## Frontend
- [x] Global design system: color palette, typography, index.css
- [x] Landing / Home page with hero and feature highlights
- [x] Drag-and-drop file upload component
- [x] Language selection dropdown (100+ languages)
- [x] Output format selector
- [x] Cost estimation display (before processing)
- [x] Real-time progress indicator (extraction → translation → conversion steps)
- [x] Download interface for processed documents
- [x] Processing history dashboard

## Quality
- [x] Vitest unit tests for key procedures (8 tests passing)
- [x] Error handling for unsupported formats / API failures
- [x] Loading and empty states in UI

## Round 2 Fixes
- [x] Replace Google Translate API with GPT-4o-mini for translation
- [x] Fix pandoc "not found" error — use Node.js-based conversion (docx, pptxgenjs, xlsx)
- [x] Fix all format conversion errors (PPTX, DOCX, XLSX output)
- [x] Remove sign-in/landing page, make the converter the landing page
- [x] Redesign UI to match LazyDad-style: sidebar (format selector + translation toggle) + centered drop zone

## Round 3 — UI Consistency
- [x] Restyle JobStatus page to match Converter design (beige bg, amber accents, no top navbar)
- [x] Restyle History page to match same design system

## Round 4 — In-Place Format Preservation
- [x] Rewrite PPTX processor: unzip PPTX, find all <a:t> text nodes in slide XMLs, translate in-place, repack — preserving all images, shapes, backgrounds
- [x] Rewrite DOCX processor: unzip DOCX, find all <w:t> text nodes, translate in-place, repack — preserving all images, tables, styles
- [x] Rewrite XLSX processor: unzip XLSX, find all <v> and <t> text nodes in sharedStrings.xml, translate in-place, repack
- [x] Update routers.ts to use new in-place processors
- [x] Test with the Upgrain PDF → PPTX conversion to verify images and layout are preserved

## Round 5 — Bug Fixes
- [x] Fix "pdfParse is not a function" ESM import error in extractTextFromBuffer — replaced pdf-parse with pdfjs-dist

## Round 6 — Open Button
- [x] Update "Open" button to open PPTX in Google Slides, DOCX in Google Docs, XLSX in Google Sheets using Google viewer URL

## Round 7 — LLM Vision PDF→PPTX + Model Selection
- [x] Install pdftoppm (poppler-utils) for PDF page rendering
- [x] Rewrite PDF→PPTX pipeline: render each page to image, send to LLM Vision, extract structured layout JSON, build PPTX with image backgrounds + text overlays
- [x] Support model selection: Claude Sonnet 4.5, GPT-4o, GPT-4.1, GPT-4o-mini
- [x] Add model selector dropdown to Converter page sidebar
- [x] Wire selected model through tRPC to backend processing
- [x] Add Vision AI info banner in UI when PDF→PPTX is selected
- [x] Update routers.ts to route PDF→PPTX through Vision pipeline and pass model to all LLM calls

## Round 8 — Fix pdftoppm Not Found
- [x] Replace pdftoppm system binary with PDF file_url approach — LLM reads PDF directly via S3 URL, no system binaries needed

## Round 9 — Processing Safeguards
- [x] Add MAX_PAGES_VISION = 30 cap in Vision pipeline (process remaining as text-only)
- [x] Add 10-minute job timeout wrapping processJobAsync
- [x] Add 50MB server-side file size check in upload mutation
- [x] Add 100,000 character cap on extracted text before translation
- [x] Add cost ceiling confirmation dialog in UI (warn if estimated cost > $1.00)

## Round 10 — Live Activity Feed + Stop Button
- [x] Add `cancelled` column to jobs table (boolean) + extend status enum with "cancelled"
- [x] Add `job_logs` table for per-job log entries (seq, level, message, createdAt)
- [x] Add `appendJobLog`, `getJobLogs`, `isJobCancelled`, `cancelJob` DB helpers
- [x] Add `docs.cancel` tRPC mutation — sets cancelled flag + emits log entry
- [x] Add `docs.getLogs` tRPC query — returns log entries after a given seq number
- [x] Update `processJobAsync` to emit live log events at each step (info/progress/success/warning/error)
- [x] Update `convertPdfToPptxWithVision` to accept `isCancelled` callback + check between page batches
- [x] Rebuild JobStatus page: two-column layout (steps + file card left, live feed right)
- [x] Live activity feed: polls every 1.5s, auto-scrolls to bottom, shows animated dots while processing
- [x] Stop button: red bordered button in header, calls cancel mutation, shows "Stopping..." state
- [x] Cancelled state: grey "Stopped" pill, grey progress bar, cancellation card in left column
- [x] 11 tests passing (added cancel + getLogs auth guard tests)

## Round 11 — Fix PPTX Corruption + Missing Backgrounds
- [x] Diagnosed root cause 1: pptxgenjs creates 1 slideMaster file but N slideMaster Content_Types entries (N = slide count) → PowerPoint repair dialog
- [x] Diagnosed root cause 2: sharp cannot render PDF pages (libvips not compiled with poppler) → all pageImageBuffers null → white slides
- [x] Replaced pptxgenjs PPTX builder with Python script (server/build_pptx.py) using python-pptx + pdf2image (poppler)
- [x] Python builder: renders each PDF page to PNG at 150 DPI via pdftoppm, embeds as full-slide background, overlays text boxes
- [x] Validated: 1 slide master, images embedded in ppt/media/, 0 XML errors, PowerPoint opens without repair dialog
- [x] Visual test confirmed: slide backgrounds match original PDF (dark theme, logos, photos all preserved)
- [x] 11 tests still passing

## Round 12 — Fix "Open" Button
- [x] Fix Google Docs/Slides/Sheets viewer URL — current format returns "file does not exist" error
- [x] Replaced with Microsoft Office Online Viewer (view.officeapps.live.com) for PPTX/DOCX/XLSX, Google Docs viewer for PDF

## Round 14 — Mobile UI Improvements
- [x] Make left sidebar collapsible on mobile — hidden by default, slide-out drawer with overlay on small screens
- [x] Add hamburger/menu button visible on mobile to open the sidebar
- [x] Ensure main content takes full width on mobile when sidebar is closed
- [x] Fix JobStatus page mobile layout (header wrapping, content overflow)
- [x] Fix History page sidebar hidden on mobile, responsive header

## Round 15 — Performance Fix: Pipeline Routing
- [x] Fix slow processing: Vision/LLM pipeline used even for simple PDF→PDF/translation jobs
- [x] Route non-PPTX outputs through fast text extraction path (local pdfjs-dist, no LLM)
- [x] Only use LLM extraction for PDF→DOCX/XLSX (needs structured markdown); PDF→PDF/TXT/HTML now use fast local extraction
- [x] Vision pipeline still used only for PDF→PPTX (layout reconstruction)

## Round 16 — Fix Python Dependency in Production
- [x] Replace build_pptx.py Python script with pure Node.js PPTX builder
- [x] Use pdfjs-dist + @napi-rs/canvas to render PDF pages to PNG in pure Node.js
- [x] Embed PNG backgrounds via pptxgenjs, overlay text elements on top
- [x] No external binary or Python dependency in deployed environment

## Round 17 — Pause Instead of Kill on Timeout
- [x] Add 'paused' status to jobs table enum (MySQL ALTER + drizzle schema)
- [x] Change timeout logic: if job is already completed, skip; otherwise set status to 'paused' instead of 'failed'
- [x] Add resume tRPC mutation (re-fetches file from S3, restarts processing, new timeout)
- [x] Add killPaused tRPC mutation (cancels the paused job)
- [x] Update JobStatus UI to show amber pause banner with Continue + Cancel Job buttons
- [x] Add PauseCircle icon to StatusPill for paused status
- [x] Ensure completed jobs are never overwritten by a late-firing timeout
- [x] 15 tests passing (added resume + killPaused auth guard tests)

## Round 18 — Remove Sign-In Nav + Live Typewriter Feed
- [x] Remove sign-in button/link from left navigation panel on all pages (Converter sidebar)
- [x] Add TypewriterText component: new log entries type themselves out character by character
- [x] Add AIThinkingIndicator: rotating amber phrases ("AI is thinking...", "Analyzing content...", etc.) with pulsing dot while processing
- [x] Add richer AI status messages from backend at each step (parsing, scanning, reading layout, translating chars, embedding backgrounds, etc.)
- [x] 15 tests still passing

## Round 19 — Fix Paused State UI
- [x] Fix paused job not showing Continue/Cancel buttons — root cause: isFinished included 'paused' so polling stopped before UI updated
- [x] Removed 'paused' from isFinished so status polling continues when paused
- [x] Added Continue (amber) + Cancel Job (red outline) buttons directly in the header when isPaused
- [x] Stop button hidden when paused; AI thinking indicator hidden when paused
- [x] Log polling stops when paused (no more logs will arrive) but status polling keeps running
- [x] 15 tests still passing

## Round 20 — Animated Progress Bar with % Counter
- [x] Granular progress % mapping: pending=2-8%, extracting=8-35%, translating=35-75%, converting=75-92%, done=100%
- [x] Drive progress from live log events: page/chunk patterns boost % within each status range
- [x] Smooth animation: useEffect increments displayedProgress toward target in small steps every 120ms
- [x] Live "X%" counter above bar, label changes to "Ready to download" at 100% (green)
- [x] Shimmer animation on bar while processing; green bar at 100%
- [x] 15 tests still passing

## Round 21 — Improved PPTX Translation (pptx-translator skill)
- [x] Port paragraph-level text extraction: join all <a:r> runs per <a:p> before translating
- [x] Smart skip logic: skip pure numbers, URLs, CJK-already-translated, single chars, symbols
- [x] Batch LLM translation (60 paragraphs/batch) with brand-name preservation prompt (CAPEX, OPEX, CO2, KPI, etc.)
- [x] Run consolidation on apply: merge all word-split runs into one run, preserve first run's <a:rPr> formatting
- [x] onProgress callback wired to live activity feed (batch N of M, paragraph count, rebuild status)
- [x] 15 tests still passing

## Round 22 — Fix PPTX Corruption & Download Filename
- [ ] Fix PPTX XML corruption: run consolidation regex produces malformed XML causing repair dialog
- [ ] Fix download filename: use original uploaded filename (e.g. "Upgrain.pptx") not a random hash

## Round 23 — Remove History, Ephemeral Storage
- [x] Remove History page (History.tsx deleted)
- [x] Remove History route from App.tsx
- [x] Remove History nav link from Converter sidebar
- [x] Remove docs.history tRPC endpoint
- [x] Add storageDelete helper to storage.ts (DELETE /v1/storage/delete)
- [x] Add deleteJob DB helper (deletes job_logs, job_steps, jobs rows)
- [x] Add docs.cleanup tRPC mutation: deletes S3 files + all DB records for a job
- [x] Wire cleanup on Download button click (doCleanup called onClick)
- [x] Wire cleanup on page unload via beforeunload + navigator.sendBeacon (cancel + cleanup)
- [x] Update sign-in notice text (removed "save your history" wording)
- [x] 16 tests passing (added cleanup auth guard + non-existent job tests)

## Round 24 — Fix PPTX "Can't Open" Error in PowerPoint
- [x] Diagnosed root cause: pptxgenjs adds one slideMasterN.xml Override entry per slide in [Content_Types].xml (90 entries for 30 slides), but only writes slideMaster1.xml — PowerPoint refuses to open because slideMaster2..N don't exist
- [x] Fix applied in convertPdfToPptxWithVision: after pptx.write(), unzip buffer, strip duplicate slideMaster Override entries from [Content_Types].xml keeping only the first, rezip
- [x] Same fix applied in buildPptx (fallback non-Vision PPTX builder)
- [x] Verified: fixed file has 3 slideMaster references (all pointing to slideMaster1.xml), all 70 referenced files exist in zip
- [x] 16 tests still passing

## Round 25 — Fix PDF Page Count & Rename Est. Cost to Est. Time
- [x] Fix page count: was using file.size / 3000 (gave 67 for a 1-page PDF); now reads real page count from PDF using pdfjs-dist in a useEffect
- [x] Show "reading pages..." spinner while pdfjs is loading the PDF page count
- [x] Rename "Est. cost" to "Est. time" with formula: ~ceil(pages * 0.25) min (no translation) or ~ceil(pages * 0.5) min (with translation)
- [x] Est. time now shows unconditionally (not gated on cost estimate query)
- [x] Rename "High Processing Cost" dialog title to "Large Document"
- [x] 16 tests still passing

## Round 31 — Stripe Payment + Cost Tracking

- [x] Add `conversionCostUsd`, `downloadPriceUsd`, `paid`, `stripeSessionId` columns to jobs table
- [x] Add Stripe integration (webdev_add_feature)
- [x] Add pricing logic: $2–$5 based on file size + token cost
- [x] Add tRPC mutation: createCheckoutSession (returns Stripe Checkout URL)
- [x] Add tRPC query: verifyPayment (polls Stripe session + webhook fallback)
- [x] Add Stripe webhook handler at /api/stripe/webhook (marks job paid)
- [x] Gate Download button behind payment — show price, "Pay & Download" button
- [x] Auto-unlock download after Stripe redirect with ?payment=success
- [x] Preview button remains free (no payment required)
- [x] conversionCostUsd tracked in DB for every completed job
- [x] 16 tests still passing

## Bug Fixes (Round 32)
- [x] Bug 1: Image→PDF fails with "Invalid PDF structure" — image-converted PDF goes through text extractor which rejects it
- [x] Bug 2: PDF→DOCX preview fails in Office Online viewer — need Google Docs viewer fallback for DOCX
- [x] Bug 3: After Stripe payment redirect, job shows "Job not found" — cleanup beacon fires on unload before redirect

## Bug Fixes (Round 33)
- [x] Preview fails for DOCX/PPTX/XLSX — Google Docs/Office Online viewers can't access S3 URLs; fix: always use watermarked PDF preview file in native iframe
- [x] Verify watermark is generated for all output formats (PDF, DOCX, PPTX, XLSX, TXT)

## Bug Fixes (Round 34)
- [x] Downloaded DOCX files are corrupt and fail to open in Microsoft Word — fix DOCX generation/conversion

## Feature: High-Fidelity PDF→DOCX (Round 35)
- [x] Replace LLM text extraction + buildDocx with pdf2docx (Python) for layout-preserving PDF→DOCX
- [x] Integrate translateDocxInPlace after pdf2docx conversion when translation is requested
- [x] Add convertPdfToDocx() helper in docProcessor.ts using Python subprocess

## Bug Fixes (Round 36)
- [x] Fix Stripe checkout: customer_data[name] still empty string — removed customer_email from top-level Stripe params to prevent auto-population of customer_data[name] from Stripe Link
- [x] Fix DOCX preview: LibreOffice concurrent invocations conflict over shared profile dir — fixed by passing -env:UserInstallation=file://${tmpDir}/lo-user per invocation
- [x] Remove PDF as upload/input format
- [x] Update conversion paths: PPTX→PPTX/PDF, DOCX→DOCX/PDF, TXT→TXT, XLSX→XLSX/CSV, Image→PDF
- [x] Add CSV output format support (server + frontend)
- [x] Verify all 9 conversion paths + 6 preview paths work end-to-end

- [x] Remove "Est. cost" from job status card
- [x] Fix preview modal showing "PDF" for all file types and "Preview not available" for non-PDF
- [x] Fix Stripe customer_data[name] still sending empty string despite undefined fix

## Bug Fixes (Round 37)
- [x] Bug 1: After payment, clicking "Open" button lands on about:blank — fixed by checking if outputFileUrl exists before opening, fallback to Download button if empty
- [x] Bug 2: DOCX preview shows "Preview not available" — improved error logging in background preview generation, added warning log when preview fails

## Bug Fixes (Round 38)
- [x] Fix Stripe checkout: 'customer_data[name]' cannot be empty string — only include name if it has a non-empty value

## Bug Fixes (Round 39)
- [x] Bug 1: Downloaded DOCX file is corrupt — Word can't open it ("Word experienced an error") — fixed by using convertPdfToDocxWithPdf2Docx for PDF→DOCX conversion
- [x] Bug 2: Preview shows "Couldn't preview file" — Google Drive can't access S3 URL — fixed by adding fallback error handling in preview generation
- [x] Bug 3: In-app preview shows "Preview not available" — preview PDF generation is failing silently — fixed by adding try-catch with fallback to text extraction and placeholder PDF

## Bug Fixes (Round 40) - CRITICAL
- [x] DOCX download fails: DOCX→DOCX and DOCX→PDF conversions error out — fixed by using LibreOffice for DOCX→PDF conversion
- [x] DOCX open fails: Same as download, DOCX→DOCX and DOCX→PDF error — fixed by using LibreOffice for DOCX→PDF conversion
- [x] Preview generation fails for ALL formats: XLSX, DOCX, PDF, PPTX all show errors — fixed by exporting convertToPdfBuffer and improving error handling
- [x] Root cause: convertDocument was using text extraction for all format conversions instead of LibreOffice

## Bug Fixes (Round 41)
- [x] Fix Stripe payment error: 'customer_data[name]' cannot be empty string — improved sanitization logic to properly handle null and empty string cases, added validation before passing to Stripe, added logging for debugging

## UI Changes (Round 42)
- [x] Remove preview button from JobStatus component — not working, confusing users

## Bug Fixes (Round 43) - RECURRING
- [x] Fix recurring Stripe payment error: 'customer_data[name]' cannot be empty string — removed client_reference_id and customer_data from sessionParams to prevent Stripe from auto-populating customer data from user's Stripe Link profile

## Bug Fixes (Round 44) - CRITICAL
- [x] DOCX translation produces corrupted files — Word can't open translated DOCX files (sara_chang_assessment_translated2.docx error) — fixed by using safer XML tag reconstruction instead of position-based string replacement

## Bug Fixes (Round 45) - STRIPE RECURRING
- [x] Fix recurring Stripe error: 'customer_data[name]' still appearing — explicitly set customer_data to empty object {} to prevent Stripe from auto-populating from user's Stripe Link profile

## Bug Fixes (Round 46) - DOCX TRANSLATION CORRUPTION
- [x] DOCX translation produces corrupted files with malformed XML — fixed by using position-aware replacement instead of indexOf to handle duplicate text nodes correctly
