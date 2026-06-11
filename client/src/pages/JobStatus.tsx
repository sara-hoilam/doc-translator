import { useParams, Link } from "wouter";
import { useEffect, useRef, useState, useCallback } from "react";
import { useLocation } from "wouter";
import { trpc } from "@/lib/trpc";
import { toast } from "sonner";
import {
  CheckCircle2, XCircle, Loader2, Clock, Download,
  FileText, Globe, RefreshCw, Upload, ArrowLeft,
  ExternalLink, Presentation, FileSpreadsheet,
  Square, AlertTriangle, Info, ChevronRight, PauseCircle, Play, Trash2, Eye,
} from "lucide-react";

// ── Typewriter text component ───────────────────────────────────────────────
function TypewriterText({ text, isNew }: { text: string; isNew: boolean }) {
  const [displayed, setDisplayed] = useState(isNew ? "" : text);
  const [done, setDone] = useState(!isNew);

  useEffect(() => {
    if (!isNew) { setDisplayed(text); setDone(true); return; }
    let i = 0;
    setDisplayed("");
    setDone(false);
    const speed = text.length > 60 ? 10 : text.length > 30 ? 16 : 22;
    const timer = setInterval(() => {
      i++;
      setDisplayed(text.slice(0, i));
      if (i >= text.length) { clearInterval(timer); setDone(true); }
    }, speed);
    return () => clearInterval(timer);
  }, [text, isNew]);

  return (
    <span>
      {displayed}
      {!done && <span className="inline-block w-0.5 h-3 bg-current ml-0.5 animate-pulse" />}
    </span>
  );
}

// ── AI thinking phrases ──────────────────────────────────────────────────────
const AI_THINKING_PHRASES = [
  "AI is thinking...",
  "Processing document...",
  "Analyzing content...",
  "Working on it...",
  "Almost there...",
  "Crunching data...",
  "Translating content...",
  "Optimizing output...",
];

function AIThinkingIndicator() {
  const [phraseIdx, setPhraseIdx] = useState(0);
  useEffect(() => {
    const t = setInterval(() => setPhraseIdx(i => (i + 1) % AI_THINKING_PHRASES.length), 2800);
    return () => clearInterval(t);
  }, []);
  return (
    <div className="flex items-center gap-2.5 px-3 py-2.5 rounded-lg bg-amber-50 border border-amber-100">
      <div className="flex gap-1 shrink-0">
        <span className="w-1.5 h-1.5 rounded-full bg-amber-400 animate-bounce" style={{ animationDelay: "0ms" }} />
        <span className="w-1.5 h-1.5 rounded-full bg-amber-400 animate-bounce" style={{ animationDelay: "150ms" }} />
        <span className="w-1.5 h-1.5 rounded-full bg-amber-400 animate-bounce" style={{ animationDelay: "300ms" }} />
      </div>
      <span className="text-amber-700 text-[11px] font-sans italic">{AI_THINKING_PHRASES[phraseIdx]}</span>
      <span className="ml-auto shrink-0">
        <span className="inline-block w-2 h-2 rounded-full bg-amber-400 animate-ping opacity-75" />
      </span>
    </div>
  );
}

const STEP_LABELS: Record<string, string> = {
  upload: "File Upload",
  extract: "Text Extraction",
  translate: "Translation",
  convert: "Format Conversion",
};

const STEP_ICONS: Record<string, React.ElementType> = {
  upload: Upload,
  extract: FileText,
  translate: Globe,
  convert: RefreshCw,
};

const FORMAT_LABELS: Record<string, string> = {
  pdf: "PDF", docx: "Word (DOCX)", pptx: "PowerPoint (PPTX)",
  xlsx: "Excel (XLSX)", txt: "Plain Text", html: "HTML",
};

const FORMAT_ICONS: Record<string, React.ElementType> = {
  pdf: FileText, docx: FileText, pptx: Presentation,
  xlsx: FileSpreadsheet, txt: FileText, html: Globe,
};

const FORMAT_COLORS: Record<string, string> = {
  pdf: "text-red-500", docx: "text-blue-500", pptx: "text-orange-500",
  xlsx: "text-green-500", txt: "text-gray-500", html: "text-purple-500",
};

function getOnlineViewerUrl(format: string, fileUrl: string): string {
  const encoded = encodeURIComponent(fileUrl);
  // Office formats: Microsoft Office Online Viewer (reliable, handles DOCX/PPTX/XLSX natively)
  if (["docx", "pptx", "xlsx"].includes(format)) {
    return `https://view.officeapps.live.com/op/view.aspx?src=${encoded}`;
  }
  // PDF: open the URL directly — all browsers have a native PDF viewer
  // TXT and others: direct link is the simplest and most reliable option
  return fileUrl;
}

// ── Log level config ─────────────────────────────────────────────────────────
const LOG_LEVEL_CONFIG = {
  info:     { icon: Info,          color: "text-gray-500",  bg: "bg-gray-50",    dot: "bg-gray-300"   },
  progress: { icon: ChevronRight,  color: "text-amber-600", bg: "bg-amber-50",   dot: "bg-amber-300"  },
  success:  { icon: CheckCircle2,  color: "text-green-600", bg: "bg-green-50",   dot: "bg-green-400"  },
  warning:  { icon: AlertTriangle, color: "text-amber-600", bg: "bg-amber-50",   dot: "bg-amber-400"  },
  error:    { icon: XCircle,       color: "text-red-500",   bg: "bg-red-50",     dot: "bg-red-400"    },
} as const;

type LogLevel = keyof typeof LOG_LEVEL_CONFIG;

interface LogEntry {
  id: number;
  seq: number;
  level: LogLevel;
  message: string;
  createdAt: number;
}

export default function JobStatus() {
  const params = useParams<{ id: string }>();
  const jobId = parseInt(params.id ?? "0");

  // ── Job status polling ───────────────────────────────────────────────────
  const { data, isLoading } = trpc.docs.status.useQuery(
    { jobId },
    {
      enabled: !!jobId,
      refetchInterval: (query) => {
        const status = query.state.data?.job?.status;
        if (status === "done" || status === "error" || status === "cancelled") return false;
        // Keep polling when paused so the UI can detect the paused state
        if (status === "paused") return 2000;
        return 2000;
      },
    }
  );

  // ── Live log polling ──────────────────────────────────────────────────────
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [lastSeq, setLastSeq] = useState(0);
  const [newLogSeqs, setNewLogSeqs] = useState<Set<number>>(new Set());
  const logEndRef = useRef<HTMLDivElement>(null);
  const isFinished = ["done", "error", "cancelled"].includes(data?.job?.status ?? "");

  const isLogPollingDone = isFinished || (data?.job?.status === "paused");
  const { data: logData } = trpc.docs.getLogs.useQuery(
    { jobId, afterSeq: lastSeq },
    {
      enabled: !!jobId,
      refetchInterval: isLogPollingDone ? false : 1500,
    }
  );

  useEffect(() => {
    if (!logData?.logs?.length) return;
    setLogs(prev => {
      const newLogs = logData.logs.filter(l => !prev.some(p => p.seq === l.seq));
      if (!newLogs.length) return prev;
      const merged = [...prev, ...newLogs].sort((a, b) => a.seq - b.seq);
      setLastSeq(merged[merged.length - 1].seq);
      // Mark newly arrived entries for typewriter animation
      setNewLogSeqs(s => {
        const next = new Set(s);
        newLogs.forEach(l => next.add(l.seq));
        return next;
      });
      return merged;
    });
  }, [logData]);

  // Auto-scroll log feed to bottom
  useEffect(() => {
    logEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [logs]);

  // ── Cancel mutation ──────────────────────────────────────────────────────
  const utils = trpc.useUtils();
  const [cancelling, setCancelling] = useState(false);
  const cancelMutation = trpc.docs.cancel.useMutation({
    onSuccess: () => {
      utils.docs.status.invalidate({ jobId });
      setCancelling(false);
    },
    onError: () => setCancelling(false),
  });

  const handleCancel = () => {
    setCancelling(true);
    cancelMutation.mutate({ jobId });
  };

  // ── Resume mutation (for paused jobs) ────────────────────────────────────
  const [resuming, setResuming] = useState(false);
  const resumeMutation = trpc.docs.resume.useMutation({
    onSuccess: () => {
      utils.docs.status.invalidate({ jobId });
      setResuming(false);
    },
    onError: () => setResuming(false),
  });

  const handleResume = () => {
    setResuming(true);
    resumeMutation.mutate({ jobId });
  };

  // ── Kill paused job mutation ──────────────────────────────────────────────
  const [killing, setKilling] = useState(false);
  const killMutation = trpc.docs.killPaused.useMutation({
    onSuccess: () => {
      utils.docs.status.invalidate({ jobId });
      setKilling(false);
    },
    onError: () => setKilling(false),
  });

  const handleKill = () => {
    setKilling(true);
    killMutation.mutate({ jobId });
  };

  const [, setLocation] = useLocation();

  // ── Ephemeral cleanup mutation ────────────────────────────────────────────
  // Deletes S3 files + DB records so nothing is stored after the user is done.
  const cleanupMutation = trpc.docs.cleanup.useMutation();

  const doCleanup = () => {
    if (jobId) cleanupMutation.mutate({ jobId });
  };

  // ── Filename dialog state ───────────────────────────────────────────────
  const [filenameDialogOpen, setFilenameDialogOpen] = useState(false);
  const [customFilename, setCustomFilename] = useState("");

  const openFilenameDialog = () => {
    if (!job) return;
    const base = job.originalFileName.replace(/\.[^.]+$/, '');
    const ext = job.outputFormat ?? job.originalFormat;
    setCustomFilename(`${base}_translated.${ext}`);
    setFilenameDialogOpen(true);
  };

  const [isDownloading, setIsDownloading] = useState(false);

  const doDownload = async () => {
    if (!job) return;
    setFilenameDialogOpen(false);
    setIsDownloading(true);
    try {
      // Use server-side download proxy — avoids CORS issues with GCS signed URLs
      // and ensures cleanup only happens after the file is fully delivered.
      const proxyUrl = `/api/download/${job.id}`;
      const response = await fetch(proxyUrl);
      if (!response.ok) throw new Error(`Download failed: ${response.status}`);
      const blob = await response.blob();
      const objectUrl = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = objectUrl;
      // Use filename from Content-Disposition if available, otherwise build one
      const ext = job.outputFormat ?? job.originalFormat;
      const filename = customFilename.trim() || `download.${ext}`;
      a.download = filename.endsWith(`.${ext}`) ? filename : `${filename}.${ext}`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(objectUrl), 5000);
    } catch (err: any) {
      toast.error("Download failed. Please try again.");
      console.error("[Download]", err);
    } finally {
      setIsDownloading(false);
    }
  };

  // Cancel + cleanup when the user closes the tab or navigates away
  useEffect(() => {
    const handleUnload = () => {
      if (!jobId) return;
      const body = JSON.stringify({ json: { jobId } });
      const blob = new Blob([body], { type: "application/json" });
      navigator.sendBeacon("/api/trpc/docs.cancel", blob);
      navigator.sendBeacon("/api/trpc/docs.cleanup", blob);
    };
    window.addEventListener("beforeunload", handleUnload);
    return () => window.removeEventListener("beforeunload", handleUnload);
  }, [jobId]);

  const job = data?.job;
  const steps = data?.steps ?? [];
  const isPaused = job?.status === "paused";
  const isProcessing = job && !isFinished && !isPaused;

  const getStepStatus = (stepName: string) => steps.find(s => s.step === stepName);

  // ── Granular progress derived from job status + live log events ──────────
  const getTargetProgress = (): number => {
    if (!job) return 0;
    if (job.status === "done") return 100;
    if (job.status === "error") return 100;
    if (job.status === "cancelled") return 100;

    // Base range per status
    const baseMap: Record<string, [number, number]> = {
      pending:    [2,   8],
      extracting: [8,  35],
      translating:[35, 75],
      converting: [75, 92],
      paused:     [50, 50],
    };
    const [base, cap] = baseMap[job.status] ?? [2, 8];

    // Scan logs for page-level progress hints
    // e.g. "Analyzed pages 1–3 of 12" or "Page 4 of 12"
    let logBoost = 0;
    for (const log of logs) {
      // Vision: "Analyzed pages X–Y of N"
      const visionMatch = log.message.match(/Analyzed pages? (\d+)[–-](\d+) of (\d+)/i);
      if (visionMatch) {
        const done = parseInt(visionMatch[2]);
        const total = parseInt(visionMatch[3]);
        if (total > 0) logBoost = Math.max(logBoost, done / total);
      }
      // Translation chunks: "Translating chunk X of N"
      const chunkMatch = log.message.match(/chunk (\d+) of (\d+)/i);
      if (chunkMatch) {
        const done = parseInt(chunkMatch[1]);
        const total = parseInt(chunkMatch[2]);
        if (total > 0) logBoost = Math.max(logBoost, done / total);
      }
      // Page-level: "Page X of N"
      const pageMatch = log.message.match(/[Pp]age (\d+) of (\d+)/);
      if (pageMatch) {
        const done = parseInt(pageMatch[1]);
        const total = parseInt(pageMatch[2]);
        if (total > 0) logBoost = Math.max(logBoost, done / total);
      }
      // "Uploading output file" → nearly done
      if (log.message.toLowerCase().includes("uploading output")) logBoost = Math.max(logBoost, 0.95);
    }

    const range = cap - base;
    return Math.round(base + range * logBoost);
  };

  // Smoothly animate displayed progress toward target
  const [displayedProgress, setDisplayedProgress] = useState(0);
  useEffect(() => {
    const target = getTargetProgress();
    if (target <= displayedProgress) {
      // Always allow jump to 100 immediately
      if (target === 100) setDisplayedProgress(100);
      return;
    }
    // Animate in small steps toward target
    const step = Math.max(1, Math.round((target - displayedProgress) / 8));
    const timer = setTimeout(() => {
      setDisplayedProgress(prev => Math.min(target, prev + step));
    }, 120);
    return () => clearTimeout(timer);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [job?.status, logs, displayedProgress]);

  // ── Loading ──────────────────────────────────────────────────────────────
  if (isLoading) {
    return (
      <div className="min-h-screen flex bg-[#f5f4f0]">
        <SidebarShell />
        <main className="flex-1 flex items-center justify-center">
          <div className="text-center">
            <Loader2 className="h-8 w-8 animate-spin text-amber-500 mx-auto mb-3" />
            <p className="text-gray-500 text-sm">Loading job status...</p>
          </div>
        </main>
      </div>
    );
  }

  if (!job) {
    return (
      <div className="min-h-screen flex bg-[#f5f4f0]">
        <SidebarShell />
        <main className="flex-1 flex items-center justify-center">
          <div className="text-center">
            <XCircle className="h-10 w-10 text-red-400 mx-auto mb-3" />
            <p className="font-semibold text-gray-800 mb-1">Job not found</p>
            <p className="text-sm text-gray-500 mb-5">This job may not exist or you don't have access.</p>
            <Link href="/">
              <button className="px-4 py-2 rounded-xl bg-gray-800 text-white text-sm font-medium hover:bg-gray-900 transition-colors">
                New Document
              </button>
            </Link>
          </div>
        </main>
      </div>
    );
  }

  const progress = displayedProgress;
  const FormatIcon = FORMAT_ICONS[job.outputFormat ?? job.originalFormat] ?? FileText;
  const formatColor = FORMAT_COLORS[job.outputFormat ?? job.originalFormat] ?? "text-gray-500";

  return (
    <div className="min-h-screen flex bg-[#f5f4f0]">
      {/* ── Sidebar ─────────────────────────────────────────────────────── */}
      <SidebarShell />

      {/* ── Main ────────────────────────────────────────────────────────── */}
      <main className="flex-1 flex flex-col min-w-0">
        {/* Header bar */}
        <div className="px-4 sm:px-8 pt-5 sm:pt-7 pb-4">
          <Link href="/">
            <button className="inline-flex items-center gap-1.5 text-sm text-gray-500 hover:text-gray-800 transition-colors mb-4">
              <ArrowLeft className="h-3.5 w-3.5" />
              New document
            </button>
          </Link>
          <div className="flex items-start justify-between gap-2">
            <div className="min-w-0">
              <h1 className="text-lg sm:text-xl font-bold text-gray-900">Processing Status</h1>
              <p className="text-xs sm:text-sm text-gray-500 mt-0.5">Job #{job.id} · {new Date(job.createdAt).toLocaleString()}</p>
            </div>
            <div className="flex items-center gap-2 shrink-0">
              {/* Paused — show Continue + Cancel buttons prominently in header */}
              {isPaused && (
                <>
                  <button
                    onClick={handleResume}
                    disabled={resuming || killing}
                    className="inline-flex items-center gap-1.5 px-2.5 sm:px-3 py-1.5 rounded-xl bg-amber-500 hover:bg-amber-600 text-white text-xs sm:text-sm font-semibold transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    {resuming ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : <Play className="h-3.5 w-3.5" />}
                    {resuming ? "Resuming..." : "Continue"}
                  </button>
                  <button
                    onClick={handleKill}
                    disabled={resuming || killing}
                    className="inline-flex items-center gap-1.5 px-2.5 sm:px-3 py-1.5 rounded-xl border border-red-200 bg-white hover:bg-red-50 text-red-600 text-xs sm:text-sm font-medium transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    {killing ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : <Trash2 className="h-3.5 w-3.5" />}
                    {killing ? "Cancelling..." : "Cancel Job"}
                  </button>
                </>
              )}
              {/* Stop button — only shown while actively processing (not paused) */}
              {isProcessing && !isPaused && (
                <button
                  onClick={handleCancel}
                  disabled={cancelling}
                  className="inline-flex items-center gap-1.5 px-2.5 sm:px-3 py-1.5 rounded-xl border border-red-200 bg-white hover:bg-red-50 text-red-600 text-xs sm:text-sm font-medium transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  {cancelling ? (
                    <Loader2 className="h-3.5 w-3.5 animate-spin" />
                  ) : (
                    <Square className="h-3.5 w-3.5 fill-red-500" />
                  )}
                  {cancelling ? "Stopping..." : "Stop"}
                </button>
              )}
              <StatusPill status={job.status} />
            </div>
          </div>
        </div>

        {/* Responsive layout: stacked on mobile, side-by-side on lg+ */}
        <div className="flex-1 px-4 sm:px-8 pb-6 sm:pb-8 flex flex-col lg:flex-row gap-4 sm:gap-5 items-start">
          {/* ── Left column ─────────────────────────────────────────────── */}
          <div className="flex flex-col gap-4 w-full lg:w-72 lg:shrink-0">
            {/* File card */}
            <div className="bg-white rounded-2xl border border-[#e5e3dc] p-5">
              <div className="flex items-center gap-3 mb-4">
                <div className="w-10 h-10 rounded-xl bg-amber-50 flex items-center justify-center shrink-0">
                  <FormatIcon className={`h-5 w-5 ${formatColor}`} />
                </div>
                <div className="flex-1 min-w-0">
                  <p className="font-bold text-gray-900 truncate text-sm">{job.originalFileName}</p>
                  <p className="text-xs text-gray-500 mt-0.5">{new Date(job.createdAt).toLocaleString()}</p>
                </div>
              </div>

              {/* Progress bar with % counter */}
              <div className="mb-4">
                <div className="flex items-center justify-between mb-1.5">
                  <span className="text-[11px] font-medium text-gray-500">
                    {job.status === "done" ? "Ready to download" :
                     job.status === "error" ? "Failed" :
                     job.status === "cancelled" ? "Cancelled" :
                     job.status === "paused" ? "Paused" :
                     "Converting..."}
                  </span>
                  <span
                    className={`text-[13px] font-bold tabular-nums transition-all duration-300 ${
                      job.status === "done" ? "text-green-600" :
                      job.status === "error" ? "text-red-500" :
                      job.status === "cancelled" ? "text-gray-400" :
                      "text-amber-600"
                    }`}
                  >
                    {progress}%
                  </span>
                </div>
                <div className="w-full bg-[#f0ede6] rounded-full h-2.5 overflow-hidden">
                  <div
                    className={`h-2.5 rounded-full transition-all duration-500 ease-out ${
                      job.status === "error" ? "bg-red-400" :
                      job.status === "cancelled" ? "bg-gray-400" :
                      job.status === "done" ? "bg-green-500" :
                      "bg-amber-500"
                    } ${
                      isProcessing ? "relative" : ""
                    }`}
                    style={{ width: `${progress}%` }}
                  >
                    {/* Shimmer animation while processing */}
                    {isProcessing && progress > 0 && progress < 100 && (
                      <div
                        className="absolute inset-0 rounded-full"
                        style={{
                          background: "linear-gradient(90deg, transparent 0%, rgba(255,255,255,0.4) 50%, transparent 100%)",
                          animation: "shimmer 1.8s infinite",
                        }}
                      />
                    )}
                  </div>
                </div>
              </div>

              {/* Metadata grid */}
              <div className="grid grid-cols-2 gap-3 text-sm">
                {[
                  { label: "Input", value: job.originalFormat.toUpperCase() },
                  { label: "Output", value: FORMAT_LABELS[job.outputFormat ?? ""] ?? (job.outputFormat ?? "—").toUpperCase() },
                  { label: "Language", value: job.targetLanguageName ?? "Original" },
                  { label: "Pages", value: job.pageCount?.toString() ?? "—" },
                ].map(({ label, value }) => (
                  <div key={label}>
                    <p className="text-[10px] text-gray-400 uppercase tracking-wider mb-0.5">{label}</p>
                    <p className="font-semibold text-gray-800 text-sm truncate">{value}</p>
                  </div>
                ))}
              </div>


            </div>

            {/* Processing Steps */}
            <div className="bg-white rounded-2xl border border-[#e5e3dc] p-5">
              <h2 className="font-semibold text-gray-900 mb-3 text-sm">Steps</h2>
              <div className="space-y-2">
                {(["upload", "extract", "translate", "convert"] as const).map((stepName) => {
                  const step = getStepStatus(stepName);
                  if (!step) return null;
                  const Icon = STEP_ICONS[stepName] ?? FileText;

                  const stateStyles = {
                    running: { wrap: "bg-amber-50 border-amber-100", dot: "bg-amber-100", icon: "text-amber-600", label: "text-amber-600" },
                    done:    { wrap: "bg-[#f5f4f0] border-[#e5e3dc]", dot: "bg-green-100", icon: "text-green-600", label: "text-green-600" },
                    error:   { wrap: "bg-red-50 border-red-100", dot: "bg-red-100", icon: "text-red-500", label: "text-red-500" },
                    pending: { wrap: "bg-[#f5f4f0] border-transparent", dot: "bg-[#e5e3dc]", icon: "text-gray-400", label: "text-gray-400" },
                  }[step.status] ?? { wrap: "bg-[#f5f4f0] border-transparent", dot: "bg-[#e5e3dc]", icon: "text-gray-400", label: "text-gray-400" };

                  return (
                    <div key={stepName} className={`flex items-start gap-2.5 rounded-xl p-2.5 border ${stateStyles.wrap} transition-colors`}>
                      <div className={`mt-0.5 rounded-full p-1.5 shrink-0 ${stateStyles.dot}`}>
                        {step.status === "running" ? (
                          <Loader2 className={`h-3 w-3 ${stateStyles.icon} animate-spin`} />
                        ) : step.status === "done" ? (
                          <CheckCircle2 className={`h-3 w-3 ${stateStyles.icon}`} />
                        ) : step.status === "error" ? (
                          <XCircle className={`h-3 w-3 ${stateStyles.icon}`} />
                        ) : (
                          <Clock className={`h-3 w-3 ${stateStyles.icon}`} />
                        )}
                      </div>
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center gap-1.5">
                          <Icon className="h-3 w-3 text-gray-400 shrink-0" />
                          <span className="text-xs font-medium text-gray-800">{STEP_LABELS[stepName]}</span>
                          <span className={`text-[10px] capitalize ml-auto font-medium ${stateStyles.label}`}>
                            {step.status === "running" ? "Running" : step.status === "done" ? "Done" : step.status === "error" ? "Error" : "Waiting"}
                          </span>
                        </div>
                        {step.message && (
                          <p className="text-[10px] text-gray-500 mt-0.5 leading-tight">{step.message}</p>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Paused — ask user to continue or cancel */}
            {isPaused && (
              <div className="bg-amber-50 rounded-2xl border border-amber-200 p-4">
                <div className="flex items-start gap-2.5 mb-3">
                  <PauseCircle className="h-4 w-4 text-amber-500 shrink-0 mt-0.5" />
                  <div>
                    <p className="font-semibold text-amber-800 mb-1 text-sm">Job paused</p>
                    <p className="text-xs text-amber-700">
                      Processing has been running for a while. Would you like to continue or cancel?
                    </p>
                  </div>
                </div>
                <div className="flex gap-2">
                  <button
                    onClick={handleResume}
                    disabled={resuming || killing}
                    className="flex-1 flex items-center justify-center gap-1.5 bg-amber-500 hover:bg-amber-600 text-white text-xs font-semibold py-2 rounded-xl transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    {resuming ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : <Play className="h-3.5 w-3.5" />}
                    {resuming ? "Resuming..." : "Continue"}
                  </button>
                  <button
                    onClick={handleKill}
                    disabled={resuming || killing}
                    className="flex-1 flex items-center justify-center gap-1.5 border border-red-200 bg-white hover:bg-red-50 text-red-600 text-xs font-semibold py-2 rounded-xl transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    {killing ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : <Trash2 className="h-3.5 w-3.5" />}
                    {killing ? "Cancelling..." : "Cancel Job"}
                  </button>
                </div>
              </div>
            )}

            {/* Error */}
            {job.status === "error" && job.errorMessage && (
              <div className="bg-red-50 rounded-2xl border border-red-200 p-4">
                <div className="flex items-start gap-2.5">
                  <XCircle className="h-4 w-4 text-red-400 shrink-0 mt-0.5" />
                  <div>
                    <p className="font-semibold text-red-800 mb-1 text-sm">Processing failed</p>
                    <p className="text-xs text-red-700">{job.errorMessage}</p>
                  </div>
                </div>
              </div>
            )}

            {/* Cancelled */}
            {job.status === "cancelled" && (
              <div className="bg-gray-50 rounded-2xl border border-gray-200 p-4">
                <div className="flex items-start gap-2.5">
                  <Square className="h-4 w-4 text-gray-400 shrink-0 mt-0.5 fill-gray-400" />
                  <div>
                    <p className="font-semibold text-gray-700 mb-1 text-sm">Processing stopped</p>
                    <p className="text-xs text-gray-500">You cancelled this job before it completed.</p>
                  </div>
                </div>
              </div>
            )}

            {/* Download */}
            {job.status === "done" && job.outputFileUrl && (
              <>

                <div className="bg-white rounded-2xl border border-[#e5e3dc] p-4">
                  <div className="flex items-center gap-2.5 mb-3">
                    <div className="w-8 h-8 rounded-xl bg-green-50 flex items-center justify-center">
                      <CheckCircle2 className="h-4 w-4 text-green-500" />
                    </div>
                    <div>
                      <p className="font-semibold text-gray-900 text-sm">Complete!</p>
                      <p className="text-xs text-gray-500">Your document is ready.</p>
                    </div>
                  </div>

                  {job.previewFileUrl ? (
                    <button
                      onClick={() => {
                        const url = job.previewFileUrl!;
                        const ext = (url.split("?")[0].split(".").pop() ?? "").toLowerCase();
                        if (ext === "pptx") {
                          const embedUrl = `https://view.officeapps.live.com/op/embed.aspx?src=${encodeURIComponent(url)}`;
                          window.open(embedUrl, "docPreview", "width=960,height=720,menubar=no,toolbar=no,location=no,status=no");
                        } else {
                          window.open(url, "_blank", "noopener,noreferrer");
                        }
                      }}
                      className="w-full flex items-center justify-center gap-1.5 border border-[#e5e3dc] bg-[#f9f8f5] hover:bg-[#f0efe9] text-gray-700 text-xs font-medium py-2 rounded-xl transition-colors mb-2.5"
                    >
                      <Eye className="h-3.5 w-3.5" />
                      Preview ({job.previewPageCount ?? 3} pages)
                    </button>
                  ) : (
                    <div className="w-full flex items-center justify-center gap-1.5 text-xs text-gray-400 py-2 mb-2.5">
                      <Loader2 className="h-3 w-3 animate-spin" />
                      Generating preview…
                    </div>
                  )}

                  <div className="flex gap-2">
                    <button
                      onClick={openFilenameDialog}
                      disabled={isDownloading}
                      className="flex-1 flex items-center justify-center gap-1.5 bg-gray-800 hover:bg-gray-900 disabled:opacity-60 text-white text-xs font-medium py-2 rounded-xl transition-colors"
                    >
                      <Download className="h-3.5 w-3.5" />
                      {isDownloading ? "Downloading..." : "Download"}
                    </button>
                    {job.outputFileUrl ? (
                      <a
                        href={getOnlineViewerUrl(job.outputFormat ?? job.originalFormat, job.outputFileUrl)}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="flex items-center justify-center gap-1.5 border border-[#e5e3dc] bg-white hover:bg-[#f5f4f0] text-gray-700 text-xs font-medium px-3 py-2 rounded-xl transition-colors"
                      >
                        <ExternalLink className="h-3.5 w-3.5" />
                        Open
                      </a>
                    ) : null}
                  </div>
                </div>
              </>
            )}
          </div>

          {/* ── Right column: Live Activity Feed ────────────────────────── */}
          <div className="flex-1 min-w-0 w-full">
            <div className="bg-white rounded-2xl border border-[#e5e3dc] flex flex-col" style={{ minHeight: "320px", maxHeight: "calc(100vh - 200px)" }}>
              {/* Feed header */}
              <div className="px-5 pt-4 pb-3 border-b border-[#f0ede6] flex items-center justify-between shrink-0">
                <div className="flex items-center gap-2">
                  <div className={`w-2 h-2 rounded-full ${isProcessing ? "bg-amber-400 animate-pulse" : job.status === "done" ? "bg-green-400" : "bg-gray-300"}`} />
                  <h2 className="font-semibold text-gray-900 text-sm">Live Activity</h2>
                </div>
                <span className="text-xs text-gray-400">{logs.length} event{logs.length !== 1 ? "s" : ""}</span>
              </div>

              {/* Log entries */}
              <div className="flex-1 overflow-y-auto px-4 py-3 space-y-1.5 font-mono text-xs">
                {logs.length === 0 && (
                  <div className="flex items-center justify-center h-32 text-gray-400 text-xs font-sans">
                    {isProcessing ? (
                      <div className="flex items-center gap-2">
                        <Loader2 className="h-4 w-4 animate-spin text-amber-400" />
                        Waiting for activity...
                      </div>
                    ) : (
                      <span>No activity recorded</span>
                    )}
                  </div>
                )}

                {logs.map((entry) => {
                  // When job is done, upgrade info/progress entries to success styling
                  const effectiveLevel = (job?.status === "done" && (entry.level === "info" || entry.level === "progress"))
                    ? "success"
                    : entry.level;
                  const cfg = LOG_LEVEL_CONFIG[effectiveLevel] ?? LOG_LEVEL_CONFIG.info;
                  const Icon = cfg.icon;
                  const isNew = newLogSeqs.has(entry.seq);
                  return (
                    <div
                      key={entry.seq}
                      className={`flex items-start gap-2.5 rounded-lg px-3 py-2 ${cfg.bg}`}
                    >
                      <Icon className={`h-3.5 w-3.5 shrink-0 mt-0.5 ${cfg.color}`} />
                      <span className={`flex-1 leading-relaxed ${cfg.color} break-words`}>
                        <TypewriterText text={entry.message} isNew={isNew} />
                      </span>
                      <span className="text-gray-300 shrink-0 text-[10px] mt-0.5 font-sans">
                        {new Date(entry.createdAt).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" })}
                      </span>
                    </div>
                  );
                })}

                {/* AI thinking indicator while processing */}
                {isProcessing && <AIThinkingIndicator />}

                <div ref={logEndRef} />
              </div>
            </div>
          </div>
        </div>

        {/* Footer actions */}
        <div className="border-t border-[#e5e3dc] bg-white/60 px-4 sm:px-8 py-4 flex gap-3">
          <Link href="/">
            <button className="flex items-center gap-2 border border-[#e5e3dc] bg-white hover:bg-[#f5f4f0] text-gray-700 text-sm font-medium px-4 py-2.5 rounded-xl transition-colors">
              <Upload className="h-4 w-4" />
              New Document
            </button>
          </Link>
        </div>
      </main>

      {/* ── Filename Dialog ────────────────────────────────────────────── */}
      {filenameDialogOpen && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 px-4"
          onClick={(e) => { if (e.target === e.currentTarget) setFilenameDialogOpen(false); }}
        >
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-sm p-6">
            <h3 className="text-base font-semibold text-gray-900 mb-1">Save as</h3>
            <p className="text-xs text-gray-500 mb-4">Enter a filename for your downloaded file.</p>
            <input
              type="text"
              value={customFilename}
              onChange={(e) => setCustomFilename(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter") doDownload(); if (e.key === "Escape") setFilenameDialogOpen(false); }}
              className="w-full border border-[#e5e3dc] rounded-xl px-3 py-2.5 text-sm text-gray-900 focus:outline-none focus:ring-2 focus:ring-amber-400 mb-4"
              autoFocus
              spellCheck={false}
            />
            <div className="flex gap-2">
              <button
                onClick={() => setFilenameDialogOpen(false)}
                className="flex-1 border border-[#e5e3dc] bg-white hover:bg-[#f5f4f0] text-gray-700 text-sm font-medium py-2.5 rounded-xl transition-colors"
              >
                Cancel
              </button>
              <button
                onClick={doDownload}
                className="flex-1 bg-gray-900 hover:bg-black text-white text-sm font-semibold py-2.5 rounded-xl transition-colors flex items-center justify-center gap-1.5"
              >
                <Download className="h-4 w-4" />
                Download
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Preview opens in a popup window via Office Online embed — no modal needed */}
    </div>
  );
}

// ── Shared sidebar shell ──────────────────────────────────────────────────────
function SidebarShell() {
  return (
    <aside className="hidden md:flex w-52 shrink-0 bg-[#f5f4f0] border-r border-[#e5e3dc] flex-col">
      <Link href="/">
        <div className="px-4 pt-5 pb-4 border-b border-[#e5e3dc] cursor-pointer hover:bg-white/40 transition-colors">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-amber-500 flex items-center justify-center">
              <FileText className="h-4 w-4 text-white" />
            </div>
            <div>
              <p className="text-sm font-bold text-gray-900 leading-none">PDFGodWork</p>
              <p className="text-[10px] text-gray-500 mt-0.5">AI Converter</p>
            </div>
          </div>
        </div>
      </Link>
      <div className="flex-1" />
    </aside>
  );
}

// ── Status pill ───────────────────────────────────────────────────────────────
function StatusPill({ status }: { status: string }) {
  const config: Record<string, { bg: string; text: string; label: string; spin: boolean }> = {
    pending:    { bg: "bg-gray-100",   text: "text-gray-600",   label: "Pending",    spin: false },
    extracting: { bg: "bg-amber-100",  text: "text-amber-700",  label: "Extracting", spin: true  },
    translating:{ bg: "bg-amber-100",  text: "text-amber-700",  label: "Translating",spin: true  },
    converting: { bg: "bg-amber-100",  text: "text-amber-700",  label: "Converting", spin: true  },
    done:       { bg: "bg-green-100",  text: "text-green-700",  label: "Complete",   spin: false },
    error:      { bg: "bg-red-100",    text: "text-red-700",    label: "Failed",     spin: false },
    cancelled:  { bg: "bg-gray-100",   text: "text-gray-600",   label: "Stopped",    spin: false },
    paused:     { bg: "bg-amber-100",  text: "text-amber-700",  label: "Paused",     spin: false },
  };
  const c = config[status] ?? config.pending;
  return (
    <span className={`inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold ${c.bg} ${c.text}`}>
      {c.spin
        ? <Loader2 className="h-3 w-3 animate-spin" />
        : status === "done"
          ? <CheckCircle2 className="h-3 w-3" />
          : status === "error"
            ? <XCircle className="h-3 w-3" />
            : status === "cancelled"
              ? <Square className="h-3 w-3 fill-gray-500" />
              : <Clock className="h-3 w-3" />
      }
      {c.label}
    </span>
  );
}
