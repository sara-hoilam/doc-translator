import { useState, useEffect, useRef, useCallback } from "react";
import { useLocation } from "wouter";
import { trpc } from "@/lib/trpc";

import { toast } from "sonner";
import { Switch } from "@/components/ui/switch";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Button } from "@/components/ui/button";
import {
  FileText, FileSpreadsheet, Presentation, Globe,
  Settings, Upload, Download, CheckCircle2,
  Loader2, X, ChevronRight, Sparkles, Info, TriangleAlert,
  Menu, ChevronLeft, ImageIcon, GripVertical, Plus,
} from "lucide-react";

// ─── Types ────────────────────────────────────────────────────────────────────
type OutputFormat = "pptx" | "docx" | "pdf" | "txt" | "xlsx" | "csv";

// Allowed output formats per input format
const ALLOWED_OUTPUT_FORMATS: Record<string, OutputFormat[]> = {
  pptx: ["pptx", "pdf"],
  docx: ["docx", "pdf"],
  xlsx: ["xlsx", "csv"],
  txt:  ["txt"],
  pdf:  ["docx"],
  // image formats → PDF only
  png:  ["pdf"],
  jpg:  ["pdf"],
  jpeg: ["pdf"],
  webp: ["pdf"],
  gif:  ["pdf"],
};

interface FormatOption {
  id: OutputFormat;
  label: string;
  icon: React.ElementType;
  color: string;
  bg: string;
}

const FORMAT_OPTIONS: FormatOption[] = [
  { id: "pptx", label: "PowerPoint", icon: Presentation, color: "text-orange-600", bg: "bg-orange-50" },
  { id: "docx", label: "Word Document", icon: FileText, color: "text-blue-600", bg: "bg-blue-50" },
  { id: "pdf", label: "PDF", icon: FileText, color: "text-red-600", bg: "bg-red-50" },
  { id: "txt", label: "Plain Text", icon: FileText, color: "text-gray-600", bg: "bg-gray-50" },
  { id: "xlsx", label: "Excel Spreadsheet", icon: FileSpreadsheet, color: "text-green-600", bg: "bg-green-50" },
  { id: "csv", label: "CSV", icon: FileSpreadsheet, color: "text-teal-600", bg: "bg-teal-50" },
];

const INPUT_FORMAT_ICONS: Record<string, string> = {
  docx: "📝", pptx: "📊", xlsx: "📋", txt: "📃", pdf: "📄",
  png: "🖼️", jpg: "🖼️", jpeg: "🖼️", webp: "🖼️", gif: "🖼️",
};

const IMAGE_FORMATS = ["png", "jpg", "jpeg", "webp", "gif"];


function getExtension(filename: string) {
  return filename.split(".").pop()?.toLowerCase() ?? "txt";
}

function formatBytes(bytes: number) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

// ─── Sidebar content (shared between desktop and mobile drawer) ───────────────
function SidebarContent({
  selectedFormat,
  setSelectedFormat,
  translateEnabled,
  setTranslateEnabled,
  targetLanguage,
  setTargetLanguage,
  config,
  file,
  imageFiles,
  onClose,
}: {
  selectedFormat: OutputFormat;
  setSelectedFormat: (f: OutputFormat) => void;
  translateEnabled: boolean;
  setTranslateEnabled: (v: boolean) => void;
  targetLanguage: string;
  setTargetLanguage: (v: string) => void;
  config: any;
  file: File | null;
  imageFiles: File[];
  onClose?: () => void;
}) {
  // Determine input format for filtering
  const isMultiImage = imageFiles.length > 0;
  const inputExt = isMultiImage
    ? "png"
    : file ? getExtension(file.name) : null;
  const allowedFormats = inputExt
    ? (ALLOWED_OUTPUT_FORMATS[inputExt] ?? FORMAT_OPTIONS.map(f => f.id))
    : FORMAT_OPTIONS.map(f => f.id);

  return (
    <>
      {/* Logo + close button (mobile only) */}
      <div className="px-4 pt-5 pb-4 border-b border-[#e5e3dc] flex items-center justify-between">
        <div className="flex items-center gap-2.5">
          <div className="w-8 h-8 rounded-lg bg-amber-500 flex items-center justify-center">
            <FileText className="h-4 w-4 text-white" />
          </div>
          <div>
            <p className="text-sm font-bold text-gray-900 leading-none">PDFGodWork</p>
            <p className="text-[10px] text-gray-500 mt-0.5">AI Translator</p>
          </div>
        </div>
        {onClose && (
          <button
            onClick={onClose}
            className="p-1.5 rounded-lg hover:bg-gray-100 text-gray-500 transition-colors"
            aria-label="Close menu"
          >
            <ChevronLeft className="h-5 w-5" />
          </button>
        )}
      </div>


      <div className="px-3 pt-4 flex-1 overflow-y-auto">
        {/* Feature description */}
        <div className="space-y-3 px-1">
          <div>
            <p className="text-sm font-semibold text-gray-800 leading-snug">Still google translating line by line?</p>
            <p className="text-xs text-gray-500 mt-1 leading-relaxed">Upload any document and instantly convert or translate it — preserving your original formatting.</p>
          </div>

          <div className="space-y-2">
            <div className="flex items-start gap-2.5">
              <div className="w-7 h-7 rounded-lg bg-amber-100 flex items-center justify-center shrink-0 mt-0.5">
                <span className="text-sm">🌍</span>
              </div>
              <div>
                <p className="text-xs font-semibold text-gray-700">109 languages supported</p>
                <p className="text-[11px] text-gray-400 leading-relaxed">Translate to Spanish, French, Chinese, Arabic, Japanese and 104 more.</p>
              </div>
            </div>

            <div className="flex items-start gap-2.5">
              <div className="w-7 h-7 rounded-lg bg-green-50 flex items-center justify-center shrink-0 mt-0.5">
                <span className="text-sm">✨</span>
              </div>
              <div>
                <p className="text-xs font-semibold text-gray-700">Formatting preserved</p>
                <p className="text-[11px] text-gray-400 leading-relaxed">Fonts, tables, images, and layout stay intact through every conversion.</p>
              </div>
            </div>

            <div className="flex items-start gap-2.5">
              <div className="w-7 h-7 rounded-lg bg-purple-50 flex items-center justify-center shrink-0 mt-0.5">
                <span className="text-sm">🔒</span>
              </div>
              <div>
                <p className="text-xs font-semibold text-gray-700">Private & secure</p>
                <p className="text-[11px] text-gray-400 leading-relaxed">Files are processed in an isolated environment and deleted after conversion.</p>
              </div>
            </div>
          </div>
        </div>

        {/* Translation */}
        <div className="mt-5 pt-4 border-t border-[#e5e3dc]">
          <p className="text-[10px] font-semibold text-gray-400 uppercase tracking-wider px-1 mb-3">Translation</p>
          <div className="flex items-center justify-between px-1 mb-3">
            <span className="text-sm text-gray-600">Translate output</span>
            <Switch checked={translateEnabled} onCheckedChange={setTranslateEnabled} className="scale-90" />
          </div>
          {translateEnabled && (
            <Select value={targetLanguage} onValueChange={setTargetLanguage}>
              <SelectTrigger className="h-8 text-xs bg-white border-[#e5e3dc]">
                <SelectValue />
              </SelectTrigger>
              <SelectContent className="max-h-64">
                {config?.languages.map((lang: any) => (
                  <SelectItem key={lang.code} value={lang.code} className="text-xs">{lang.name}</SelectItem>
                ))}
              </SelectContent>
            </Select>
          )}
        </div>

        {/* Powered by Anthropic badge */}
        <div className="mt-5 pt-4 border-t border-[#e5e3dc]">
          <div className="flex items-center gap-2.5 px-1 py-2 rounded-xl bg-gradient-to-r from-amber-50 to-orange-50 border border-amber-100">
            <div className="w-7 h-7 rounded-lg bg-amber-500 flex items-center justify-center shrink-0">
              <Sparkles className="h-3.5 w-3.5 text-white" />
            </div>
            <div>
              <p className="text-xs font-semibold text-gray-800 leading-none">Powered by Anthropic</p>
              <p className="text-[10px] text-gray-500 mt-0.5 leading-relaxed">Claude — the world's most intelligent AI for precise, natural-sounding translations.</p>
            </div>
          </div>
        </div>
      </div>
    </>
  );
}

// ─── Image thumbnail item (draggable) ─────────────────────────────────────────
function ImageThumb({
  file,
  index,
  total,
  onRemove,
  onMoveUp,
  onMoveDown,
  objectUrl,
}: {
  file: File;
  index: number;
  total: number;
  onRemove: () => void;
  onMoveUp: () => void;
  onMoveDown: () => void;
  objectUrl: string;
}) {
  return (
    <div className="flex items-center gap-3 bg-white rounded-xl border border-[#e5e3dc] p-2.5 group">
      {/* Thumbnail */}
      <div className="w-12 h-12 rounded-lg overflow-hidden bg-gray-100 shrink-0 border border-[#e5e3dc]">
        <img src={objectUrl} alt={file.name} className="w-full h-full object-cover" />
      </div>

      {/* Info */}
      <div className="flex-1 min-w-0">
        <p className="text-xs font-medium text-gray-800 truncate">{file.name}</p>
        <p className="text-[10px] text-gray-400 mt-0.5">{formatBytes(file.size)}</p>
      </div>

      {/* Page number badge */}
      <span className="text-[10px] font-semibold text-amber-700 bg-amber-50 border border-amber-200 rounded-md px-1.5 py-0.5 shrink-0">
        p.{index + 1}
      </span>

      {/* Reorder arrows */}
      <div className="flex flex-col gap-0.5 shrink-0">
        <button
          onClick={onMoveUp}
          disabled={index === 0}
          className="w-5 h-5 flex items-center justify-center rounded text-gray-400 hover:text-gray-700 hover:bg-gray-100 disabled:opacity-20 disabled:cursor-not-allowed transition-colors"
          title="Move up"
        >
          <svg className="h-3 w-3" viewBox="0 0 12 12" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M2 8l4-4 4 4" />
          </svg>
        </button>
        <button
          onClick={onMoveDown}
          disabled={index === total - 1}
          className="w-5 h-5 flex items-center justify-center rounded text-gray-400 hover:text-gray-700 hover:bg-gray-100 disabled:opacity-20 disabled:cursor-not-allowed transition-colors"
          title="Move down"
        >
          <svg className="h-3 w-3" viewBox="0 0 12 12" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M2 4l4 4 4-4" />
          </svg>
        </button>
      </div>

      {/* Remove */}
      <button
        onClick={onRemove}
        className="w-6 h-6 flex items-center justify-center rounded-lg text-gray-300 hover:text-red-500 hover:bg-red-50 transition-colors shrink-0"
        title="Remove"
      >
        <X className="h-3.5 w-3.5" />
      </button>
    </div>
  );
}

// ─── Main component ───────────────────────────────────────────────────────────
export default function Converter() {
  const [, navigate] = useLocation();

  // ── Single-file state (non-image docs) ──────────────────────────────────────
  const [file, setFile] = useState<File | null>(null);

  // ── Multi-image state ────────────────────────────────────────────────────────
  const [imageFiles, setImageFiles] = useState<File[]>([]);
  const [imageObjectUrls, setImageObjectUrls] = useState<string[]>([]);

  const isMultiImageMode = imageFiles.length > 0;

  // Revoke object URLs when images change to avoid memory leaks
  useEffect(() => {
    return () => {
      imageObjectUrls.forEach(url => URL.revokeObjectURL(url));
    };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const addImages = useCallback((newFiles: File[]) => {
    const valid = newFiles.filter(f => {
      const ext = getExtension(f.name);
      if (!IMAGE_FORMATS.includes(ext)) return false;
      if (f.size > 50 * 1024 * 1024) { toast.error(`${f.name} is too large (max 50 MB)`); return false; }
      return true;
    });
    if (!valid.length) return;
    const newUrls = valid.map(f => URL.createObjectURL(f));
    setImageFiles(prev => [...prev, ...valid]);
    setImageObjectUrls(prev => [...prev, ...newUrls]);
    // Switching to image mode clears any single file
    setFile(null);
    // Images always output to PDF
    setSelectedFormat("pdf");
  }, []);

  const removeImage = useCallback((index: number) => {
    setImageFiles(prev => {
      const next = [...prev];
      next.splice(index, 1);
      return next;
    });
    setImageObjectUrls(prev => {
      const next = [...prev];
      URL.revokeObjectURL(next[index]);
      next.splice(index, 1);
      return next;
    });
  }, []);

  const moveImage = useCallback((from: number, to: number) => {
    setImageFiles(prev => {
      const next = [...prev];
      const [item] = next.splice(from, 1);
      next.splice(to, 0, item);
      return next;
    });
    setImageObjectUrls(prev => {
      const next = [...prev];
      const [item] = next.splice(from, 1);
      next.splice(to, 0, item);
      return next;
    });
  }, []);

  const clearAll = useCallback(() => {
    imageObjectUrls.forEach(url => URL.revokeObjectURL(url));
    setImageFiles([]);
    setImageObjectUrls([]);
    setFile(null);
    if (fileInputRef.current) fileInputRef.current.value = "";
    if (multiImageInputRef.current) multiImageInputRef.current.value = "";
  }, [imageObjectUrls]);

  const [isDragging, setIsDragging] = useState(false);
  const [selectedFormat, setSelectedFormat] = useState<OutputFormat>("pptx");
  const [translateEnabled, setTranslateEnabled] = useState(false);
  const [targetLanguage, setTargetLanguage] = useState("zh-CN");
  const [isUploading, setIsUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(0); // 0-100
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const multiImageInputRef = useRef<HTMLInputElement>(null);

  const { data: config } = trpc.docs.config.useQuery();
  const inputFormat = isMultiImageMode ? "png" : (file ? getExtension(file.name) : "");


  const estimatedPages = isMultiImageMode
    ? imageFiles.length
    : file
    ? Math.max(1, Math.ceil(file.size / 3000))
    : 1;

  const { data: estimate } = trpc.docs.estimate.useQuery(
    { pageCount: estimatedPages, doTranslate: translateEnabled, doConvert: true },
    { enabled: !!(file || isMultiImageMode) }
  );

  const uploadMutation = trpc.docs.upload.useMutation({
    onSuccess: (data) => {
      toast.success("Processing started!");
      navigate(`/job/${data.jobId}`);
    },
    onError: (err) => {
      toast.error(`Upload failed: ${err.message}`);
      setIsUploading(false);
    },
  });

  // ── Drag & drop handlers ─────────────────────────────────────────────────────
  const handleDragOver = useCallback((e: React.DragEvent) => { e.preventDefault(); setIsDragging(true); }, []);
  const handleDragLeave = useCallback(() => setIsDragging(false), []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const files = Array.from(e.dataTransfer.files);
    if (!files.length) return;

    const imageDrops = files.filter(f => IMAGE_FORMATS.includes(getExtension(f.name)));
    const docDrops = files.filter(f => !IMAGE_FORMATS.includes(getExtension(f.name)));

    if (imageDrops.length > 0) {
      // All images → multi-image mode
      if (docDrops.length > 0) {
        toast.warning("Mixed drop detected — only images were added. Drop documents separately.");
      }
      addImages(imageDrops);
    } else if (docDrops.length > 0) {
      // Single document
      handleFileSelect(docDrops[0]);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [addImages]);

  // Drop onto the existing image list (add more images)
  const handleImageListDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    const files = Array.from(e.dataTransfer.files);
    const imageDrops = files.filter(f => IMAGE_FORMATS.includes(getExtension(f.name)));
    if (imageDrops.length) addImages(imageDrops);
  }, [addImages]);

  const handleFileSelect = (f: File) => {
    const ext = getExtension(f.name);
    const SUPPORTED = ["docx", "pptx", "xlsx", "txt", "pdf", "png", "jpg", "jpeg", "webp", "gif"];
    if (!SUPPORTED.includes(ext)) {
      toast.error(`Unsupported format: .${ext}`); return;
    }
    if (f.size > 50 * 1024 * 1024) { toast.error("File too large (max 50 MB)"); return; }

    if (IMAGE_FORMATS.includes(ext)) {
      // Single image → add to multi-image list
      addImages([f]);
    } else {
      // Non-image doc → single file mode, clear images
      imageObjectUrls.forEach(url => URL.revokeObjectURL(url));
      setImageFiles([]);
      setImageObjectUrls([]);
      setFile(f);
      const allowed = ALLOWED_OUTPUT_FORMATS[ext] ?? ["txt"];
      if (!allowed.includes(selectedFormat)) {
        setSelectedFormat(allowed[0]);
      }
    }
  };

  // Paste handler for images (Cmd+V / Ctrl+V)
  useEffect(() => {
    const handlePaste = (e: ClipboardEvent) => {
      const items = e.clipboardData?.items;
      if (!items) return;
      const imageItems: File[] = [];
      for (const item of Array.from(items)) {
        if (item.type.startsWith("image/")) {
          const blob = item.getAsFile();
          if (blob) {
            const rawExt = item.type.split("/")[1] ?? "png";
            const ext = rawExt === "jpeg" ? "jpg" : rawExt;
            imageItems.push(new File([blob], `pasted-image.${ext}`, { type: item.type }));
          }
        }
      }
      if (imageItems.length) {
        addImages(imageItems);
        toast.success(`${imageItems.length} image${imageItems.length > 1 ? "s" : ""} pasted — will merge into PDF`);
      }
    };
    window.addEventListener("paste", handlePaste);
    return () => window.removeEventListener("paste", handlePaste);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [addImages]);

  // ── Convert / upload ─────────────────────────────────────────────────────────
  const doConvert = () => {
    if (isMultiImageMode) {
      doConvertMultiImage();
    } else if (file) {
      doConvertSingleFile(file);
    }
  };

  const doConvertSingleFile = (f: File) => {
    setIsUploading(true);
    setUploadProgress(0);
    const reader = new FileReader();
    reader.onprogress = (e) => {
      if (e.lengthComputable) setUploadProgress(Math.round((e.loaded / e.total) * 30));
    };
    reader.onload = (e) => {
      setUploadProgress(35);
      const base64 = (e.target?.result as string).split(",")[1];
      const lang = config?.languages.find((l: any) => l.code === targetLanguage);
      let simProgress = 35;
      const simInterval = setInterval(() => {
        simProgress = Math.min(90, simProgress + Math.random() * 8 + 2);
        setUploadProgress(Math.round(simProgress));
      }, 300);
      uploadMutation.mutate(
        {
          filename: f.name,
          mimeType: f.type || "application/octet-stream",
          base64Data: base64,
          outputFormat: selectedFormat,
          targetLanguage: translateEnabled ? targetLanguage : undefined,
          targetLanguageName: translateEnabled ? (lang?.name ?? targetLanguage) : undefined,
        },
        {
          onSuccess: () => { clearInterval(simInterval); setUploadProgress(100); },
          onError: () => { clearInterval(simInterval); setUploadProgress(0); },
        }
      );
    };
    reader.readAsDataURL(f);
  };

  const doConvertMultiImage = async () => {
    if (!imageFiles.length) return;
    setIsUploading(true);
    setUploadProgress(0);

    try {
      // Read all images as base64 in parallel, updating progress
      const total = imageFiles.length;
      const base64List: string[] = new Array(total).fill("");
      let done = 0;

      await Promise.all(imageFiles.map((imgFile, i) =>
        new Promise<void>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = (e) => {
            base64List[i] = (e.target?.result as string).split(",")[1];
            done++;
            setUploadProgress(Math.round((done / total) * 30));
            resolve();
          };
          reader.onerror = reject;
          reader.readAsDataURL(imgFile);
        })
      ));

      setUploadProgress(35);
      const lang = config?.languages.find((l: any) => l.code === targetLanguage);
      let simProgress = 35;
      const simInterval = setInterval(() => {
        simProgress = Math.min(90, simProgress + Math.random() * 8 + 2);
        setUploadProgress(Math.round(simProgress));
      }, 300);

      uploadMutation.mutate(
        {
          filename: "merged-images.pdf",
          mimeType: "application/octet-stream",
          base64Data: "MULTI_IMAGE", // sentinel — real data in multiImageBase64
          multiImageBase64: base64List,
          multiImageMimeTypes: imageFiles.map(f => f.type || "image/png"),
          outputFormat: "pdf",
          targetLanguage: translateEnabled ? targetLanguage : undefined,
          targetLanguageName: translateEnabled ? (lang?.name ?? targetLanguage) : undefined,
        },
        {
          onSuccess: () => { clearInterval(simInterval); setUploadProgress(100); },
          onError: () => { clearInterval(simInterval); setUploadProgress(0); },
        }
      );
    } catch {
      toast.error("Failed to read images");
      setIsUploading(false);
    }
  };

  const handleConvert = () => {
    if (!file && !isMultiImageMode) { toast.error("Please upload a file first"); return; }
    doConvert();
  };

  const selectedFormatInfo = FORMAT_OPTIONS.find(f => f.id === selectedFormat)!;

  const sidebarProps = {
    selectedFormat,
    setSelectedFormat,
    translateEnabled,
    setTranslateEnabled,
    targetLanguage,
    setTargetLanguage,
    config,
    file,
    imageFiles,
  };

  const hasInput = !!file || isMultiImageMode;

  return (
    <div className="min-h-screen flex bg-[#f5f4f0]">

      {/* ── Mobile overlay ──────────────────────────────────────────────────── */}
      {sidebarOpen && (
        <div
          className="fixed inset-0 bg-black/40 z-30 md:hidden"
          onClick={() => setSidebarOpen(false)}
        />
      )}

      {/* ── Sidebar (desktop: always visible, mobile: drawer) ───────────────── */}
      <aside className={`
        fixed md:sticky top-0 left-0 h-screen w-80 bg-[#faf9f6] border-r border-[#e5e3dc] flex flex-col z-40 transition-transform duration-200
        ${sidebarOpen ? "translate-x-0" : "-translate-x-full md:translate-x-0"}
      `}>
        <SidebarContent {...sidebarProps} onClose={() => setSidebarOpen(false)} />
      </aside>

      {/* ── Main content ────────────────────────────────────────────────────── */}
      <main className="flex-1 flex flex-col min-h-screen min-w-0">
        {/* Header */}
        <div className="px-4 sm:px-8 pt-5 sm:pt-8 pb-4 sm:pb-6 border-b border-[#e5e3dc] bg-white/60 flex items-center justify-between gap-4">
          <div className="min-w-0">
            <div className="flex items-center gap-2">
              <button
                className="md:hidden p-1.5 rounded-lg hover:bg-gray-100 text-gray-500 transition-colors mr-1"
                onClick={() => setSidebarOpen(true)}
              >
                <Menu className="h-5 w-5" />
              </button>
              <h1 className="text-lg sm:text-2xl font-bold text-gray-900 leading-tight truncate">
                AI Translates Any Document
              </h1>
            </div>
            <p className="text-xs sm:text-sm text-gray-500 mt-0.5 hidden sm:block">
              Upload a file, AI translates to any language
            </p>
          </div>
          <div className="flex items-center gap-1.5 shrink-0">
            <div className="w-2 h-2 rounded-full bg-green-500" />
            <span className="text-xs sm:text-sm text-gray-500 font-medium whitespace-nowrap">AI Ready</span>
          </div>
        </div>
        {/* Mobile subtitle */}
        {isMultiImageMode && (
          <p className="text-xs text-gray-500 px-4 pt-2 sm:hidden">
            {imageFiles.length} image{imageFiles.length !== 1 ? "s" : ""} — drag to reorder, then convert to PDF
          </p>
        )}

        {/* Drop zone / file area */}
        <div className="flex-1 flex flex-col items-center justify-center px-4 sm:px-8 pb-4 pt-4">
          {!hasInput ? (
            <>
              {/* Empty drop zone */}
              <div
                className={`w-full max-w-xl border-2 border-dashed rounded-2xl p-8 sm:p-14 text-center cursor-pointer transition-all ${
                  isDragging
                    ? "border-amber-400 bg-amber-50/50"
                    : "border-[#d5d3cc] bg-white/50 hover:border-amber-300 hover:bg-white/80"
                }`}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
                onClick={() => fileInputRef.current?.click()}
              >
                <input
                  ref={fileInputRef}
                  type="file"
                  className="hidden"
                  accept=".docx,.pptx,.xlsx,.txt,.pdf,.png,.jpg,.jpeg,.webp,.gif"
                  onChange={(e) => e.target.files?.[0] && handleFileSelect(e.target.files[0])}
                />
                {/* Hidden multi-image input */}
                <input
                  ref={multiImageInputRef}
                  type="file"
                  className="hidden"
                  accept=".png,.jpg,.jpeg,.webp,.gif"
                  multiple
                  onChange={(e) => {
                    if (e.target.files) addImages(Array.from(e.target.files));
                  }}
                />
                <div className="w-14 h-14 sm:w-16 sm:h-16 rounded-2xl bg-amber-100 flex items-center justify-center mx-auto mb-4 sm:mb-5">
                  <Upload className="h-7 w-7 sm:h-8 sm:w-8 text-amber-600" />
                </div>
                <p className="text-base font-semibold text-gray-800 mb-1">Drop your file here</p>
                <p className="text-sm text-amber-600 font-medium mb-2">or browse to upload</p>
                <p className="text-xs text-gray-400">Supports DOCX, PPTX, XLSX, TXT, PNG, JPG · up to 50MB</p>
                <p className="text-xs text-gray-400 mt-1">Drop multiple images to merge into a single PDF</p>
                <p className="text-xs text-gray-400 mt-0.5">or paste an image with Cmd+V / Ctrl+V</p>
              </div>

              {/* How it works */}
              <div className="grid grid-cols-2 gap-2 sm:gap-3 w-full max-w-xl mt-4 sm:mt-6">
                {[
                  { icon: Upload, title: "Upload file", desc: "Drag & drop or browse" },
                  { icon: Settings, title: "Choose format", desc: "PowerPoint, Word & more" },
                  { icon: Globe, title: "Translate", desc: `${config?.languages.length ?? 109} languages supported` },
                  { icon: Download, title: "Download", desc: "Instant export" },
                ].map(({ icon: Icon, title, desc }) => (
                  <div key={title} className="bg-white/70 rounded-xl border border-[#e5e3dc] p-3 sm:p-4 flex items-start gap-2 sm:gap-3">
                    <div className="w-7 h-7 sm:w-8 sm:h-8 rounded-lg bg-amber-50 flex items-center justify-center shrink-0">
                      <Icon className="h-3.5 w-3.5 sm:h-4 sm:w-4 text-amber-600" />
                    </div>
                    <div className="min-w-0">
                      <p className="text-xs sm:text-sm font-semibold text-gray-800">{title}</p>
                      <p className="text-[10px] sm:text-xs text-gray-500 mt-0.5">{desc}</p>
                    </div>
                  </div>
                ))}
              </div>
            </>
          ) : isMultiImageMode ? (
            /* ── Multi-image mode ─────────────────────────────────────────── */
            <div className="w-full max-w-xl space-y-3">
              {/* Header row */}
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <ImageIcon className="h-4 w-4 text-amber-600" />
                  <span className="text-sm font-semibold text-gray-800">
                    {imageFiles.length} image{imageFiles.length !== 1 ? "s" : ""} → PDF
                  </span>
                  <span className="text-[10px] text-gray-400 bg-gray-100 rounded-md px-1.5 py-0.5">
                    {formatBytes(imageFiles.reduce((s, f) => s + f.size, 0))} total
                  </span>
                </div>
                <button
                  onClick={clearAll}
                  className="text-xs text-gray-400 hover:text-red-500 transition-colors flex items-center gap-1"
                >
                  <X className="h-3.5 w-3.5" /> Clear all
                </button>
              </div>

              {/* Image list (drop target for adding more) */}
              <div
                className={`space-y-2 rounded-2xl border-2 border-dashed p-3 transition-all ${
                  isDragging ? "border-amber-400 bg-amber-50/40" : "border-[#e5e3dc] bg-white/30"
                }`}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleImageListDrop}
              >
                {imageFiles.map((imgFile, i) => (
                  <ImageThumb
                    key={`${imgFile.name}-${i}`}
                    file={imgFile}
                    index={i}
                    total={imageFiles.length}
                    objectUrl={imageObjectUrls[i] ?? ""}
                    onRemove={() => removeImage(i)}
                    onMoveUp={() => moveImage(i, i - 1)}
                    onMoveDown={() => moveImage(i, i + 1)}
                  />
                ))}

                {/* Drop hint inside list */}
                <div
                  className="flex items-center justify-center gap-2 py-2 text-xs text-gray-400 cursor-pointer hover:text-amber-600 transition-colors"
                  onClick={() => multiImageInputRef.current?.click()}
                >
                  <Plus className="h-3.5 w-3.5" />
                  Drop more images here or click to browse
                </div>
              </div>

              {/* Conversion settings card */}
              <div className="bg-white rounded-2xl border border-[#e5e3dc] p-4 sm:p-5 space-y-3">
                {/* From → To */}
                <div className="flex items-center gap-2 text-sm flex-wrap">
                  <span className="text-gray-500 shrink-0">From:</span>
                  <span className="font-medium text-gray-800 shrink-0">
                    {imageFiles.length} Image{imageFiles.length !== 1 ? "s" : ""}
                  </span>
                  <ChevronRight className="h-4 w-4 text-gray-400 shrink-0" />
                  <span className="text-gray-500 shrink-0">To:</span>
                  <span className="font-semibold text-amber-800 bg-amber-50 border border-amber-200 rounded-lg px-2.5 py-1 text-xs">
                    PDF (merged)
                  </span>
                </div>

                {/* Translate toggle */}
                <div className="pt-3 border-t border-[#f0ede6] flex items-center gap-2 flex-wrap">
                  <span className="text-xs text-gray-500 shrink-0">Translate:</span>
                  <Switch checked={translateEnabled} onCheckedChange={setTranslateEnabled} className="scale-90 shrink-0" />
                  {translateEnabled && (
                    <Select value={targetLanguage} onValueChange={setTargetLanguage}>
                      <SelectTrigger className="h-7 text-xs bg-white border-[#e5e3dc] w-auto min-w-[130px] px-2.5">
                        <Globe className="h-3 w-3 text-amber-500 shrink-0 mr-1" />
                        <SelectValue />
                      </SelectTrigger>
                      <SelectContent className="max-h-64">
                        {config?.languages.map((lang: any) => (
                          <SelectItem key={lang.code} value={lang.code} className="text-xs">{lang.name}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  )}
                </div>

                {/* Powered by Anthropic */}
                <div className="pt-3 border-t border-[#f0ede6] flex items-center gap-2">
                  <Sparkles className="h-3.5 w-3.5 text-amber-500 shrink-0" />
                  <span className="text-xs text-gray-500">Powered by <span className="font-semibold text-gray-700">Anthropic Claude</span></span>
                </div>
              </div>
            </div>
          ) : (
            /* ── Single file mode ─────────────────────────────────────────── */
            <div className="w-full max-w-xl space-y-3">
              {/* File card */}
              <div className="bg-white rounded-2xl border border-[#e5e3dc] p-4 sm:p-5 flex items-center gap-3 sm:gap-4">
                <div className="w-10 h-10 sm:w-12 sm:h-12 rounded-xl bg-amber-50 flex items-center justify-center text-xl sm:text-2xl shrink-0">
                  {INPUT_FORMAT_ICONS[inputFormat] ?? "📄"}
                </div>
                <div className="flex-1 min-w-0">
                  <p className="font-semibold text-gray-900 truncate text-sm sm:text-base">{file!.name}</p>
                  <p className="text-xs text-gray-500 mt-0.5">{formatBytes(file!.size)}</p>
                </div>
                <button onClick={clearAll} className="text-gray-400 hover:text-gray-600 transition-colors shrink-0">
                  <X className="h-4 w-4" />
                </button>
              </div>

              {/* Conversion summary card with inline controls */}
              <div className="bg-white rounded-2xl border border-[#e5e3dc] p-4 sm:p-5 space-y-3">

                {/* Row 1: From → To (dropdown) */}
                <div className="flex items-center gap-2 text-sm flex-wrap">
                  <span className="text-gray-500 shrink-0">From:</span>
                  <span className="font-medium text-gray-800 shrink-0">{inputFormat.toUpperCase()}</span>
                  <ChevronRight className="h-4 w-4 text-gray-400 shrink-0" />
                  <span className="text-gray-500 shrink-0">To:</span>
                  <Select value={selectedFormat} onValueChange={(v) => setSelectedFormat(v as OutputFormat)}>
                    <SelectTrigger className="h-7 text-xs bg-amber-50 border-amber-200 text-amber-800 font-semibold w-auto min-w-[110px] px-2.5">
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      {FORMAT_OPTIONS.filter(fmt => {
                        const allowed = ALLOWED_OUTPUT_FORMATS[inputFormat] ?? FORMAT_OPTIONS.map(f => f.id);
                        return allowed.includes(fmt.id);
                      }).map((fmt) => {
                        const Icon = fmt.icon;
                        return (
                          <SelectItem key={fmt.id} value={fmt.id} className="text-xs">
                            <div className="flex items-center gap-2">
                              <Icon className={`h-3.5 w-3.5 ${fmt.color}`} />
                              <span>{fmt.label}</span>
                            </div>
                          </SelectItem>
                        );
                      })}
                    </SelectContent>
                  </Select>
                </div>

                {/* Row 2: Translate toggle + language dropdown */}
                <div className="pt-3 border-t border-[#f0ede6] flex items-center gap-2 flex-wrap">
                  <span className="text-xs text-gray-500 shrink-0">Translate:</span>
                  <Switch checked={translateEnabled} onCheckedChange={setTranslateEnabled} className="scale-90 shrink-0" />
                  {translateEnabled && (
                    <Select value={targetLanguage} onValueChange={setTargetLanguage}>
                      <SelectTrigger className="h-7 text-xs bg-white border-[#e5e3dc] w-auto min-w-[130px] px-2.5">
                        <Globe className="h-3 w-3 text-amber-500 shrink-0 mr-1" />
                        <SelectValue />
                      </SelectTrigger>
                      <SelectContent className="max-h-64">
                        {config?.languages.map((lang: any) => (
                          <SelectItem key={lang.code} value={lang.code} className="text-xs">{lang.name}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  )}
                </div>

                {/* Powered by Anthropic */}
                <div className="pt-3 border-t border-[#f0ede6] flex items-center gap-2">
                  <Sparkles className="h-3.5 w-3.5 text-amber-500 shrink-0" />
                  <span className="text-xs text-gray-500">Powered by <span className="font-semibold text-gray-700">Anthropic Claude</span></span>
                </div>

              </div>


            </div>
          )}
        </div>

        {/* Convert button footer */}
        <div className="border-t border-[#e5e3dc] bg-white/60 px-4 sm:px-8 py-4">
          {/* Upload progress bar */}
          {(isUploading || uploadMutation.isPending) && (
            <div className="mb-3">
              <div className="flex items-center justify-between mb-1.5">
                <span className="text-xs text-gray-500 font-medium">
                  {uploadProgress < 35 ? "Reading files..." : uploadProgress < 95 ? "Uploading..." : "Starting conversion..."}
                </span>
                <span className="text-xs font-semibold text-amber-600">{uploadProgress}%</span>
              </div>
              <div className="w-full h-2 bg-gray-100 rounded-full overflow-hidden">
                <div
                  className="h-full bg-amber-500 rounded-full transition-all duration-300 ease-out"
                  style={{ width: `${uploadProgress}%` }}
                />
              </div>
            </div>
          )}
          <Button
            className="w-full bg-gray-800 hover:bg-gray-900 text-white h-11 text-sm font-medium rounded-xl gap-2"
            onClick={handleConvert}
            disabled={!hasInput || isUploading || uploadMutation.isPending}
          >
            {isUploading || uploadMutation.isPending ? (
              <><Loader2 className="h-4 w-4 animate-spin" /> Uploading...</>
            ) : isMultiImageMode ? (
              <><ImageIcon className="h-4 w-4" /> Merge {imageFiles.length} Image{imageFiles.length !== 1 ? "s" : ""} into PDF</>
            ) : (
              <><Download className="h-4 w-4" /> Convert Document</>
            )}
          </Button>
        </div>
      </main>
    </div>
  );
}
