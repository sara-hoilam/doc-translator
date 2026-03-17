import { useState, useRef, useCallback } from "react";
import { useLocation } from "wouter";
import Navbar from "@/components/Navbar";
import { Button } from "@/components/ui/button";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Badge } from "@/components/ui/badge";
import { trpc } from "@/lib/trpc";
import { useAuth } from "@/_core/hooks/useAuth";
import { getLoginUrl } from "@/const";
import { toast } from "sonner";
import {
  Upload, FileText, Globe, RefreshCw, DollarSign,
  X, AlertCircle, ArrowRight, Loader2, CheckCircle2
} from "lucide-react";

const FORMAT_LABELS: Record<string, string> = {
  pdf: "PDF", docx: "Word (DOCX)", pptx: "PowerPoint (PPTX)",
  xlsx: "Excel (XLSX)", txt: "Plain Text (TXT)", html: "HTML",
};

const FORMAT_ICONS: Record<string, string> = {
  pdf: "📄", docx: "📝", pptx: "📊", xlsx: "📋", txt: "📃", html: "🌐",
};

function formatBytes(bytes: number) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function getExtension(filename: string) {
  return filename.split(".").pop()?.toLowerCase() ?? "txt";
}

export default function Translate() {
  const [, navigate] = useLocation();
  const { isAuthenticated } = useAuth();

  // File state
  const [file, setFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Options state
  const [targetLanguage, setTargetLanguage] = useState<string>("");
  const [targetLanguageName, setTargetLanguageName] = useState<string>("");
  const [outputFormat, setOutputFormat] = useState<string>("");
  const [isUploading, setIsUploading] = useState(false);

  // Queries
  const { data: config } = trpc.docs.config.useQuery();
  const inputFormat = file ? getExtension(file.name) : "";
  const estimatedPages = file ? Math.max(1, Math.ceil(file.size / 3000)) : 1;

  const { data: estimate } = trpc.docs.estimate.useQuery(
    { pageCount: estimatedPages, doTranslate: !!targetLanguage, doConvert: !!(outputFormat && outputFormat !== inputFormat) },
    { enabled: !!file }
  );

  const uploadMutation = trpc.docs.upload.useMutation({
    onSuccess: (data) => {
      toast.success("Document uploaded! Processing started.");
      navigate(`/job/${data.jobId}`);
    },
    onError: (err) => {
      toast.error(`Upload failed: ${err.message}`);
      setIsUploading(false);
    },
  });

  // Drag and drop handlers
  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback(() => setIsDragging(false), []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const dropped = e.dataTransfer.files[0];
    if (dropped) handleFileSelect(dropped);
  }, []);

  const handleFileSelect = (f: File) => {
    const ext = getExtension(f.name);
    const supported = ["pdf", "docx", "pptx", "xlsx", "txt", "html"];
    if (!supported.includes(ext)) {
      toast.error(`Unsupported format: .${ext}. Please upload PDF, DOCX, PPTX, XLSX, TXT, or HTML.`);
      return;
    }
    if (f.size > 50 * 1024 * 1024) {
      toast.error("File too large. Maximum size is 50 MB.");
      return;
    }
    setFile(f);
    setOutputFormat(ext); // default same format
  };

  const handleSubmit = async () => {
    if (!file) return;
    if (!isAuthenticated) {
      window.location.href = getLoginUrl();
      return;
    }
    setIsUploading(true);
    const reader = new FileReader();
    reader.onload = (e) => {
      const base64 = (e.target?.result as string).split(",")[1];
      uploadMutation.mutate({
        filename: file.name,
        mimeType: file.type || "application/octet-stream",
        base64Data: base64,
        outputFormat: outputFormat || getExtension(file.name),
        targetLanguage: targetLanguage || undefined,
        targetLanguageName: targetLanguageName || undefined,
      });
    };
    reader.readAsDataURL(file);
  };

  const clearFile = () => {
    setFile(null);
    setTargetLanguage("");
    setOutputFormat("");
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  return (
    <div className="min-h-screen flex flex-col bg-background">
      <Navbar />
      <main className="flex-1 container py-10">
        <div className="max-w-2xl mx-auto">
          <div className="mb-8">
            <h1 className="text-2xl font-bold text-foreground mb-1">Translate & Convert Document</h1>
            <p className="text-muted-foreground text-sm">Upload a document, choose translation language and output format, then download the result.</p>
          </div>

          {/* Step 1: Upload */}
          <div className="rounded-xl border border-border bg-card p-6 mb-5">
            <div className="flex items-center gap-2 mb-4">
              <div className="w-6 h-6 rounded-full bg-primary text-primary-foreground flex items-center justify-center text-xs font-bold">1</div>
              <h2 className="font-semibold text-foreground">Upload Document</h2>
            </div>

            {!file ? (
              <div
                className={`relative border-2 border-dashed rounded-lg p-10 text-center cursor-pointer transition-colors ${isDragging ? "border-primary bg-primary/5 drop-active" : "border-border hover:border-primary/50 hover:bg-muted/30"}`}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
                onClick={() => fileInputRef.current?.click()}
              >
                <input
                  ref={fileInputRef}
                  type="file"
                  className="hidden"
                  accept=".pdf,.docx,.pptx,.xlsx,.txt,.html"
                  onChange={(e) => e.target.files?.[0] && handleFileSelect(e.target.files[0])}
                />
                <Upload className="h-10 w-10 text-muted-foreground mx-auto mb-3" />
                <p className="font-medium text-foreground mb-1">
                  {isDragging ? "Drop your file here" : "Drag & drop or click to upload"}
                </p>
                <p className="text-sm text-muted-foreground mb-3">PDF, DOCX, PPTX, XLSX, TXT, HTML · Max 50 MB</p>
                <div className="flex flex-wrap justify-center gap-2">
                  {Object.entries(FORMAT_ICONS).map(([fmt, icon]) => (
                    <span key={fmt} className="text-xs bg-muted rounded px-2 py-0.5 text-muted-foreground">
                      {icon} {fmt.toUpperCase()}
                    </span>
                  ))}
                </div>
              </div>
            ) : (
              <div className="flex items-center gap-3 rounded-lg border border-border bg-muted/30 p-4">
                <div className="text-2xl">{FORMAT_ICONS[getExtension(file.name)] ?? "📄"}</div>
                <div className="flex-1 min-w-0">
                  <p className="font-medium text-foreground truncate">{file.name}</p>
                  <p className="text-xs text-muted-foreground">{formatBytes(file.size)} · ~{estimatedPages} page{estimatedPages !== 1 ? "s" : ""} estimated</p>
                </div>
                <Button variant="ghost" size="icon" onClick={clearFile} className="shrink-0">
                  <X className="h-4 w-4" />
                </Button>
              </div>
            )}
          </div>

          {/* Step 2: Options */}
          {file && (
            <div className="rounded-xl border border-border bg-card p-6 mb-5">
              <div className="flex items-center gap-2 mb-5">
                <div className="w-6 h-6 rounded-full bg-primary text-primary-foreground flex items-center justify-center text-xs font-bold">2</div>
                <h2 className="font-semibold text-foreground">Configure Options</h2>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-5">
                {/* Target Language */}
                <div>
                  <label className="text-sm font-medium text-foreground mb-1.5 flex items-center gap-1.5">
                    <Globe className="h-3.5 w-3.5 text-primary" />
                    Translate to
                    <span className="text-xs text-muted-foreground font-normal">(optional)</span>
                  </label>
                  <Select
                    value={targetLanguage}
                    onValueChange={(val) => {
                      setTargetLanguage(val === "none" ? "" : val);
                      const lang = config?.languages.find(l => l.code === val);
                      setTargetLanguageName(lang?.name ?? "");
                    }}
                  >
                    <SelectTrigger>
                      <SelectValue placeholder="No translation" />
                    </SelectTrigger>
                    <SelectContent className="max-h-72">
                      <SelectItem value="none">No translation (keep original)</SelectItem>
                      {config?.languages.map(lang => (
                        <SelectItem key={lang.code} value={lang.code}>{lang.name}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>

                {/* Output Format */}
                <div>
                  <label className="text-sm font-medium text-foreground mb-1.5 flex items-center gap-1.5">
                    <RefreshCw className="h-3.5 w-3.5 text-primary" />
                    Output format
                  </label>
                  <Select value={outputFormat} onValueChange={setOutputFormat}>
                    <SelectTrigger>
                      <SelectValue placeholder="Same as input" />
                    </SelectTrigger>
                    <SelectContent>
                      {config?.outputFormats.map(fmt => (
                        <SelectItem key={fmt} value={fmt}>
                          {FORMAT_ICONS[fmt]} {FORMAT_LABELS[fmt] ?? fmt.toUpperCase()}
                          {fmt === inputFormat && " (same as input)"}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              </div>

              {/* Summary badges */}
              <div className="flex flex-wrap gap-2 mt-4">
                <Badge variant="secondary" className="gap-1">
                  <FileText className="h-3 w-3" />
                  {FORMAT_LABELS[inputFormat] ?? inputFormat.toUpperCase()}
                </Badge>
                {outputFormat && outputFormat !== inputFormat && (
                  <>
                    <ArrowRight className="h-4 w-4 text-muted-foreground self-center" />
                    <Badge variant="secondary" className="gap-1">
                      <RefreshCw className="h-3 w-3" />
                      {FORMAT_LABELS[outputFormat] ?? outputFormat.toUpperCase()}
                    </Badge>
                  </>
                )}
                {targetLanguage && (
                  <>
                    <ArrowRight className="h-4 w-4 text-muted-foreground self-center" />
                    <Badge className="gap-1">
                      <Globe className="h-3 w-3" />
                      {targetLanguageName}
                    </Badge>
                  </>
                )}
              </div>
            </div>
          )}

          {/* Step 3: Cost estimate */}
          {file && estimate && (
            <div className="rounded-xl border border-border bg-card p-6 mb-5">
              <div className="flex items-center gap-2 mb-4">
                <div className="w-6 h-6 rounded-full bg-primary text-primary-foreground flex items-center justify-center text-xs font-bold">3</div>
                <h2 className="font-semibold text-foreground">Cost Estimate</h2>
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                <div className="rounded-lg bg-muted/40 p-3 text-center">
                  <p className="text-xs text-muted-foreground mb-1">Pages</p>
                  <p className="text-xl font-bold text-foreground">{estimatedPages}</p>
                </div>
                <div className="rounded-lg bg-muted/40 p-3 text-center">
                  <p className="text-xs text-muted-foreground mb-1">Extraction</p>
                  <p className="text-xl font-bold text-foreground">${estimate.extractionCost.toFixed(4)}</p>
                </div>
                {targetLanguage && (
                  <div className="rounded-lg bg-muted/40 p-3 text-center">
                    <p className="text-xs text-muted-foreground mb-1">Translation</p>
                    <p className="text-xl font-bold text-foreground">${estimate.translationCost.toFixed(4)}</p>
                  </div>
                )}
                <div className="rounded-lg bg-primary/10 p-3 text-center">
                  <p className="text-xs text-primary mb-1 font-medium">Total Est.</p>
                  <p className="text-xl font-bold text-primary">${estimate.total.toFixed(4)}</p>
                </div>
              </div>
              <p className="text-xs text-muted-foreground mt-3 flex items-center gap-1">
                <AlertCircle className="h-3 w-3" />
                Estimates based on ~{estimatedPages} pages. Actual cost may vary based on document content.
              </p>
            </div>
          )}

          {/* Submit */}
          {file && (
            <div className="rounded-xl border border-border bg-card p-6">
              <div className="flex items-center gap-2 mb-4">
                <div className="w-6 h-6 rounded-full bg-primary text-primary-foreground flex items-center justify-center text-xs font-bold">4</div>
                <h2 className="font-semibold text-foreground">Process Document</h2>
              </div>

              {!isAuthenticated ? (
                <div className="rounded-lg bg-amber-50 border border-amber-200 p-4 flex items-start gap-3 mb-4">
                  <AlertCircle className="h-5 w-5 text-amber-500 shrink-0 mt-0.5" />
                  <div>
                    <p className="text-sm font-medium text-amber-800">Sign in required</p>
                    <p className="text-xs text-amber-700 mt-0.5">You need to sign in to process documents and save your history.</p>
                  </div>
                </div>
              ) : null}

              <div className="flex flex-col sm:flex-row gap-3">
                <Button
                  className="flex-1 gap-2"
                  size="lg"
                  onClick={handleSubmit}
                  disabled={isUploading || uploadMutation.isPending}
                >
                  {isUploading || uploadMutation.isPending ? (
                    <><Loader2 className="h-4 w-4 animate-spin" /> Processing...</>
                  ) : !isAuthenticated ? (
                    <><ArrowRight className="h-4 w-4" /> Sign in & Process</>
                  ) : (
                    <><CheckCircle2 className="h-4 w-4" /> Start Processing</>
                  )}
                </Button>
                <Button variant="outline" onClick={clearFile} disabled={isUploading}>
                  Clear
                </Button>
              </div>
            </div>
          )}
        </div>
      </main>
    </div>
  );
}
