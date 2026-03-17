import { Link } from "wouter";
import { Button } from "@/components/ui/button";
import Navbar from "@/components/Navbar";
import { FileText, Globe, RefreshCw, Shield, Zap, BarChart2, ArrowRight, CheckCircle } from "lucide-react";
import { trpc } from "@/lib/trpc";

const FEATURES = [
  { icon: FileText, title: "6 File Formats", desc: "PDF, DOCX, PPTX, XLSX, TXT, HTML — upload any document type with drag-and-drop." },
  { icon: Globe, title: "100+ Languages", desc: "Powered by Google Translate for accurate, natural-sounding translations across 100+ languages." },
  { icon: RefreshCw, title: "Format Conversion", desc: "Convert between any supported formats using LibreOffice and Pandoc — free and open-source." },
  { icon: Shield, title: "Layout Preserved", desc: "GPT-4o-mini extracts structure, headings, tables, and lists so your formatting stays intact." },
  { icon: Zap, title: "Real-time Progress", desc: "Watch each step — extraction, translation, conversion — complete in real time." },
  { icon: BarChart2, title: "Cost Transparency", desc: "See estimated per-page costs before processing. No surprises, no hidden fees." },
];

const STEPS = [
  { n: "1", title: "Upload your document", desc: "Drag and drop or browse to upload PDF, DOCX, PPTX, XLSX, TXT, or HTML files." },
  { n: "2", title: "Choose options", desc: "Select target language (optional), output format, and review the cost estimate." },
  { n: "3", title: "Process & download", desc: "Watch real-time progress as your document is extracted, translated, and converted." },
];

export default function Home() {
  const { data: config } = trpc.docs.config.useQuery();

  return (
    <div className="min-h-screen flex flex-col bg-background">
      <Navbar />

      {/* Hero */}
      <section className="relative overflow-hidden border-b border-border">
        <div className="absolute inset-0 bg-gradient-to-br from-primary/5 via-background to-background pointer-events-none" />
        <div className="container py-20 md:py-28 relative">
          <div className="max-w-3xl">
            <div className="inline-flex items-center gap-2 rounded-full bg-primary/10 px-3 py-1 text-sm text-primary font-medium mb-6">
              <Globe className="h-3.5 w-3.5" />
              {config ? `${config.languages.length}+ languages supported` : "100+ languages supported"}
            </div>
            <h1 className="text-4xl md:text-5xl font-bold tracking-tight text-foreground mb-5 leading-tight">
              Translate & Convert Documents<br />
              <span className="text-primary">Without Losing Formatting</span>
            </h1>
            <p className="text-lg text-muted-foreground mb-8 max-w-2xl">
              Upload any document, translate it to 100+ languages, and convert between formats — all while preserving your original layout, tables, headings, and structure.
            </p>
            <div className="flex flex-wrap gap-3">
              <Button size="lg" asChild className="gap-2">
                <Link href="/translate">
                  Get Started Free
                  <ArrowRight className="h-4 w-4" />
                </Link>
              </Button>
              <Button size="lg" variant="outline" asChild>
                <Link href="/history">View History</Link>
              </Button>
            </div>
            <div className="flex flex-wrap gap-4 mt-8">
              {["PDF", "DOCX", "PPTX", "XLSX", "TXT", "HTML"].map(fmt => (
                <span key={fmt} className="inline-flex items-center gap-1 text-sm text-muted-foreground">
                  <CheckCircle className="h-3.5 w-3.5 text-primary" />
                  {fmt}
                </span>
              ))}
            </div>
          </div>
        </div>
      </section>

      {/* Cost callout */}
      {config && (
        <section className="border-b border-border bg-muted/30">
          <div className="container py-6">
            <div className="flex flex-wrap items-center gap-6 text-sm">
              <span className="font-semibold text-foreground">Estimated costs:</span>
              <span className="text-muted-foreground">
                Extraction: <strong className="text-foreground">${config.costPerPageExtraction.toFixed(4)}/page</strong>
              </span>
              <span className="text-muted-foreground">
                Translation: <strong className="text-foreground">${config.costPerPageTranslation.toFixed(4)}/page</strong>
              </span>
              <span className="text-muted-foreground">
                Total: <strong className="text-primary">${config.costPerPageTotal.toFixed(4)}/page</strong>
              </span>
              <span className="text-xs text-muted-foreground">(GPT-4o-mini + Google Translate NMT)</span>
            </div>
          </div>
        </section>
      )}

      {/* Features */}
      <section className="container py-16">
        <h2 className="text-2xl font-bold text-center mb-10">Everything you need for document processing</h2>
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6">
          {FEATURES.map(({ icon: Icon, title, desc }) => (
            <div key={title} className="rounded-lg border border-border p-5 bg-card hover:shadow-sm transition-shadow">
              <div className="flex items-center gap-3 mb-3">
                <div className="rounded-md bg-primary/10 p-2">
                  <Icon className="h-5 w-5 text-primary" />
                </div>
                <h3 className="font-semibold text-foreground">{title}</h3>
              </div>
              <p className="text-sm text-muted-foreground leading-relaxed">{desc}</p>
            </div>
          ))}
        </div>
      </section>

      {/* How it works */}
      <section className="border-t border-border bg-muted/20">
        <div className="container py-16">
          <h2 className="text-2xl font-bold text-center mb-10">How it works</h2>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-8 max-w-4xl mx-auto">
            {STEPS.map(({ n, title, desc }) => (
              <div key={n} className="flex flex-col items-center text-center">
                <div className="w-10 h-10 rounded-full bg-primary text-primary-foreground flex items-center justify-center font-bold text-lg mb-4">
                  {n}
                </div>
                <h3 className="font-semibold text-foreground mb-2">{title}</h3>
                <p className="text-sm text-muted-foreground leading-relaxed">{desc}</p>
              </div>
            ))}
          </div>
          <div className="text-center mt-10">
            <Button size="lg" asChild className="gap-2">
              <Link href="/translate">
                Start Translating
                <ArrowRight className="h-4 w-4" />
              </Link>
            </Button>
          </div>
        </div>
      </section>

      {/* Footer */}
      <footer className="border-t border-border mt-auto">
        <div className="container py-6 flex flex-col sm:flex-row items-center justify-between gap-2 text-sm text-muted-foreground">
          <div className="flex items-center gap-2">
            <FileText className="h-4 w-4 text-primary" />
            <span className="font-medium text-foreground">PDFGodWork</span>
          </div>
          <span>Powered by GPT-4o-mini · Google Translate · LibreOffice · Pandoc</span>
        </div>
      </footer>
    </div>
  );
}
