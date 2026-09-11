import React from "react";
import { 
  Menu, 
  Save, 
  Printer, 
  Download, 
  RefreshCw, 
  Check, 
  Sparkles,
  Smartphone,
  Laptop
} from "lucide-react";

export interface ErpTopNavProps {
  activeView: "dashboard" | "editor" | "saved-docs";
  docType?: "quotation" | "challan" | "invoice";
  onSelectDocType?: (type: "quotation" | "challan" | "invoice") => void;
  currentDocName?: string;
  saveStatus?: "idle" | "saving" | "saved" | "error";
  lastSavedTime?: string | null;
  onSaveDoc: () => void;
  onExportExcel: () => void;
  isGeneratingExcel: boolean;
  onPrint: () => void;
  onDownloadPDF?: () => void;
  isGeneratingPDF?: boolean;
  onOpenExcelModal?: () => void;
  onToggleMobileSidebar: () => void;
  isInstallable?: boolean;
  onInstallClick?: () => void;
}

export default function ErpTopNav({
  activeView,
  docType,
  onSelectDocType,
  currentDocName,
  saveStatus,
  lastSavedTime,
  onSaveDoc,
  onExportExcel,
  isGeneratingExcel,
  onPrint,
  onOpenExcelModal,
  onToggleMobileSidebar,
  isInstallable,
  onInstallClick,
}: ErpTopNavProps) {
  return (
    <header
      id="erp-top-nav"
      className="sticky top-0 z-30 h-13 bg-white border-b border-slate-200 px-4 flex items-center justify-between no-print print:hidden shadow-2xs"
    >
      {/* Left: Mobile Menu Button + Breadcrumb */}
      <div className="flex items-center gap-3 min-w-0">
        <button
          type="button"
          onClick={onToggleMobileSidebar}
          className="lg:hidden p-1.5 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-md cursor-pointer"
          title="Toggle Navigation Menu"
        >
          <Menu className="h-5 w-5" />
        </button>

        <div className="flex items-center gap-2 text-xs truncate">
          <span className="font-semibold text-slate-500 hidden sm:inline">CT Maritime</span>
          <span className="text-slate-300 hidden sm:inline">/</span>
          <span className="font-bold text-slate-900 truncate">
            {activeView === "dashboard"
              ? "Operations Dashboard"
              : activeView === "saved-docs"
              ? "Records Archive"
              : currentDocName || "Document Canvas"}
          </span>
        </div>
      </div>

      {/* Right: Actions */}
      <div className="flex items-center gap-2 shrink-0">
        {/* Editor Action Buttons */}
        {activeView === "editor" && (
          <div className="flex items-center gap-1">
            {/* Export Excel (.xlsx) */}
            <button
              type="button"
              id="topnav-btn-export-excel"
              onClick={onExportExcel}
              disabled={isGeneratingExcel}
              className="h-8 px-2.5 bg-white hover:bg-slate-50 text-slate-700 hover:text-slate-900 border border-slate-300 rounded-md text-xs font-semibold flex items-center gap-1.5 shadow-2xs transition-colors cursor-pointer disabled:opacity-50"
              title="Download clean formatted Excel file"
            >
              {isGeneratingExcel ? (
                <RefreshCw className="h-3.5 w-3.5 animate-spin text-slate-500" />
              ) : (
                <Download className="h-3.5 w-3.5 text-slate-600" />
              )}
              <span className="hidden xl:inline">Excel</span>
            </button>

            {/* Print / Save PDF */}
            <button
              type="button"
              id="topnav-btn-print"
              onClick={onPrint}
              className="h-8 px-2.5 bg-white hover:bg-slate-50 text-slate-700 hover:text-slate-900 border border-slate-300 rounded-md text-xs font-semibold flex items-center gap-1.5 shadow-2xs transition-colors cursor-pointer"
              title="Print Document or Save as PDF (Ctrl+P)"
            >
              <Printer className="h-3.5 w-3.5 text-slate-700" />
              <span className="hidden sm:inline">Print / PDF</span>
            </button>

            {/* Save to Cloud Button */}
            <button
              type="button"
              id="topnav-btn-save"
              onClick={onSaveDoc}
              disabled={saveStatus === "saving"}
              className="h-8 px-3 bg-blue-600 hover:bg-blue-700 text-white rounded-md text-xs font-semibold flex items-center gap-1.5 shadow-2xs transition-colors cursor-pointer disabled:opacity-60"
              title="Save Document Record"
            >
              <Save className="h-3.5 w-3.5" />
              <span>Save</span>
            </button>
          </div>
        )}

        {/* Install PWA Button if available */}
        {isInstallable && onInstallClick && (
          <button
            type="button"
            onClick={onInstallClick}
            className="h-8 px-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-md text-xs font-semibold flex items-center gap-1.5 transition-colors cursor-pointer"
            title="Install Standalone Desktop App"
          >
            <Laptop className="h-3.5 w-3.5 text-slate-600" />
            <span className="hidden sm:inline">Install</span>
          </button>
        )}

      </div>
    </header>
  );
}
