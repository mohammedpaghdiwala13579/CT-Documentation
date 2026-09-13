import React, { useState } from "react";
import { 
  X, 
  FileSpreadsheet, 
  Download, 
  CheckCircle2, 
  Layers, 
  FileText, 
  Sparkles,
  Printer
} from "lucide-react";
import { QuotationRow, CompanyProfile } from "../types";
import { generateExcelWorkbook, downloadExcelFile, partitionItemsByPageCapacity } from "../utils/excelGenerator";
import { stripHtml } from "../utils/textFormatter";

export interface ExcelGeneratorModalProps {
  isOpen: boolean;
  onClose: () => void;
  docType: "quotation" | "challan" | "invoice";
  companyProfile: CompanyProfile;
  dateVal: string;
  messers: string;
  address: string;
  vesselName?: string;
  portBerth?: string;
  includeVesselName?: boolean;
  includePortBerth?: boolean;
  quotationNo?: string;
  challanNo?: string;
  invoiceNo?: string;
  requisitionNo?: string;
  poNumber?: string;
  includeInvoiceNo?: boolean;
  includeChallanNo?: boolean;
  includeQuotationNo?: boolean;
  includeRequisitionNo?: boolean;
  includePoNumber?: boolean;
  rows: QuotationRow[];
  currency?: string;
  vatPercent?: string | number;
  transportationFee?: string | number;
  includeDiscount?: boolean;
  discountType?: "percentage" | "fixed";
  discountValue?: string | number;
}

export default function ExcelGeneratorModal({
  isOpen,
  onClose,
  docType,
  companyProfile,
  dateVal,
  messers,
  address,
  vesselName = "",
  portBerth = "",
  includeVesselName = true,
  includePortBerth = true,
  quotationNo = "",
  challanNo = "",
  invoiceNo = "",
  requisitionNo = "",
  poNumber = "",
  includeInvoiceNo = true,
  includeChallanNo = true,
  includeQuotationNo = true,
  includeRequisitionNo = true,
  includePoNumber = true,
  rows,
  currency = "Taka",
  vatPercent = "0",
  transportationFee = "0",
  includeDiscount = false,
  discountType = "percentage",
  discountValue = "0",
}: ExcelGeneratorModalProps) {
  const [rowsPerPage, setRowsPerPage] = useState<number>(0); // 0 = Auto Capacity (dynamic based on available page space)
  const [padEmptyRows, setPadEmptyRows] = useState<boolean>(true);
  const [isGenerating, setIsGenerating] = useState<boolean>(false);
  const [isGenerated, setIsGenerated] = useState<boolean>(false);

  if (!isOpen) return null;

  // Active items calculation
  let lastNonEmptyIndex = -1;
  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i];
    if (
      (r.desc && stripHtml(r.desc).trim().length > 0) ||
      (r.qty && String(r.qty).trim().length > 0) ||
      (r.price && String(r.price).trim().length > 0) ||
      (r.amount && r.amount > 0)
    ) {
      lastNonEmptyIndex = i;
      break;
    }
  }

  const activeRows = lastNonEmptyIndex >= 0 ? rows.slice(0, lastNonEmptyIndex + 1) : rows.slice(0, 1);
  const activeCount = activeRows.length;

  // Calculate pages based on selected mode (Auto Capacity vs fixed)
  const autoSlices = partitionItemsByPageCapacity(activeRows, docType, includeDiscount);
  const totalSheets = rowsPerPage === 0 
    ? autoSlices.length 
    : Math.max(1, Math.ceil(activeCount / rowsPerPage));

  // Default filename
  const formatName = docType.toUpperCase();
  const refNo =
    docType === "invoice"
      ? invoiceNo || challanNo || "INV"
      : docType === "challan"
      ? challanNo || "CH"
      : quotationNo || "QT";
  const defaultFilename = `${formatName}_${refNo ? refNo.replace(/[/\\?%*:|"<>]/g, "-") + "_" : ""}${dateVal || "Document"}.xlsx`;

  const handleGenerateAndDownload = () => {
    setIsGenerating(true);
    try {
      const excelData = generateExcelWorkbook({
        docType,
        companyProfile,
        dateVal,
        messers,
        address,
        vesselName,
        portBerth,
        includeVesselName,
        includePortBerth,
        quotationNo,
        challanNo,
        invoiceNo,
        requisitionNo,
        poNumber,
        includeInvoiceNo,
        includeChallanNo,
        includeQuotationNo,
        includeRequisitionNo,
        includePoNumber,
        rows,
        currency,
        vatPercent,
        transportationFee,
        includeDiscount,
        discountType,
        discountValue,
        rowsPerPage,
        padEmptyRows,
      });

      downloadExcelFile(excelData, defaultFilename);
      setIsGenerated(true);
      setTimeout(() => {
        setIsGenerated(false);
        onClose();
      }, 1200);
    } catch (error) {
      console.error("Excel generation error:", error);
      alert("Failed to generate Excel file. Please check your data and try again.");
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-in fade-in duration-150">
      <div 
        className="bg-white rounded-xl shadow-2xl border border-slate-200 max-w-lg w-full overflow-hidden flex flex-col max-h-[90vh]"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div className="bg-emerald-700 px-5 py-4 text-white flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="h-9 w-9 bg-emerald-800 rounded-lg flex items-center justify-center shadow-inner">
              <FileSpreadsheet className="h-5 w-5 text-emerald-100" />
            </div>
            <div>
              <h3 className="font-bold text-base leading-tight flex items-center gap-2">
                <span>Excel Generator</span>
                <span className="text-[10px] uppercase font-mono px-2 py-0.5 bg-emerald-600 rounded-full tracking-wider border border-emerald-500">
                  {docType}
                </span>
              </h3>
              <p className="text-xs text-emerald-100">
                Letterhead-ready formatted Excel generator
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="text-emerald-200 hover:text-white p-1 rounded-md hover:bg-emerald-600/50 transition-colors cursor-pointer"
            title="Close"
          >
            <X className="h-5 w-5" />
          </button>
        </div>

        {/* Content */}
        <div className="p-5 space-y-4 overflow-y-auto">
          {/* Format Specification Checklist */}
          <div className="bg-slate-50 border border-slate-200 rounded-lg p-3.5 space-y-2">
            <div className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5 mb-1">
              <Sparkles className="w-3.5 h-3.5 text-emerald-600" />
              <span>Applied Format Rules</span>
            </div>
            <ul className="text-xs text-slate-600 space-y-1.5">
              <li className="flex items-start gap-2">
                <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                <span>
                  <strong>Row 1:</strong> Contains format name (<code>{docType.toUpperCase()}</code>)
                </span>
              </li>
              <li className="flex items-start gap-2">
                <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                <span>
                  <strong>Rows 2–11:</strong> Left completely blank (reserved space for pre-printed letterhead pad)
                </span>
              </li>
              <li className="flex items-start gap-2">
                <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                <span>
                  <strong>Listed Boxes:</strong> Customer, vessel, address, dates, and reference numbers neatly arranged
                </span>
              </li>
              <li className="flex items-start gap-2">
                <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                <span>
                  <strong>Description Column:</strong> Vertically compacted with text wrapping (no hidden or overwritten text)
                </span>
              </li>
              <li className="flex items-start gap-2">
                <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                <span>
                  <strong>Signature Block:</strong> Receiver & Authorized Signature placed on every page according to format
                </span>
              </li>
              <li className="flex items-start gap-2">
                <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                <span>
                  <strong>Multi-Sheet Pagination:</strong> Each page created as a new sheet once the current page is filled
                </span>
              </li>
            </ul>
          </div>

          {/* Pagination & Sheets Structure */}
          <div className="border border-slate-200 rounded-lg p-3.5 bg-white space-y-3">
            <div className="flex items-center justify-between">
              <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                <Layers className="w-3.5 h-3.5 text-blue-600" />
                <span>Page Capacity & Sheet Layout</span>
              </label>
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={() => setRowsPerPage(0)}
                  title="Automatically packs maximum items per sheet according to text heights"
                  className={`px-2.5 py-1 text-xs font-semibold rounded-md border transition-colors cursor-pointer ${
                    rowsPerPage === 0
                      ? "bg-emerald-600 text-white border-emerald-600 shadow-2xs font-bold"
                      : "bg-white text-slate-700 border-slate-300 hover:bg-slate-50"
                  }`}
                >
                  ⚡ Auto Fit
                </button>
                {[15, 18, 20, 25].map((val) => (
                  <button
                    key={val}
                    type="button"
                    onClick={() => setRowsPerPage(val)}
                    className={`px-2.5 py-1 text-xs font-semibold rounded-md border transition-colors cursor-pointer ${
                      rowsPerPage === val
                        ? "bg-emerald-600 text-white border-emerald-600 shadow-2xs font-bold"
                        : "bg-white text-slate-700 border-slate-300 hover:bg-slate-50"
                    }`}
                  >
                    {val}
                  </button>
                ))}
              </div>
            </div>

            {/* Sheets Preview Indicator */}
            <div className="bg-emerald-50/70 border border-emerald-200 rounded-md p-2.5 text-xs text-emerald-950">
              <div className="font-bold flex items-center justify-between mb-1">
                <span>Active Document Items: {activeCount}</span>
                <span className="bg-emerald-200/80 text-emerald-900 px-2 py-0.5 rounded font-mono text-[11px]">
                  {totalSheets} {totalSheets === 1 ? "Sheet" : "Sheets"}
                </span>
              </div>
              <div className="text-[11px] text-emerald-800 space-y-0.5 mt-1.5">
                {rowsPerPage === 0 ? (
                  autoSlices.map((slice, idx) => {
                    const pNum = idx + 1;
                    const startItem = slice.startIndex + 1;
                    const endItem = slice.endIndex;
                    const isLast = pNum === autoSlices.length;
                    return (
                      <div key={pNum} className="flex items-center justify-between">
                        <span className="font-semibold">
                          Sheet {pNum} (&quot;Page {pNum}&quot;):
                        </span>
                        <span>
                          Items {startItem}–{endItem} ({slice.items.length} items) {isLast ? "+ Totals & Signatures" : "+ Page Subtotal & Signatures"}
                        </span>
                      </div>
                    );
                  })
                ) : (
                  Array.from({ length: totalSheets }, (_, idx) => {
                    const pNum = idx + 1;
                    const startItem = idx * rowsPerPage + 1;
                    const endItem = Math.min(activeCount, pNum * rowsPerPage);
                    return (
                      <div key={pNum} className="flex items-center justify-between">
                        <span className="font-semibold">
                          Sheet {pNum} (&quot;Page {pNum}&quot;):
                        </span>
                        <span>
                          Items {startItem}–{endItem} {pNum === totalSheets ? "+ Grand Total & Signatures" : "+ Page Total & Signatures"}
                        </span>
                      </div>
                    );
                  })
                )}
              </div>
            </div>

            {/* Pad empty rows option (only relevant when fixed rowsPerPage is selected) */}
            {rowsPerPage > 0 && (
              <label className="flex items-center gap-2 text-xs text-slate-700 cursor-pointer pt-1">
                <input
                  type="checkbox"
                  checked={padEmptyRows}
                  onChange={(e) => setPadEmptyRows(e.target.checked)}
                  className="rounded border-slate-300 text-emerald-600 focus:ring-emerald-500"
                />
                <span>Fill empty rows on final page to maintain uniform table size</span>
              </label>
            )}
          </div>

          {/* Filename Preview */}
          <div className="text-xs text-slate-500">
            <span className="font-semibold text-slate-700">Output File: </span>
            <span className="font-mono bg-slate-100 px-1.5 py-0.5 rounded text-slate-800 border border-slate-200">
              {defaultFilename}
            </span>
          </div>
        </div>

        {/* Footer Actions */}
        <div className="p-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between gap-3">
          <button
            type="button"
            onClick={onClose}
            className="px-3.5 py-2 text-xs font-semibold text-slate-700 hover:text-slate-900 hover:bg-slate-200/60 rounded-md transition-colors cursor-pointer"
          >
            Cancel
          </button>
          
          <button
            type="button"
            id="btn-confirm-generate-excel"
            onClick={handleGenerateAndDownload}
            disabled={isGenerating || isGenerated}
            className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-md text-xs font-bold flex items-center gap-2 shadow-sm transition-all cursor-pointer disabled:opacity-75"
          >
            {isGenerating ? (
              <>
                <div className="w-3.5 h-3.5 border-2 border-white border-t-transparent rounded-full animate-spin" />
                <span>Generating Workbook...</span>
              </>
            ) : isGenerated ? (
              <>
                <CheckCircle2 className="w-4 h-4 text-emerald-200" />
                <span>Downloaded Successfully!</span>
              </>
            ) : (
              <>
                <Download className="w-4 h-4" />
                <span>Generate Excel Workbook (.xlsx)</span>
              </>
            )}
          </button>
        </div>
      </div>
    </div>
  );
}
