import XLSX from "xlsx-js-style";
import { QuotationRow, CompanyProfile } from "../types";
import { numberToWords } from "./numberToWords";
import { stripHtml, parseNumericInput } from "./textFormatter";

export interface ExcelGeneratorOptions {
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
  rowsPerPage?: number;
  padEmptyRows?: boolean;
}

// Border presets
const thinBorder = {
  top: { style: "thin", color: { rgb: "000000" } },
  bottom: { style: "thin", color: { rgb: "000000" } },
  left: { style: "thin", color: { rgb: "000000" } },
  right: { style: "thin", color: { rgb: "000000" } },
};

const topSignatureBorder = {
  top: { style: "medium", color: { rgb: "000000" } },
  bottom: { style: "none" },
  left: { style: "none" },
  right: { style: "none" },
};

// Set single cell value and style
function setCell(
  ws: Record<string, any>,
  r: number,
  c: number,
  value: any,
  style?: any,
  type: string = "s"
) {
  const cellRef = XLSX.utils.encode_cell({ r, c });
  ws[cellRef] = {
    v: value ?? "",
    t: type,
    s: style || {},
  };
}

// Set merged cell range and apply borders to all cells in the rectangle
function mergeAndStyleRange(
  ws: Record<string, any>,
  merges: XLSX.Range[],
  sR: number,
  sC: number,
  eR: number,
  eC: number,
  value: any,
  style: any,
  type: string = "s"
) {
  merges.push({ s: { r: sR, c: sC }, e: { r: eR, c: eC } });
  for (let r = sR; r <= eR; r++) {
    for (let c = sC; c <= eC; c++) {
      const isTopLeft = r === sR && c === sC;
      setCell(ws, r, c, isTopLeft ? value : "", style, type);
    }
  }
}

// Calculate dynamic row height strictly according to actual text content
// Single-line text is compacted to 18pt, multi-line calculates visual wrap lines
export function calculateDescriptionRowHeight(text: string, colWidthChars: number = 48): number {
  if (!text || text.trim() === "") return 18;

  const lines = text.split(/\r\n|\r|\n/);
  let totalVisualLines = 0;

  for (const line of lines) {
    const visual = Math.max(1, Math.ceil(line.length / colWidthChars));
    totalVisualLines += visual;
  }

  // Single-line items are vertically compacted to 18pt (clean, crisp, no empty space)
  if (totalVisualLines <= 1) {
    return 18;
  }

  // Multi-line expands tightly according to content: ~14.5pt per line + 4pt padding
  return Math.min(140, Math.max(18, Math.round(totalVisualLines * 14.5 + 4)));
}

// Fixed heights (in points) of sections according to the exact layout format:
// Top row (title): 24pt
// Rows 2-11 (letterhead padding): 10 rows * 13pt = 130pt
// Metadata boxes (Rows 12-14): ~64pt
// Spacer row 15: 6pt
// Table header row: 22pt
// Total fixed header space = 246pt
const FIXED_HEADER_SPACE_PT = 246;

// Fixed Signature Block on EVERY page:
// Spacer before sig: 8pt
// "For <Company>" (in quotation/invoice): 16pt (0pt in challan)
// Signature line (Receiver & Authorized Signature): 36pt
// Spacer: 4pt
// Non-returnable notice: 14pt
// Total fixed signature space = ~78pt (62pt for challan)
const FIXED_SIGNATURE_SPACE_PT_QUOTATION = 78;
const FIXED_SIGNATURE_SPACE_PT_CHALLAN = 62;

// Totals section space (on final page for quotation/invoice):
// Quotation: 1 row = 22pt
// Invoice: 4 to 5 rows = 4 * 18 + 22 = 94pt to 112pt
// Challan: 0pt (no prices/totals)
// Intermediate page subtotal row: 20pt

// Usable printable vertical points on an A4 sheet with 0.25in margins:
// A4 height = 842 pt. Printable area = 842 - 36 (margins) ≈ 790 pt.
// Available space for items + totals on a page = 790 - FIXED_HEADER_SPACE_PT - FIXED_SIGNATURE_SPACE_PT ≈ 466 pt (Quotation) / 482 pt (Challan).
const TOTAL_PAGE_USABLE_PT = 765;

export interface PageItemSlice {
  startIndex: number;
  endIndex: number;
  items: QuotationRow[];
}

// Dynamic capacity-based placement: calculates the available vertical space on each page
// and packs the maximum possible items without exceeding page limit or splitting items.
export function partitionItemsByPageCapacity(
  items: QuotationRow[],
  docType: "quotation" | "challan" | "invoice",
  includeDiscount: boolean = false
): PageItemSlice[] {
  if (items.length === 0) {
    return [{ startIndex: 0, endIndex: 0, items: [] }];
  }

  const isChallan = docType === "challan";
  const sigSpace = isChallan ? FIXED_SIGNATURE_SPACE_PT_CHALLAN : FIXED_SIGNATURE_SPACE_PT_QUOTATION;
  const colWidth = isChallan ? 62 : 48;

  // Space reserved for totals on the LAST page:
  const lastPageTotalsSpace = isChallan
    ? 0
    : docType === "invoice"
    ? includeDiscount ? 104 : 86
    : 24;

  // Space reserved for page subtotal on INTERMEDIATE pages:
  const intermediateTotalsSpace = isChallan ? 0 : 20;

  const pages: PageItemSlice[] = [];
  let itemIdx = 0;
  const totalItems = items.length;

  while (itemIdx < totalItems) {
    // Check if remaining items can all fit onto this page as the last page
    let testH = FIXED_HEADER_SPACE_PT + sigSpace + lastPageTotalsSpace;
    let canFitAllRemaining = true;
    let tempIdx = itemIdx;

    while (tempIdx < totalItems) {
      const h = calculateDescriptionRowHeight(stripHtml(items[tempIdx].desc || ""), colWidth);
      if (testH + h <= TOTAL_PAGE_USABLE_PT) {
        testH += h;
        tempIdx++;
      } else {
        canFitAllRemaining = false;
        break;
      }
    }

    if (canFitAllRemaining) {
      // All remaining fit neatly onto this final page!
      pages.push({
        startIndex: itemIdx,
        endIndex: totalItems,
        items: items.slice(itemIdx, totalItems),
      });
      break;
    }

    // Otherwise, this is an intermediate page: pack maximum possible items
    let pageH = FIXED_HEADER_SPACE_PT + sigSpace + intermediateTotalsSpace;
    const startOfPage = itemIdx;

    while (itemIdx < totalItems) {
      const h = calculateDescriptionRowHeight(stripHtml(items[itemIdx].desc || ""), colWidth);
      // If adding this item exceeds the page limit, stop here and move the entire item to the next page
      if (pageH + h > TOTAL_PAGE_USABLE_PT && itemIdx > startOfPage) {
        break;
      }
      pageH += h;
      itemIdx++;
    }

    pages.push({
      startIndex: startOfPage,
      endIndex: itemIdx,
      items: items.slice(startOfPage, itemIdx),
    });
  }

  return pages;
}

export function generateExcelWorkbook(options: ExcelGeneratorOptions): Uint8Array {
  const {
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
    rowsPerPage = 20,
    padEmptyRows = true,
  } = options;

  const isChallan = docType === "challan";
  const numCols = isChallan ? 4 : 6; // Challan: A-D (4), Quotation/Invoice: A-F (6)
  const lastColIdx = numCols - 1;

  // Clean rows
  const cleanMessers = stripHtml(messers || "");
  const cleanAddress = stripHtml(address || "");
  const cleanVessel = stripHtml(vesselName || "");
  const cleanPort = stripHtml(portBerth || "");

  // Find actual active rows (rows with content)
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

  // Financial calculations matching QuotationBuilder exactly
  const rowsTotal = activeRows.reduce(
    (acc, r) => acc + (isChallan ? 0 : r.amount || 0),
    0
  );
  const parsedDiscountValue = parseNumericInput(discountValue);
  const discountAmount = !includeDiscount
    ? 0
    : discountType === "percentage"
    ? (rowsTotal * parsedDiscountValue) / 100
    : parsedDiscountValue;
  const parsedVatPercent = parseNumericInput(vatPercent);
  const vatAmount = docType === "invoice" ? ((rowsTotal - discountAmount) * parsedVatPercent) / 100 : 0;
  const parsedTransportationFee = docType === "invoice" ? parseNumericInput(transportationFee) : 0;
  const grandTotal = Math.max(0, rowsTotal - discountAmount + vatAmount + parsedTransportationFee);

  // Automatic Capacity-Based Pagination vs Fixed Rows Per Page:
  // If rowsPerPage is 0 or negative, use dynamic capacity placement to maximize items per sheet cleanly.
  // Otherwise, if a positive rowsPerPage is provided, partition according to that count while still
  // respecting page boundaries and moving entire items without cutting them.
  let pageSlices: PageItemSlice[] = [];
  if (!rowsPerPage || rowsPerPage <= 0) {
    pageSlices = partitionItemsByPageCapacity(activeRows, docType, includeDiscount);
  } else {
    // Fixed partition mode
    const totalPagesFixed = Math.max(1, Math.ceil(activeRows.length / rowsPerPage));
    for (let p = 0; p < totalPagesFixed; p++) {
      const s = p * rowsPerPage;
      const e = Math.min(activeRows.length, (p + 1) * rowsPerPage);
      pageSlices.push({
        startIndex: s,
        endIndex: e,
        items: activeRows.slice(s, e),
      });
    }
  }

  const totalPages = pageSlices.length;

  const wb = XLSX.utils.book_new();

  for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
    const ws: Record<string, any> = {};
    const merges: XLSX.Range[] = [];
    const rowHeights: { hpt: number }[] = [];

    // Column widths setup
    if (isChallan) {
      ws["!cols"] = [
        { wch: 8 },  // Col A: SL
        { wch: 62 }, // Col B: Description (spacious & compacted)
        { wch: 12 }, // Col C: Qty
        { wch: 12 }, // Col D: Unit
      ];
    } else {
      ws["!cols"] = [
        { wch: 7 },  // Col A: SL
        { wch: 48 }, // Col B: Description (spacious & compacted)
        { wch: 9 },  // Col C: Qty
        { wch: 9 },  // Col D: Unit
        { wch: 13 }, // Col E: Price
        { wch: 16 }, // Col F: Amount
      ];
    }

    // A4 Portrait Print setup
    ws["!pageSetup"] = {
      orientation: "portrait",
      paperSize: 9, // A4
      fitToWidth: 1,
      fitToHeight: 0,
    };
    ws["!margins"] = {
      left: 0.25,
      right: 0.25,
      top: 0.35,
      bottom: 0.35,
      header: 0.15,
      footer: 0.15,
    };

    // ==========================================
    // 1. ROW 1: NAME OF THE FORMAT (NO BUSINESS HEADER)
    // ==========================================
    const formatTitle =
      (docType === "challan" ? "CHALLAN" : docType === "invoice" ? "INVOICE" : "QUOTATION") +
      (totalPages > 1 ? ` (PAGE ${pageNum} OF ${totalPages})` : "");

    mergeAndStyleRange(
      ws,
      merges,
      0,
      0,
      0,
      lastColIdx,
      formatTitle,
      {
        font: { name: "Arial", sz: 14, bold: true, color: { rgb: "000000" } },
        alignment: { horizontal: "center", vertical: "center" },
      }
    );
    rowHeights[0] = { hpt: 26 };

    // ==========================================
    // 2. ROWS 2-11: LEAVE SPACE FOR 1-11 ROWS BLANK
    // (Indices 1 to 10 = exactly 10 blank spacer rows)
    // ==========================================
    for (let r = 1; r <= 10; r++) {
      rowHeights[r] = { hpt: 18 };
      // Leave cells unpopulated/blank for official pre-printed letterhead pad space
    }

    // ==========================================
    // 3. ROWS 12-15: THE LISTED BOXES OF THE FORMAT
    // ==========================================
    // Left Box: Columns A to B (Messers, Vessel/Port, Address)
    // Right Box: Columns C to lastColIdx (Date, Quotation/Challan/Invoice No, Requisition, PO)

    // Row 12 (Index 11):
    // Left: Messers
    setCell(
      ws,
      11,
      0,
      "Messers:",
      {
        font: { name: "Arial", sz: 8.5, bold: true },
        fill: { fgColor: { rgb: "F1F5F9" } },
        alignment: { horizontal: "left", vertical: "center" },
        border: thinBorder,
      }
    );
    setCell(
      ws,
      11,
      1,
      cleanMessers,
      {
        font: { name: "Arial", sz: 9, bold: true },
        alignment: { horizontal: "left", vertical: "center", wrapText: true },
        border: thinBorder,
      }
    );

    // Right Box Row 12:
    if (isChallan) {
      setCell(ws, 11, 2, "Date:", {
        font: { name: "Arial", sz: 8.5, bold: true },
        fill: { fgColor: { rgb: "F1F5F9" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
      setCell(ws, 11, 3, dateVal, {
        font: { name: "Arial", sz: 9 },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
    } else if (docType === "invoice") {
      setCell(ws, 11, 2, "Inv No:", {
        font: { name: "Arial", sz: 8, bold: true },
        fill: { fgColor: { rgb: "F1F5F9" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
      setCell(ws, 11, 3, includeInvoiceNo ? invoiceNo : "", {
        font: { name: "Arial", sz: 8.5, bold: true },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
      setCell(ws, 11, 4, "Date:", {
        font: { name: "Arial", sz: 8.5, bold: true },
        fill: { fgColor: { rgb: "F1F5F9" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
      setCell(ws, 11, 5, dateVal, {
        font: { name: "Arial", sz: 9 },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
    } else {
      // Quotation
      setCell(ws, 11, 2, "Date:", {
        font: { name: "Arial", sz: 8.5, bold: true },
        fill: { fgColor: { rgb: "F1F5F9" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
      setCell(ws, 11, 3, dateVal, {
        font: { name: "Arial", sz: 9 },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
      setCell(ws, 11, 4, "Quote No:", {
        font: { name: "Arial", sz: 8, bold: true },
        fill: { fgColor: { rgb: "F1F5F9" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
      setCell(ws, 11, 5, includeQuotationNo ? quotationNo : "", {
        font: { name: "Arial", sz: 8.5 },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
    }
    rowHeights[11] = { hpt: 22 };

    // Row 13 (Index 12):
    // Left: Vessel / Port / Berth
    const vesselPortText = [
      includeVesselName && cleanVessel ? `Vessel: ${cleanVessel}` : "",
      includePortBerth && cleanPort ? `Port/Berth: ${cleanPort}` : "",
    ]
      .filter(Boolean)
      .join("  |  ");

    setCell(
      ws,
      12,
      0,
      "Vessel/Port:",
      {
        font: { name: "Arial", sz: 8, bold: true },
        fill: { fgColor: { rgb: "F1F5F9" } },
        alignment: { horizontal: "left", vertical: "center" },
        border: thinBorder,
      }
    );
    setCell(
      ws,
      12,
      1,
      vesselPortText || "-",
      {
        font: { name: "Arial", sz: 8.5 },
        alignment: { horizontal: "left", vertical: "center", wrapText: true },
        border: thinBorder,
      }
    );

    // Right Box Row 13:
    if (isChallan) {
      setCell(ws, 12, 2, "Challan No:", {
        font: { name: "Arial", sz: 8, bold: true },
        fill: { fgColor: { rgb: "F1F5F9" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
      setCell(ws, 12, 3, includeChallanNo ? challanNo : "", {
        font: { name: "Arial", sz: 8.5 },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
    } else if (docType === "invoice") {
      setCell(ws, 12, 2, "Challan:", {
        font: { name: "Arial", sz: 8, bold: true },
        fill: { fgColor: { rgb: "F1F5F9" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
      setCell(ws, 12, 3, includeChallanNo ? challanNo : "", {
        font: { name: "Arial", sz: 8.5 },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
      setCell(ws, 12, 4, "Req No:", {
        font: { name: "Arial", sz: 8, bold: true },
        fill: { fgColor: { rgb: "F1F5F9" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
      setCell(ws, 12, 5, includeRequisitionNo ? requisitionNo : "", {
        font: { name: "Arial", sz: 8.5 },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
    } else {
      // Quotation
      setCell(ws, 12, 2, "Req No:", {
        font: { name: "Arial", sz: 8, bold: true },
        fill: { fgColor: { rgb: "F1F5F9" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
      mergeAndStyleRange(
        ws,
        merges,
        12,
        3,
        12,
        5,
        includeRequisitionNo ? requisitionNo : "",
        {
          font: { name: "Arial", sz: 8.5 },
          alignment: { horizontal: "left", vertical: "center" },
          border: thinBorder,
        }
      );
    }
    rowHeights[12] = { hpt: 20 };

    // Row 14 (Index 13):
    // Left: Address
    setCell(
      ws,
      13,
      0,
      "Address:",
      {
        font: { name: "Arial", sz: 8.5, bold: true },
        fill: { fgColor: { rgb: "F1F5F9" } },
        alignment: { horizontal: "left", vertical: "center" },
        border: thinBorder,
      }
    );
    setCell(
      ws,
      13,
      1,
      cleanAddress,
      {
        font: { name: "Arial", sz: 8.5 },
        alignment: { horizontal: "left", vertical: "center", wrapText: true },
        border: thinBorder,
      }
    );

    // Right Box Row 14:
    if (isChallan) {
      setCell(ws, 13, 2, "Req No:", {
        font: { name: "Arial", sz: 8, bold: true },
        fill: { fgColor: { rgb: "F1F5F9" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
      setCell(ws, 13, 3, includeRequisitionNo ? requisitionNo : "", {
        font: { name: "Arial", sz: 8.5 },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
    } else if (docType === "invoice") {
      setCell(ws, 13, 2, "PO No:", {
        font: { name: "Arial", sz: 8, bold: true },
        fill: { fgColor: { rgb: "F1F5F9" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });
      mergeAndStyleRange(
        ws,
        merges,
        13,
        3,
        13,
        5,
        includePoNumber ? poNumber : "",
        {
          font: { name: "Arial", sz: 8.5 },
          alignment: { horizontal: "left", vertical: "center" },
          border: thinBorder,
        }
      );
    } else {
      // Quotation: blank/border
      mergeAndStyleRange(
        ws,
        merges,
        13,
        2,
        13,
        5,
        "",
        {
          border: thinBorder,
        }
      );
    }
    // Calculate address row height so long delivery address is not truncated
    const addressHeight = Math.max(22, calculateDescriptionRowHeight(cleanAddress, 48));
    rowHeights[13] = { hpt: addressHeight };

    // Row 15 (Index 14): Spacer between metadata boxes and table
    rowHeights[14] = { hpt: 8 };

    // ==========================================
    // 4. TABLE HEADER (Index 15)
    // ==========================================
    const tableHeaderRowIndex = 15;
    const headerStyle = {
      font: { name: "Arial", sz: 9.5, bold: true, color: { rgb: "000000" } },
      fill: { fgColor: { rgb: "E2E8F0" } },
      alignment: { horizontal: "center", vertical: "center" },
      border: thinBorder,
    };

    setCell(ws, tableHeaderRowIndex, 0, "SL", headerStyle);
    setCell(
      ws,
      tableHeaderRowIndex,
      1,
      "Description of Marine Items / Spare Parts",
      {
        ...headerStyle,
        alignment: { horizontal: "left", vertical: "center" },
      }
    );
    setCell(ws, tableHeaderRowIndex, 2, "Qty", headerStyle);
    setCell(ws, tableHeaderRowIndex, 3, "Unit", headerStyle);

    if (!isChallan) {
      setCell(ws, tableHeaderRowIndex, 4, "Price", headerStyle);
      setCell(ws, tableHeaderRowIndex, 5, "Amount", headerStyle);
    }
    rowHeights[tableHeaderRowIndex] = { hpt: 24 };

    // ==========================================
    // 5. TABLE DATA ROWS FOR THIS PAGE
    // ==========================================
    const currentSlice = pageSlices[pageNum - 1] || { startIndex: 0, endIndex: 0, items: [] };
    const startIndex = currentSlice.startIndex;
    const pageItems = currentSlice.items;

    let currentRow = tableHeaderRowIndex + 1;
    let pageItemsSubtotal = 0;

    for (let i = 0; i < pageItems.length; i++) {
      const item = pageItems[i];
      const slNumber = startIndex + i + 1;
      const cleanDesc = stripHtml(item.desc || "");
      const cleanQty = stripHtml(item.qty || "");
      const cleanUnit = stripHtml(item.unit || "");
      const cleanPrice = stripHtml(item.price || "");
      const amountVal = item.amount || 0;
      pageItemsSubtotal += amountVal;

      // SL
      setCell(ws, currentRow, 0, slNumber, {
        font: { name: "Arial", sz: 9 },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      }, "n");

      // Description: compacted vertically and text wrapped so nothing is hidden
      const descHeight = calculateDescriptionRowHeight(cleanDesc, isChallan ? 60 : 48);
      setCell(ws, currentRow, 1, cleanDesc, {
        font: { name: "Arial", sz: 9 },
        alignment: { horizontal: "left", vertical: "center", wrapText: true },
        border: thinBorder,
      });

      // Qty
      setCell(ws, currentRow, 2, cleanQty, {
        font: { name: "Arial", sz: 9 },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });

      // Unit
      setCell(ws, currentRow, 3, cleanUnit, {
        font: { name: "Arial", sz: 9 },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder,
      });

      if (!isChallan) {
        // Price
        const parsedPrice = parseNumericInput(cleanPrice);
        if (!isNaN(parsedPrice) && parsedPrice > 0) {
          setCell(ws, currentRow, 4, parsedPrice, {
            font: { name: "Arial", sz: 9 },
            alignment: { horizontal: "right", vertical: "center" },
            numFmt: "#,##0.00",
            border: thinBorder,
          }, "n");
        } else {
          setCell(ws, currentRow, 4, cleanPrice, {
            font: { name: "Arial", sz: 9 },
            alignment: { horizontal: "right", vertical: "center" },
            border: thinBorder,
          });
        }

        // Amount
        setCell(ws, currentRow, 5, amountVal, {
          font: { name: "Arial", sz: 9, bold: true },
          alignment: { horizontal: "right", vertical: "center" },
          numFmt: "#,##0.00",
          border: thinBorder,
        }, "n");
      }

      rowHeights[currentRow] = { hpt: descHeight };
      currentRow++;
    }

    // Pad empty rows if requested and fixed rowsPerPage is enabled so every page is filled to its capacity
    if (padEmptyRows && rowsPerPage > 0 && pageItems.length < rowsPerPage) {
      const neededPadding = rowsPerPage - pageItems.length;
      for (let pad = 0; pad < neededPadding; pad++) {
        const slNumber = startIndex + pageItems.length + pad + 1;
        setCell(ws, currentRow, 0, slNumber, {
          font: { name: "Arial", sz: 9, color: { rgb: "94A3B8" } },
          alignment: { horizontal: "center", vertical: "center" },
          border: thinBorder,
        }, "n");

        setCell(ws, currentRow, 1, "", { border: thinBorder });
        setCell(ws, currentRow, 2, "", { border: thinBorder });
        setCell(ws, currentRow, 3, "", { border: thinBorder });

        if (!isChallan) {
          setCell(ws, currentRow, 4, "", { border: thinBorder });
          setCell(ws, currentRow, 5, "", { border: thinBorder });
        }

        rowHeights[currentRow] = { hpt: 18 };
        currentRow++;
      }
    }

    // ==========================================
    // 6. TOTALS / CLOSING SECTION (FOR FINAL PAGE OR PAGE SUBTOTAL)
    // ==========================================
    const isLastPage = pageNum === totalPages;

    if (!isChallan) {
      if (isLastPage) {
        // FINAL PAGE TOTALS
        if (docType === "invoice") {
          // INVOICE SUMMARY
          const totalRowsCount = includeDiscount ? 5 : 4;
          const wordsStartRow = currentRow;
          const wordsEndRow = currentRow + totalRowsCount - 1;

          // Amount in words (merged across Cols A to D)
          const wordsText = `Amount in Words: ${numberToWords(grandTotal, currency)}`;
          mergeAndStyleRange(
            ws,
            merges,
            wordsStartRow,
            0,
            wordsEndRow,
            3,
            wordsText,
            {
              font: { name: "Arial", sz: 8.5, bold: true, italic: true },
              fill: { fgColor: { rgb: "F8FAFC" } },
              alignment: { horizontal: "left", vertical: "center", wrapText: true },
              border: thinBorder,
            }
          );

          // Subtotal
          setCell(ws, currentRow, 4, "SUBTOTAL =", {
            font: { name: "Arial", sz: 8.5, bold: true },
            fill: { fgColor: { rgb: "F1F5F9" } },
            alignment: { horizontal: "right", vertical: "center" },
            border: thinBorder,
          });
          setCell(ws, currentRow, 5, rowsTotal, {
            font: { name: "Arial", sz: 9, bold: true },
            alignment: { horizontal: "right", vertical: "center" },
            numFmt: "#,##0.00",
            border: thinBorder,
          }, "n");
          rowHeights[currentRow] = { hpt: 20 };
          currentRow++;

          // Discount (if included)
          if (includeDiscount) {
            const discLabel = `DISCOUNT${discountType === "percentage" ? ` (${parsedDiscountValue}%)` : ""} =`;
            setCell(ws, currentRow, 4, discLabel, {
              font: { name: "Arial", sz: 8.5, bold: true },
              fill: { fgColor: { rgb: "F1F5F9" } },
              alignment: { horizontal: "right", vertical: "center" },
              border: thinBorder,
            });
            setCell(ws, currentRow, 5, -discountAmount, {
              font: { name: "Arial", sz: 9, bold: true, color: { rgb: "DC2626" } },
              alignment: { horizontal: "right", vertical: "center" },
              numFmt: "#,##0.00",
              border: thinBorder,
            }, "n");
            rowHeights[currentRow] = { hpt: 20 };
            currentRow++;
          }

          // VAT
          setCell(ws, currentRow, 4, `VAT (${parsedVatPercent}%) =`, {
            font: { name: "Arial", sz: 8.5, bold: true },
            fill: { fgColor: { rgb: "F1F5F9" } },
            alignment: { horizontal: "right", vertical: "center" },
            border: thinBorder,
          });
          setCell(ws, currentRow, 5, vatAmount, {
            font: { name: "Arial", sz: 9 },
            alignment: { horizontal: "right", vertical: "center" },
            numFmt: "#,##0.00",
            border: thinBorder,
          }, "n");
          rowHeights[currentRow] = { hpt: 20 };
          currentRow++;

          // Transportation
          setCell(ws, currentRow, 4, "TRANS. =", {
            font: { name: "Arial", sz: 8.5, bold: true },
            fill: { fgColor: { rgb: "F1F5F9" } },
            alignment: { horizontal: "right", vertical: "center" },
            border: thinBorder,
          });
          setCell(ws, currentRow, 5, parsedTransportationFee, {
            font: { name: "Arial", sz: 9 },
            alignment: { horizontal: "right", vertical: "center" },
            numFmt: "#,##0.00",
            border: thinBorder,
          }, "n");
          rowHeights[currentRow] = { hpt: 20 };
          currentRow++;

          // Grand Total
          setCell(ws, currentRow, 4, "GRAND TOTAL =", {
            font: { name: "Arial", sz: 9.5, bold: true, color: { rgb: "1E1B4B" } },
            fill: { fgColor: { rgb: "EEF2FF" } },
            alignment: { horizontal: "right", vertical: "center" },
            border: thinBorder,
          });
          setCell(ws, currentRow, 5, grandTotal, {
            font: { name: "Arial", sz: 10, bold: true, color: { rgb: "1E1B4B" } },
            fill: { fgColor: { rgb: "EEF2FF" } },
            alignment: { horizontal: "right", vertical: "center" },
            numFmt: "#,##0.00",
            border: thinBorder,
          }, "n");
          rowHeights[currentRow] = { hpt: 24 };
          currentRow++;
        } else {
          // QUOTATION SUMMARY
          const wordsText = `Amount in Words: ${numberToWords(grandTotal, currency)}`;
          mergeAndStyleRange(
            ws,
            merges,
            currentRow,
            0,
            currentRow,
            3,
            wordsText,
            {
              font: { name: "Arial", sz: 8.5, bold: true, italic: true },
              fill: { fgColor: { rgb: "F8FAFC" } },
              alignment: { horizontal: "left", vertical: "center", wrapText: true },
              border: thinBorder,
            }
          );
          setCell(ws, currentRow, 4, "TOTAL =", {
            font: { name: "Arial", sz: 9, bold: true },
            fill: { fgColor: { rgb: "F1F5F9" } },
            alignment: { horizontal: "right", vertical: "center" },
            border: thinBorder,
          });
          setCell(ws, currentRow, 5, grandTotal, {
            font: { name: "Arial", sz: 10, bold: true },
            alignment: { horizontal: "right", vertical: "center" },
            numFmt: "#,##0.00",
            border: thinBorder,
          }, "n");
          rowHeights[currentRow] = { hpt: 24 };
          currentRow++;
        }
      } else {
        // INTERMEDIATE PAGE: SHOW PAGE TOTAL & CARRY FORWARD
        mergeAndStyleRange(
          ws,
          merges,
          currentRow,
          0,
          currentRow,
          3,
          `(Items carried forward to Page ${pageNum + 1}...)`,
          {
            font: { name: "Arial", sz: 8.5, italic: true },
            fill: { fgColor: { rgb: "F8FAFC" } },
            alignment: { horizontal: "left", vertical: "center" },
            border: thinBorder,
          }
        );
        setCell(ws, currentRow, 4, "PAGE TOTAL =", {
          font: { name: "Arial", sz: 8.5, bold: true },
          fill: { fgColor: { rgb: "F1F5F9" } },
          alignment: { horizontal: "right", vertical: "center" },
          border: thinBorder,
        });
        setCell(ws, currentRow, 5, pageItemsSubtotal, {
          font: { name: "Arial", sz: 9, bold: true },
          alignment: { horizontal: "right", vertical: "center" },
          numFmt: "#,##0.00",
          border: thinBorder,
        }, "n");
        rowHeights[currentRow] = { hpt: 22 };
        currentRow++;
      }
    }

    // ==========================================
    // 7. SIGNATURE BLOCK (ACCORDING TO FORMAT ON EVERY PAGE)
    // ==========================================
    // 1 Spacer row
    rowHeights[currentRow] = { hpt: 12 };
    currentRow++;

    // "For <Company Name>" row
    if (!isChallan) {
      const companyForLabel = `For ${companyProfile.name}`;
      mergeAndStyleRange(
        ws,
        merges,
        currentRow,
        isChallan ? 2 : 4,
        currentRow,
        lastColIdx,
        companyForLabel,
        {
          font: { name: "Arial", sz: 9, bold: true },
          alignment: { horizontal: "center", vertical: "center" },
        }
      );
      rowHeights[currentRow] = { hpt: 18 };
      currentRow++;
    }

    // Signature lines with top border for physical signature
    const sigLineRow = currentRow;
    rowHeights[sigLineRow] = { hpt: 38 };

    // Receiver Signature on left (Cols A-B)
    mergeAndStyleRange(
      ws,
      merges,
      sigLineRow,
      0,
      sigLineRow,
      1,
      "Receiver's Signature",
      {
        font: { name: "Arial", sz: 8.5, bold: true },
        alignment: { horizontal: "center", vertical: "bottom" },
        border: topSignatureBorder,
      }
    );

    // Authorized Signature on right (Cols E-F or C-D)
    mergeAndStyleRange(
      ws,
      merges,
      sigLineRow,
      isChallan ? 2 : 4,
      sigLineRow,
      lastColIdx,
      "Authorized Signature",
      {
        font: { name: "Arial", sz: 8.5, bold: true },
        alignment: { horizontal: "center", vertical: "bottom" },
        border: topSignatureBorder,
      }
    );
    currentRow++;

    // Non-returnable & non-exchangeable notice
    rowHeights[currentRow] = { hpt: 6 };
    currentRow++;

    mergeAndStyleRange(
      ws,
      merges,
      currentRow,
      0,
      currentRow,
      lastColIdx,
      "ITEMS ONCE SOLD ARE NON-RETURNABLE AND NON-EXCHANGEABLE.",
      {
        font: { name: "Arial", sz: 7.5, bold: true, color: { rgb: "475569" } },
        alignment: { horizontal: "center", vertical: "center" },
      }
    );
    rowHeights[currentRow] = { hpt: 16 };

    // Finalize sheet
    ws["!merges"] = merges;
    ws["!rows"] = rowHeights;
    ws["!ref"] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: currentRow, c: lastColIdx },
    });

    const sheetName = totalPages === 1 ? "Page 1" : `Page ${pageNum}`;
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  }

  // Generate binary XLSX output
  const wbOut = XLSX.write(wb, {
    bookType: "xlsx",
    type: "array",
  });

  return new Uint8Array(wbOut);
}

// Download helper to trigger browser download
export function downloadExcelFile(data: Uint8Array, filename: string) {
  const blob = new Blob([data], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename.endsWith(".xlsx") ? filename : `${filename}.xlsx`;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 200);
}
