import ExcelJS from "exceljs";
import { CompanyProfile, QuotationRow } from "../types";
import { numberToWords } from "./numberToWords";

export interface ExcelGeneratorOptions {
  currentCompany: CompanyProfile;
  docType: "quotation" | "challan" | "invoice";
  messers: string;
  address: string;
  vesselName?: string;
  portBerth?: string;
  includeVesselName?: boolean;
  includePortBerth?: boolean;
  invoiceNo?: string;
  challanNo?: string;
  quotationNo?: string;
  requisitionNo?: string;
  poNumber?: string;
  includeInvoiceNo?: boolean;
  includeChallanNo?: boolean;
  includeQuotationNo?: boolean;
  includeRequisitionNo?: boolean;
  includePoNumber?: boolean;
  dateVal: string;
  rows: QuotationRow[];
  includeDiscount?: boolean;
  discountType?: "percentage" | "fixed";
  discountValue?: string | number;
  vatPercent?: string | number;
  transportationFee?: string | number;
  currency?: string;
  currencySymbol?: string;
}

/**
 * Strips HTML tags, decodes standard HTML entities, and normalizes line breaks
 * so multi-line text wraps cleanly in Excel cells.
 */
export function cleanHtmlText(html: string): string {
  if (!html) return "";
  let text = html
    .replace(/<br\s*[\/]?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<li>/gi, "• ")
    .replace(/<\/li>/gi, "\n")
    .replace(/<[^>]+>/g, "");

  text = text
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");

  return text.trim();
}

/**
 * Safely parses a numeric input string or number
 */
function parseNum(val: string | number | undefined): number {
  if (val === undefined || val === null) return 0;
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  const cleaned = val.toString().replace(/,/g, "").trim();
  const parsed = parseFloat(cleaned);
  return isNaN(parsed) ? 0 : parsed;
}

/**
 * Fetches an image from URL and converts to ArrayBuffer for embedding in ExcelJS
 */
async function fetchImageBuffer(url: string): Promise<ArrayBuffer | null> {
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    return await res.arrayBuffer();
  } catch (err) {
    console.warn("Could not fetch image for Excel embedding:", err);
    return null;
  }
}

const THIN_BORDER: Partial<ExcelJS.Borders> = {
  top: { style: "thin", color: { argb: "FF000000" } },
  bottom: { style: "thin", color: { argb: "FF000000" } },
  left: { style: "thin", color: { argb: "FF000000" } },
  right: { style: "thin", color: { argb: "FF000000" } },
};

interface PreparedItem {
  row: QuotationRow;
  originalIndex: number;
  cleanDesc: string;
  height: number;
}

interface PageData {
  pageNumber: number;
  items: PreparedItem[];
  isFirstPage: boolean;
  isLastPage: boolean;
}

/**
 * Generates an Excel document (.xlsx) with:
 * - Row 1: Document format title (e.g. "QUOTATION", "DELIVERY CHALLAN", "INVOICE") with NO BORDERS and clean transparent background
 * - Rows 2 to 12: 11 blank rows for pre-printed company letterhead stationery
 * - Utilizes the FULL printable height of A4 paper (~800pt) exclusively to fit maximum items per page
 * - Stamp image placed ABOVE the Authorized Signature text line for Comilla Traders (matching PDF)
 * - Signature section present on EVERY single page according to the format
 * - Dynamic pagination dividing items when the A4 printable area is full across worksheets
 */
export async function generateExcelDocument(options: ExcelGeneratorOptions): Promise<void> {
  const {
    currentCompany,
    docType,
    messers,
    address,
    vesselName = "",
    portBerth = "",
    includeVesselName = false,
    includePortBerth = false,
    invoiceNo = "",
    challanNo = "",
    quotationNo = "",
    requisitionNo = "",
    poNumber = "",
    includeInvoiceNo = true,
    includeChallanNo = true,
    includeQuotationNo = true,
    includeRequisitionNo = true,
    includePoNumber = true,
    dateVal,
    rows,
    includeDiscount = false,
    discountType = "percentage",
    discountValue = "0",
    vatPercent = "0",
    transportationFee = "0",
    currency = "BDT",
    currencySymbol = "Tk",
  } = options;

  const isChallan = docType === "challan";
  const isInvoice = docType === "invoice";
  const colCount = isChallan ? 4 : 6;
  const lastColLetter = isChallan ? "D" : "F";
  const docTitleText = isChallan ? "DELIVERY CHALLAN" : isInvoice ? "INVOICE" : "QUOTATION";
  const baseSheetName = isChallan ? "Challan" : isInvoice ? "Invoice" : "Quotation";

  // Filter active rows or fallback
  const activeRows = rows.filter((r) => {
    const hasDesc = (r.desc || "").trim() !== "";
    const hasQty = (r.qty || "").trim() !== "";
    const hasPrice = (r.price || "").trim() !== "";
    const hasAmount = (r.amount || 0) > 0;
    return hasDesc || hasQty || hasPrice || hasAmount;
  });

  const rowsToExport = activeRows.length > 0 ? activeRows : rows.slice(0, 10);

  // Compact height per item to utilize the whole page exclusively and maximize item count
  // Standard single line: 13.5pt. Each wrapped line adds ~9.5pt.
  const preparedItems: PreparedItem[] = rowsToExport.map((r, idx) => {
    const cleanDesc = cleanHtmlText(r.desc);
    const lineBreaks = (cleanDesc.match(/\n/g) || []).length + 1;
    const charWrapLines = Math.ceil(cleanDesc.length / (isChallan ? 58 : 46));
    const lines = Math.max(lineBreaks, charWrapLines, 1);
    const height = lines === 1 ? 13.5 : Math.max(13.5, lines * 9.5 + 2);
    return {
      row: r,
      originalIndex: idx + 1,
      cleanDesc,
      height,
    };
  });

  // Calculate summary counts for Quotation & Invoice
  const rowsTotal = rowsToExport.reduce((sum, r) => sum + (r.amount || 0), 0);
  const parsedVat = parseNum(vatPercent);
  const parsedTransport = parseNum(transportationFee);
  const parsedDiscVal = parseNum(discountValue);

  let discountAmount = 0;
  if (isInvoice && includeDiscount && parsedDiscVal > 0) {
    if (discountType === "percentage") {
      discountAmount = (rowsTotal * parsedDiscVal) / 100;
    } else {
      discountAmount = parsedDiscVal;
    }
  }
  discountAmount = Math.min(rowsTotal, Math.max(0, discountAmount));
  const netAfterDiscount = Math.max(0, rowsTotal - discountAmount);
  const vatAmount = isInvoice ? (netAfterDiscount * parsedVat) / 100 : 0;
  const grandTotal = isInvoice ? netAfterDiscount + vatAmount + parsedTransport : rowsTotal;

  // Build Metadata Left and Right lines for Page 1
  const cleanMessers = cleanHtmlText(messers) || " ";
  const cleanAddress = cleanHtmlText(address) || " ";

  const leftLines: { label: string; value: string; bold?: boolean }[] = [
    { label: "Messers:", value: cleanMessers, bold: true },
  ];
  if (includeVesselName && vesselName) {
    leftLines.push({ label: "Vessel Name:", value: vesselName, bold: true });
  }
  if (includePortBerth && portBerth) {
    leftLines.push({ label: "Port / Berth:", value: portBerth });
  }
  leftLines.push({ label: "Address:", value: cleanAddress });

  const rightLines: { label: string; value: string; bold?: boolean }[] = [];
  if (isInvoice) {
    if (includeInvoiceNo && invoiceNo) rightLines.push({ label: "Invoice No.:", value: invoiceNo, bold: true });
    if (includeChallanNo && challanNo) rightLines.push({ label: "Challan No.:", value: challanNo, bold: true });
    rightLines.push({ label: "Date:", value: dateVal || new Date().toLocaleDateString("en-GB"), bold: true });
    if (includeRequisitionNo && requisitionNo) rightLines.push({ label: "Requisition No.:", value: requisitionNo, bold: true });
    if (includePoNumber && poNumber) rightLines.push({ label: "PO Number:", value: poNumber, bold: true });
  } else if (isChallan) {
    if (includeChallanNo && challanNo) rightLines.push({ label: "Challan No.:", value: challanNo, bold: true });
    rightLines.push({ label: "Date:", value: dateVal || new Date().toLocaleDateString("en-GB"), bold: true });
    if (includeRequisitionNo && requisitionNo) rightLines.push({ label: "Requisition No.:", value: requisitionNo, bold: true });
  } else {
    if (includeQuotationNo && quotationNo) rightLines.push({ label: "Quotation No.:", value: quotationNo, bold: true });
    rightLines.push({ label: "Date:", value: dateVal || new Date().toLocaleDateString("en-GB"), bold: true });
    if (includeRequisitionNo && requisitionNo) rightLines.push({ label: "Requisition No.:", value: requisitionNo, bold: true });
    if (includePoNumber && poNumber) rightLines.push({ label: "PO Number:", value: poNumber, bold: true });
  }

  const maxMetaRows = Math.max(leftLines.length, rightLines.length, 3);
  const leftColEnd = isChallan ? 2 : 3;
  const rightColStart = leftColEnd + 1;

  // Number of rows in summary block
  let summaryRowsCount = 2; // Subtotal + Grand Total
  if (isInvoice && includeDiscount && discountAmount > 0) summaryRowsCount++;
  if (isInvoice && vatAmount > 0) summaryRowsCount++;
  if (isInvoice && parsedTransport > 0) summaryRowsCount++;

  // FULL PAGE USAGE IN EXCEL (A4 height with 0.2in margins = ~800pt printable)
  const PAGE_LIMIT = 800;
  const ROW1_TITLE_HEIGHT = 16;
  const BLANK_11_ROWS_HEIGHT = 11 * 12; // 132pt (11 blank rows at 12pt each)
  const META_BOX_HEIGHT = maxMetaRows * 12.5 + 3; // ~40-52pt
  const TABLE_HEADER_HEIGHT = 14;

  const P1_PRE_HEIGHT = ROW1_TITLE_HEIGHT + BLANK_11_ROWS_HEIGHT + META_BOX_HEIGHT + TABLE_HEADER_HEIGHT; // ~202pt
  const CONT_PRE_HEIGHT = ROW1_TITLE_HEIGHT + BLANK_11_ROWS_HEIGHT + TABLE_HEADER_HEIGHT + 3; // ~165pt

  // Signature block is now present on EVERY page!
  // Top gap (4) + For Company (12.5) + Stamp clearance (3 * 13 = 39) + Sig lines (13) + Notice gap (3) + Notice (11) = 82.5pt
  const SIGNATURE_BLOCK_HEIGHT = 82.5;

  // Summary block (Subtotal, VAT, Grand Total) is placed on the final page
  const SUMMARY_BLOCK_HEIGHT = isChallan ? 0 : (summaryRowsCount * 13 + 5);

  // Every page reserves room for the signature block
  const INTERMEDIATE_FOOTER_HEIGHT = SIGNATURE_BLOCK_HEIGHT;
  const FINAL_FOOTER_HEIGHT = SUMMARY_BLOCK_HEIGHT + SIGNATURE_BLOCK_HEIGHT;

  // =========================================================================
  // PAGINATION / FULL PAGE USAGE ALGORITHM
  // =========================================================================
  const pages: PageData[] = [];
  const totalItemsHeight = preparedItems.reduce((s, it) => s + it.height, 0);

  if (P1_PRE_HEIGHT + totalItemsHeight + FINAL_FOOTER_HEIGHT <= PAGE_LIMIT) {
    pages.push({
      pageNumber: 1,
      items: preparedItems,
      isFirstPage: true,
      isLastPage: true,
    });
  } else {
    let remaining = [...preparedItems];
    let pageNum = 1;

    while (remaining.length > 0) {
      const isFirst = pageNum === 1;
      const preHeight = isFirst ? P1_PRE_HEIGHT : CONT_PRE_HEIGHT;
      const maxAvailableFinal = PAGE_LIMIT - preHeight - FINAL_FOOTER_HEIGHT;
      const maxAvailableIntermediate = PAGE_LIMIT - preHeight - INTERMEDIATE_FOOTER_HEIGHT;

      // Check if ALL remaining items plus the final footer fit on this page
      const remainingHeight = remaining.reduce((s, it) => s + it.height, 0);
      if (remainingHeight <= maxAvailableFinal) {
        pages.push({
          pageNumber: pageNum,
          items: remaining,
          isFirstPage: isFirst,
          isLastPage: true,
        });
        remaining = [];
        break;
      }

      // Maximize items placed onto this page (using intermediate footer budget)
      const pageItems: PreparedItem[] = [];
      let currentHeight = 0;

      while (remaining.length > 0) {
        const nextItem = remaining[0];
        if (pageItems.length > 0 && currentHeight + nextItem.height > maxAvailableIntermediate) {
          break;
        }
        currentHeight += nextItem.height;
        pageItems.push(remaining.shift()!);
      }

      // If all items fit on this page but the final footer cannot, shift minimal items to next page
      if (remaining.length === 0 && pageItems.length > 2) {
        const moved = pageItems.splice(pageItems.length - 2, 2);
        remaining = moved;
      }

      pages.push({
        pageNumber: pageNum,
        items: pageItems,
        isFirstPage: isFirst,
        isLastPage: false,
      });

      pageNum++;
    }

    if (remaining.length > 0) {
      pages.push({
        pageNumber: pageNum,
        items: remaining,
        isFirstPage: false,
        isLastPage: true,
      });
    }
  }

  // =========================================================================
  // PRE-FETCH COMILLA TRADERS STAMP IMAGE
  // =========================================================================
  let stampImageId: number | null = null;
  const isComilla = currentCompany.id === "comilla" || (currentCompany.hasStamp && !!currentCompany.stampUrl);
  let stampBuffer: ArrayBuffer | null = null;

  if (isComilla && currentCompany.stampUrl) {
    stampBuffer = await fetchImageBuffer(currentCompany.stampUrl);
  }

  // =========================================================================
  // WORKBOOK CREATION & PER-PAGE SHEET GENERATION
  // =========================================================================
  const workbook = new ExcelJS.Workbook();
  workbook.creator = currentCompany.name;
  workbook.lastModifiedBy = currentCompany.name;
  workbook.created = new Date();
  workbook.modified = new Date();

  // Register stamp image with workbook if available
  if (stampBuffer) {
    try {
      stampImageId = workbook.addImage({
        buffer: stampBuffer,
        extension: "png",
      });
    } catch (e) {
      console.warn("Could not register stamp in Excel workbook:", e);
    }
  }

  pages.forEach((page) => {
    const sheetName = pages.length === 1 
      ? baseSheetName 
      : `${baseSheetName} - Page ${page.pageNumber}`.substring(0, 31);

    const ws = workbook.addWorksheet(sheetName, {
      pageSetup: {
        paperSize: 9, // A4
        orientation: "portrait",
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 1, // Fit exactly to 1 page tall
        margins: {
          left: 0.2,
          right: 0.2,
          top: 0.2,
          bottom: 0.2,
          header: 0.05,
          footer: 0.05,
        },
      },
      views: [{ showGridLines: true }],
    });

    // Optimized column widths for maximum content visibility & print fitting
    if (isChallan) {
      ws.columns = [
        { key: "sl", width: 5.5 },
        { key: "desc", width: 63 },
        { key: "qty", width: 12 },
        { key: "remarks", width: 18 },
      ];
    } else {
      ws.columns = [
        { key: "sl", width: 5 },
        { key: "desc", width: 53 },
        { key: "qty", width: 8 },
        { key: "unit", width: 8 },
        { key: "rate", width: 14 },
        { key: "total", width: 16.5 },
      ];
    }

    // =========================================================================
    // ROW 1: FORMAT NAME ON THE FIRST ROW (NO BORDERS, CLEAN DISPLAY)
    // =========================================================================
    const pageFormatTitle = pages.length > 1 && !page.isFirstPage
      ? `${docTitleText}  —  PAGE ${page.pageNumber} OF ${pages.length}`
      : docTitleText;

    const row1 = ws.addRow([pageFormatTitle]);
    row1.height = ROW1_TITLE_HEIGHT;
    ws.mergeCells(`A1:${lastColLetter}1`);
    const row1Cell = ws.getCell("A1");
    row1Cell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF0F172A" } };
    row1Cell.alignment = { horizontal: "center", vertical: "middle" };
    // Explicitly remove all borders around row 1 format title as requested
    for (let c = 1; c <= colCount; c++) {
      ws.getCell(1, c).border = {};
    }

    // =========================================================================
    // ROWS 2 TO 12: 11 BLANK ROWS (For pre-printed letterhead stationery)
    // =========================================================================
    for (let r = 2; r <= 12; r++) {
      const blankRow = ws.addRow([]);
      blankRow.height = 12;
    }

    let currentRow = 13;

    // =========================================================================
    // METADATA BOXES (Page 1 Only)
    // =========================================================================
    if (page.isFirstPage) {
      for (let i = 0; i < maxMetaRows; i++) {
        const lItem = leftLines[i];
        const rItem = rightLines[i];

        const row = ws.addRow([]);
        row.height = 12.5;

        // Left Box
        if (lItem) {
          const leftCell = ws.getCell(currentRow, 1);
          leftCell.value = `${lItem.label} ${lItem.value}`;
          leftCell.font = { name: "Arial", size: 7.5, bold: !!lItem.bold, color: { argb: "FF000000" } };
          leftCell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
        }
        ws.mergeCells(currentRow, 1, currentRow, leftColEnd);

        // Right Box
        if (rItem) {
          const rightCell = ws.getCell(currentRow, rightColStart);
          rightCell.value = `${rItem.label} ${rItem.value}`;
          rightCell.font = { name: "Arial", size: 7.5, bold: !!rItem.bold, color: { argb: "FF000000" } };
          rightCell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
        }
        ws.mergeCells(currentRow, rightColStart, currentRow, colCount);

        // Thin borders around metadata boxes
        for (let c = 1; c <= leftColEnd; c++) {
          const cCell = ws.getCell(currentRow, c);
          cCell.border = {
            top: i === 0 ? { style: "thin" } : undefined,
            bottom: i === maxMetaRows - 1 ? { style: "thin" } : undefined,
            left: c === 1 ? { style: "thin" } : undefined,
            right: c === leftColEnd ? { style: "thin" } : undefined,
          };
        }
        for (let c = rightColStart; c <= colCount; c++) {
          const cCell = ws.getCell(currentRow, c);
          cCell.border = {
            top: i === 0 ? { style: "thin" } : undefined,
            bottom: i === maxMetaRows - 1 ? { style: "thin" } : undefined,
            left: c === rightColStart ? { style: "thin" } : undefined,
            right: c === colCount ? { style: "thin" } : undefined,
          };
        }

        currentRow++;
      }

      // Compact gap before table
      const gapRow = ws.addRow([]);
      gapRow.height = 3;
      currentRow++;
    }

    // =========================================================================
    // TABLE HEADERS (Compact height)
    // =========================================================================
    const headerValues = isChallan
      ? ["SL", "Description of Marine Items / Spare Parts", "Qty", "Remarks / Unit"]
      : [
          "SL",
          "Description of Marine Items / Spare Parts",
          "Qty",
          "Unit",
          `Unit Price (${currencySymbol || "Tk"})`,
          `Total Price (${currencySymbol || "Tk"})`,
        ];

    const headerRow = ws.addRow(headerValues);
    headerRow.height = TABLE_HEADER_HEIGHT;
    headerRow.eachCell((cell, colNumber) => {
      cell.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "FF000000" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
      cell.border = THIN_BORDER;

      if (colNumber === 1 || colNumber === 3 || colNumber === 4) {
        cell.alignment = { horizontal: "center", vertical: "middle" };
      } else if (colNumber === 2) {
        cell.alignment = { horizontal: "left", vertical: "middle" };
      } else {
        cell.alignment = { horizontal: "right", vertical: "middle" };
      }
    });
    currentRow++;

    // =========================================================================
    // TABLE ITEMS (Compact vertical cell height)
    // =========================================================================
    page.items.forEach((item) => {
      const r = item.row;
      const qtyVal = parseNum(r.qty);
      const priceVal = parseNum(r.price);
      const amountVal = r.amount || (qtyVal * priceVal);

      let rowData: any[];
      if (isChallan) {
        rowData = [
          item.originalIndex,
          item.cleanDesc,
          qtyVal > 0 ? qtyVal : (r.qty || ""),
          r.unit || "",
        ];
      } else {
        rowData = [
          item.originalIndex,
          item.cleanDesc,
          qtyVal > 0 ? qtyVal : (r.qty || ""),
          r.unit || "",
          priceVal > 0 ? priceVal : "",
          amountVal > 0 ? amountVal : "",
        ];
      }

      const row = ws.addRow(rowData);
      row.height = item.height;

      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        cell.border = THIN_BORDER;
        cell.font = { name: "Arial", size: 7.5, color: { argb: "FF000000" } };

        if (colNumber === 1) {
          cell.alignment = { horizontal: "center", vertical: "middle" };
        } else if (colNumber === 2) {
          cell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
        } else if (colNumber === 3) {
          cell.alignment = { horizontal: "center", vertical: "middle" };
          if (typeof cell.value === "number") {
            cell.numFmt = "#,##0.##";
          }
        } else if (colNumber === 4) {
          cell.alignment = { horizontal: "center", vertical: "middle" };
        } else if (colNumber === 5) {
          cell.alignment = { horizontal: "right", vertical: "middle" };
          if (typeof cell.value === "number") {
            cell.numFmt = "#,##0.00";
          }
        } else if (colNumber === 6) {
          cell.alignment = { horizontal: "right", vertical: "middle" };
          if (typeof cell.value === "number") {
            cell.numFmt = "#,##0.00";
          }
        }
      });

      currentRow++;
    });

    // If intermediate page, add continuation notice row above the signature section
    if (!page.isLastPage) {
      const contRow = ws.addRow([`Continued on Page ${page.pageNumber + 1}...`]);
      contRow.height = 12;
      ws.mergeCells(`A${currentRow}:${lastColLetter}${currentRow}`);
      const cCell = ws.getCell(`A${currentRow}`);
      cCell.font = { name: "Arial", size: 7.5, italic: true, bold: true, color: { argb: "FF64748B" } };
      cCell.alignment = { horizontal: "right", vertical: "middle" };
      currentRow++;
    }

    // =========================================================================
    // TOTALS & SUMMARY SECTION (Only on Last Page for Quotation & Invoice)
    // =========================================================================
    if (page.isLastPage && !isChallan) {
      const wordsText = numberToWords(grandTotal, currency);
      const summaryStartRow = currentRow;

      // Sub Total
      const subTotalRow = ws.addRow(["", "", "", "", "Sub Total:", rowsTotal]);
      subTotalRow.height = 13;
      currentRow++;

      // Discount (if any)
      if (isInvoice && includeDiscount && discountAmount > 0) {
        const discLabel = discountType === "percentage" ? `Less Discount (${discountValue}%):` : "Less Discount:";
        const discRow = ws.addRow(["", "", "", "", discLabel, -discountAmount]);
        discRow.height = 13;
        currentRow++;
      }

      // VAT (if any)
      if (isInvoice && vatAmount > 0) {
        const vatRow = ws.addRow(["", "", "", "", `Add VAT (${vatPercent}%):`, vatAmount]);
        vatRow.height = 13;
        currentRow++;
      }

      // Transportation (if any)
      if (isInvoice && parsedTransport > 0) {
        const transRow = ws.addRow(["", "", "", "", "Transportation:", parsedTransport]);
        transRow.height = 13;
        currentRow++;
      }

      // Grand Total / Net Payable
      const grandTotalLabel = isInvoice ? "Net Payable:" : "Grand Total:";
      const grandTotalRow = ws.addRow(["", "", "", "", grandTotalLabel, grandTotal]);
      grandTotalRow.height = 15;
      const grandTotalRowIndex = currentRow;
      currentRow++;

      const summaryEndRow = currentRow - 1;

      // Merge columns A-D across summary rows for "In Words" section
      ws.mergeCells(`A${summaryStartRow}:D${summaryEndRow}`);
      const wordsCell = ws.getCell(`A${summaryStartRow}`);
      wordsCell.value = `In Words:\n${wordsText}`;
      wordsCell.font = { name: "Arial", size: 7.5, bold: true, italic: true, color: { argb: "FF0F172A" } };
      wordsCell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
      wordsCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };

      // Apply borders to the In Words merged block
      for (let r = summaryStartRow; r <= summaryEndRow; r++) {
        for (let c = 1; c <= 4; c++) {
          const cell = ws.getCell(r, c);
          cell.border = {
            top: r === summaryStartRow ? { style: "thin" } : undefined,
            bottom: r === summaryEndRow ? { style: "thin" } : undefined,
            left: c === 1 ? { style: "thin" } : undefined,
            right: c === 4 ? { style: "thin" } : undefined,
          };
        }
      }

      // Style totals cells (Columns E & F)
      for (let r = summaryStartRow; r <= summaryEndRow; r++) {
        const labelCell = ws.getCell(r, 5);
        const valCell = ws.getCell(r, 6);
        const isGrandTotal = r === grandTotalRowIndex;

        labelCell.font = {
          name: "Arial",
          size: isGrandTotal ? 8 : 7.5,
          bold: true,
          color: { argb: "FF000000" },
        };
        labelCell.alignment = { horizontal: "right", vertical: "middle" };

        valCell.font = {
          name: "Arial",
          size: isGrandTotal ? 8.5 : 7.5,
          bold: true,
          color: { argb: "FF000000" },
        };
        valCell.alignment = { horizontal: "right", vertical: "middle" };
        valCell.numFmt = "#,##0.00";

        if (isGrandTotal) {
          labelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
          valCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
          labelCell.border = {
            top: { style: "thin", color: { argb: "FF000000" } },
            bottom: { style: "double", color: { argb: "FF000000" } },
            left: { style: "thin", color: { argb: "FF000000" } },
            right: { style: "thin", color: { argb: "FF000000" } },
          };
          valCell.border = {
            top: { style: "thin", color: { argb: "FF000000" } },
            bottom: { style: "double", color: { argb: "FF000000" } },
            left: { style: "thin", color: { argb: "FF000000" } },
            right: { style: "thin", color: { argb: "FF000000" } },
          };
        } else {
          labelCell.border = THIN_BORDER;
          valCell.border = THIN_BORDER;
        }
      }
    }

    // =========================================================================
    // SIGNATURES & STAMP SECTION (Now present on EVERY PAGE according to format)
    // =========================================================================
    // Top spacing above signature block
    const sigGap = ws.addRow([]);
    sigGap.height = 4;
    currentRow++;

    // Row 1: "For [Company Name]" above Authorized Signature on the right
    const forCompRow = ws.addRow([]);
    forCompRow.height = 12.5;
    if (!isChallan) {
      const rightSigColStart = isChallan ? "C" : "E";
      ws.mergeCells(`${rightSigColStart}${currentRow}:${lastColLetter}${currentRow}`);
      const forCompCell = ws.getCell(`${rightSigColStart}${currentRow}`);
      forCompCell.value = `For ${currentCompany.name}`;
      forCompCell.font = { name: "Arial", size: 8, bold: true, color: { argb: "FF000000" } };
      forCompCell.alignment = { horizontal: "center", vertical: "middle" };
    }
    currentRow++;

    // Rows 2, 3, 4: Clear Stamp and Pen Signature Space (3 rows of 13pt = 39pt)
    const stampStartRowIndex = currentRow;
    for (let s = 0; s < 3; s++) {
      const spacer = ws.addRow([]);
      spacer.height = 13;
      currentRow++;
    }

    // Embed Stamp Image for Comilla Traders ABOVE the "Authorized Signature" text line
    if (isComilla && stampImageId !== null && !isChallan) {
      // Anchored inside the 3 spacer rows directly above the signature line (rows stampStartRowIndex to currentRow-1)
      const startCol = isChallan ? 2.3 : 4.3;
      ws.addImage(stampImageId, {
        tl: { col: startCol, row: stampStartRowIndex - 1.2 },
        ext: { width: 85, height: 85 },
      });
    }

    // Row 5: Signature Lines & Headings
    const sigLineRow = ws.addRow([]);
    sigLineRow.height = 13;

    // Left Box: Receiver's Signature (Cols A-B)
    ws.mergeCells(`A${currentRow}:B${currentRow}`);
    const recCell = ws.getCell(`A${currentRow}`);
    recCell.value = "Receiver's Signature";
    recCell.font = { name: "Arial", size: 8, bold: true, color: { argb: "FF000000" } };
    recCell.alignment = { horizontal: "center", vertical: "top" };
    ws.getCell(`A${currentRow}`).border = { top: { style: "medium", color: { argb: "FF000000" } } };
    ws.getCell(`B${currentRow}`).border = { top: { style: "medium", color: { argb: "FF000000" } } };

    // Right Box: Authorized Signature (Cols E-F, or C-D for challan)
    if (!isChallan) {
      const rightSigColStart = isChallan ? "C" : "E";
      ws.mergeCells(`${rightSigColStart}${currentRow}:${lastColLetter}${currentRow}`);
      const authCell = ws.getCell(`${rightSigColStart}${currentRow}`);
      authCell.value = "Authorized Signature";
      authCell.font = { name: "Arial", size: 8, bold: true, color: { argb: "FF000000" } };
      authCell.alignment = { horizontal: "center", vertical: "top" };

      const startColNum = isChallan ? 3 : 5;
      for (let c = startColNum; c <= colCount; c++) {
        ws.getCell(currentRow, c).border = { top: { style: "medium", color: { argb: "FF000000" } } };
      }
    }
    currentRow++;

    // Small gap before legal notice
    const preNoticeGap = ws.addRow([]);
    preNoticeGap.height = 3;
    currentRow++;

    // Legal notice footer
    const noticeRow = ws.addRow(["ITEMS ONCE SOLD ARE NON-RETURNABLE AND NON-EXCHANGEABLE."]);
    noticeRow.height = 11;
    ws.mergeCells(`A${currentRow}:${lastColLetter}${currentRow}`);
    const noticeCell = ws.getCell(`A${currentRow}`);
    noticeCell.font = { name: "Arial", size: 6.5, bold: true, color: { argb: "FF000000" } };
    noticeCell.alignment = { horizontal: "center", vertical: "middle" };
    currentRow++;
  });

  // =========================================================================
  // FILE DOWNLOAD TRIGGER
  // =========================================================================
  const filePrefix = isChallan ? "Challan" : isInvoice ? "Invoice" : "Quotation";
  const identifier = isChallan
    ? (challanNo || "NEW")
    : isInvoice
    ? (invoiceNo || "NEW")
    : (quotationNo || requisitionNo || "NEW");

  const filename = `${filePrefix}_${identifier.replace(/[\/\\?%*:|"<>\s]/g, "_")}.xlsx`;

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  URL.revokeObjectURL(url);
}
