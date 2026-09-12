import ExcelJS from "exceljs";
import { QuotationRow, MergedRegion, CellFormatMap } from "../types";

export interface ExcelMetadataOptions {
  vesselName?: string;
  portBerth?: string;
  includeVesselName?: boolean;
  includePortBerth?: boolean;
  includeInvoiceNo?: boolean;
  includeChallanNo?: boolean;
  includeRequisitionNo?: boolean;
  includePoNumber?: boolean;
}

export interface ExcelPageChunk {
  rows: QuotationRow[];
  startSlIndex: number;
  isLastPage: boolean;
}

/**
 * Strips HTML tags and entities to produce clean plain text.
 */
export const htmlToPlainText = (html: string): string => {
  if (!html) return "";
  return html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<p[^>]*>/gi, "")
    .replace(/<[^>]+>/gi, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
};

/**
 * Converts various color representations into Excel ARGB hex format (e.g. FFRRGGBB).
 */
export function colorToArgb(color: string | null | undefined): string | null {
  if (!color || color === "transparent" || color === "inherit") return null;
  const c = color.trim().toLowerCase();
  if (c.startsWith("#")) {
    const hex = c.substring(1);
    if (hex.length === 3) {
      const r = hex[0] + hex[0];
      const g = hex[1] + hex[1];
      const b = hex[2] + hex[2];
      return `FF${r}${g}${b}`.toUpperCase();
    }
    if (hex.length === 6) {
      return `FF${hex}`.toUpperCase();
    }
  }
  const rgbMatch = c.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  if (rgbMatch) {
    const r = parseInt(rgbMatch[1], 10).toString(16).padStart(2, "0");
    const g = parseInt(rgbMatch[2], 10).toString(16).padStart(2, "0");
    const b = parseInt(rgbMatch[3], 10).toString(16).padStart(2, "0");
    return `FF${r}${g}${b}`.toUpperCase();
  }
  const named: Record<string, string> = {
    yellow: "FFFFFF00",
    red: "FFFF0000",
    blue: "FF0000FF",
    green: "FF008000",
    black: "FF000000",
    white: "FFFFFFFF",
    orange: "FFFFA500",
  };
  return named[c] || null;
}

/**
 * Parses simple inline HTML (b, strong, i, em, u, font color, span background) into ExcelJS rich text.
 */
export function parseHtmlToExcelRuns(
  html: string,
  baseFont: Partial<ExcelJS.Font>
): {
  richText: ExcelJS.RichText[];
  plainText: string;
  hasFormatting: boolean;
  highlightColor: string | null;
} {
  const plain = htmlToPlainText(html);
  if (!html || !html.includes("<")) {
    return {
      richText: [{ text: plain, font: { ...baseFont } }],
      plainText: plain,
      hasFormatting: false,
      highlightColor: null,
    };
  }

  // Detect overall background highlight color if entire cell has background
  let highlightColor: string | null = null;
  const bgMatch = html.match(/background(?:-color)?:\s*([^;"'>]+)/i);
  if (bgMatch) {
    highlightColor = colorToArgb(bgMatch[1]);
  }

  // Simplified robust run extractor
  const runs: ExcelJS.RichText[] = [];
  const tagRegex = /(<[^>]+>|[^<]+)/g;
  let match: RegExpExecArray | null;
  let isBold = Boolean(baseFont.bold);
  let isItalic = Boolean(baseFont.italic);
  let isUnderline = Boolean(baseFont.underline);
  let currentColor: string | null = null;

  while ((match = tagRegex.exec(html)) !== null) {
    const token = match[1];
    if (token.startsWith("<")) {
      const lower = token.toLowerCase();
      if (lower.startsWith("<b") || lower.startsWith("<strong")) isBold = true;
      else if (lower.startsWith("</b") || lower.startsWith("</strong")) isBold = Boolean(baseFont.bold);
      else if (lower.startsWith("<i") || lower.startsWith("<em")) isItalic = true;
      else if (lower.startsWith("</i") || lower.startsWith("</em")) isItalic = Boolean(baseFont.italic);
      else if (lower.startsWith("<u")) isUnderline = true;
      else if (lower.startsWith("</u")) isUnderline = Boolean(baseFont.underline);
      else if (lower.includes("color:")) {
        const cMatch = token.match(/color:\s*([^;"'>]+)/i);
        if (cMatch) currentColor = colorToArgb(cMatch[1]);
      } else if (lower.startsWith("</font") || lower.startsWith("</span")) {
        currentColor = null;
      }
    } else {
      const decoded = token
        .replace(/&nbsp;/g, " ")
        .replace(/&amp;/g, "&")
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'");

      const runFont: Partial<ExcelJS.Font> = {
        name: baseFont.name || "Arial",
        size: baseFont.size || 8.5,
        bold: isBold,
        italic: isItalic,
        underline: isUnderline,
      };
      if (currentColor) {
        runFont.color = { argb: currentColor };
      } else if (baseFont.color) {
        runFont.color = baseFont.color;
      }
      runs.push({ text: decoded, font: runFont });
    }
  }

  if (runs.length === 0) {
    runs.push({ text: plain, font: { ...baseFont } });
  }

  return {
    richText: runs,
    plainText: plain,
    hasFormatting: runs.length > 1 || isBold || isItalic || isUnderline || Boolean(currentColor),
    highlightColor,
  };
}

/**
 * Calculates visual text lines in a cell given text, column width, and font size.
 */
export const calculateItemVisualLines = (
  text: string,
  colWidth: number,
  fontSize: number = 8.5
): number => {
  if (!text) return 1;
  const plain = htmlToPlainText(text).trim();
  if (!plain) return 1;

  const rawParagraphs = plain.split(/\r\n|\r|\n/);
  const charsPerLine = Math.max(12, Math.floor(colWidth * (8.5 / fontSize) * 1.05));

  let totalLines = 0;
  for (const para of rawParagraphs) {
    if (!para.trim()) {
      totalLines += 1;
      continue;
    }
    const words = para.split(/\s+/);
    let currentLineLength = 0;
    let linesInPara = 1;

    for (const word of words) {
      const wordLen = word.length;
      if (currentLineLength === 0) {
        currentLineLength = wordLen;
      } else if (currentLineLength + 1 + wordLen <= charsPerLine) {
        currentLineLength += 1 + wordLen;
      } else {
        linesInPara++;
        currentLineLength = wordLen;
      }
    }
    totalLines += linesInPara;
  }
  return Math.max(1, totalLines);
};

/**
 * Calculates compact row height in points.
 * Sized at 14.5pt for 1 line, allowing 35+ items per A4 page.
 */
export const getItemRowHeight = (visualLines: number, fontSize: number = 8.5): number => {
  if (visualLines <= 1) {
    return Math.max(14.5, fontSize * 1.45);
  }
  return Math.max(14.5, Math.ceil(visualLines * (fontSize * 1.35) + 3));
};

/**
 * Paginates rows dynamically to maximize items per page (32-35+ on page 1).
 */
export const paginateRowsForExcel = (
  allRows: QuotationRow[],
  docType: "quotation" | "challan" | "invoice",
  cellFormats?: CellFormatMap
): ExcelPageChunk[] => {
  if (allRows.length === 0) {
    return [{ rows: [], startSlIndex: 1, isLastPage: true }];
  }

  const isChallan = docType === "challan";
  const descColWidth = isChallan ? 52 : 42;

  // Maximum vertical height budget for table rows in points:
  // Challan: ~550pt (no totals/vat/in-words block) -> fits up to 38 single-line items
  // Quotation / Invoice: ~480pt on Page 1 -> fits 32-35 single-line items
  const REGULAR_PAGE_BUDGET = 550;
  const LAST_PAGE_BUDGET = isChallan ? 550 : 480;

  // Check if ALL rows fit comfortably on a single page
  let totalHeight = 0;
  for (let i = 0; i < allRows.length; i++) {
    const fs = cellFormats?.[`${i}_0`]?.fontSize || 8.5;
    const lines = calculateItemVisualLines(allRows[i].desc, descColWidth, fs);
    totalHeight += getItemRowHeight(lines, fs);
  }

  if (totalHeight <= LAST_PAGE_BUDGET) {
    return [{ rows: allRows, startSlIndex: 1, isLastPage: true }];
  }

  // Multi-page chunking
  const chunks: ExcelPageChunk[] = [];
  let currentChunk: QuotationRow[] = [];
  let currentHeight = 0;
  let currentSl = 1;

  for (let i = 0; i < allRows.length; i++) {
    const row = allRows[i];
    const fs = cellFormats?.[`${i}_0`]?.fontSize || 8.5;
    const lines = calculateItemVisualLines(row.desc, descColWidth, fs);
    const rowHeight = getItemRowHeight(lines, fs);

    if (currentHeight + rowHeight > REGULAR_PAGE_BUDGET && currentChunk.length > 0) {
      chunks.push({
        rows: currentChunk,
        startSlIndex: currentSl,
        isLastPage: false,
      });
      currentSl += currentChunk.length;
      currentChunk = [];
      currentHeight = 0;
    }

    currentChunk.push(row);
    currentHeight += rowHeight;
  }

  if (currentChunk.length > 0) {
    chunks.push({
      rows: currentChunk,
      startSlIndex: currentSl,
      isLastPage: true,
    });
  }

  return chunks;
};

/**
 * Formats a number to words in Bangladeshi Taka style.
 */
const numberToWords = (num: number): string => {
  if (!num || isNaN(num) || num <= 0) return "Zero Taka Only";
  const a = [
    "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
    "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
    "Seventeen", "Eighteen", "Nineteen"
  ];
  const b = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];

  function inWords(n: number): string {
    if (n === 0) return "";
    if (n < 20) return a[n];
    if (n < 100) return b[Math.floor(n / 10)] + (n % 10 !== 0 ? " " + a[n % 10] : "");
    if (n < 1000) return a[Math.floor(n / 100)] + " Hundred" + (n % 100 !== 0 ? " and " + inWords(n % 100) : "");
    if (n < 100000) return inWords(Math.floor(n / 1000)) + " Thousand" + (n % 1000 !== 0 ? " " + inWords(n % 1000) : "");
    if (n < 10000000) return inWords(Math.floor(n / 100000)) + " Lakh" + (n % 100000 !== 0 ? " " + inWords(n % 100000) : "");
    return inWords(Math.floor(n / 10000000)) + " Crore" + (n % 10000000 !== 0 ? " " + inWords(n % 10000000) : "");
  }

  const integerPart = Math.floor(num);
  const decimalPart = Math.round((num - integerPart) * 100);
  let words = inWords(integerPart).trim() + " Taka";
  if (decimalPart > 0) {
    words += " and " + inWords(decimalPart).trim() + " Paisa";
  }
  return words + " Only";
};

/**
 * Builds an individual Excel page matching the exact PDF print template.
 */
const renderPageWorksheet = (
  worksheet: ExcelJS.Worksheet,
  pageRows: QuotationRow[],
  startSlIndex: number,
  isLastPage: boolean,
  pageNumber: number,
  totalPages: number,
  docType: "quotation" | "challan" | "invoice",
  messers: string,
  address: string,
  challanNo: string,
  dateVal: string,
  requisitionNo: string,
  mergedRegions: MergedRegion[],
  invoiceNo: string = "",
  poNumber: string = "",
  vatPercent: number = 0,
  transportationFee: number = 0,
  cellFormats?: CellFormatMap,
  includeDiscount: boolean = false,
  discountType: "percentage" | "fixed" = "percentage",
  discountValue: number = 0,
  discountAmount: number = 0,
  companyName: string = "Comilla Traders",
  metaOptions?: ExcelMetadataOptions
) => {
  const isChallan = docType === "challan";
  const isInvoice = docType === "invoice";
  const totalCols = isChallan ? 4 : 6;
  const lastColLetter = isChallan ? "D" : "F";

  // 1. Page Setup for pristine A4 printing
  worksheet.pageSetup = {
    paperSize: 9, // A4
    orientation: "portrait",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: totalPages === 1 ? 1 : 0,
    margins: {
      left: 0.25,
      right: 0.25,
      top: 0.22,
      bottom: 0.3,
      header: 0.1,
      footer: 0.1,
    },
    horizontalCentered: true,
    showGridLines: false,
  };

  worksheet.views = [{ showGridLines: true }];

  // Column Widths
  if (isChallan) {
    worksheet.columns = [
      { key: "A", width: 6 },  // SL
      { key: "B", width: 52 }, // Description
      { key: "C", width: 10 }, // Qty
      { key: "D", width: 12 }, // Unit
    ];
  } else {
    worksheet.columns = [
      { key: "A", width: 5.5 },  // SL
      { key: "B", width: 42 },   // Description
      { key: "C", width: 7.5 },  // Qty
      { key: "D", width: 8.5 },  // Unit
      { key: "E", width: 11.5 }, // Price
      { key: "F", width: 13.5 }, // Amount
    ];
  }

  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: "thin", color: { argb: "FF000000" } },
    bottom: { style: "thin", color: { argb: "FF000000" } },
    left: { style: "thin", color: { argb: "FF000000" } },
    right: { style: "thin", color: { argb: "FF000000" } },
  };

  // =========================================================================
  // ROW 1: Document Title Header
  // =========================================================================
  worksheet.getRow(1).height = 20;
  worksheet.mergeCells(`A1:${lastColLetter}1`);
  const titleCell = worksheet.getCell("A1");
  titleCell.value = docType.toUpperCase();
  titleCell.font = { name: "Arial", size: 12, bold: true, color: { argb: "FF000000" } };
  titleCell.alignment = { vertical: "middle", horizontal: "center" };

  // =========================================================================
  // ROWS 2 to 11: 10 Letterhead Gap Rows (for pre-printed pad alignment)
  // =========================================================================
  for (let r = 2; r <= 11; r++) {
    worksheet.getRow(r).height = 14.5;
  }

  // =========================================================================
  // ROWS 12 to 16: Metadata Information Boxes
  // =========================================================================
  for (let r = 12; r <= 16; r++) {
    worksheet.getRow(r).height = 14;
  }

  const hasVessel = Boolean(metaOptions?.includeVesselName && metaOptions?.vesselName?.trim());
  const hasPort = Boolean(metaOptions?.includePortBerth && metaOptions?.portBerth?.trim());
  const vesselText = metaOptions?.vesselName?.trim() || "";
  const portText = metaOptions?.portBerth?.trim() || "";

  // LEFT BOX: Cols A & B (Rows 12-16)
  worksheet.mergeCells("A12:B12");
  const messersLbl = worksheet.getCell("A12");
  messersLbl.value = "MESSERS:";
  messersLbl.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
  messersLbl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
  messersLbl.alignment = { vertical: "middle", horizontal: "left" };

  worksheet.mergeCells("A13:B13");
  const messersCell = worksheet.getCell("A13");
  const messersParsed = parseHtmlToExcelRuns(messers || "", {
    name: "Arial",
    size: 8.5,
    bold: true,
    color: { argb: "FF000000" },
  });
  messersCell.value = messersParsed.hasFormatting && messersParsed.richText.length > 0
    ? { richText: messersParsed.richText }
    : messersParsed.plainText;
  messersCell.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "FF000000" } };
  messersCell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
  messersCell.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

  if (hasVessel || hasPort) {
    worksheet.mergeCells("A14:B14");
    const vesselCell = worksheet.getCell("A14");
    const vesselParts = [
      hasVessel ? `VESSEL: ${vesselText}` : "",
      hasPort ? `BERTH: ${portText}` : "",
    ].filter(Boolean);
    vesselCell.value = vesselParts.join(" | ");
    vesselCell.font = { name: "Arial", size: 8.0, bold: true, color: { argb: "FF000000" } };
    vesselCell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    vesselCell.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    worksheet.mergeCells("A15:B15");
    const addrLbl = worksheet.getCell("A15");
    addrLbl.value = "ADDRESS:";
    addrLbl.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    addrLbl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
    addrLbl.alignment = { vertical: "middle", horizontal: "left" };

    worksheet.mergeCells("A16:B16");
    const addrCell = worksheet.getCell("A16");
    const addrParsed = parseHtmlToExcelRuns(address || "", { name: "Arial", size: 8.0, color: { argb: "FF000000" } });
    addrCell.value = addrParsed.hasFormatting && addrParsed.richText.length > 0 ? { richText: addrParsed.richText } : addrParsed.plainText;
    addrCell.font = { name: "Arial", size: 8.0, color: { argb: "FF000000" } };
    addrCell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    addrCell.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };
  } else {
    worksheet.mergeCells("A14:B14");
    const addrLbl = worksheet.getCell("A14");
    addrLbl.value = "ADDRESS:";
    addrLbl.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    addrLbl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
    addrLbl.alignment = { vertical: "middle", horizontal: "left" };

    worksheet.mergeCells("A15:B16");
    const addrCell = worksheet.getCell("A15");
    const addrParsed = parseHtmlToExcelRuns(address || "", { name: "Arial", size: 8.0, color: { argb: "FF000000" } });
    addrCell.value = addrParsed.hasFormatting && addrParsed.richText.length > 0 ? { richText: addrParsed.richText } : addrParsed.plainText;
    addrCell.font = { name: "Arial", size: 8.0, color: { argb: "FF000000" } };
    addrCell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    addrCell.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };
  }

  // Left Box Outer Borders
  for (let r = 12; r <= 16; r++) {
    const cA = worksheet.getCell(`A${r}`);
    const cB = worksheet.getCell(`B${r}`);
    cA.border = { ...cA.border, left: { style: "thin", color: { argb: "FF000000" } } };
    cB.border = { ...cB.border, right: { style: "thin", color: { argb: "FF000000" } } };
    if (r === 12) {
      cA.border = { ...cA.border, top: { style: "thin", color: { argb: "FF000000" } } };
      cB.border = { ...cB.border, top: { style: "thin", color: { argb: "FF000000" } } };
    }
    if (r === 16) {
      cA.border = { ...cA.border, bottom: { style: "thin", color: { argb: "FF000000" } } };
      cB.border = { ...cB.border, bottom: { style: "thin", color: { argb: "FF000000" } } };
    }
  }

  // RIGHT BOX: Rows 12 to 16
  if (isChallan) {
    // Cols C & D for Challan
    const setRightField = (rowNum: number, label: string, val: string) => {
      const lblCell = worksheet.getCell(`C${rowNum}`);
      lblCell.value = label;
      lblCell.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
      lblCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
      lblCell.alignment = { vertical: "middle", horizontal: "left" };

      const valCell = worksheet.getCell(`D${rowNum}`);
      valCell.value = val;
      valCell.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "FF000000" } };
      valCell.alignment = { vertical: "middle", horizontal: "left" };
      valCell.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };
    };

    setRightField(12, "DATE:", dateVal || "");
    setRightField(13, "CHALLAN NO.:", metaOptions?.includeChallanNo !== false ? (challanNo || "") : "");
    setRightField(14, "REQUISITION NO.:", metaOptions?.includeRequisitionNo !== false ? (requisitionNo || "") : "");
    setRightField(15, "P.O. NUMBER:", metaOptions?.includePoNumber !== false ? (poNumber || "") : "");

    // Outer border for right box (C12:D16)
    for (let r = 12; r <= 16; r++) {
      const cC = worksheet.getCell(`C${r}`);
      const cD = worksheet.getCell(`D${r}`);
      cC.border = { ...cC.border, left: { style: "thin", color: { argb: "FF000000" } } };
      cD.border = { ...cD.border, right: { style: "thin", color: { argb: "FF000000" } } };
      if (r === 12) {
        cC.border = { ...cC.border, top: { style: "thin", color: { argb: "FF000000" } } };
        cD.border = { ...cD.border, top: { style: "thin", color: { argb: "FF000000" } } };
      }
      if (r === 16) {
        cC.border = { ...cC.border, bottom: { style: "thin", color: { argb: "FF000000" } } };
        cD.border = { ...cD.border, bottom: { style: "thin", color: { argb: "FF000000" } } };
      }
    }
  } else {
    // Cols D, E, F for Quotation / Invoice
    const setRightField = (rowNum: number, label: string, val: string) => {
      const lblCell = worksheet.getCell(`D${rowNum}`);
      lblCell.value = label;
      lblCell.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
      lblCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
      lblCell.alignment = { vertical: "middle", horizontal: "left" };

      worksheet.mergeCells(`E${rowNum}:F${rowNum}`);
      const valCell = worksheet.getCell(`E${rowNum}`);
      valCell.value = val;
      valCell.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "FF000000" } };
      valCell.alignment = { vertical: "middle", horizontal: "left" };
      valCell.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };
    };

    setRightField(12, "DATE:", dateVal || "");
    if (isInvoice) {
      setRightField(13, "INVOICE NO.:", metaOptions?.includeInvoiceNo !== false ? (invoiceNo || "") : "");
      setRightField(14, "CHALLAN NO.:", metaOptions?.includeChallanNo !== false ? (challanNo || "") : "");
      setRightField(15, "REQUISITION NO.:", metaOptions?.includeRequisitionNo !== false ? (requisitionNo || "") : "");
      setRightField(16, "P.O. NUMBER:", metaOptions?.includePoNumber !== false ? (poNumber || "") : "");
    } else {
      setRightField(13, "REQUISITION NO.:", metaOptions?.includeRequisitionNo !== false ? (requisitionNo || "") : "");
      setRightField(14, "P.O. NUMBER:", metaOptions?.includePoNumber !== false ? (poNumber || "") : "");
    }

    // Outer border for right box (D12:F16)
    for (let r = 12; r <= 16; r++) {
      const cD = worksheet.getCell(`D${r}`);
      const cF = worksheet.getCell(`F${r}`);
      cD.border = { ...cD.border, left: { style: "thin", color: { argb: "FF000000" } } };
      cF.border = { ...cF.border, right: { style: "thin", color: { argb: "FF000000" } } };
      if (r === 12) {
        worksheet.getCell(`D${r}`).border = { ...worksheet.getCell(`D${r}`).border, top: { style: "thin", color: { argb: "FF000000" } } };
        worksheet.getCell(`E${r}`).border = { ...worksheet.getCell(`E${r}`).border, top: { style: "thin", color: { argb: "FF000000" } } };
        worksheet.getCell(`F${r}`).border = { ...worksheet.getCell(`F${r}`).border, top: { style: "thin", color: { argb: "FF000000" } } };
      }
      if (r === 16) {
        worksheet.getCell(`D${r}`).border = { ...worksheet.getCell(`D${r}`).border, bottom: { style: "thin", color: { argb: "FF000000" } } };
        worksheet.getCell(`E${r}`).border = { ...worksheet.getCell(`E${r}`).border, bottom: { style: "thin", color: { argb: "FF000000" } } };
        worksheet.getCell(`F${r}`).border = { ...worksheet.getCell(`F${r}`).border, bottom: { style: "thin", color: { argb: "FF000000" } } };
      }
    }
  }

  // =========================================================================
  // ROW 17: Table Header
  // =========================================================================
  worksheet.getRow(17).height = 18;
  const thHeaders = isChallan
    ? [
        { col: 1, text: "SL", align: "center" as const },
        { col: 2, text: "Description of Marine Items / Spare Parts", align: "left" as const },
        { col: 3, text: "Qty", align: "center" as const },
        { col: 4, text: "Unit", align: "center" as const },
      ]
    : [
        { col: 1, text: "SL", align: "center" as const },
        { col: 2, text: "Description of Marine Items / Spare Parts", align: "left" as const },
        { col: 3, text: "Qty", align: "center" as const },
        { col: 4, text: "Unit", align: "center" as const },
        { col: 5, text: "Price", align: "center" as const },
        { col: 6, text: "Amount", align: "center" as const },
      ];

  thHeaders.forEach((th) => {
    const c = worksheet.getRow(17).getCell(th.col);
    c.value = th.text;
    c.font = { name: "Arial", size: 8, bold: true, color: { argb: "FF000000" } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
    c.alignment = { vertical: "middle", horizontal: th.align, wrapText: true };
    c.border = thinBorder;
  });

  // =========================================================================
  // ROWS 18+: Table Items
  // =========================================================================
  const startItemRow = 18;
  let currentRowNum = startItemRow;
  const descColWidth = isChallan ? 52 : 42;

  pageRows.forEach((row, rowIdx) => {
    const r = worksheet.getRow(currentRowNum);
    const slNumber = startSlIndex + rowIdx;

    const descLines = calculateItemVisualLines(row.desc, descColWidth, 8.5);
    const dynamicRowHeight = getItemRowHeight(descLines, 8.5);
    r.height = dynamicRowHeight;

    // Col 1: SL
    const cellSl = r.getCell(1);
    cellSl.value = slNumber;
    cellSl.font = { name: "Arial", size: 8, color: { argb: "FF000000" } };
    cellSl.alignment = { vertical: "middle", horizontal: "center" };
    cellSl.border = thinBorder;

    // Col 2: Description
    const cellDesc = r.getCell(2);
    const parsedDesc = parseHtmlToExcelRuns(row.desc, { name: "Arial", size: 8.5, color: { argb: "FF000000" } });
    cellDesc.value = parsedDesc.hasFormatting && parsedDesc.richText.length > 0 ? { richText: parsedDesc.richText } : parsedDesc.plainText;
    cellDesc.font = { name: "Arial", size: 8.5, color: { argb: "FF000000" } };
    cellDesc.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    cellDesc.border = thinBorder;
    if (parsedDesc.highlightColor) {
      cellDesc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: parsedDesc.highlightColor } };
    }

    // Col 3: Qty
    const cellQty = r.getCell(3);
    const cleanQty = htmlToPlainText(row.qty || "").trim();
    const numQty = parseFloat(cleanQty.replace(/,/g, ""));
    const isNumQty = cleanQty !== "" && !isNaN(numQty);
    cellQty.value = isNumQty ? numQty : cleanQty;
    cellQty.font = { name: "Arial", size: 8.5, color: { argb: "FF000000" } };
    cellQty.alignment = { vertical: "middle", horizontal: "center" };
    cellQty.border = thinBorder;
    if (isNumQty) {
      cellQty.numFmt = Number.isInteger(numQty) ? "#,##0" : "#,##0.##";
    }

    // Col 4: Unit
    const cellUnit = r.getCell(4);
    cellUnit.value = htmlToPlainText(row.unit || "").trim();
    cellUnit.font = { name: "Arial", size: 8.5, color: { argb: "FF000000" } };
    cellUnit.alignment = { vertical: "middle", horizontal: "center" };
    cellUnit.border = thinBorder;

    // Cols 5 & 6: Price & Amount (for Quotation and Invoice)
    if (!isChallan) {
      const cellPrice = r.getCell(5);
      const cleanPrice = htmlToPlainText(row.price || "").trim();
      const numPrice = parseFloat(cleanPrice.replace(/,/g, ""));
      const isNumPrice = cleanPrice !== "" && !isNaN(numPrice);
      cellPrice.value = isNumPrice ? numPrice : cleanPrice;
      cellPrice.font = { name: "Arial", size: 8.5, color: { argb: "FF000000" } };
      cellPrice.alignment = { vertical: "middle", horizontal: "right" };
      cellPrice.border = thinBorder;
      if (isNumPrice) {
        cellPrice.numFmt = "#,##0.00";
      }

      const cellAmount = r.getCell(6);
      const hasContent = cleanQty || cleanPrice || parsedDesc.plainText.trim();
      if (hasContent && isNumQty && isNumPrice) {
        cellAmount.value = { formula: `C${currentRowNum}*E${currentRowNum}`, result: numQty * numPrice };
      } else if (hasContent && typeof row.amount === "number" && row.amount > 0) {
        cellAmount.value = row.amount;
      } else {
        cellAmount.value = "";
      }
      cellAmount.font = { name: "Arial", size: 8.5, color: { argb: "FF000000" } };
      cellAmount.alignment = { vertical: "middle", horizontal: "right" };
      cellAmount.border = thinBorder;
      cellAmount.numFmt = "#,##0.00";
    }

    currentRowNum++;
  });

  const lastItemRowNum = currentRowNum - 1;

  // =========================================================================
  // TOTALS & SUMMARY SECTION (Last page only for Quotation / Invoice)
  // =========================================================================
  if (isLastPage && !isChallan && pageRows.length > 0) {
    // 1. Subtotal Row
    const subTotalRow = currentRowNum;
    worksheet.getRow(subTotalRow).height = 15;
    worksheet.mergeCells(`A${subTotalRow}:E${subTotalRow}`);
    const subLbl = worksheet.getCell(`A${subTotalRow}`);
    subLbl.value = "SUB TOTAL";
    subLbl.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "FF000000" } };
    subLbl.alignment = { vertical: "middle", horizontal: "right" };
    subLbl.border = thinBorder;

    const subVal = worksheet.getCell(`F${subTotalRow}`);
    subVal.value = { formula: `SUM(F${startItemRow}:F${lastItemRowNum})` };
    subVal.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "FF000000" } };
    subVal.alignment = { vertical: "middle", horizontal: "right" };
    subVal.border = thinBorder;
    subVal.numFmt = "#,##0.00";
    currentRowNum++;

    // 2. Discount Row (if applicable)
    let discountRow = 0;
    if (includeDiscount && discountAmount > 0) {
      discountRow = currentRowNum;
      worksheet.getRow(discountRow).height = 15;
      worksheet.mergeCells(`A${discountRow}:E${discountRow}`);
      const dLbl = worksheet.getCell(`A${discountRow}`);
      dLbl.value = `DISCOUNT ${discountType === "percentage" ? `(${discountValue}%)` : ""}`;
      dLbl.font = { name: "Arial", size: 8.5, color: { argb: "FF000000" } };
      dLbl.alignment = { vertical: "middle", horizontal: "right" };
      dLbl.border = thinBorder;

      const dVal = worksheet.getCell(`F${discountRow}`);
      if (discountType === "percentage") {
        dVal.value = { formula: `F${subTotalRow}*${discountValue / 100}`, result: discountAmount };
      } else {
        dVal.value = discountAmount;
      }
      dVal.font = { name: "Arial", size: 8.5, color: { argb: "FF000000" } };
      dVal.alignment = { vertical: "middle", horizontal: "right" };
      dVal.border = thinBorder;
      dVal.numFmt = "#,##0.00";
      currentRowNum++;
    }

    // 3. VAT Row (if applicable)
    let vatRow = 0;
    if (vatPercent > 0) {
      vatRow = currentRowNum;
      worksheet.getRow(vatRow).height = 15;
      worksheet.mergeCells(`A${vatRow}:E${vatRow}`);
      const vLbl = worksheet.getCell(`A${vatRow}`);
      vLbl.value = `VAT (${vatPercent}%)`;
      vLbl.font = { name: "Arial", size: 8.5, color: { argb: "FF000000" } };
      vLbl.alignment = { vertical: "middle", horizontal: "right" };
      vLbl.border = thinBorder;

      const vVal = worksheet.getCell(`F${vatRow}`);
      const baseCell = discountRow > 0 ? `(F${subTotalRow}-F${discountRow})` : `F${subTotalRow}`;
      vVal.value = { formula: `${baseCell}*${vatPercent / 100}` };
      vVal.font = { name: "Arial", size: 8.5, color: { argb: "FF000000" } };
      vVal.alignment = { vertical: "middle", horizontal: "right" };
      vVal.border = thinBorder;
      vVal.numFmt = "#,##0.00";
      currentRowNum++;
    }

    // 4. Transportation Fee Row (if applicable)
    let transportRow = 0;
    if (transportationFee > 0) {
      transportRow = currentRowNum;
      worksheet.getRow(transportRow).height = 15;
      worksheet.mergeCells(`A${transportRow}:E${transportRow}`);
      const tLbl = worksheet.getCell(`A${transportRow}`);
      tLbl.value = "TRANSPORTATION / CARRYING COST";
      tLbl.font = { name: "Arial", size: 8.5, color: { argb: "FF000000" } };
      tLbl.alignment = { vertical: "middle", horizontal: "right" };
      tLbl.border = thinBorder;

      const tVal = worksheet.getCell(`F${transportRow}`);
      tVal.value = transportationFee;
      tVal.font = { name: "Arial", size: 8.5, color: { argb: "FF000000" } };
      tVal.alignment = { vertical: "middle", horizontal: "right" };
      tVal.border = thinBorder;
      tVal.numFmt = "#,##0.00";
      currentRowNum++;
    }

    // 5. Grand Total Row
    const grandRow = currentRowNum;
    worksheet.getRow(grandRow).height = 18;
    worksheet.mergeCells(`A${grandRow}:E${grandRow}`);
    const gLbl = worksheet.getCell(`A${grandRow}`);
    gLbl.value = "GRAND TOTAL (BDT)";
    gLbl.font = { name: "Arial", size: 9, bold: true, color: { argb: "FF000000" } };
    gLbl.alignment = { vertical: "middle", horizontal: "right" };
    gLbl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
    gLbl.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "double", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };

    const gVal = worksheet.getCell(`F${grandRow}`);
    let grandFormula = `F${subTotalRow}`;
    if (discountRow > 0) grandFormula += `-F${discountRow}`;
    if (vatRow > 0) grandFormula += `+F${vatRow}`;
    if (transportRow > 0) grandFormula += `+F${transportRow}`;

    gVal.value = { formula: grandFormula };
    gVal.font = { name: "Arial", size: 9.5, bold: true, color: { argb: "FF000000" } };
    gVal.alignment = { vertical: "middle", horizontal: "right" };
    gVal.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
    gVal.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "double", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };
    gVal.numFmt = "#,##0.00";
    currentRowNum++;

    // 6. In Words Row
    const wordsRow = currentRowNum;
    worksheet.getRow(wordsRow).height = 16;
    worksheet.mergeCells(`A${wordsRow}:${lastColLetter}${wordsRow}`);
    const wordsCell = worksheet.getCell(`A${wordsRow}`);

    // Compute approximate numeric grand total for In Words display
    let approxSub = 0;
    pageRows.forEach((r) => {
      const q = parseFloat(htmlToPlainText(r.qty || "").replace(/,/g, ""));
      const p = parseFloat(htmlToPlainText(r.price || "").replace(/,/g, ""));
      if (!isNaN(q) && !isNaN(p)) approxSub += q * p;
    });
    const approxDiscount = discountType === "percentage" ? approxSub * (discountValue / 100) : discountAmount;
    const afterDiscount = Math.max(0, approxSub - approxDiscount);
    const approxVat = afterDiscount * (vatPercent / 100);
    const approxGrand = Math.round(afterDiscount + approxVat + transportationFee);

    wordsCell.value = `IN WORDS: ${numberToWords(approxGrand)}`;
    wordsCell.font = { name: "Arial", size: 8, bold: true, italic: true, color: { argb: "FF334155" } };
    wordsCell.alignment = { vertical: "middle", horizontal: "left" };
    wordsCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
    wordsCell.border = thinBorder;
    currentRowNum++;
  }

  // =========================================================================
  // SIGNATURE SECTION (Last Page only)
  // =========================================================================
  if (isLastPage) {
    // Spacer row
    worksheet.getRow(currentRowNum).height = 18;
    currentRowNum++;

    const sigLineRow = currentRowNum;
    worksheet.getRow(sigLineRow).height = 15;

    // Customer Acceptance on left
    worksheet.mergeCells(`A${sigLineRow}:B${sigLineRow}`);
    const custCell = worksheet.getCell(`A${sigLineRow}`);
    custCell.value = "Customer's Acceptance";
    custCell.font = { name: "Arial", size: 8, bold: true, color: { argb: "FF000000" } };
    custCell.alignment = { vertical: "top", horizontal: "center" };
    custCell.border = { top: { style: "thin", color: { argb: "FF000000" } } };

    // Authorized Signature on right
    const sigColStart = isChallan ? "C" : "E";
    worksheet.mergeCells(`${sigColStart}${sigLineRow}:${lastColLetter}${sigLineRow}`);
    const authCell = worksheet.getCell(`${sigColStart}${sigLineRow}`);
    authCell.value = "For " + companyName;
    authCell.font = { name: "Arial", size: 8, bold: true, color: { argb: "FF000000" } };
    authCell.alignment = { vertical: "top", horizontal: "center" };
    authCell.border = { top: { style: "thin", color: { argb: "FF000000" } } };

    currentRowNum++;
    const authSubRow = currentRowNum;
    worksheet.getRow(authSubRow).height = 14;
    worksheet.mergeCells(`${sigColStart}${authSubRow}:${lastColLetter}${authSubRow}`);
    const authSubCell = worksheet.getCell(`${sigColStart}${authSubRow}`);
    authSubCell.value = "Authorized Signature";
    authSubCell.font = { name: "Arial", size: 7.5, color: { argb: "475569" } };
    authSubCell.alignment = { vertical: "top", horizontal: "center" };
  }
};

/**
 * Main export function called by QuotationBuilder to generate Excel workbooks.
 */
export const generateExcelWorkbook = async (
  docType: "quotation" | "challan" | "invoice",
  messers: string,
  address: string,
  challanNo: string,
  dateVal: string,
  requisitionNo: string,
  rows: QuotationRow[],
  mergedRegions: MergedRegion[],
  invoiceNo?: string,
  poNumber?: string,
  vatPercent?: number,
  transportationFee?: number,
  cellFormats?: CellFormatMap,
  includeDiscount?: boolean,
  discountType?: "percentage" | "fixed",
  discountValue?: number,
  discountAmount?: number,
  companyId: "comilla" | "zainee" = "comilla",
  metaOptions?: ExcelMetadataOptions
): Promise<ExcelJS.Workbook> => {
  const isZainee = companyId === "zainee";
  const companyName = isZainee ? "Zainee Enterprise" : "Comilla Traders";

  const workbook = new ExcelJS.Workbook();
  workbook.creator = companyName;
  workbook.lastModifiedBy = companyName;
  workbook.created = new Date();
  workbook.modified = new Date();

  // If all rows are completely blank, supply at least 35 default empty rows so Excel is ready with plenty of items
  let effectiveRows = rows;
  const hasAnyData = rows.some(
    (r) => (r.desc && r.desc.trim() !== "") || (r.qty && r.qty.trim() !== "") || (r.price && r.price.trim() !== "")
  );

  if (!hasAnyData && rows.length < 35) {
    effectiveRows = [];
    for (let i = 1; i <= 35; i++) {
      effectiveRows.push({
        sl: i,
        desc: "",
        qty: "",
        unit: "",
        price: "",
        amount: 0,
      });
    }
  }

  // Dynamic capacity-aware pagination maximizing A4 page utilization (32-35+ items per page)
  const pageChunks = paginateRowsForExcel(effectiveRows, docType, cellFormats);
  const totalPages = pageChunks.length;

  const baseSheetNames: Record<string, string> = {
    quotation: "Quotation",
    challan: "Challan",
    invoice: "Invoice",
  };

  pageChunks.forEach((chunk, pageIndex) => {
    const sheetName =
      totalPages === 1
        ? baseSheetNames[docType] || "Quotation"
        : `${baseSheetNames[docType] || "Quotation"} - Page ${pageIndex + 1}`;

    const worksheet = workbook.addWorksheet(sheetName);

    renderPageWorksheet(
      worksheet,
      chunk.rows,
      chunk.startSlIndex,
      chunk.isLastPage,
      pageIndex + 1,
      totalPages,
      docType,
      messers,
      address,
      challanNo,
      dateVal,
      requisitionNo,
      mergedRegions,
      invoiceNo || "",
      poNumber || "",
      vatPercent || 0,
      transportationFee || 0,
      cellFormats,
      includeDiscount || false,
      discountType || "percentage",
      discountValue || 0,
      discountAmount || 0,
      companyName,
      metaOptions
    );
  });

  return workbook;
};
