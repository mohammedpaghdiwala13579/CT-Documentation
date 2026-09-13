import ExcelJS from "exceljs";
import { numberToWords } from "./numberToWords";

export interface ExcelExportRow {
  desc: string;
  qty: string;
  unit: string;
  price: string;
  amount: number;
}

export interface ExcelExportOptions {
  docType: "quotation" | "invoice" | "challan";
  companyName: string;
  // Format details toggles
  includeVesselName: boolean;
  includePortBerth: boolean;
  includeInvoiceNo: boolean;
  includeChallanNo: boolean;
  includeRequisitionNo: boolean;
  includePoNumber: boolean;
  includeDiscount: boolean;
  // Format details values
  messers: string;
  vesselName: string;
  portBerth: string;
  address: string;
  invoiceNo: string;
  challanNo: string;
  dateVal: string;
  requisitionNo: string;
  poNumber: string;
  // Table rows
  rows: ExcelExportRow[];
  // Calculations / Summary
  currency?: string;
  discountType?: "percentage" | "fixed";
  discountValue?: string;
  discountAmount?: number;
  vatPercent?: string;
  vatAmount?: number;
  transportationFee?: string;
  rowsTotal: number;
  grandTotal: number;
}

function stripHtmlPreservingBreaks(html: string): string {
  if (!html) return "";
  let text = html
    .replace(/<br\s*[\/]?>/gi, "\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<div>/gi, "")
    .replace(/<\/p>/gi, "\n")
    .replace(/<p>/gi, "")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/<[^>]+>/g, "");
  
  text = text.replace(/\r\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  return text;
}

function parseNumeric(val: string | number | undefined | null): number {
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  if (!val) return 0;
  const cleaned = String(val).replace(/[^0-9.-]+/g, "");
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

interface PageChunk {
  pageNum: number;
  start: number;
  count: number;
  isFirst: boolean;
  isLast: boolean;
}

function chunkRowsForA4(totalCount: number): PageChunk[] {
  const chunks: PageChunk[] = [];
  let remaining = totalCount;
  let offset = 0;
  let pageNum = 1;

  // Up to 14 items comfortably fit on Page 1 along with Format Details, Totals & Signatures
  if (remaining <= 14) {
    chunks.push({ pageNum: 1, start: 0, count: remaining, isFirst: true, isLast: true });
    return chunks;
  }

  // Multi-page document:
  // Page 1 without totals takes up to 14 items
  const p1Count = 14;
  chunks.push({ pageNum: 1, start: 0, count: p1Count, isFirst: true, isLast: false });
  remaining -= p1Count;
  offset += p1Count;
  pageNum++;

  while (remaining > 0) {
    // If remaining items fit on the final page with Totals (up to 17 items)
    if (remaining <= 17) {
      chunks.push({ pageNum, start: offset, count: remaining, isFirst: false, isLast: true });
      break;
    } else {
      // Continuation page without totals can take 18 items
      const pCount = 18;
      const thisCount = Math.min(remaining, pCount);
      const isLast = (remaining - thisCount === 0);
      chunks.push({ pageNum, start: offset, count: thisCount, isFirst: false, isLast });
      remaining -= thisCount;
      offset += thisCount;
      pageNum++;
      if (isLast) break;
    }
  }

  return chunks;
}

export async function generateExcelDocument(options: ExcelExportOptions): Promise<void> {
  const {
    docType,
    companyName,
    includeVesselName,
    includePortBerth,
    includeInvoiceNo,
    includeChallanNo,
    includeRequisitionNo,
    includePoNumber,
    includeDiscount,
    messers,
    vesselName,
    portBerth,
    address,
    invoiceNo,
    challanNo,
    dateVal,
    requisitionNo,
    poNumber,
    rows,
    currency = "Taka",
    discountType = "percentage",
    discountValue = "0",
    discountAmount = 0,
    vatPercent = "0",
    vatAmount = 0,
    transportationFee = "0",
    rowsTotal,
    grandTotal,
  } = options;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = companyName || "Comilla Traders";
  workbook.created = new Date();

  const isChallan = docType === "challan";
  const numCols = isChallan ? 5 : 6;
  const lastColLetter = isChallan ? "E" : "F";

  const sheetName = isChallan ? "Challan" : docType === "invoice" ? "Invoice" : "Quotation";
  const worksheet = workbook.addWorksheet(sheetName, {
    pageSetup: {
      paperSize: 9, // A4
      orientation: "portrait",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0, // Automatic vertical scaling - respects manual row breaks
      horizontalCentered: true,
      margins: {
        left: 0.35,
        right: 0.35,
        top: 0.35,
        bottom: 0.35,
        header: 0.1,
        footer: 0.1,
      },
    },
    views: [{ state: "normal", showGridLines: true }],
  });

  // Set column widths matching the PDF proportions
  if (isChallan) {
    worksheet.columns = [
      { key: "sl", width: 7 },
      { key: "desc", width: 56 },
      { key: "qty", width: 11 },
      { key: "unit", width: 11 },
      { key: "remarks", width: 18 },
    ];
  } else {
    worksheet.columns = [
      { key: "sl", width: 7 },
      { key: "desc", width: 50 },
      { key: "qty", width: 10 },
      { key: "unit", width: 10 },
      { key: "price", width: 14 },
      { key: "amount", width: 16 },
    ];
  }

  const slateFill: ExcelJS.Fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFF8FAFC" },
  };

  const headerFill: ExcelJS.Fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFF1F5F9" },
  };

  const headerBorder: Partial<ExcelJS.Borders> = {
    top: { style: "thin", color: { argb: "FF000000" } },
    bottom: { style: "thin", color: { argb: "FF000000" } },
    left: { style: "thin", color: { argb: "FF000000" } },
    right: { style: "thin", color: { argb: "FF000000" } },
  };

  const formatTitle = isChallan ? "CHALLAN" : docType === "invoice" ? "INVOICE" : "QUOTATION";
  const cleanMessers = stripHtmlPreservingBreaks(messers);
  const cleanAddress = stripHtmlPreservingBreaks(address);

  // Filter active non-empty rows
  const activeRows = rows.filter((r) => r.desc || r.qty || r.unit || r.price || r.amount > 0);
  const totalItemsCount = activeRows.length > 0 ? activeRows.length : 0;

  // Split items into A4 pages
  const pageChunks = chunkRowsForA4(totalItemsCount);
  const totalPages = pageChunks.length;

  // Track all item amount ranges for the final Subtotal formula: e.g. ["F18:F31", "F52:F62"]
  const amountRangeParts: string[] = [];

  // Helper to render the Signature Section on every page according to format
  const renderSignatureSection = (startRow: number): number => {
    // 1. Spacing row above signature box
    const spacerRow = startRow;
    worksheet.getRow(spacerRow).height = 10;

    // 2. Company Name title above signature (Right side - for Quotation & Invoice)
    const sigTitleRow = spacerRow + 1;
    if (!isChallan) {
      worksheet.getRow(sigTitleRow).height = 18;
      const rightSigCol = isChallan ? 4 : 5;
      worksheet.mergeCells(sigTitleRow, rightSigCol, sigTitleRow, numCols);
      const cellForCompany = worksheet.getCell(sigTitleRow, rightSigCol);
      cellForCompany.value = `For ${companyName || "COMILLA TRADERS"}`;
      cellForCompany.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "FF000000" } };
      cellForCompany.alignment = { horizontal: "center", vertical: "middle" };
    } else {
      worksheet.getRow(sigTitleRow).height = 8;
    }

    // 3. Gap row for physical hand signature
    const sigGapRow = sigTitleRow + 1;
    worksheet.getRow(sigGapRow).height = isChallan ? 30 : 22;

    // 4. Signature Line row with top border
    const sigLineRow = sigGapRow + 1;
    worksheet.getRow(sigLineRow).height = 20;

    // Left signature line: Receiver's Signature
    worksheet.mergeCells(sigLineRow, 1, sigLineRow, 2);
    const cellReceiver = worksheet.getCell(sigLineRow, 1);
    cellReceiver.value = "Receiver's Signature";
    cellReceiver.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "FF000000" } };
    cellReceiver.alignment = { horizontal: "center", vertical: "top" };
    cellReceiver.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
    };

    // Right signature line: Authorized Signature
    const rightSigCol = isChallan ? 4 : 5;
    worksheet.mergeCells(sigLineRow, rightSigCol, sigLineRow, numCols);
    const cellAuth = worksheet.getCell(sigLineRow, rightSigCol);
    cellAuth.value = "Authorized Signature";
    cellAuth.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "FF000000" } };
    cellAuth.alignment = { horizontal: "center", vertical: "top" };
    cellAuth.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
    };

    // 5. Gap row before notice
    const noticeGapRow = sigLineRow + 1;
    worksheet.getRow(noticeGapRow).height = 6;

    // 6. Non-returnable notice (matching PDF footer)
    const noticeRow = noticeGapRow + 1;
    worksheet.getRow(noticeRow).height = 18;
    worksheet.mergeCells(noticeRow, 1, noticeRow, numCols);
    const cellNotice = worksheet.getCell(noticeRow, 1);
    cellNotice.value = "ITEMS ONCE SOLD ARE NON-RETURNABLE AND NON-EXCHANGEABLE.";
    cellNotice.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "FF475569" } };
    cellNotice.alignment = { horizontal: "center", vertical: "middle" };

    return noticeRow;
  };

  let currentRowPointer = 1;

  // Process and render each page chunk
  for (let pIdx = 0; pIdx < pageChunks.length; pIdx++) {
    const chunk = pageChunks[pIdx];
    const pageNum = chunk.pageNum;
    const isFirstPage = chunk.isFirst;
    const isLastPage = chunk.isLast;

    if (isFirstPage) {
      // -------------------------------------------------------------
      // PAGE 1 LAYOUT
      // -------------------------------------------------------------
      // 1. TOP ROW: Format Name written on the top row (Row 1)
      const row1 = worksheet.getRow(1);
      row1.height = 28;
      worksheet.mergeCells(`A1:${lastColLetter}1`);
      const titleCell = worksheet.getCell("A1");
      titleCell.value = formatTitle;
      titleCell.font = {
        name: "Arial",
        size: 15,
        bold: true,
        color: { argb: "FF000000" },
      };
      titleCell.alignment = { horizontal: "center", vertical: "middle" };

      // 2. LEAVE 10 ROWS BLANK (Rows 2 to 11) for pre-printed letterhead pad
      for (let r = 2; r <= 11; r++) {
        const blankRow = worksheet.getRow(r);
        blankRow.height = 16;
      }

      // 3. FORMAT DETAILS (METADATA) SECTION - STRICTLY STARTS AT ROW 12
      // Ensures "Row 12 must not be blank" requirement is met directly!
      const leftItems: { label: string; value: string; isBold?: boolean }[] = [];
      leftItems.push({ label: "Messers:", value: cleanMessers || "", isBold: true });

      if (includeVesselName) {
        leftItems.push({ label: "Vessel Name:", value: vesselName || "", isBold: true });
      }
      if (includePortBerth) {
        leftItems.push({ label: "Port / Berth:", value: portBerth || "" });
      }
      leftItems.push({ label: "Address:", value: cleanAddress || "" });

      const rightItems: { label: string; value: string; isBold?: boolean }[] = [];
      if (docType === "invoice" && includeInvoiceNo) {
        rightItems.push({ label: "Invoice No.:", value: invoiceNo || "", isBold: true });
      }
      if ((docType === "invoice" || docType === "challan") && includeChallanNo) {
        rightItems.push({ label: "Challan No.:", value: challanNo || "", isBold: true });
      }
      // Date is standard
      rightItems.push({ label: "Date:", value: dateVal || "", isBold: true });

      if (includeRequisitionNo) {
        rightItems.push({ label: "Requisition No.:", value: requisitionNo || "" });
      }
      if (docType === "invoice" && includePoNumber) {
        rightItems.push({ label: "PO Number:", value: poNumber || "" });
      }

      const numMetaRows = Math.max(leftItems.length, rightItems.length, 3);
      const metaStartRow = 12; // Row 12 is NOT blank!
      const metaEndRow = metaStartRow + numMetaRows - 1;

      // Populate Left Box (Cols A to C) and Right Box (Cols D to F/E)
      for (let i = 0; i < numMetaRows; i++) {
        const r = metaStartRow + i;
        const row = worksheet.getRow(r);
        row.height = 20;

        // LEFT BOX
        const leftItem = leftItems[i];
        const cellA = worksheet.getCell(r, 1);
        cellA.fill = slateFill;
        cellA.font = { name: "Arial", size: 8, bold: true, color: { argb: "FF334155" } };
        cellA.alignment = { horizontal: "left", vertical: "middle" };

        worksheet.mergeCells(r, 2, r, 3);
        const cellBC = worksheet.getCell(r, 2);
        cellBC.fill = slateFill;
        cellBC.font = {
          name: "Arial",
          size: 8.5,
          bold: leftItem?.isBold ?? false,
          color: { argb: "FF000000" },
        };
        cellBC.alignment = { horizontal: "left", vertical: "middle", wrapText: true };

        if (leftItem) {
          cellA.value = leftItem.label;
          cellBC.value = leftItem.value;
          cellBC.border = {
            bottom: { style: "hair", color: { argb: "FF94A3B8" } },
          };
        }

        // RIGHT BOX
        const rightItem = rightItems[i];
        const cellD = worksheet.getCell(r, 4);
        cellD.fill = slateFill;
        cellD.font = { name: "Arial", size: 8, bold: true, color: { argb: "FF334155" } };
        cellD.alignment = { horizontal: "left", vertical: "middle" };

        const rightValueColEnd = numCols;
        if (rightValueColEnd > 4) {
          worksheet.mergeCells(r, 5, r, rightValueColEnd);
        }
        const cellRightVal = worksheet.getCell(r, 5);
        cellRightVal.fill = slateFill;
        cellRightVal.font = {
          name: "Arial",
          size: 8.5,
          bold: rightItem?.isBold ?? false,
          color: { argb: "FF000000" },
        };
        cellRightVal.alignment = { horizontal: "left", vertical: "middle", wrapText: true };

        if (rightItem) {
          cellD.value = rightItem.label;
          cellRightVal.value = rightItem.value;
          cellRightVal.border = {
            bottom: { style: "hair", color: { argb: "FF94A3B8" } },
          };
        }
      }

      // Outer border for Left Box (Columns A to C)
      for (let r = metaStartRow; r <= metaEndRow; r++) {
        for (let c = 1; c <= 3; c++) {
          const cell = worksheet.getCell(r, c);
          const b: Partial<ExcelJS.Borders> = { ...cell.border };
          if (r === metaStartRow) b.top = { style: "thin", color: { argb: "FF000000" } };
          if (r === metaEndRow) b.bottom = { style: "thin", color: { argb: "FF000000" } };
          if (c === 1) b.left = { style: "thin", color: { argb: "FF000000" } };
          if (c === 3) b.right = { style: "thin", color: { argb: "FF000000" } };
          cell.border = b;
        }
      }

      // Outer border for Right Box (Columns D to numCols)
      for (let r = metaStartRow; r <= metaEndRow; r++) {
        for (let c = 4; c <= numCols; c++) {
          const cell = worksheet.getCell(r, c);
          const b: Partial<ExcelJS.Borders> = { ...cell.border };
          if (r === metaStartRow) b.top = { style: "thin", color: { argb: "FF000000" } };
          if (r === metaEndRow) b.bottom = { style: "thin", color: { argb: "FF000000" } };
          if (c === 4) b.left = { style: "thin", color: { argb: "FF000000" } };
          if (c === numCols) b.right = { style: "thin", color: { argb: "FF000000" } };
          cell.border = b;
        }
      }

      // Spacer row before table
      const spacerRow1 = worksheet.getRow(metaEndRow + 1);
      spacerRow1.height = 6;

      currentRowPointer = metaEndRow + 2;
    } else {
      // -------------------------------------------------------------
      // CONTINUATION PAGE LAYOUT (PAGE 2, PAGE 3, ETC.)
      // -------------------------------------------------------------
      // 1. Top row: Format Name (Contd.)
      const pageStartRow = currentRowPointer;
      const topContRow = worksheet.getRow(pageStartRow);
      topContRow.height = 26;
      worksheet.mergeCells(`A${pageStartRow}:${lastColLetter}${pageStartRow}`);
      const titleContCell = worksheet.getCell(`A${pageStartRow}`);
      titleContCell.value = `${formatTitle} (Contd. - Page ${pageNum} of ${totalPages})`;
      titleContCell.font = {
        name: "Arial",
        size: 13,
        bold: true,
        color: { argb: "FF000000" },
      };
      titleContCell.alignment = { horizontal: "center", vertical: "middle" };

      // 2. Leave 10 rows blank for pre-printed letterhead pad clearance
      for (let r = pageStartRow + 1; r <= pageStartRow + 10; r++) {
        const blankRow = worksheet.getRow(r);
        blankRow.height = 16;
      }

      // 3. 12th row of this continuation page: Reference Header (Not blank!)
      const contRefRow = pageStartRow + 11;
      const refRowObj = worksheet.getRow(contRefRow);
      refRowObj.height = 20;

      // Col A: Messers
      const cellRefA = worksheet.getCell(contRefRow, 1);
      cellRefA.value = "Messers:";
      cellRefA.fill = slateFill;
      cellRefA.font = { name: "Arial", size: 8, bold: true, color: { argb: "FF334155" } };
      cellRefA.alignment = { horizontal: "left", vertical: "middle" };

      worksheet.mergeCells(contRefRow, 2, contRefRow, 3);
      const cellRefBC = worksheet.getCell(contRefRow, 2);
      cellRefBC.value = cleanMessers || "";
      cellRefBC.fill = slateFill;
      cellRefBC.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "FF000000" } };
      cellRefBC.alignment = { horizontal: "left", vertical: "middle" };
      cellRefBC.border = { bottom: { style: "hair", color: { argb: "FF94A3B8" } } };

      // Right: Document Identifier & Page reference
      const cellRefD = worksheet.getCell(contRefRow, 4);
      cellRefD.value = docType === "invoice" ? "Invoice No.:" : isChallan ? "Challan No.:" : "Date:";
      cellRefD.fill = slateFill;
      cellRefD.font = { name: "Arial", size: 8, bold: true, color: { argb: "FF334155" } };
      cellRefD.alignment = { horizontal: "left", vertical: "middle" };

      if (numCols > 4) {
        worksheet.mergeCells(contRefRow, 5, contRefRow, numCols);
      }
      const cellRefRightVal = worksheet.getCell(contRefRow, 5);
      const docIdentifier = docType === "invoice" ? invoiceNo : isChallan ? challanNo : dateVal;
      cellRefRightVal.value = `${docIdentifier || "NEW"} (Page ${pageNum})`;
      cellRefRightVal.fill = slateFill;
      cellRefRightVal.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "FF000000" } };
      cellRefRightVal.alignment = { horizontal: "left", vertical: "middle" };
      cellRefRightVal.border = { bottom: { style: "hair", color: { argb: "FF94A3B8" } } };

      // Border around reference boxes
      for (let c = 1; c <= 3; c++) {
        const cell = worksheet.getCell(contRefRow, c);
        const b: Partial<ExcelJS.Borders> = { ...cell.border, top: { style: "thin", color: { argb: "FF000000" } }, bottom: { style: "thin", color: { argb: "FF000000" } } };
        if (c === 1) b.left = { style: "thin", color: { argb: "FF000000" } };
        if (c === 3) b.right = { style: "thin", color: { argb: "FF000000" } };
        cell.border = b;
      }
      for (let c = 4; c <= numCols; c++) {
        const cell = worksheet.getCell(contRefRow, c);
        const b: Partial<ExcelJS.Borders> = { ...cell.border, top: { style: "thin", color: { argb: "FF000000" } }, bottom: { style: "thin", color: { argb: "FF000000" } } };
        if (c === 4) b.left = { style: "thin", color: { argb: "FF000000" } };
        if (c === numCols) b.right = { style: "thin", color: { argb: "FF000000" } };
        cell.border = b;
      }

      // Spacer row
      worksheet.getRow(contRefRow + 1).height = 6;
      currentRowPointer = contRefRow + 2;
    }

    // -------------------------------------------------------------
    // DATA TABLE HEADER FOR THIS PAGE
    // -------------------------------------------------------------
    const tableHeaderRow = currentRowPointer;
    const headerRowObj = worksheet.getRow(tableHeaderRow);
    headerRowObj.height = 22;

    const headers = isChallan
      ? [
          { col: 1, text: "SL", align: "center" as const },
          { col: 2, text: "Description of Marine Items / Spare Parts", align: "left" as const },
          { col: 3, text: "Qty", align: "center" as const },
          { col: 4, text: "Unit", align: "center" as const },
          { col: 5, text: "Remarks", align: "center" as const },
        ]
      : [
          { col: 1, text: "SL", align: "center" as const },
          { col: 2, text: "Description of Marine Items / Spare Parts", align: "left" as const },
          { col: 3, text: "Qty", align: "center" as const },
          { col: 4, text: "Unit", align: "center" as const },
          { col: 5, text: "Price", align: "right" as const },
          { col: 6, text: "Amount", align: "right" as const },
        ];

    headers.forEach((h) => {
      const cell = worksheet.getCell(tableHeaderRow, h.col);
      cell.value = h.text;
      cell.fill = headerFill;
      cell.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "FF000000" } };
      cell.alignment = { horizontal: h.align, vertical: "middle" };
      cell.border = headerBorder;
    });

    // -------------------------------------------------------------
    // DATA ROWS FOR THIS PAGE
    // -------------------------------------------------------------
    const pageItems = activeRows.slice(chunk.start, chunk.start + chunk.count);
    
    // If single page with very few items, pad to at least 6 rows so table looks visually complete
    const minPaddingCount = isFirstPage && isLastPage ? Math.max(6, pageItems.length) : pageItems.length;
    const paddedItems: (ExcelExportRow | null)[] = [...pageItems];
    while (paddedItems.length < minPaddingCount) {
      paddedItems.push(null);
    }

    const dataStartRow = tableHeaderRow + 1;
    let currentRow = dataStartRow;

    paddedItems.forEach((item, itemIdx) => {
      const globalSl = chunk.start + itemIdx + 1;
      const rowObj = worksheet.getRow(currentRow);
      const cleanDesc = item ? stripHtmlPreservingBreaks(item.desc) : "";
      const lineCount = cleanDesc.split("\n").length;
      rowObj.height = Math.max(20, lineCount * 15);

      // Col 1: SL
      const cellSl = worksheet.getCell(currentRow, 1);
      cellSl.value = item ? globalSl : "";
      cellSl.font = { name: "Arial", size: 8.5, color: { argb: "FF334155" } };
      cellSl.alignment = { horizontal: "center", vertical: "middle" };
      cellSl.border = headerBorder;

      // Col 2: Description
      const cellDesc = worksheet.getCell(currentRow, 2);
      cellDesc.value = cleanDesc;
      cellDesc.font = { name: "Arial", size: 8.5, color: { argb: "FF000000" } };
      cellDesc.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
      cellDesc.border = headerBorder;

      // Col 3: Qty
      const cellQty = worksheet.getCell(currentRow, 3);
      if (item && item.qty) {
        if (!isNaN(Number(item.qty.trim()))) {
          cellQty.value = Number(item.qty.trim());
          cellQty.numFmt = Number.isInteger(Number(item.qty.trim())) ? "#,##0" : "#,##0.00";
        } else {
          const numQty = parseNumeric(item.qty);
          cellQty.value = item.qty || (numQty > 0 ? numQty : "");
        }
      } else {
        cellQty.value = "";
      }
      cellQty.font = { name: "Arial", size: 8.5, color: { argb: "FF000000" } };
      cellQty.alignment = { horizontal: "center", vertical: "middle" };
      cellQty.border = headerBorder;

      // Col 4: Unit
      const cellUnit = worksheet.getCell(currentRow, 4);
      cellUnit.value = item ? (item.unit || "").toUpperCase() : "";
      cellUnit.font = { name: "Arial", size: 8.5, color: { argb: "FF000000" } };
      cellUnit.alignment = { horizontal: "center", vertical: "middle" };
      cellUnit.border = headerBorder;

      if (!isChallan) {
        // Col 5: Price
        const cellPrice = worksheet.getCell(currentRow, 5);
        if (item && item.price) {
          const numPrice = parseNumeric(item.price);
          if (!isNaN(Number(item.price.trim()))) {
            cellPrice.value = Number(item.price.trim());
          } else {
            cellPrice.value = numPrice > 0 ? numPrice : "";
          }
          cellPrice.numFmt = "#,##0.00";
        } else {
          cellPrice.value = "";
        }
        cellPrice.font = { name: "Arial", size: 8.5, color: { argb: "FF000000" } };
        cellPrice.alignment = { horizontal: "right", vertical: "middle" };
        cellPrice.border = headerBorder;

        // Col 6: Amount (Formula = C{row} * E{row} for editable live calculation)
        const cellAmount = worksheet.getCell(currentRow, 6);
        if (item) {
          cellAmount.value = {
            formula: `IF(AND(ISNUMBER(C${currentRow}),ISNUMBER(E${currentRow})),C${currentRow}*E${currentRow},${item.amount || 0})`,
            result: item.amount || 0,
          };
          cellAmount.numFmt = "#,##0.00";
        } else {
          cellAmount.value = "";
        }
        cellAmount.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "FF000000" } };
        cellAmount.alignment = { horizontal: "right", vertical: "middle" };
        cellAmount.border = headerBorder;
      } else {
        // Col 5: Remarks (for Challan)
        const cellRemarks = worksheet.getCell(currentRow, 5);
        cellRemarks.value = "";
        cellRemarks.font = { name: "Arial", size: 8.5, color: { argb: "FF000000" } };
        cellRemarks.alignment = { horizontal: "center", vertical: "middle" };
        cellRemarks.border = headerBorder;
      }

      currentRow++;
    });

    const dataEndRow = currentRow - 1;
    if (pageItems.length > 0 && !isChallan) {
      amountRangeParts.push(`F${dataStartRow}:F${dataStartRow + pageItems.length - 1}`);
    }

    let lastContentRowOfPage = dataEndRow;

    // -------------------------------------------------------------
    // SUMMARY / TOTALS SECTION (ONLY ON THE FINAL PAGE)
    // -------------------------------------------------------------
    if (isLastPage && !isChallan) {
      const summaryRows: {
        label: string;
        valueFormula?: string;
        valueNumber?: number;
        isGrandTotal?: boolean;
        isDiscount?: boolean;
      }[] = [];

      const sumRangesFormula = amountRangeParts.length > 0 ? `SUM(${amountRangeParts.join(",")})` : "0";

      if (docType === "quotation") {
        summaryRows.push({
          label: "TOTAL =",
          valueFormula: sumRangesFormula,
          valueNumber: rowsTotal,
          isGrandTotal: true,
        });
      } else {
        // Invoice summary rows
        const subtotalRowIdx = dataEndRow + 1;
        summaryRows.push({
          label: "SUBTOTAL =",
          valueFormula: sumRangesFormula,
          valueNumber: rowsTotal,
        });

        let discRowIdx: number | null = null;
        const parsedDisc = parseNumeric(discountValue);
        if (includeDiscount && parsedDisc > 0) {
          discRowIdx = subtotalRowIdx + summaryRows.length;
          const discLabel = discountType === "percentage" ? `DISCOUNT (${parsedDisc}%) =` : "DISCOUNT =";
          const discFormula =
            discountType === "percentage"
              ? `ROUND(F${subtotalRowIdx}*${parsedDisc / 100},2)`
              : `${discountAmount}`;
          summaryRows.push({
            label: discLabel,
            valueFormula: discFormula,
            valueNumber: -Math.abs(discountAmount),
            isDiscount: true,
          });
        }

        let vatRowIdx: number | null = null;
        const parsedVat = parseNumeric(vatPercent);
        if (parsedVat > 0) {
          vatRowIdx = subtotalRowIdx + summaryRows.length;
          const baseRowForVat = discRowIdx ? `(F${subtotalRowIdx}-F${discRowIdx})` : `F${subtotalRowIdx}`;
          summaryRows.push({
            label: `VAT (${parsedVat}%) =`,
            valueFormula: `ROUND(${baseRowForVat}*${parsedVat / 100},2)`,
            valueNumber: vatAmount,
          });
        }

        let transRowIdx: number | null = null;
        const parsedTrans = parseNumeric(transportationFee);
        if (parsedTrans > 0) {
          transRowIdx = subtotalRowIdx + summaryRows.length;
          summaryRows.push({
            label: "TRANS. =",
            valueNumber: parsedTrans,
          });
        }

        // Grand Total Formula
        let grandTotalFormula = `F${subtotalRowIdx}`;
        if (discRowIdx) grandTotalFormula += `-F${discRowIdx}`;
        if (vatRowIdx) grandTotalFormula += `+F${vatRowIdx}`;
        if (transRowIdx) grandTotalFormula += `+F${transRowIdx}`;

        summaryRows.push({
          label: "GRAND TOTAL =",
          valueFormula: grandTotalFormula,
          valueNumber: grandTotal,
          isGrandTotal: true,
        });
      }

      const summaryStartRow = dataEndRow + 1;
      const summaryEndRow = summaryStartRow + summaryRows.length - 1;

      // Merge Amount in Words across Columns A to D for all summary rows
      worksheet.mergeCells(`A${summaryStartRow}:D${summaryEndRow}`);
      const cellWords = worksheet.getCell(`A${summaryStartRow}`);
      const wordsText = numberToWords(grandTotal, currency);
      cellWords.value = `Amount in Words:\n${wordsText}`;
      cellWords.font = { name: "Arial", size: 8, bold: true, color: { argb: "FF000000" } };
      cellWords.fill = slateFill;
      cellWords.alignment = { horizontal: "left", vertical: "middle", wrapText: true };

      // Outer border for Amount in Words block
      for (let r = summaryStartRow; r <= summaryEndRow; r++) {
        for (let c = 1; c <= 4; c++) {
          const cell = worksheet.getCell(r, c);
          const b: Partial<ExcelJS.Borders> = { ...cell.border };
          if (r === summaryStartRow) b.top = { style: "thin", color: { argb: "FF000000" } };
          if (r === summaryEndRow) b.bottom = { style: "thin", color: { argb: "FF000000" } };
          if (c === 1) b.left = { style: "thin", color: { argb: "FF000000" } };
          if (c === 4) b.right = { style: "thin", color: { argb: "FF000000" } };
          cell.border = b;
        }
      }

      // Populate Right Summary rows (Columns E & F)
      summaryRows.forEach((item, idx) => {
        const r = summaryStartRow + idx;
        const rowObj = worksheet.getRow(r);
        rowObj.height = item.isGrandTotal ? 22 : 20;

        const cellLbl = worksheet.getCell(r, 5);
        cellLbl.value = item.label;
        cellLbl.font = {
          name: "Arial",
          size: item.isGrandTotal ? 8.5 : 8,
          bold: true,
          color: item.isGrandTotal ? { argb: "FF1E1B4B" } : { argb: "FF000000" },
        };
        cellLbl.alignment = { horizontal: "right", vertical: "middle" };
        cellLbl.fill = item.isGrandTotal
          ? { type: "pattern", pattern: "solid", fgColor: { argb: "FFEEF2FF" } }
          : slateFill;
        cellLbl.border = headerBorder;

        const cellVal = worksheet.getCell(r, 6);
        if (item.valueFormula) {
          cellVal.value = {
            formula: item.valueFormula,
            result: item.valueNumber || 0,
          };
        } else {
          cellVal.value = item.valueNumber || 0;
        }
        cellVal.numFmt = "#,##0.00";
        cellVal.font = {
          name: "Arial",
          size: item.isGrandTotal ? 9.5 : 8.5,
          bold: true,
          color: item.isDiscount
            ? { argb: "FFB91C1C" }
            : item.isGrandTotal
            ? { argb: "FF1E1B4B" }
            : { argb: "FF000000" },
        };
        cellVal.alignment = { horizontal: "right", vertical: "middle" };
        cellVal.fill = item.isGrandTotal
          ? { type: "pattern", pattern: "solid", fgColor: { argb: "FFEEF2FF" } }
          : { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
        cellVal.border = headerBorder;
      });

      lastContentRowOfPage = summaryEndRow;
    }

    // -------------------------------------------------------------
    // SIGNATURE SECTION ON EVERY PAGE ACCORDING TO FORMAT
    // -------------------------------------------------------------
    const endNoticeRow = renderSignatureSection(lastContentRowOfPage + 1);

    // If there is another page, insert a manual page break right after the notice
    if (!isLastPage) {
      worksheet.getRow(endNoticeRow).addPageBreak();
    }

    // Advance pointer for the next page
    currentRowPointer = endNoticeRow + 1;
  }

  // Generate binary buffer and download in browser
  const buffer = await workbook.xlsx.writeBuffer();
  const filePrefix = isChallan ? "Challan" : docType === "invoice" ? "Invoice" : "Quotation";
  const identifier = isChallan
    ? challanNo || "NEW"
    : docType === "invoice"
    ? invoiceNo || "NEW"
    : requisitionNo || "NEW";
  const filename = `${filePrefix}_${identifier.replace(/[\/\\?%*:|"<>\s]/g, "_")}.xlsx`;

  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = window.URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  window.URL.revokeObjectURL(url);
}
