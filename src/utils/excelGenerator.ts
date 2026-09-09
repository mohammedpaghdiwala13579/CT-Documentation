import ExcelJS from "exceljs";
import { QuotationRow, MergedRegion, CellFormatMap, CellFormat, CellBorders } from "../types";
import { numberToWords } from "./numberToWords";
import { cleanCellText } from "./tsvParser";
import { parseNumericInput, stripHtml } from "./textFormatter";

/**
 * Converts CSS color (hex, rgb, rgba, hsl, named colors) to 8-character ARGB for ExcelJS.
 * Supports modern and legacy CSS color syntax, with DOM fallback in browser.
 */
export function colorToArgb(color: string | null | undefined): string | null {
  if (!color) return null;
  const c = color.trim().toLowerCase();
  if (c === "transparent" || c === "none" || c === "initial" || c === "inherit" || c === "") return null;

  // Hex format #RGB, #RGBA, #RRGGBB, #RRGGBBAA
  if (c.startsWith("#")) {
    const hex = c.slice(1);
    if (hex.length === 3) {
      return (
        "FF" +
        hex[0] + hex[0] +
        hex[1] + hex[1] +
        hex[2] + hex[2]
      ).toUpperCase();
    }
    if (hex.length === 4) {
      const a = hex[3] + hex[3];
      return (
        a +
        hex[0] + hex[0] +
        hex[1] + hex[1] +
        hex[2] + hex[2]
      ).toUpperCase();
    }
    if (hex.length === 6) {
      return ("FF" + hex).toUpperCase();
    }
    if (hex.length === 8) {
      return (hex.slice(6, 8) + hex.slice(0, 6)).toUpperCase();
    }
  }

  // rgb(r, g, b) or rgba(r, g, b, a) - legacy and modern space/slash syntax
  const rgbMatch = c.match(/^rgba?\(\s*([0-9]+)\s*[, ]\s*([0-9]+)\s*[, ]\s*([0-9]+)(?:\s*[,/]\s*([0-9.]+))?\s*\)$/);
  if (rgbMatch) {
    const r = parseInt(rgbMatch[1], 10);
    const g = parseInt(rgbMatch[2], 10);
    const b = parseInt(rgbMatch[3], 10);
    const a = rgbMatch[4] !== undefined ? parseFloat(rgbMatch[4]) : 1;
    if (a <= 0) return null;
    const aHex = Math.round(Math.max(0, Math.min(1, a)) * 255).toString(16).padStart(2, "0");
    const rHex = Math.max(0, Math.min(255, r)).toString(16).padStart(2, "0");
    const gHex = Math.max(0, Math.min(255, g)).toString(16).padStart(2, "0");
    const bHex = Math.max(0, Math.min(255, b)).toString(16).padStart(2, "0");
    return (aHex + rHex + gHex + bHex).toUpperCase();
  }

  const NAMED_COLORS: Record<string, string> = {
    black: "FF000000",
    white: "FFFFFFFF",
    red: "FFFF0000",
    green: "FF008000",
    blue: "FF0000FF",
    yellow: "FFFFFF00",
    amber: "FFF59E0B",
    gray: "FF808080",
    grey: "FF808080",
    darkgray: "FFA9A9A9",
    darkgrey: "FFA9A9A9",
    lightgray: "FFD3D3D3",
    lightgrey: "FFD3D3D3",
    slate: "FF64748B",
    orange: "FFF97316",
    purple: "FFA855F7",
    teal: "FF14B8A6",
    cyan: "FF06B6D4",
    gold: "FFFFD700",
    navy: "FF000080",
    maroon: "FF800000",
    olive: "FF808000",
    lime: "FF00FF00",
    aqua: "FF00FFFF",
    fuchsia: "FFFF00FF",
    silver: "FFC0C0C0",
  };
  if (NAMED_COLORS[c]) {
    return NAMED_COLORS[c];
  }

  if (/^[0-9a-f]{6}$/i.test(c)) {
    return ("FF" + c).toUpperCase();
  }
  if (/^[0-9a-f]{8}$/i.test(c)) {
    return c.toUpperCase();
  }

  // Browser DOM fallback for any valid CSS color string (e.g. HSL or Tailwind palette)
  if (typeof document !== "undefined") {
    try {
      const dummy = document.createElement("div");
      dummy.style.color = color;
      const comp = dummy.style.color;
      if (comp && comp !== color) {
        const m = comp.match(/^rgba?\(\s*([0-9]+)\s*,\s*([0-9]+)\s*,\s*([0-9]+)(?:\s*,\s*([0-9.]+))?\s*\)$/);
        if (m) {
          const r = parseInt(m[1], 10);
          const g = parseInt(m[2], 10);
          const b = parseInt(m[3], 10);
          const a = m[4] !== undefined ? parseFloat(m[4]) : 1;
          const aHex = Math.round(a * 255).toString(16).padStart(2, "0");
          const rHex = Math.max(0, Math.min(255, r)).toString(16).padStart(2, "0");
          const gHex = Math.max(0, Math.min(255, g)).toString(16).padStart(2, "0");
          const bHex = Math.max(0, Math.min(255, b)).toString(16).padStart(2, "0");
          return (aHex + rHex + gHex + bHex).toUpperCase();
        }
      }
    } catch (e) {}
  }

  return null;
}

export type ExcelBorderDef = Partial<ExcelJS.Border>;

/**
 * Parses CSS border string (e.g. "1px solid black", "2px medium black", "none") into ExcelJS Border
 */
export function parseBorderSide(
  borderStr?: string,
  defaultStyle?: ExcelJS.BorderStyle
): ExcelBorderDef | undefined {
  if (borderStr === undefined) {
    if (!defaultStyle) return undefined;
    return { style: defaultStyle, color: { argb: "FF000000" } };
  }
  const s = borderStr.trim().toLowerCase();
  if (s === "none" || s === "0" || s === "0px" || s === "hidden" || s === "") {
    return undefined;
  }

  let style: ExcelJS.BorderStyle = "thin";
  if (s.includes("double")) {
    style = "double";
  } else if (s.includes("3px") || s.includes("thick")) {
    style = "thick";
  } else if (s.includes("2px") || s.includes("medium")) {
    style = "medium";
  } else if (s.includes("dotted")) {
    style = "dotted";
  } else if (s.includes("dashed")) {
    style = "dashed";
  } else {
    style = "thin";
  }

  let argb = "FF000000";
  const hexMatch = s.match(/#([0-9a-f]{3,8})/i);
  if (hexMatch) {
    const parsed = colorToArgb(hexMatch[0]);
    if (parsed) argb = parsed;
  } else {
    const rgbMatch = s.match(/rgba?\([^)]+\)/i);
    if (rgbMatch) {
      const parsed = colorToArgb(rgbMatch[0]);
      if (parsed) argb = parsed;
    }
  }

  return { style, color: { argb } };
}

/**
 * Converts HTML into plain text while preserving intentional line breaks (<br>, </p>, </div>)
 * and unescaping HTML entities cleanly.
 */
export const htmlToPlainText = (html: string): string => {
  if (!html) return "";
  const replaced = html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<\/tr>/gi, "\n")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'");

  if (typeof DOMParser !== "undefined") {
    try {
      const doc = new DOMParser().parseFromString(replaced, "text/html");
      return (doc.body.textContent || "").replace(/\u00A0/g, " ");
    } catch (e) {}
  }
  return replaced.replace(/<[^>]*>/g, "").replace(/\u00A0/g, " ");
};

interface TextRunStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | "single" | "double";
  strike?: boolean;
  color?: string; // ARGB
  bgColor?: string; // ARGB
  size?: number; // pt
  fontFamily?: string;
}

export interface ParsedHtmlResult {
  richText: ExcelJS.RichText[];
  plainText: string;
  hasFormatting: boolean;
  highlightColor?: string; // ARGB
}

/**
 * Parses an HTML string (from contenteditable RichTextCell) into ExcelJS richText runs
 * preserving specific word highlight, color, size, bold, italic, underline, and fonts.
 */
export function parseHtmlToExcelRuns(
  htmlOrText: string,
  baseFont: Partial<ExcelJS.Font> = {}
): ParsedHtmlResult {
  if (!htmlOrText) {
    return { richText: [], plainText: "", hasFormatting: false };
  }

  // If no HTML tags and no entities, return single run with base font
  if (!htmlOrText.includes("<") && !htmlOrText.includes("&")) {
    return {
      richText: [
        {
          text: htmlOrText,
          font: { ...baseFont },
        },
      ],
      plainText: htmlOrText,
      hasFormatting: false,
    };
  }

  if (typeof DOMParser === "undefined") {
    const clean = stripHtml(htmlOrText);
    return {
      richText: [{ text: clean, font: { ...baseFont } }],
      plainText: clean,
      hasFormatting: false,
    };
  }

  const normalizedHtml = htmlOrText
    .replace(/<br\s*\/?>/gi, "<br>")
    .replace(/&nbsp;/gi, " ");

  const doc = new DOMParser().parseFromString(`<div>${normalizedHtml}</div>`, "text/html");
  const root = doc.body.firstElementChild || doc.body;

  const rawRuns: { text: string; style: TextRunStyle }[] = [];
  let foundHighlight: string | undefined = undefined;
  let hasSpecialFormatting = false;

  const traverse = (node: Node, currentStyle: TextRunStyle) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent || "";
      if (text.length > 0) {
        rawRuns.push({ text, style: { ...currentStyle } });
      }
      return;
    }

    if (node.nodeType === Node.ELEMENT_NODE) {
      const el = node as HTMLElement;
      const tag = el.tagName.toUpperCase();

      if (tag === "BR") {
        rawRuns.push({ text: "\n", style: { ...currentStyle } });
        return;
      }

      const nextStyle: TextRunStyle = { ...currentStyle };

      // HTML Tags
      if (tag === "B" || tag === "STRONG") {
        nextStyle.bold = true;
        hasSpecialFormatting = true;
      }
      if (tag === "I" || tag === "EM") {
        nextStyle.italic = true;
        hasSpecialFormatting = true;
      }
      if (tag === "U") {
        nextStyle.underline = true;
        hasSpecialFormatting = true;
      }
      if (tag === "S" || tag === "STRIKE" || tag === "DEL") {
        nextStyle.strike = true;
        hasSpecialFormatting = true;
      }
      if (tag === "MARK") {
        const bg = colorToArgb(el.style.backgroundColor || "#fef08a");
        if (bg) {
          nextStyle.bgColor = bg;
          foundHighlight = bg;
          hasSpecialFormatting = true;
        }
      }
      if (tag === "FONT") {
        const fontColor = el.getAttribute("color");
        if (fontColor) {
          const argb = colorToArgb(fontColor);
          if (argb) {
            nextStyle.color = argb;
            hasSpecialFormatting = true;
          }
        }
        const face = el.getAttribute("face");
        if (face) {
          nextStyle.fontFamily = face;
          hasSpecialFormatting = true;
        }
        const sizeAttr = el.getAttribute("size");
        if (sizeAttr) {
          const num = parseInt(sizeAttr, 10);
          if (!isNaN(num)) {
            const sizeMap = [8, 9, 10, 11, 14, 18, 24, 36];
            nextStyle.size = sizeMap[Math.min(num, sizeMap.length - 1)];
            hasSpecialFormatting = true;
          }
        }
      }

      // Inline Styles
      if (el.style) {
        if (el.style.fontWeight) {
          const fw = el.style.fontWeight.toLowerCase();
          if (fw === "bold" || fw === "700" || fw === "800" || fw === "900" || fw === "bolder") {
            nextStyle.bold = true;
            hasSpecialFormatting = true;
          } else if (fw === "normal" || fw === "400") {
            nextStyle.bold = false;
          }
        }

        if (el.style.fontStyle) {
          const fs = el.style.fontStyle.toLowerCase();
          if (fs === "italic" || fs === "oblique") {
            nextStyle.italic = true;
            hasSpecialFormatting = true;
          } else if (fs === "normal") {
            nextStyle.italic = false;
          }
        }

        if (el.style.textDecoration || el.style.textDecorationLine) {
          const td = (el.style.textDecoration || el.style.textDecorationLine).toLowerCase();
          if (td.includes("underline")) {
            nextStyle.underline = td.includes("double") || el.style.textDecorationStyle === "double" ? "double" : true;
            hasSpecialFormatting = true;
          } else if (td.includes("line-through")) {
            nextStyle.strike = true;
            hasSpecialFormatting = true;
          } else if (td === "none") {
            nextStyle.underline = false;
            nextStyle.strike = false;
          }
        }

        if (el.style.color) {
          const argb = colorToArgb(el.style.color);
          if (argb) {
            nextStyle.color = argb;
            hasSpecialFormatting = true;
          }
        }

        if (el.style.backgroundColor || el.style.background) {
          const bg = colorToArgb(el.style.backgroundColor || el.style.background);
          if (bg) {
            nextStyle.bgColor = bg;
            foundHighlight = bg;
            hasSpecialFormatting = true;
          }
        }

        if (el.style.fontSize) {
          const fsStr = el.style.fontSize.trim().toLowerCase();
          if (fsStr.endsWith("pt")) {
            nextStyle.size = parseFloat(fsStr);
            hasSpecialFormatting = true;
          } else if (fsStr.endsWith("px")) {
            nextStyle.size = Math.round(parseFloat(fsStr) * 0.75);
            hasSpecialFormatting = true;
          } else {
            const num = parseFloat(fsStr);
            if (!isNaN(num) && num > 0) {
              nextStyle.size = num;
              hasSpecialFormatting = true;
            }
          }
        }

        if (el.style.fontFamily) {
          const cleaned = el.style.fontFamily.replace(/['"]/g, "").split(",")[0].trim();
          if (cleaned) {
            nextStyle.fontFamily = cleaned;
            hasSpecialFormatting = true;
          }
        }
      }

      const isBlock = tag === "DIV" || tag === "P" || tag === "TR";
      if (isBlock && rawRuns.length > 0 && !rawRuns[rawRuns.length - 1].text.endsWith("\n")) {
        rawRuns.push({ text: "\n", style: { ...currentStyle } });
      }

      for (let i = 0; i < node.childNodes.length; i++) {
        traverse(node.childNodes[i], nextStyle);
      }

      if (isBlock && rawRuns.length > 0 && !rawRuns[rawRuns.length - 1].text.endsWith("\n")) {
        rawRuns.push({ text: "\n", style: { ...currentStyle } });
      }
    }
  };

  const baseArgb = baseFont.color?.argb ? String(baseFont.color.argb) : undefined;
  traverse(root, {
    bold: baseFont.bold,
    italic: baseFont.italic,
    underline: baseFont.underline === "double" ? "double" : !!baseFont.underline,
    color: baseArgb,
    size: baseFont.size,
    fontFamily: baseFont.name,
  });

  // Consolidate adjacent runs with identical styling
  const consolidatedRuns: { text: string; style: TextRunStyle }[] = [];
  for (const run of rawRuns) {
    if (!run.text) continue;
    if (consolidatedRuns.length === 0) {
      consolidatedRuns.push({ ...run });
    } else {
      const prev = consolidatedRuns[consolidatedRuns.length - 1];
      const sameStyle =
        prev.style.bold === run.style.bold &&
        prev.style.italic === run.style.italic &&
        prev.style.underline === run.style.underline &&
        prev.style.strike === run.style.strike &&
        prev.style.color === run.style.color &&
        prev.style.bgColor === run.style.bgColor &&
        prev.style.size === run.style.size &&
        prev.style.fontFamily === run.style.fontFamily;
      if (sameStyle) {
        prev.text += run.text;
      } else {
        consolidatedRuns.push({ ...run });
      }
    }
  }

  // Convert to ExcelJS.RichText
  const richText: ExcelJS.RichText[] = consolidatedRuns.map((r) => {
    const font: Partial<ExcelJS.Font> = {
      name: r.style.fontFamily || baseFont.name || "Arial",
      size: r.style.size !== undefined ? r.style.size : (baseFont.size || 8.5),
      bold: r.style.bold !== undefined ? r.style.bold : (baseFont.bold || false),
      italic: r.style.italic !== undefined ? r.style.italic : (baseFont.italic || false),
      underline: r.style.underline !== undefined ? (r.style.underline === "double" ? "double" : !!r.style.underline) : (baseFont.underline || false),
      strike: r.style.strike !== undefined ? r.style.strike : (baseFont.strike || false),
      color: r.style.color ? { argb: r.style.color } : (baseFont.color || { argb: "FF000000" }),
    };

    return {
      text: r.text,
      font,
    };
  });

  const plainText = consolidatedRuns.map((r) => r.text).join("");

  return {
    richText,
    plainText,
    hasFormatting: hasSpecialFormatting,
    highlightColor: foundHighlight,
  };
}

const getBase64Image = async (url: string): Promise<{ base64: string; ext: string } | null> => {
  try {
    const res = await fetch(url, { referrerPolicy: "no-referrer" });
    if (!res.ok) return null;
    const blob = await res.blob();
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        const base64 = reader.result as string;
        const matches = base64.match(/^data:image\/([a-zA-Z+]+);base64,(.+)$/);
        if (matches && matches.length === 3) {
          resolve({ ext: matches[1], base64: matches[2] });
        } else {
          resolve(null);
        }
      };
      reader.onerror = () => resolve(null);
      reader.readAsDataURL(blob);
    });
  } catch (e) {
    console.warn("Could not load image for Excel:", url, e);
    return null;
  }
};

/**
 * Calculates the optimal Description column width dynamically according to the longest sentence across the rows.
 * Sizing accommodates uppercase characters and cell padding so sentences fit on 1 line or wrap cleanly without clipping.
 */
export const getOptimalDescColWidth = (
  rows: QuotationRow[],
  isChallan: boolean
): number => {
  let maxChars = 0;
  for (const r of rows) {
    if (!r.desc) continue;
    const plain = htmlToPlainText(r.desc);
    const lines = plain.split(/\r?\n/);
    for (const l of lines) {
      const len = l.trim().length;
      if (len > maxChars) maxChars = len;
    }
  }

  // If no items or very short descriptions
  if (maxChars <= 20) {
    return isChallan ? 56 : 50;
  }

  // Sizing according to sentence length:
  // Uppercase characters take ~1.0 column width unit in Arial, plus ~3 units for cell padding.
  const neededWidth = Math.ceil(maxChars * 1.05 + 3);

  if (isChallan) {
    // Challan has 4 columns (SL, Desc, Qty, Unit), allowing Description to comfortably range from 56 to 72
    return Math.min(72, Math.max(56, neededWidth));
  } else {
    // Quotation/Invoice has 6 columns, allowing Description to comfortably range from 50 to 65
    return Math.min(65, Math.max(50, neededWidth));
  }
};

/**
 * Calculates visual text lines in an item's cell based on column width, line breaks, and font size.
 * Accurately models word-wrapping so no lines of text are clipped, while preventing false line wrapping.
 */
export const calculateItemVisualLines = (
  text: string,
  colWidthOrIsChallan: number | boolean = 52,
  fontSize: number = 8.5
): number => {
  if (!text) return 1;
  const plain = htmlToPlainText(text).trim();
  if (!plain) return 1;

  const colWidth = typeof colWidthOrIsChallan === "boolean"
    ? (colWidthOrIsChallan ? 60 : 52)
    : colWidthOrIsChallan;

  // In Excel with Arial 8.5pt font, accounting for uppercase letters and cell padding (~2.5 units):
  const scale = fontSize > 0 ? 8.5 / fontSize : 1;
  const maxChars = Math.max(12, Math.floor((colWidth - 2.5) * 0.94 * scale));

  const paragraphs = plain.split(/\r?\n/);
  let totalLines = 0;

  for (const para of paragraphs) {
    const trimmed = para.trim();
    if (!trimmed) {
      totalLines += 1;
      continue;
    }

    // Wrap words within maxChars
    const words = trimmed.split(/\s+/);
    let currentLineLen = 0;
    let paraLines = 1;

    for (const w of words) {
      if (w.length > maxChars) {
        // Break long unbreakable word across lines
        if (currentLineLen > 0) {
          paraLines += 1;
          currentLineLen = 0;
        }
        const brokenLines = Math.ceil(w.length / maxChars);
        paraLines += brokenLines - 1;
        currentLineLen = w.length % maxChars;
        if (currentLineLen === 0 && brokenLines > 1) currentLineLen = maxChars;
      } else if (currentLineLen === 0) {
        currentLineLen = w.length;
      } else if (currentLineLen + 1 + w.length <= maxChars) {
        currentLineLen += 1 + w.length;
      } else {
        paraLines += 1;
        currentLineLen = w.length;
      }
    }

    totalLines += paraLines;
  }

  return Math.max(1, totalLines);
};

/**
 * Calculates dynamic row height in points for item rows.
 * Provides comfortable vertical centering clearance (at least 19.5pt) so upper ascenders
 * and lower descenders of text never get cut off or hidden behind cell borders.
 */
export const getItemRowHeight = (visualLines: number, fontSize: number = 8.5): number => {
  const lineRate = Math.max(12, fontSize * 1.35 + 1.0);
  if (visualLines <= 1) {
    // Sized precisely according to sentence size so maximum items fit per page cleanly and clearly
    return Math.max(16.5, Math.round(fontSize * 1.35 + 4.0));
  }
  // Multi-line: calculates exact height for all wrapped lines to fit sentences cleanly
  return Math.max(16.5, Math.round(visualLines * lineRate + 4.0));
};

export interface ExcelPageChunk {
  rows: QuotationRow[];
  startSlIndex: number;
  isLastPage: boolean;
}

/**
 * Dynamically paginates items to maximize space utilization on each A4 page.
 * Uses exact vertical height budgets based on item text.
 * When the page reaches its maximum vertical capacity, remaining items are transferred to the next page.
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
  const isInvoice = docType === "invoice";
  const descColWidth = getOptimalDescColWidth(allRows, isChallan);
  
  // Maximum usable vertical height budget for items on A4 page (in points)
  // Sized to allow high density while keeping each cell crisp and unclipped
  const REGULAR_PAGE_BUDGET = isChallan ? 520 : 500;
  const LAST_PAGE_BUDGET = isChallan ? 520 : isInvoice ? 425 : 475;

  const chunks: ExcelPageChunk[] = [];
  let currentChunk: QuotationRow[] = [];
  let currentHeight = 0;
  let currentSl = 1;

  for (let i = 0; i < allRows.length; i++) {
    const row = allRows[i];
    const descFmt = cellFormats ? cellFormats[`${i}_0`] : undefined;
    const fontSize = descFmt?.fontSize || 8.5;
    const lines = calculateItemVisualLines(row.desc, descColWidth, fontSize);
    const rowHeight = getItemRowHeight(lines, fontSize);

    const isPotentialLastItem = i === allRows.length - 1;
    const budgetForCurrentPage = isPotentialLastItem ? LAST_PAGE_BUDGET : REGULAR_PAGE_BUDGET;

    // Check if adding this item exceeds the maximum A4 vertical budget
    if (currentChunk.length > 0 && currentHeight + rowHeight > budgetForCurrentPage) {
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
    // If the last page exceeds the LAST_PAGE_BUDGET (which includes totals & signatures),
    // cleanly split the overflow items to the next page
    if (!isChallan && currentHeight > LAST_PAGE_BUDGET && currentChunk.length > 1) {
      let splitIdx = currentChunk.length - 1;
      let remHeight = currentHeight;
      while (splitIdx > 0 && remHeight > LAST_PAGE_BUDGET) {
        const itemIdx = currentSl - 1 + splitIdx;
        const fmt = cellFormats ? cellFormats[`${itemIdx}_0`] : undefined;
        const fs = fmt?.fontSize || 8.5;
        const lines = calculateItemVisualLines(currentChunk[splitIdx].desc, descColWidth, fs);
        remHeight -= getItemRowHeight(lines, fs);
        splitIdx--;
      }

      const page1Rows = currentChunk.slice(0, splitIdx + 1);
      const page2Rows = currentChunk.slice(splitIdx + 1);

      if (page1Rows.length > 0) {
        chunks.push({
          rows: page1Rows,
          startSlIndex: currentSl,
          isLastPage: false,
        });
        currentSl += page1Rows.length;
      }
      if (page2Rows.length > 0) {
        chunks.push({
          rows: page2Rows,
          startSlIndex: currentSl,
          isLastPage: true,
        });
      }
    } else {
      chunks.push({
        rows: currentChunk,
        startSlIndex: currentSl,
        isLastPage: true,
      });
    }
  }

  chunks.forEach((chunk, idx) => {
    chunk.isLastPage = idx === chunks.length - 1;
  });

  return chunks;
};

/**
 * Builds a compact, space-maximized worksheet matching the exact print format.
 * - Leaves the starting 11 cells/rows at the top instead of the business header for pre-printed letterhead.
 * - Cell A1 contains the page format (e.g. "QUOTATION", "DELIVERY CHALLAN", "BILL / INVOICE").
 * - Everything is aligned exactly with the print format.
 * - Cell sizes increase dynamically according to text content.
 * - Formatting such as highlight, bold, color, fonts, and borders are fully applied.
 */
const buildDocumentWorksheet = (
  workbook: ExcelJS.Workbook,
  worksheet: ExcelJS.Worksheet,
  docType: "quotation" | "challan" | "invoice",
  messers: string,
  address: string,
  challanNo: string,
  dateVal: string,
  requisitionNo: string,
  pageRows: QuotationRow[],
  mergedRegions: MergedRegion[],
  invoiceNo?: string,
  poNumber?: string,
  vatPercent: number = 0,
  transportationFee: number = 0,
  logoId: number | null = null,
  stampId: number | null = null,
  startSlIndex: number = 1,
  pageIndex: number = 1,
  totalPages: number = 1,
  isLastPage: boolean = true,
  allDocumentRows: QuotationRow[] = [],
  cellFormats?: CellFormatMap
) => {
  // Page Setup: Fit to 1 Page Wide and 1 Page Tall on standard A4 portrait
  worksheet.pageSetup = {
    paperSize: 9, // A4
    orientation: "portrait",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    margins: {
      left: 0.25,
      right: 0.25,
      top: 0.25,
      bottom: 0.25,
      header: 0.1,
      footer: 0.1,
    },
    horizontalCentered: true,
    showGridLines: false, // Clean sharp borders
  };

  (worksheet.properties as any).pageSetUpPr = { fitToPage: true };

  const isChallan = docType === "challan";
  const isInvoice = docType === "invoice";
  const totalCols = isChallan ? 4 : 6;
  const lastColLetter = isChallan ? "D" : "F";

  const allRowsToMeasure = allDocumentRows.length > 0 ? allDocumentRows : pageRows;
  const descColWidth = getOptimalDescColWidth(allRowsToMeasure, isChallan);

  // Set compact, well-proportioned column widths matching print document layout and sentence lengths
  if (isChallan) {
    worksheet.columns = [
      { key: "A", width: 5.5 },          // SL
      { key: "B", width: descColWidth }, // Description - dynamically sized to sentence length!
      { key: "C", width: 12.0 },         // Qty
      { key: "D", width: 14.0 },         // Unit
    ];
  } else {
    worksheet.columns = [
      { key: "A", width: 5.5 },          // SL
      { key: "B", width: descColWidth }, // Description - dynamically sized to sentence length!
      { key: "C", width: 8.0 },          // Qty
      { key: "D", width: 8.0 },          // Unit
      { key: "E", width: 11.5 },         // Price
      { key: "F", width: 13.5 },         // Amount
    ];
  }

  // Standard Excel grid view
  worksheet.views = [{ showGridLines: true }];

  // =========================================================================
  // 1. FORMAT TITLE IN ROW 1 (FIRST CELL A1), WITH 10 BLANK ROWS GAP (ROWS 2-11)
  // Replaces the business header to allow pre-printed letterhead pads.
  // Starting 11 cells/rows: Row 1 has the page format, Rows 2 to 11 are left blank.
  // =========================================================================
  worksheet.getRow(1).height = 18;
  worksheet.mergeCells(`A1:${lastColLetter}1`);
  const titleCell = worksheet.getCell("A1");
  const baseTitle = isChallan
    ? "DELIVERY CHALLAN"
    : isInvoice
    ? "BILL / INVOICE"
    : "QUOTATION";

  const pageFormatTitle = totalPages > 1
    ? `${baseTitle} (PAGE ${pageIndex} OF ${totalPages})`
    : baseTitle;

  titleCell.value = pageFormatTitle;
  titleCell.font = { name: "Arial", size: 11.0, bold: true, color: { argb: "000000" } };
  titleCell.alignment = { vertical: "middle", horizontal: "center" };

  // 10 Blank Rows Gap (Rows 2 to 11) for pre-printed letterhead business header:
  // Normal standard cell size (15pt each) and not compressed
  for (let r = 2; r <= 11; r++) {
    worksheet.getRow(r).height = 15;
  }

  // =========================================================================
  // 2. METADATA BOXES (ROWS 12 TO 16) - EXACT PRINT FORMAT ALIGNMENT
  // =========================================================================
  
  // Dynamic height calculation for Messers and Address so cell size expands to fit content
  const messersLines = calculateItemVisualLines(messers || "", isChallan ? 40 : 35, 8.5);
  worksheet.getRow(12).height = 13.5;
  worksheet.getRow(13).height = Math.max(15.5, messersLines * 13.5 + 2);
  worksheet.getRow(14).height = 13.5;

  const addrLines = calculateItemVisualLines(address || "", isChallan ? 40 : 35, 8.0);
  const neededAddrHeight = Math.max(26, addrLines * 13 + 2);
  worksheet.getRow(15).height = Math.ceil(neededAddrHeight / 2);
  worksheet.getRow(16).height = Math.floor(neededAddrHeight / 2);

  // Left Box: Messers & Address
  worksheet.mergeCells("A12:B12");
  const messersLbl = worksheet.getCell("A12");
  messersLbl.value = "MESSERS:";
  messersLbl.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
  messersLbl.alignment = { vertical: "middle", horizontal: "left" };
  messersLbl.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFF8FAFC" },
  };

  worksheet.mergeCells("A13:B13");
  const messersCell = worksheet.getCell("A13");
  const messersParsed = parseHtmlToExcelRuns(messers || "", {
    name: "Arial",
    size: 8.5,
    bold: true,
    color: { argb: "FF000000" },
  });
  if (messersParsed.hasFormatting && messersParsed.richText.length > 0) {
    messersCell.value = { richText: messersParsed.richText };
  } else {
    messersCell.value = messersParsed.plainText;
    messersCell.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "FF000000" } };
  }
  if (messersParsed.highlightColor) {
    messersCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: messersParsed.highlightColor },
    };
  }
  messersCell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
  messersCell.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

  worksheet.mergeCells("A14:B14");
  const addrLbl = worksheet.getCell("A14");
  addrLbl.value = "ADDRESS:";
  addrLbl.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
  addrLbl.alignment = { vertical: "middle", horizontal: "left" };
  addrLbl.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFF8FAFC" },
  };

  worksheet.mergeCells("A15:B16");
  const addrCell = worksheet.getCell("A15");
  const addrParsed = parseHtmlToExcelRuns(address || "", {
    name: "Arial",
    size: 8.0,
    color: { argb: "FF000000" },
  });
  if (addrParsed.hasFormatting && addrParsed.richText.length > 0) {
    addrCell.value = { richText: addrParsed.richText };
  } else {
    addrCell.value = addrParsed.plainText;
    addrCell.font = { name: "Arial", size: 8.0, color: { argb: "FF000000" } };
  }
  if (addrParsed.highlightColor) {
    addrCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: addrParsed.highlightColor },
    };
  }
  addrCell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
  addrCell.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

  // Left Box outer borders (Rows 12 to 16)
  for (let r = 12; r <= 16; r++) {
    for (let c = 1; c <= 2; c++) {
      const cell = worksheet.getCell(r, c);
      const cellBorders: any = { ...cell.border };
      if (r === 12) cellBorders.top = { style: "thin", color: { argb: "000000" } };
      if (r === 16) cellBorders.bottom = { style: "thin", color: { argb: "000000" } };
      if (c === 1) cellBorders.left = { style: "thin", color: { argb: "000000" } };
      if (c === 2) cellBorders.right = { style: "thin", color: { argb: "000000" } };
      cell.border = cellBorders;
    }
  }

  // Right Box (Rows 12 to 16) - DATE is strictly placed at the TOP
  if (isInvoice) {
    // Row 12 (TOP): DATE
    worksheet.mergeCells("C12:D12");
    const dLbl = worksheet.getCell("C12");
    dLbl.value = "DATE:";
    dLbl.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    dLbl.alignment = { vertical: "middle", horizontal: "left" };
    dLbl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };

    worksheet.mergeCells("E12:F12");
    const dVal = worksheet.getCell("E12");
    dVal.value = dateVal || "";
    dVal.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
    dVal.alignment = { vertical: "middle", horizontal: "left" };
    dVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 13: INVOICE NO
    worksheet.mergeCells("C13:D13");
    const iLbl = worksheet.getCell("C13");
    iLbl.value = "INVOICE NO.:";
    iLbl.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    iLbl.alignment = { vertical: "middle", horizontal: "left" };
    iLbl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };

    worksheet.mergeCells("E13:F13");
    const iVal = worksheet.getCell("E13");
    iVal.value = invoiceNo || "";
    iVal.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
    iVal.alignment = { vertical: "middle", horizontal: "left" };
    iVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 14: CHALLAN NO
    worksheet.mergeCells("C14:D14");
    const cLbl = worksheet.getCell("C14");
    cLbl.value = "CHALLAN NO.:";
    cLbl.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    cLbl.alignment = { vertical: "middle", horizontal: "left" };
    cLbl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };

    worksheet.mergeCells("E14:F14");
    const cVal = worksheet.getCell("E14");
    cVal.value = challanNo || "";
    cVal.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
    cVal.alignment = { vertical: "middle", horizontal: "left" };
    cVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 15: REQUISITION NO
    worksheet.mergeCells("C15:D15");
    const rLbl = worksheet.getCell("C15");
    rLbl.value = "REQUISITION NO.:";
    rLbl.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    rLbl.alignment = { vertical: "middle", horizontal: "left" };
    rLbl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };

    worksheet.mergeCells("E15:F15");
    const rVal = worksheet.getCell("E15");
    rVal.value = requisitionNo || "";
    rVal.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
    rVal.alignment = { vertical: "middle", horizontal: "left" };
    rVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 16: P.O. NUMBER
    worksheet.mergeCells("C16:D16");
    const pLbl = worksheet.getCell("C16");
    pLbl.value = "P.O. NUMBER:";
    pLbl.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    pLbl.alignment = { vertical: "middle", horizontal: "left" };
    pLbl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };

    worksheet.mergeCells("E16:F16");
    const pVal = worksheet.getCell("E16");
    pVal.value = poNumber || "";
    pVal.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
    pVal.alignment = { vertical: "middle", horizontal: "left" };
    pVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };
  } else if (isChallan) {
    // Row 12 (TOP): DATE
    const dLbl = worksheet.getCell("C12");
    dLbl.value = "DATE:";
    dLbl.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    dLbl.alignment = { vertical: "middle", horizontal: "left" };
    dLbl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };

    const dVal = worksheet.getCell("D12");
    dVal.value = dateVal || "";
    dVal.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
    dVal.alignment = { vertical: "middle", horizontal: "left" };
    dVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 13: CHALLAN NO
    const cLbl = worksheet.getCell("C13");
    cLbl.value = "CHALLAN NO.:";
    cLbl.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    cLbl.alignment = { vertical: "middle", horizontal: "left" };
    cLbl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };

    const cVal = worksheet.getCell("D13");
    cVal.value = challanNo || "";
    cVal.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
    cVal.alignment = { vertical: "middle", horizontal: "left" };
    cVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 14: REQUISITION NO
    const rLbl = worksheet.getCell("C14");
    rLbl.value = "REQUISITION NO.:";
    rLbl.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    rLbl.alignment = { vertical: "middle", horizontal: "left" };
    rLbl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };

    const rVal = worksheet.getCell("D14");
    rVal.value = requisitionNo || "";
    rVal.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
    rVal.alignment = { vertical: "middle", horizontal: "left" };
    rVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Rows 15, 16: blank dotted lines
    for (let r = 15; r <= 16; r++) {
      worksheet.getCell(`D${r}`).border = { bottom: { style: "dotted", color: { argb: "64748B" } } };
    }
  } else {
    // Quotation - Row 12 (TOP): DATE
    worksheet.mergeCells("C12:D12");
    const dLbl = worksheet.getCell("C12");
    dLbl.value = "DATE:";
    dLbl.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    dLbl.alignment = { vertical: "middle", horizontal: "left" };
    dLbl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };

    worksheet.mergeCells("E12:F12");
    const dVal = worksheet.getCell("E12");
    dVal.value = dateVal || "";
    dVal.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
    dVal.alignment = { vertical: "middle", horizontal: "left" };
    dVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 13: REQUISITION NO
    worksheet.mergeCells("C13:D13");
    const rLbl = worksheet.getCell("C13");
    rLbl.value = "REQUISITION NO.:";
    rLbl.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    rLbl.alignment = { vertical: "middle", horizontal: "left" };
    rLbl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };

    worksheet.mergeCells("E13:F13");
    const rVal = worksheet.getCell("E13");
    rVal.value = requisitionNo || "";
    rVal.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
    rVal.alignment = { vertical: "middle", horizontal: "left" };
    rVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    for (let r = 14; r <= 16; r++) {
      worksheet.mergeCells(`E${r}:F${r}`);
      worksheet.getCell(`E${r}`).border = { bottom: { style: "dotted", color: { argb: "64748B" } } };
    }
  }

  // Right Box outer borders (Rows 12 to 16)
  const rStartCol = 3;
  for (let r = 12; r <= 16; r++) {
    for (let c = rStartCol; c <= totalCols; c++) {
      const cell = worksheet.getCell(r, c);
      const cellBorders: any = { ...cell.border };
      if (r === 12) cellBorders.top = { style: "thin", color: { argb: "000000" } };
      if (r === 16) cellBorders.bottom = { style: "thin", color: { argb: "000000" } };
      if (c === rStartCol) cellBorders.left = { style: "thin", color: { argb: "000000" } };
      if (c === totalCols) cellBorders.right = { style: "thin", color: { argb: "000000" } };
      cell.border = cellBorders;
    }
  }

  // =========================================================================
  // 3. MAIN TABLE HEADER (ROW 17) - EXACT PRINT FORMAT PARITY
  // =========================================================================
  const headerRow = worksheet.getRow(17);
  headerRow.height = 18;

  const colHeaders = isChallan
    ? ["SL", "Description of Marine Items / Spare Parts", "Qty", "Unit"]
    : ["SL", "Description of Marine Items / Spare Parts", "Qty", "Unit", "Price", "Amount"];

  colHeaders.forEach((text, i) => {
    const colIdx = i + 1;
    const cell = headerRow.getCell(colIdx);
    cell.value = text;
    cell.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
    cell.alignment = {
      vertical: "middle",
      horizontal: colIdx === 2 ? "left" : colIdx === 6 ? "right" : "center",
      wrapText: true,
    };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF1F5F9" },
    };
    cell.border = {
      top: { style: "medium", color: { argb: "000000" } },
      bottom: { style: "medium", color: { argb: "000000" } },
      left: { style: "thin", color: { argb: "000000" } },
      right: { style: "thin", color: { argb: "000000" } },
    };
  });

  // =========================================================================
  // 4. HELPER TO APPLY STYLES, HIGHLIGHTS, COLORS, FONTS, AND BORDERS
  // =========================================================================
  const applyCellFormatToExcel = (
    cell: ExcelJS.Cell,
    rIndex: number,
    cIndex: number,
    defaultAlign: "left" | "center" | "right",
    rawValue?: string,
    options?: {
      isLastRow?: boolean;
      isNumeric?: boolean;
      numericVal?: number;
      formula?: string;
    }
  ) => {
    const key = `${rIndex}_${cIndex}`;
    const fmt = cellFormats ? cellFormats[key] : undefined;

    // 1. Build Base Font (incorporating toolbar settings)
    const baseFont: Partial<ExcelJS.Font> = {
      name: fmt?.fontFamily || cell.font?.name || "Arial",
      size: fmt?.fontSize !== undefined ? fmt.fontSize : (cell.font?.size || 8.5),
      bold: fmt?.bold !== undefined ? fmt.bold : cell.font?.bold,
      italic: fmt?.italic !== undefined ? fmt.italic : cell.font?.italic,
      underline: fmt?.underline ? (fmt.underline === "double" ? "double" : fmt.underline === "single" ? true : false) : cell.font?.underline,
      color: fmt?.color ? { argb: colorToArgb(fmt.color) || "FF000000" } : (cell.font?.color || { argb: "FF000000" }),
    };

    // 2. Parse Value & Inline RichText (word-level formatting)
    let inlineHighlight: string | undefined = undefined;
    if (options?.formula) {
      cell.value = { formula: options.formula } as any;
      cell.font = baseFont;
    } else if (rawValue && (rawValue.includes("<") || rawValue.includes("&"))) {
      const parsed = parseHtmlToExcelRuns(rawValue, baseFont);
      if (parsed.hasFormatting && parsed.richText.length > 0) {
        cell.value = { richText: parsed.richText };
        inlineHighlight = parsed.highlightColor;
      } else {
        if (options?.isNumeric && options.numericVal !== undefined && !isNaN(options.numericVal)) {
          cell.value = options.numericVal;
        } else {
          cell.value = parsed.plainText;
        }
        cell.font = baseFont;
      }
    } else if (options?.isNumeric && options.numericVal !== undefined && !isNaN(options.numericVal)) {
      cell.value = options.numericVal;
      cell.font = baseFont;
    } else if (rawValue !== undefined) {
      cell.value = rawValue;
      cell.font = baseFont;
    } else {
      cell.font = baseFont;
    }

    // 3. Highlight / Background Fill (respects inline word highlight or cell-level background color)
    const fillArgb = fmt?.bgColor && fmt.bgColor !== "transparent"
      ? colorToArgb(fmt.bgColor)
      : inlineHighlight;

    if (fillArgb) {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: fillArgb },
      };
    } else if (fmt?.bgColor === "transparent") {
      delete (cell as any).fill;
    }

    // 4. Alignments (horizontal, vertical, wrapText, indent, orientation)
    // Always vertically center text in the box so no upper or lower part of the sentence gets hidden
    const alignObj: Partial<ExcelJS.Alignment> = {
      vertical: "middle",
      horizontal: fmt?.align || defaultAlign,
      wrapText: true,
    };

    if (fmt?.indent) {
      alignObj.indent = fmt.indent;
    }

    if (fmt?.orientation) {
      if (fmt.orientation === "angle-up") alignObj.textRotation = 45;
      else if (fmt.orientation === "angle-down") alignObj.textRotation = -45;
      else if (fmt.orientation === "vertical") alignObj.textRotation = 255;
      else if (fmt.orientation === "rotate-up") alignObj.textRotation = 90;
      else if (fmt.orientation === "rotate-down") alignObj.textRotation = -90;
    }
    cell.alignment = alignObj;

    // 5. Borders (Respects custom toolbar borders, presets, and standard table grid)
    const isLastRow = !!options?.isLastRow;
    const defaultBottomStyle: ExcelJS.BorderStyle = isLastRow ? "medium" : "thin";

    const topBorder = parseBorderSide(fmt?.borders?.top, "thin");
    const bottomBorder = parseBorderSide(fmt?.borders?.bottom, defaultBottomStyle);
    const leftBorder = parseBorderSide(fmt?.borders?.left, "thin");
    const rightBorder = parseBorderSide(fmt?.borders?.right, "thin");

    cell.border = {
      top: topBorder,
      bottom: bottomBorder,
      left: leftBorder,
      right: rightBorder,
    };
  };

  // =========================================================================
  // 5. TABLE DATA ROWS - CELL SIZES INCREASE DYNAMICALLY ACCORDING TO TEXT
  // =========================================================================
  let currentRowNum = 18;
  const numItemsOnThisPage = pageRows.length;
  const displayRowCount = Math.max(numItemsOnThisPage, 1);

  for (let idx = 0; idx < displayRowCount; idx++) {
    const r = worksheet.getRow(currentRowNum);
    const rowData = pageRows[idx];

    const slVal = startSlIndex + idx;
    const rawDesc = rowData ? rowData.desc : "";
    const rawQty = rowData ? rowData.qty : "";
    const rawUnit = rowData ? rowData.unit : "";
    const rawPrice = rowData ? rowData.price : "";

    const cleanQtyStr = rawQty ? htmlToPlainText(rawQty).trim() : "";
    const qtyVal = cleanQtyStr ? parseNumericInput(cleanQtyStr) : "";
    const isNumericQty = typeof qtyVal === "number" && !isNaN(qtyVal) && qtyVal !== 0;

    const cleanPriceStr = rawPrice ? htmlToPlainText(rawPrice).trim() : "";
    const priceVal = cleanPriceStr ? parseNumericInput(cleanPriceStr) : "";
    const isNumericPrice = typeof priceVal === "number" && !isNaN(priceVal) && priceVal !== 0;

    // Calculate maximum visual lines across all cells in this row
    const itemRowIdx = startSlIndex - 1 + idx;
    const descFmt = cellFormats ? cellFormats[`${itemRowIdx}_0`] : undefined;
    const qtyFmt = cellFormats ? cellFormats[`${itemRowIdx}_1`] : undefined;
    const unitFmt = cellFormats ? cellFormats[`${itemRowIdx}_2`] : undefined;
    const priceFmt = cellFormats ? cellFormats[`${itemRowIdx}_3`] : undefined;

    const descFontSize = descFmt?.fontSize || 8.5;
    const maxFontSize = Math.max(
      descFontSize,
      qtyFmt?.fontSize || 8.5,
      unitFmt?.fontSize || 8.5,
      priceFmt?.fontSize || 8.5
    );

    const descLines = calculateItemVisualLines(rawDesc, descColWidth, descFontSize);
    const qtyLines = calculateItemVisualLines(rawQty, isChallan ? 12.0 : 8.0, qtyFmt?.fontSize || 8.5);
    const unitLines = calculateItemVisualLines(rawUnit, isChallan ? 14.0 : 8.0, unitFmt?.fontSize || 8.5);
    const priceLines = calculateItemVisualLines(rawPrice, 11.5, priceFmt?.fontSize || 8.5);

    const maxVisualLines = Math.max(descLines, qtyLines, unitLines, priceLines, 1);

    // Dynamic row height strictly proportional to sentence lines and font size
    const dynamicRowHeight = getItemRowHeight(maxVisualLines, maxFontSize);
    r.height = dynamicRowHeight;

    const isLastItemRow = idx === displayRowCount - 1;

    // Col 1: SL
    const cellSL = r.getCell(1);
    applyCellFormatToExcel(cellSL, itemRowIdx, -1, "center", rowData ? String(slVal) : "", {
      isLastRow: isLastItemRow,
    });

    // Col 2: Description
    const cellDesc = r.getCell(2);
    applyCellFormatToExcel(cellDesc, itemRowIdx, 0, "left", rawDesc, {
      isLastRow: isLastItemRow,
    });

    // Col 3: Qty
    const cellQty = r.getCell(3);
    applyCellFormatToExcel(cellQty, itemRowIdx, 1, "center", rawQty, {
      isLastRow: isLastItemRow,
      isNumeric: isNumericQty,
      numericVal: typeof qtyVal === "number" ? qtyVal : undefined,
    });
    if (isNumericQty) {
      cellQty.numFmt = "#,##0.00";
    }

    // Col 4: Unit
    const cellUnit = r.getCell(4);
    applyCellFormatToExcel(cellUnit, itemRowIdx, 2, "center", rawUnit, {
      isLastRow: isLastItemRow,
    });

    // Cols 5 & 6: Price & Amount for Quotation / Invoice
    if (!isChallan) {
      const cellPrice = r.getCell(5);
      applyCellFormatToExcel(cellPrice, itemRowIdx, 3, "right", rawPrice, {
        isLastRow: isLastItemRow,
        isNumeric: isNumericPrice,
        numericVal: typeof priceVal === "number" ? priceVal : undefined,
      });
      if (isNumericPrice) {
        cellPrice.numFmt = "#,##0.00";
      }

      const cellAmount = r.getCell(6);
      const hasContent = rowData && (htmlToPlainText(rawDesc).trim() || cleanQtyStr || cleanPriceStr);
      const amountFormula = hasContent
        ? `=IF(OR(C${currentRowNum}="", E${currentRowNum}=""), 0, C${currentRowNum}*E${currentRowNum})`
        : undefined;

      applyCellFormatToExcel(cellAmount, itemRowIdx, 4, "right", undefined, {
        isLastRow: isLastItemRow,
        formula: amountFormula,
      });
      cellAmount.numFmt = "#,##0.00";
    }

    currentRowNum++;
  }

  // Apply Cell Merging to Excel Worksheet for this page
  const pageStartRowIdx = startSlIndex - 1;
  const pageEndRowIdx = pageStartRowIdx + numItemsOnThisPage - 1;

  mergedRegions.forEach((region) => {
    if (region.startRow >= pageStartRowIdx && region.endRow <= pageEndRowIdx) {
      const excelStartRow = (region.startRow - pageStartRowIdx) + 18;
      const excelEndRow = (region.endRow - pageStartRowIdx) + 18;
      const excelStartCol = region.startCol + 2;
      const excelEndCol = region.endCol + 2;

      if (
        excelStartRow >= 18 &&
        excelStartCol >= 1 &&
        excelEndRow < 18 + displayRowCount &&
        excelEndCol <= totalCols
      ) {
        try {
          worksheet.mergeCells(excelStartRow, excelStartCol, excelEndRow, excelEndCol);
          const masterCell = worksheet.getCell(excelStartRow, excelStartCol);
          masterCell.alignment = {
            ...masterCell.alignment,
            vertical: "middle",
            wrapText: true,
          };
        } catch (err) {
          console.warn("Could not merge cells in Excel workbook:", region, err);
        }
      }
    }
  });

  // =========================================================================
  // 6. TOTALS & WORDS BLOCK (Rendered strictly on the LAST page of Invoice/Quotation)
  // =========================================================================
  if (!isChallan && isLastPage) {
    const totalRow = currentRowNum;
    const numTotalRows = isInvoice ? 4 : 1;

    for (let rOffset = 0; rOffset < numTotalRows; rOffset++) {
      worksheet.getRow(totalRow + rOffset).height = 16.5;
    }

    if (isInvoice) {
      worksheet.mergeCells(`A${totalRow}:D${totalRow + 3}`);
    } else {
      worksheet.mergeCells(`A${totalRow}:D${totalRow}`);
    }

    const wordCell = worksheet.getCell(`A${totalRow}`);

    const effectiveRowsForTotal = allDocumentRows.length > 0 ? allDocumentRows : pageRows;
    const subtotalValue = effectiveRowsForTotal.reduce((sum, r) => sum + r.amount, 0);
    const calculatedVat = isInvoice ? (subtotalValue * (vatPercent || 0)) / 100 : 0;
    const finalGrandTotal = isInvoice ? subtotalValue + calculatedVat + (transportationFee || 0) : subtotalValue;

    const words = numberToWords(Math.round(finalGrandTotal));
    const wordsStr = words ? words.toUpperCase() : "ZERO ONLY";
    wordCell.value = `AMOUNT IN WORDS: ${wordsStr}`;

    wordCell.font = { name: "Arial", size: 7.5, bold: true, italic: true, color: { argb: "000000" } };
    wordCell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    wordCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF8FAFC" },
    };

    // If quotation, adjust row height according to length of words sentence
    if (!isInvoice) {
      const mergedWordWidth = 5.5 + descColWidth + 8.0 + 8.0;
      const wordCharsPerLine = Math.floor(mergedWordWidth * 1.15);
      const wordLines = Math.max(1, Math.ceil(wordsStr.length / wordCharsPerLine));
      worksheet.getRow(totalRow).height = wordLines <= 1 ? 17 : Math.max(17, wordLines * 13.5 + 3);
    }

    for (let rOffset = 0; rOffset < numTotalRows; rOffset++) {
      const rNum = totalRow + rOffset;
      for (let c = 1; c <= 4; c++) {
        const cell = worksheet.getCell(rNum, c);
        cell.border = {
          top: rOffset === 0 ? { style: "medium", color: { argb: "000000" } } : undefined,
          bottom: rOffset === numTotalRows - 1 ? { style: "medium", color: { argb: "000000" } } : undefined,
          left: c === 1 ? { style: "medium", color: { argb: "000000" } } : undefined,
          right: c === 4 ? { style: "medium", color: { argb: "000000" } } : undefined,
        };
      }
    }

    const sumRange = `F18:F${totalRow - 1}`;

    if (isInvoice) {
      // Row 1: SUBTOTAL
      const subtotalLbl = worksheet.getCell(`E${totalRow}`);
      subtotalLbl.value = "SUBTOTAL";
      subtotalLbl.font = { name: "Arial", size: 8.0, bold: true };
      subtotalLbl.alignment = { vertical: "middle", horizontal: "right" };
      subtotalLbl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };

      const subtotalValCell = worksheet.getCell(`F${totalRow}`);
      if (totalPages === 1) {
        subtotalValCell.value = { formula: `=SUM(${sumRange})` } as any;
      } else {
        subtotalValCell.value = subtotalValue;
      }
      subtotalValCell.font = { name: "Arial", size: 8.5, bold: true };
      subtotalValCell.alignment = { vertical: "middle", horizontal: "right" };
      subtotalValCell.numFmt = "#,##0.00";

      // Row 2: VAT
      const vatLbl = worksheet.getCell(`E${totalRow + 1}`);
      vatLbl.value = `VAT (${vatPercent || 0}%)`;
      vatLbl.font = { name: "Arial", size: 8.0, bold: true };
      vatLbl.alignment = { vertical: "middle", horizontal: "right" };
      vatLbl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };

      const vatValCell = worksheet.getCell(`F${totalRow + 1}`);
      vatValCell.value = calculatedVat;
      vatValCell.font = { name: "Arial", size: 8.5, bold: true };
      vatValCell.alignment = { vertical: "middle", horizontal: "right" };
      vatValCell.numFmt = "#,##0.00";

      // Row 3: TRANS.
      const transLbl = worksheet.getCell(`E${totalRow + 2}`);
      transLbl.value = "TRANS.";
      transLbl.font = { name: "Arial", size: 8.0, bold: true };
      transLbl.alignment = { vertical: "middle", horizontal: "right" };
      transLbl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };

      const transValCell = worksheet.getCell(`F${totalRow + 2}`);
      transValCell.value = transportationFee || 0;
      transValCell.font = { name: "Arial", size: 8.5, bold: true };
      transValCell.alignment = { vertical: "middle", horizontal: "right" };
      transValCell.numFmt = "#,##0.00";

      // Row 4: GRAND TOTAL (decreased font size by 1)
      const grandLbl = worksheet.getCell(`E${totalRow + 3}`);
      grandLbl.value = "GRAND TOTAL";
      grandLbl.font = { name: "Arial", size: 7.5, bold: true };
      grandLbl.alignment = { vertical: "middle", horizontal: "right" };
      grandLbl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE0E7FF" } }; // Subtle Indigo

      const grandValCell = worksheet.getCell(`F${totalRow + 3}`);
      grandValCell.value = finalGrandTotal;
      grandValCell.font = { name: "Arial", size: 8.5, bold: true };
      grandValCell.alignment = { vertical: "middle", horizontal: "right" };
      grandValCell.numFmt = "#,##0.00";

      for (let rOffset = 0; rOffset < numTotalRows; rOffset++) {
        const rNum = totalRow + rOffset;
        for (let c = 5; c <= 6; c++) {
          const cell = worksheet.getCell(rNum, c);
          cell.border = {
            top: rOffset === 0 ? { style: "medium", color: { argb: "000000" } } : { style: "thin", color: { argb: "000000" } },
            bottom: rOffset === numTotalRows - 1 ? { style: "medium", color: { argb: "000000" } } : { style: "thin", color: { argb: "000000" } },
            left: c === 5 ? { style: "medium", color: { argb: "000000" } } : { style: "thin", color: { argb: "000000" } },
            right: c === 6 ? { style: "medium", color: { argb: "000000" } } : { style: "thin", color: { argb: "000000" } },
          };
        }
      }
    } else {
      // Quotation TOTAL
      const totalLbl = worksheet.getCell(`E${totalRow}`);
      totalLbl.value = totalPages > 1 ? "GRAND TOTAL" : "TOTAL";
      totalLbl.font = { name: "Arial", size: 7.5, bold: true };
      totalLbl.alignment = { vertical: "middle", horizontal: "right" };
      totalLbl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };

      const valCell = worksheet.getCell(`F${totalRow}`);
      if (totalPages === 1) {
        valCell.value = { formula: `=SUM(${sumRange})` } as any;
      } else {
        valCell.value = subtotalValue;
      }
      valCell.font = { name: "Arial", size: 8.0, bold: true };
      valCell.alignment = { vertical: "middle", horizontal: "right" };
      valCell.numFmt = "#,##0.00";

      for (let c = 5; c <= 6; c++) {
        const cell = worksheet.getCell(totalRow, c);
        cell.border = {
          top: { style: "medium", color: { argb: "000000" } },
          bottom: { style: "medium", color: { argb: "000000" } },
          left: c === 5 ? { style: "medium", color: { argb: "000000" } } : { style: "thin", color: { argb: "000000" } },
          right: c === 6 ? { style: "medium", color: { argb: "000000" } } : { style: "thin", color: { argb: "000000" } },
        };
      }
    }

    currentRowNum = totalRow + numTotalRows;
  }

  // =========================================================================
  // 7. SIGNATURES & STAMPS AREA - REPEATED ON EVERY PAGE (MATCHING PRINT TFOOT)
  // For Challan: ONLY Receiver's Signature is rendered.
  // For Quotation/Invoice: Both Receiver's and Authorized Signatures + Stamp are rendered.
  // =========================================================================
  worksheet.getRow(currentRowNum).height = 8;
  currentRowNum++;

  if (!isChallan) {
    // "For Comilla Traders" row on right
    worksheet.getRow(currentRowNum).height = 14;
    worksheet.mergeCells(`E${currentRowNum}:F${currentRowNum}`);
    const authTitle = worksheet.getCell(`E${currentRowNum}`);
    authTitle.value = "For Comilla Traders";
    authTitle.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
    authTitle.alignment = { vertical: "middle", horizontal: "center" };
    currentRowNum++;
  }

  // Room for signatures and stamp
  worksheet.getRow(currentRowNum).height = 36;
  const sigRow = currentRowNum + 1;
  worksheet.getRow(sigRow).height = 16;

  // Receiver's Signature (rendered on every page for all formats)
  worksheet.mergeCells(`A${sigRow}:B${sigRow}`);
  const recSig = worksheet.getCell(`A${sigRow}`);
  recSig.value = "Receiver's Signature";
  recSig.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
  recSig.alignment = { vertical: "middle", horizontal: "center" };
  recSig.border = { top: { style: "thin", color: { argb: "000000" } } };

  // Authorized Signature & Stamp (strictly hidden for Challan, rendered for Quotation & Invoice)
  if (!isChallan) {
    worksheet.mergeCells(`E${sigRow}:F${sigRow}`);
    const authSig = worksheet.getCell(`E${sigRow}`);
    authSig.value = "Authorized Signature";
    authSig.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
    authSig.alignment = { vertical: "middle", horizontal: "center" };
    authSig.border = { top: { style: "thin", color: { argb: "000000" } } };

    // Centered Official Stamp positioned over Authorized Signature
    if (stampId !== null) {
      worksheet.addImage(stampId, {
        tl: { col: 4.64, row: sigRow - 2.15 },
        ext: { width: 84, height: 84 },
      });
    }
  }

  // =========================================================================
  // 8. FOOTER DISCLAIMER NOTICE - MATCHING PRINT TFOOT
  // =========================================================================
  const noticeRow = sigRow + 2;
  worksheet.getRow(sigRow + 1).height = 3;
  worksheet.getRow(noticeRow).height = 13;
  worksheet.mergeCells(`A${noticeRow}:${lastColLetter}${noticeRow}`);
  const noticeCell = worksheet.getCell(`A${noticeRow}`);
  noticeCell.value = "ITEMS ONCE SOLD ARE NON-RETURNABLE AND NON-EXCHANGEABLE.";
  noticeCell.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "000000" } };
  noticeCell.alignment = { vertical: "middle", horizontal: "center" };
};

/**
 * Generates complete multi-sheet Excel Workbook for the currently selected document type.
 * Maximizes A4 space utilization, uses space per text only, compact layout, centered stamp,
 * and automatic page overflow transfer matching the browser print layout exactly.
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
  cellFormats?: CellFormatMap
): Promise<ExcelJS.Workbook> => {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Comilla Traders";
  workbook.lastModifiedBy = "Comilla Traders";
  workbook.created = new Date();
  workbook.modified = new Date();

  // Fetch images in parallel
  const [logoData, stampData] = await Promise.all([
    getBase64Image("https://i.ibb.co.com/gFBkpt8B/Chat-GPT-Image-Apr-23-2026-01-10-13-PM.png"),
    getBase64Image("https://i.ibb.co.com/jZswrtn6/image-4-removebg-preview.png"),
  ]);

  let logoId: number | null = null;
  if (logoData) {
    logoId = workbook.addImage({
      base64: logoData.base64,
      extension: (logoData.ext as any) || "png",
    });
  }

  let stampId: number | null = null;
  if (stampData) {
    stampId = workbook.addImage({
      base64: stampData.base64,
      extension: (stampData.ext as any) || "png",
    });
  }

  let effectiveRows = rows;
  let lastNonEmptyIndex = -1;
  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i];
    if (r.desc.trim() !== "" || r.qty.trim() !== "" || r.price.trim() !== "") {
      lastNonEmptyIndex = i;
      break;
    }
  }
  if (lastNonEmptyIndex >= 0) {
    effectiveRows = rows.slice(0, lastNonEmptyIndex + 1);
  }

  // Dynamic capacity-aware pagination maximizing A4 page utilization
  const pageChunks = paginateRowsForExcel(effectiveRows, docType, cellFormats);
  const totalPages = pageChunks.length;

  const baseSheetNames: Record<string, string> = {
    quotation: "Quotation",
    challan: "Challan",
    invoice: "Invoice",
  };

  pageChunks.forEach((chunk, pageIdx) => {
    const pageNumber = pageIdx + 1;
    const sheetName =
      totalPages === 1
        ? baseSheetNames[docType]
        : `${baseSheetNames[docType]} - Pg ${pageNumber}`;

    const ws = workbook.addWorksheet(sheetName);
    buildDocumentWorksheet(
      workbook,
      ws,
      docType,
      messers,
      address,
      challanNo,
      dateVal,
      requisitionNo,
      chunk.rows,
      mergedRegions,
      invoiceNo,
      poNumber,
      vatPercent || 0,
      transportationFee || 0,
      logoId,
      stampId,
      chunk.startSlIndex,
      pageNumber,
      totalPages,
      chunk.isLastPage,
      effectiveRows,
      cellFormats
    );
  });

  return workbook;
};
