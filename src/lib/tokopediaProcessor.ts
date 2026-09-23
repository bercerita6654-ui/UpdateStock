import * as XLSX from 'xlsx';
import { TokopediaRowMatch, StockListItem } from '../types';
import { MatchingMaps, findStockForSku } from './stockMatcher';
import { cleanSku, parseNumber } from './sheets';
import { downloadBlob } from './excelProcessor';

export interface ParsedTokopediaSheet {
  workbook: XLSX.WorkBook;
  sheetName: string;
  fileName?: string;
  headerRowIndex: number; // 0-based (default 2 for Row 3)
  dataStartRowIndex: number; // 0-based (default 3 for Row 4)
  skuColIndex: number; // 0-based: index 3 for Kolom 4 (D)
  stockColIndex: number; // 0-based: index 8 for Kolom 9 (I)
  headers: string[];
  rows: any[][];
  rawWorkbookData: Uint8Array;
}

/**
 * Tokopedia Mass Stock Update Structure:
 * - SKU Tokopedia: Kolom 4 (Index 3 / Kolom D)
 * - Stok Tokopedia: Kolom 9 (Index 8 / Kolom I)
 * - Data Rows: Mulai Baris 4 (Index 3)
 * - Headers: Baris 3 (Index 2)
 */
export function detectTokopediaStructure(data: any[][]): {
  dataStartRowIndex: number;
  headerRowIndex: number;
  skuColIndex: number;
  stockColIndex: number;
  headers: string[];
} {
  // Standard Tokopedia Mass Update Excel template constants
  const skuColIndex = 3; // Kolom ke-4 (Index 3, Kolom D)
  const stockColIndex = 8; // Kolom ke-9 (Index 8, Kolom I)
  let headerRowIndex = 2; // Baris ke-3 (Index 2)
  let dataStartRowIndex = 3; // Baris ke-4 (Index 3)

  if (data.length <= 3) {
    headerRowIndex = 0;
    dataStartRowIndex = Math.min(1, data.length);
  }

  const rawHeaders = data[headerRowIndex] || data[0] || [];
  const headers = rawHeaders.map((c, i) => {
    const text = String(c || '').trim();
    if (text) return text;
    if (i === 3) return 'SKU Produk (Kolom 4)';
    if (i === 8) return 'Stok Produk (Kolom 9)';
    return `Kolom ${i + 1}`;
  });

  return {
    dataStartRowIndex,
    headerRowIndex,
    skuColIndex,
    stockColIndex,
    headers,
  };
}

export async function parseTokopediaXlsx(file: File): Promise<ParsedTokopediaSheet> {
  const arrayBuffer = await file.arrayBuffer();
  const rawWorkbookData = new Uint8Array(arrayBuffer);
  const workbook = XLSX.read(rawWorkbookData, {
    type: 'array',
    cellStyles: true,
    cellDates: true,
  });

  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error('File Excel Tokopedia tidak memiliki lembar kerja (sheet).');
  }

  const worksheet = workbook.Sheets[firstSheetName];
  const data: any[][] = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: '',
  });

  if (!data || data.length === 0) {
    throw new Error('File Excel Tokopedia kosong atau tidak terbaca.');
  }

  const detected = detectTokopediaStructure(data);

  return {
    workbook,
    sheetName: firstSheetName,
    fileName: file.name,
    headerRowIndex: detected.headerRowIndex,
    dataStartRowIndex: detected.dataStartRowIndex,
    skuColIndex: detected.skuColIndex,
    stockColIndex: detected.stockColIndex,
    headers: detected.headers,
    rows: data,
    rawWorkbookData,
  };
}

export function matchTokopediaFile(
  parsed: ParsedTokopediaSheet,
  maps: MatchingMaps,
  options?: {
    customSkuCol?: number;
    customStockCol?: number;
    customStartRow?: number; // 1-based (e.g. 4)
    unmatchedAction?: 'keep' | 'zero';
  }
): TokopediaRowMatch[] {
  // Data starts at row 4 (index 3) by default
  const startRowIndex = options?.customStartRow !== undefined
    ? Math.max(0, options.customStartRow - 1)
    : (parsed.dataStartRowIndex ?? 3);

  // Column 4 (index 3 / D) for SKU, Column 9 (index 8 / I) for Stock
  const skuCol = options?.customSkuCol !== undefined ? options.customSkuCol : (parsed.skuColIndex ?? 3);
  const stockCol = options?.customStockCol !== undefined ? options.customStockCol : (parsed.stockColIndex ?? 8);
  const unmatchedAction = options?.unmatchedAction ?? 'keep';

  const matches: TokopediaRowMatch[] = [];

  for (let r = startRowIndex; r < parsed.rows.length; r++) {
    const row = parsed.rows[r];
    if (!row || row.length === 0 || row.every((c) => c === '' || c === null || c === undefined)) {
      continue;
    }

    // Tokopedia Column 4 (Index 3, Kolom D)
    const rawSku = row[skuCol] !== undefined && row[skuCol] !== null ? cleanSku(row[skuCol]) : '';
    const originalStock = parseNumber(row[stockCol]);

    // Metadata columns (usually Column 2 / Index 1 is Product Name)
    const productName = (row[1] !== undefined && row[1] !== null && String(row[1]).trim() !== '')
      ? String(row[1]).trim()
      : (row[2] ? String(row[2]).trim() : (row[0] ? String(row[0]).trim() : undefined));
    const variationName = row[4] && skuCol !== 4 ? String(row[4]).trim() : undefined;

    // Match against STOCK LIST database
    const matchResult = findStockForSku(rawSku, maps);

    if (matchResult.stockItem) {
      const newStock = Number(matchResult.stockItem.qty);
      matches.push({
        rowIndex: r,
        displayRow: r + 1,
        skuCol4: rawSku,
        productName: productName || matchResult.stockItem.description,
        variationName,
        originalStock,
        newStock,
        matchStatus: 'matched',
        matchedBy: matchResult.matchedBy as any,
        stockDiff: newStock - originalStock,
        notes: matchResult.notes,
        matchedStockItem: matchResult.stockItem,
        rawRow: row,
      });
    } else {
      const newStock = unmatchedAction === 'zero' ? 0 : null;
      matches.push({
        rowIndex: r,
        displayRow: r + 1,
        skuCol4: rawSku,
        productName,
        variationName,
        originalStock,
        newStock,
        matchStatus: 'unmatched',
        stockDiff: newStock !== null ? newStock - originalStock : 0,
        notes: rawSku ? (matchResult.notes || 'SKU tidak ditemukan di database STOCK LIST') : 'SKU Kosong pada Kolom 4',
        matchedStockItem: null,
        rawRow: row,
      });
    }
  }

  return matches;
}

export function generateUpdatedTokopediaWorkbook(
  parsed: ParsedTokopediaSheet,
  matches: TokopediaRowMatch[],
  stockColIndex: number = 8, // Default Kolom 9 (Index 8, Kolom I)
  options?: {
    unmatchedAction?: 'keep' | 'zero';
  }
): Uint8Array {
  // Re-read workbook to get a pristine clone with all formats, formulas, and styles intact
  const wb = XLSX.read(parsed.rawWorkbookData, {
    type: 'array',
    cellStyles: true,
    cellDates: true,
  });

  const ws = wb.Sheets[parsed.sheetName];
  if (!ws) {
    throw new Error('Sheet target Tokopedia tidak ditemukan pada workbook.');
  }

  // Ensure sheet range includes stock column (Index 8 / Kolom I) and all data rows
  const currentRange = XLSX.utils.decode_range(ws['!ref'] || 'A1:I100');
  if (stockColIndex > currentRange.e.c) {
    currentRange.e.c = stockColIndex;
  }
  for (const m of matches) {
    if (m.rowIndex > currentRange.e.r) {
      currentRange.e.r = m.rowIndex;
    }
  }
  ws['!ref'] = XLSX.utils.encode_range(currentRange);

  // Update Column 9 (Index 8 / Kolom I) with synced stock from STOCK LIST
  for (const m of matches) {
    let finalStockVal: number | null = null;
    if (m.matchStatus === 'matched' && m.newStock !== null) {
      finalStockVal = Number(m.newStock);
    } else if (options?.unmatchedAction === 'zero') {
      finalStockVal = 0;
    }

    if (finalStockVal !== null) {
      const cellAddress = XLSX.utils.encode_cell({ r: m.rowIndex, c: stockColIndex });
      if (ws[cellAddress]) {
        ws[cellAddress].v = finalStockVal;
        ws[cellAddress].t = 'n';
        ws[cellAddress].w = String(finalStockVal);
      } else {
        ws[cellAddress] = {
          t: 'n',
          v: finalStockVal,
          w: String(finalStockVal),
        };
      }
    }
  }

  // Generate output binary buffer
  const out = XLSX.write(wb, {
    bookType: 'xlsx',
    type: 'array',
  });

  return new Uint8Array(out);
}

export function generateTokopediaFilename(
  prefix: string = 'tokopedia',
  date: Date = new Date(),
  format: 'dash' | 'readable' = 'dash'
): string {
  const cleanPrefix = (prefix || 'tokopedia').trim().replace(/\.xlsx$/i, '');
  const day = String(date.getDate()).padStart(2, '0');
  const monthNum = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();

  if (format === 'dash') {
    return `${cleanPrefix}(${day}-${monthNum}-${year}).xlsx`;
  }

  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sept', 'Okt', 'Nov', 'Des'];
  const monthName = months[date.getMonth()];
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  return `${cleanPrefix}(${day} ${monthName} ${year}-${hours}:${minutes}).xlsx`;
}

export function generateFilteredTokopediaWorkbook(
  parsed: ParsedTokopediaSheet,
  matches: TokopediaRowMatch[],
  stockColIndex: number = 8,
  filterOptions?: {
    mode: 'all' | 'category' | 'brand' | 'custom';
    selectedCategories?: string[];
    selectedBrands?: string[];
    selectedRowIndices?: number[];
    unmatchedAction?: 'keep' | 'zero';
  }
): Uint8Array {
  const mode = filterOptions?.mode ?? 'all';
  const unmatchedAction = filterOptions?.unmatchedAction ?? 'keep';

  // If mode is 'all', use the in-place workbook updater to preserve 100% of formatting
  if (mode === 'all') {
    return generateUpdatedTokopediaWorkbook(parsed, matches, stockColIndex, { unmatchedAction });
  }

  // Filter matches based on user selection
  let filteredMatches = matches;
  if (mode === 'category') {
    const set = new Set(filterOptions?.selectedCategories || []);
    filteredMatches = matches.filter((m) => {
      const cat = m.matchedStockItem?.category?.trim() || (m.matchStatus === 'matched' ? '(Tanpa Kategori)' : '(Belum Cocok)');
      return set.has(cat);
    });
  } else if (mode === 'brand') {
    const set = new Set(filterOptions?.selectedBrands || []);
    filteredMatches = matches.filter((m) => {
      const brand = m.matchedStockItem?.brand?.trim() || (m.matchStatus === 'matched' ? '(Tanpa Merk)' : '(Belum Cocok)');
      return set.has(brand);
    });
  } else if (mode === 'custom') {
    const set = new Set(filterOptions?.selectedRowIndices || []);
    filteredMatches = matches.filter((m) => set.has(m.rowIndex));
  }

  // If all rows match or filtered set is empty, fallback
  if (filteredMatches.length === 0) {
    throw new Error('Tidak ada produk yang cocok dengan filter yang dipilih.');
  }

  // Re-read workbook
  const wb = XLSX.read(parsed.rawWorkbookData, {
    type: 'array',
    cellStyles: true,
    cellDates: true,
  });

  const originalWs = wb.Sheets[parsed.sheetName];
  const startRowIndex = parsed.dataStartRowIndex ?? 3;

  // Extract header rows (rows 0 to startRowIndex - 1)
  const headerRows: any[][] = [];
  for (let r = 0; r < startRowIndex; r++) {
    headerRows.push(parsed.rows[r] ? [...parsed.rows[r]] : []);
  }

  // Extract data rows for filtered matches and write updated stock at stockColIndex
  const dataRows: any[][] = filteredMatches.map((m) => {
    const row = [...m.rawRow];
    while (row.length <= stockColIndex) {
      row.push('');
    }

    let finalStockVal: number | null = null;
    if (m.matchStatus === 'matched' && m.newStock !== null) {
      finalStockVal = Number(m.newStock);
    } else if (unmatchedAction === 'zero') {
      finalStockVal = 0;
    }

    if (finalStockVal !== null) {
      row[stockColIndex] = finalStockVal;
    }

    return row;
  });

  // Combine header rows and filtered data rows
  const fullRows = [...headerRows, ...dataRows];
  const ws = XLSX.utils.aoa_to_sheet(fullRows);

  // Preserve column widths if present
  if (originalWs && originalWs['!cols']) {
    ws['!cols'] = originalWs['!cols'];
  } else {
    ws['!cols'] = [
      { wch: 8 },  // No
      { wch: 36 }, // Nama Produk
      { wch: 18 }, // Kode Produk
      { wch: 22 }, // SKU (Kolom 4)
      { wch: 16 }, // Variasi
      { wch: 14 }, // Harga
      { wch: 14 }, // Harga Diskon
      { wch: 14 }, // Status
      { wch: 14 }, // Stok (Kolom 9)
    ];
  }

  const newWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWb, ws, parsed.sheetName || 'Template Ubah Stok');

  const out = XLSX.write(newWb, {
    bookType: 'xlsx',
    type: 'array',
  });

  return new Uint8Array(out);
}

/**
 * Creates a sample Tokopedia mass stock file for user testing
 * Column 4 (D) = SKU, Column 9 (I) = Stock
 * Data rows start on Row 4 (rows 1-3 are sample instructions/headers)
 */
export function createSampleTokopediaFile(): Uint8Array {
  const row1 = ['PETUNJUK PENGISIAN UBAH STOK MASSAL TOKOPEDIA', '', '', '', '', '', '', '', ''];
  const row2 = ['Kolom D (Kolom 4): SKU Produk', '', '', '', '', '', '', '', 'Kolom I (Kolom 9): Stok Produk'];
  const row3 = [
    'No',
    'Nama Produk Tokopedia',
    'Kode Produk Internal',
    'SKU (Kolom 4)', // Col 4 (D)
    'Variasi',
    'Harga Normal',
    'Harga Diskon',
    'Status Produk',
    'Stok (Kolom 9)', // Col 9 (I)
  ];

  const sampleRows = [
    [1, '3M Double Tape Foam Indoor 1/2', 'TKP-001', '19163', 'Standar', 20000, 19000, 'Aktif', 5],
    [2, '3M Double Tape Foam Scotch Indoor 1"', 'TKP-002', '16987', 'Standar', 71000, 68000, 'Aktif', 0],
    [3, '3M Double Tape Scotch Outdoor 19mm', 'TKP-003', '17626', 'Standar', 39000, 37000, 'Aktif', 12],
    [4, 'Acco Fastener Putih V-Tech', 'TKP-004', '00003', 'Putih', 8500, 8000, 'Aktif', 20],
    [5, 'Acrylic Colour 12 Warna Deli', 'TKP-005', '18439', '12ml', 55000, 52000, 'Aktif', 2],
    [6, 'Buku Tulis Sinar Dunia 38 Lembar', 'TKP-006', '00125', 'Standar', 35000, 33000, 'Aktif', 10],
    [7, 'Pulpen Pilot G2 0.5 Hitam', 'TKP-007', '00450', 'Hitam', 180000, 175000, 'Aktif', 15],
    [8, 'Spidol Snowman Whiteboard Hitam', 'TKP-008', '00890', 'Hitam', 96000, 92000, 'Aktif', 8],
    [9, 'Produk Tokopedia Belum Terdaftar Testing', 'TKP-009', 'TKP-NONEXIST-99', 'Custom', 15000, 15000, 'Aktif', 3],
  ];

  const ws = XLSX.utils.aoa_to_sheet([row1, row2, row3, ...sampleRows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Template Ubah Stok');

  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  return new Uint8Array(out);
}
