import * as XLSX from 'xlsx';
import { ShopeeRowMatch } from '../types';
import { MatchingMaps, findStockForSku } from './stockMatcher';
import { cleanSku, parseNumber } from './sheets';

export interface ParsedShopeeSheet {
  workbook: XLSX.WorkBook;
  sheetName: string;
  fileName?: string;
  headerRowIndex: number; // 0-based
  skuColIndex: number;
  skuParentColIndex: number;
  stockColIndex: number;
  headers: string[];
  rows: any[][];
  rawWorkbookData: Uint8Array;
}

// Heuristic to find Shopee header row and columns
export function detectShopeeColumns(data: any[][]): {
  headerRowIndex: number;
  skuColIndex: number;
  skuParentColIndex: number;
  stockColIndex: number;
  headers: string[];
} {
  // Check rows 0 to 5 for known Shopee column keywords
  for (let r = 0; r < Math.min(data.length, 6); r++) {
    const row = data[r] || [];
    const lowerRow = row.map((cell) => String(cell || '').trim().toLowerCase());

    let skuIndex = -1;
    let skuParentIndex = -1;
    let stockIndex = -1;

    for (let c = 0; c < lowerRow.length; c++) {
      const text = lowerRow[c];
      if (!text) continue;

      // Variation SKU (Kode Variasi / SKU)
      if (
        text.includes('kode variasi') ||
        text.includes('variation sku') ||
        text.includes('sku variasi') ||
        text === 'sku' ||
        text.includes('kode sku')
      ) {
        skuIndex = c;
      } else if (
        text.includes('sku induk') ||
        text.includes('parent sku') ||
        text.includes('sku referensi')
      ) {
        skuParentIndex = c;
      }

      // Stock column
      if (
        text === 'stok' ||
        text === 'stock' ||
        text.includes('stok saat ini') ||
        text.includes('current stock') ||
        text.includes('jumlah stok')
      ) {
        stockIndex = c;
      }
    }

    if (skuIndex !== -1 || stockIndex !== -1) {
      return {
        headerRowIndex: r,
        skuColIndex: skuIndex !== -1 ? skuIndex : 0,
        skuParentColIndex: skuParentIndex,
        stockColIndex: stockIndex !== -1 ? stockIndex : (row.length > 1 ? row.length - 1 : 1),
        headers: row.map((c) => String(c || '')),
      };
    }
  }

  // Fallback if not detected automatically
  const fallbackHeaders = (data[0] || []).map((c, i) => String(c || `Kolom ${i + 1}`));
  return {
    headerRowIndex: 0,
    skuColIndex: 0,
    skuParentColIndex: -1,
    stockColIndex: data[0] && data[0].length > 1 ? 1 : 0,
    headers: fallbackHeaders,
  };
}

export async function parseShopeeXlsx(file: File): Promise<ParsedShopeeSheet> {
  const arrayBuffer = await file.arrayBuffer();
  const rawWorkbookData = new Uint8Array(arrayBuffer);
  const workbook = XLSX.read(rawWorkbookData, {
    type: 'array',
    cellStyles: true,
    cellDates: true,
  });

  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error('File Excel tidak memiliki lembar kerja (sheet).');
  }

  const worksheet = workbook.Sheets[firstSheetName];
  const data: any[][] = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: '',
  });

  if (!data || data.length === 0) {
    throw new Error('File Excel kosong atau tidak terbaca.');
  }

  const detected = detectShopeeColumns(data);

  return {
    workbook,
    sheetName: firstSheetName,
    fileName: file.name,
    headerRowIndex: detected.headerRowIndex,
    skuColIndex: detected.skuColIndex,
    skuParentColIndex: detected.skuParentColIndex,
    stockColIndex: detected.stockColIndex,
    headers: detected.headers,
    rows: data,
    rawWorkbookData,
  };
}

export function matchShopeeFile(
  parsed: ParsedShopeeSheet,
  maps: MatchingMaps,
  options?: {
    customSkuCol?: number;
    customStockCol?: number;
    customHeaderRow?: number;
    unmatchedAction?: 'keep' | 'zero';
  }
): ShopeeRowMatch[] {
  const headerRow = options?.customHeaderRow ?? parsed.headerRowIndex;
  const skuCol = options?.customSkuCol ?? parsed.skuColIndex;
  const stockCol = options?.customStockCol ?? parsed.stockColIndex;
  const unmatchedAction = options?.unmatchedAction ?? 'keep';

  const matches: ShopeeRowMatch[] = [];

  for (let r = headerRow + 1; r < parsed.rows.length; r++) {
    const row = parsed.rows[r];
    if (!row || row.length === 0 || row.every((c) => c === '' || c === null || c === undefined)) {
      continue;
    }

    let rawSku = cleanSku(row[skuCol]);
    // If empty variation SKU and there's a parent SKU column, try parent SKU
    if (!rawSku && parsed.skuParentColIndex !== -1 && row[parsed.skuParentColIndex]) {
      rawSku = cleanSku(row[parsed.skuParentColIndex]);
    }

    const originalStock = parseNumber(row[stockCol]);

    // Optional metadata columns (product name, variation name)
    const productName = row[1] ? String(row[1]) : undefined;
    const variationName = row[3] ? String(row[3]) : undefined;

    const matchResult = findStockForSku(rawSku, maps);

    if (matchResult.stockItem) {
      const newStock = matchResult.stockItem.qty;
      matches.push({
        rowIndex: r,
        displayRow: r + 1,
        rawSku,
        cleanedSku: cleanSku(rawSku),
        productName,
        variationName,
        originalStock,
        newStock,
        matchStatus: 'matched',
        matchedBy: matchResult.matchedBy,
        stockDiff: newStock - originalStock,
        notes: matchResult.notes,
      });
    } else {
      const newStock = unmatchedAction === 'zero' ? 0 : null;
      matches.push({
        rowIndex: r,
        displayRow: r + 1,
        rawSku,
        cleanedSku: cleanSku(rawSku),
        productName,
        variationName,
        originalStock,
        newStock,
        matchStatus: 'unmatched',
        stockDiff: newStock !== null ? newStock - originalStock : 0,
        notes: matchResult.notes || 'SKU tidak ditemukan di database stok',
      });
    }
  }

  return matches;
}

export function generateUpdatedShopeeWorkbook(
  parsed: ParsedShopeeSheet,
  matches: ShopeeRowMatch[],
  stockColIndex: number,
  options?: {
    unmatchedAction?: 'keep' | 'zero';
  }
): Uint8Array {
  // Re-read workbook to get fresh copy
  const wb = XLSX.read(parsed.rawWorkbookData, {
    type: 'array',
    cellStyles: true,
    cellDates: true,
  });

  const ws = wb.Sheets[parsed.sheetName];
  if (!ws) {
    throw new Error('Sheet target tidak ditemukan pada workbook.');
  }

  for (const m of matches) {
    let finalStockVal: number | null = null;
    if (m.matchStatus === 'matched' && m.newStock !== null) {
      finalStockVal = m.newStock;
    } else if (options?.unmatchedAction === 'zero') {
      finalStockVal = 0;
    }

    if (finalStockVal !== null) {
      // Cell address: r = m.rowIndex, c = stockColIndex
      const cellAddress = XLSX.utils.encode_cell({ r: m.rowIndex, c: stockColIndex });
      ws[cellAddress] = {
        t: 'n',
        v: finalStockVal,
        w: String(finalStockVal),
      };
    }
  }

  // Generate output buffer
  const out = XLSX.write(wb, {
    bookType: 'xlsx',
    type: 'array',
  });

  return new Uint8Array(out);
}

// Format date-time for export filename like: balist(17 Sept 2026-14:26).xlsx or gomall(17 Sept 2026-14:26).xlsx
export function generateShopeeBalistFilename(prefix: string = 'balist', date: Date = new Date()): string {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sept', 'Okt', 'Nov', 'Des'];
  const day = date.getDate();
  const month = months[date.getMonth()];
  const year = date.getFullYear();
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');

  const cleanPrefix = (prefix || 'balist').trim().replace(/\.xlsx$/i, '');
  return `${cleanPrefix}(${day} ${month} ${year}-${hours}:${minutes}).xlsx`;
}

export interface DetectedStoreInfo {
  storePrefix: 'balist' | 'gomall' | null;
  storeName: string | null;
  storeCode: string | null;
  matchedPattern: string | null;
}

/**
 * Automatically detects store from uploaded XLSX filename:
 * - 31475604 (e.g. mass_update_sales_info_31475604...) -> balist (balistationery)
 * - 56977507 (e.g. mass_update_sales_info_56977507...) -> gomall (Gomall)
 * - balistationery / balist -> balist
 * - gomall -> gomall
 */
export function detectStoreFromFilename(fileName: string): DetectedStoreInfo {
  if (!fileName) {
    return { storePrefix: null, storeName: null, storeCode: null, matchedPattern: null };
  }
  const cleanName = fileName.toLowerCase();

  if (cleanName.includes('31475604')) {
    return {
      storePrefix: 'balist',
      storeName: 'Balistationery',
      storeCode: '31475604',
      matchedPattern: 'mass_update_sales_info_31475604 (balistationery)',
    };
  }

  if (cleanName.includes('56977507')) {
    return {
      storePrefix: 'gomall',
      storeName: 'Gomall',
      storeCode: '56977507',
      matchedPattern: 'mass_update_sales_info_56977507 (Gomall)',
    };
  }

  if (cleanName.includes('balistationery')) {
    return {
      storePrefix: 'balist',
      storeName: 'Balistationery',
      storeCode: '31475604',
      matchedPattern: 'balistationery',
    };
  }

  if (cleanName.includes('gomall')) {
    return {
      storePrefix: 'gomall',
      storeName: 'Gomall',
      storeCode: '56977507',
      matchedPattern: 'gomall',
    };
  }

  if (cleanName.includes('balist')) {
    return {
      storePrefix: 'balist',
      storeName: 'Balistationery',
      storeCode: '31475604',
      matchedPattern: 'balist',
    };
  }

  return {
    storePrefix: null,
    storeName: null,
    storeCode: null,
    matchedPattern: null,
  };
}

// Function to trigger file download in browser
export function downloadBlob(data: Uint8Array, fileName: string) {
  const blob = new Blob([data], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 100);
}

// Generate a sample template for user quick testing if needed
export function createSampleShopeeFile(): Uint8Array {
  const headers = [
    'No',
    'Nama Produk',
    'No. Variasi',
    'Nama Variasi',
    'Kode Variasi (SKU)',
    'Stok Saat Ini',
    'Harga',
  ];

  const sampleRows = [
    [1, '3M Double Tape Foam Indoor 1/2', '1001', 'Default', '19163', 5, 20000],
    [2, '3M Double Tape Foam Scotch Indoor 1"', '1002', 'Default', '16987', 0, 71000],
    [3, '3M Double Tape Scotch Outdoor 19mm', '1003', 'Default', '17626', 12, 39000],
    [4, 'Acco Fastener Putih V-Tech', '1004', 'Putih', '00003', 20, 8500],
    [5, 'Acrylic Colour 12 Warna Deli', '1005', 'Set 12ml', '18439', 2, 55000],
    [6, 'Produk Non-Eksis Testing', '1006', 'Varian A', 'NONEXIST999', 10, 15000],
  ];

  const ws = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Shopee Stock Update');

  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  return new Uint8Array(out);
}

export interface ParsedGenericXlsx {
  workbook: XLSX.WorkBook;
  sheetNames: string[];
  selectedSheetName: string;
  rows: any[][];
  headers: string[];
  totalRows: number;
  rawWorkbookData: Uint8Array;
}

export async function parseGenericXlsx(
  file: File,
  targetSheetName?: string
): Promise<ParsedGenericXlsx> {
  const arrayBuffer = await file.arrayBuffer();
  const rawWorkbookData = new Uint8Array(arrayBuffer);
  const workbook = XLSX.read(rawWorkbookData, {
    type: 'array',
    cellStyles: true,
    cellDates: true,
  });

  const sheetNames = workbook.SheetNames;
  if (!sheetNames || sheetNames.length === 0) {
    throw new Error('File Excel tidak memiliki lembar kerja (sheet).');
  }

  // Look for target sheet name case-insensitively, or default to first
  let selectedSheetName = sheetNames[0];
  if (targetSheetName) {
    const found = sheetNames.find(
      (s) => s.trim().toLowerCase() === targetSheetName.trim().toLowerCase()
    );
    if (found) {
      selectedSheetName = found;
    }
  }

  const worksheet = workbook.Sheets[selectedSheetName];
  const rows: any[][] = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: '',
  });

  const headers = rows.length > 0 ? rows[0].map((c) => String(c ?? '')) : [];

  return {
    workbook,
    sheetNames,
    selectedSheetName,
    rows,
    headers,
    totalRows: rows.length,
    rawWorkbookData,
  };
}
