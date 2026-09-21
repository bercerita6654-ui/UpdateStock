import { StockListItem, BalistShopeeItem, ShopeeRowMatch, ProcessSummary } from '../types';
import { cleanSku, parseNumber } from './sheets';

export interface MatchingMaps {
  stockListByCode: Map<string, StockListItem>;
  stockListByNormalizedCode: Map<string, StockListItem>;
  stockListByBarcode: Map<string, StockListItem>;
  stockListBy5Digits: Map<string, StockListItem>;
  stockListByPrefix5: Map<string, StockListItem>;
  fiveDigitStockEntries: { code5: string; item: StockListItem }[];
  balistByCol5: Map<string, BalistShopeeItem>;
  balistByCol6: Map<string, BalistShopeeItem>;
}

// Normalizes an SKU for exact matching: uppercase, strip quotes, trim
export const normalizeSkuKey = (sku: string): string => {
  const cleaned = cleanSku(sku).toUpperCase();
  return cleaned.replace(/[^A-Z0-9]/g, '');
};

export const stripLeadingZeros = (sku: string): string => {
  const cleaned = cleanSku(sku);
  return cleaned.replace(/^0+/, '');
};

// Format numeric code to exact 5 digits (e.g. "7945" -> "07945", "10510" -> "10510")
export const format5DigitCode = (code: string): string => {
  const cleaned = cleanSku(code);
  if (/^\d{1,5}$/.test(cleaned)) {
    return cleaned.padStart(5, '0');
  }
  return cleaned.toUpperCase();
};

/**
 * Extract exact 5-digit numerical or alphanumeric SKU tokens from a string.
 * Accurately extracts patterns like "10510", "07945", "10510-HITAM", "10510_01", "[10510] Kaos".
 */
export const extract5DigitTokens = (sku: string): string[] => {
  const cleaned = cleanSku(sku);
  if (!cleaned) return [];
  const tokens = new Set<string>();

  // 1. If entire string is 1 to 5 digits numeric
  if (/^\d{1,5}$/.test(cleaned)) {
    tokens.add(cleaned);
    tokens.add(cleaned.padStart(5, '0'));
    const noZero = cleaned.replace(/^0+/, '');
    if (noZero) tokens.add(noZero);
  } else if (cleaned.length === 5) {
    tokens.add(cleaned.toUpperCase());
  }

  // 2. Exact 5 consecutive digits with boundary or anywhere in string (e.g. "10510" from "10510-01", "ABC_10510_01", "[10510] Daster")
  const all5Digits = cleaned.match(/\d{5}/g);
  if (all5Digits) {
    for (const d of all5Digits) {
      tokens.add(d);
      const noZero = d.replace(/^0+/, '');
      if (noZero) tokens.add(noZero);
    }
  }

  // 3. Delimited parts split by space, dash, underscore, slash, dot, bracket, pipe, comma, etc.
  const parts = cleaned.split(/[\s\-_/.,|#()[\]{}]+/);
  for (const part of parts) {
    const pTrim = cleanSku(part);
    if (!pTrim) continue;
    if (/^\d{1,5}$/.test(pTrim)) {
      tokens.add(pTrim);
      tokens.add(pTrim.padStart(5, '0'));
      const noZero = pTrim.replace(/^0+/, '');
      if (noZero) tokens.add(noZero);
    } else if (pTrim.length === 5) {
      tokens.add(pTrim.toUpperCase());
    }
  }

  return Array.from(tokens).filter((t) => t.length > 0);
};

export function buildMatchingMaps(
  stockList: StockListItem[],
  balistList: BalistShopeeItem[]
): MatchingMaps {
  const stockListByCode = new Map<string, StockListItem>();
  const stockListByNormalizedCode = new Map<string, StockListItem>();
  const stockListByBarcode = new Map<string, StockListItem>();
  const stockListBy5Digits = new Map<string, StockListItem>();
  const stockListByPrefix5 = new Map<string, StockListItem>();
  const fiveDigitStockEntries: { code5: string; item: StockListItem }[] = [];
  const balistByCol5 = new Map<string, BalistShopeeItem>();
  const balistByCol6 = new Map<string, BalistShopeeItem>();

  // Index Stock List
  for (const item of stockList) {
    if (item.code) {
      const codeClean = cleanSku(item.code);
      const codeUpper = codeClean.toUpperCase();
      
      // Direct exact match
      stockListByCode.set(codeUpper, item);
      stockListByCode.set(codeClean, item);

      // 5-digit exact code
      if (/^\d{1,5}$/.test(codeClean)) {
        const code5 = codeClean.padStart(5, '0');
        stockListBy5Digits.set(code5, item);
        stockListBy5Digits.set(codeClean, item);
        stockListByCode.set(code5, item);
        fiveDigitStockEntries.push({ code5, item });
        const noZero = codeClean.replace(/^0+/, '');
        if (noZero) {
          stockListBy5Digits.set(noZero, item);
          stockListByCode.set(noZero, item);
        }
      } else if (codeClean.length === 5) {
        stockListBy5Digits.set(codeUpper, item);
      }

      const norm = normalizeSkuKey(item.code);
      if (norm) {
        stockListByNormalizedCode.set(norm, item);
      }

      const noZero = stripLeadingZeros(item.code);
      if (noZero) {
        stockListByNormalizedCode.set(noZero.toUpperCase(), item);
      }

      // Index 5-digit tokens from Stock List item code
      const tokens = extract5DigitTokens(item.code);
      for (const tok of tokens) {
        if (!stockListBy5Digits.has(tok)) {
          stockListBy5Digits.set(tok, item);
        }
      }
    }

    if (item.barcode) {
      const barcodeClean = cleanSku(item.barcode);
      const barcodeUpper = barcodeClean.toUpperCase();
      stockListByBarcode.set(barcodeUpper, item);
      stockListByBarcode.set(barcodeClean, item);

      // If barcode itself is 1 to 5 digits
      if (/^\d{1,5}$/.test(barcodeClean)) {
        const barcode5 = barcodeClean.padStart(5, '0');
        if (!stockListBy5Digits.has(barcode5)) {
          stockListBy5Digits.set(barcode5, item);
        }
        if (!stockListBy5Digits.has(barcodeClean)) {
          stockListBy5Digits.set(barcodeClean, item);
        }
      }

      const normBarcode = normalizeSkuKey(item.barcode);
      if (normBarcode) {
        stockListByNormalizedCode.set(normBarcode, item);
      }

      const tokens = extract5DigitTokens(item.barcode);
      for (const tok of tokens) {
        if (!stockListBy5Digits.has(tok)) {
          stockListBy5Digits.set(tok, item);
        }
      }
    }

    // Index tokens from item description
    if (item.description) {
      const descTokens = extract5DigitTokens(item.description);
      for (const tok of descTokens) {
        if (!stockListBy5Digits.has(tok)) {
          stockListBy5Digits.set(tok, item);
        }
      }
    }

    // Index tokens from all other row cells if available
    if (item.rawRow && Array.isArray(item.rawRow)) {
      for (let c = 0; c < Math.min(item.rawRow.length, 6); c++) {
        const cellClean = cleanSku(item.rawRow[c]);
        if (cellClean && /^\d{4,5}$/.test(cellClean)) {
          if (!stockListBy5Digits.has(cellClean)) {
            stockListBy5Digits.set(cellClean, item);
          }
          if (!stockListByCode.has(cellClean)) {
            stockListByCode.set(cellClean, item);
          }
        }
      }
    }
  }

  // Index Balistshopee (Col 5 and Col 6)
  for (const item of balistList) {
    if (item.skuCol5) {
      balistByCol5.set(item.skuCol5.toUpperCase(), item);
      const norm = normalizeSkuKey(item.skuCol5);
      if (norm) balistByCol5.set(norm, item);
    }
    if (item.skuCol6) {
      balistByCol6.set(item.skuCol6.toUpperCase(), item);
      const norm = normalizeSkuKey(item.skuCol6);
      if (norm) balistByCol6.set(norm, item);
    }
  }

  return {
    stockListByCode,
    stockListByNormalizedCode,
    stockListByBarcode,
    stockListBy5Digits,
    stockListByPrefix5,
    fiveDigitStockEntries,
    balistByCol5,
    balistByCol6,
  };
}

export interface MatchResult {
  stockItem: StockListItem | null;
  matchedBy?: 'code' | 'barcode' | 'balist_col5' | 'balist_col6' | '5digits_sku';
  notes?: string;
}

const GENERIC_VARIATION_WORDS = new Set([
  'DEFAULT', 'STANDAR', 'STANDARD', 'VARIASI', 'WARNA', 'HITAM', 'PUTIH', 'MERAH',
  'BIRU', 'KUNING', 'HIJAU', 'UNGU', 'PINK', 'COKLAT', 'ABU', 'ORANGE', 'PACK',
  'PCS', 'BOX', 'LUSIN', 'SET', 'SATUAN', 'KECIL', 'SEDANG', 'BESAR', 'S', 'M', 'L', 'XL', 'XXL', 'ALL SIZE', 'ALLSIZE'
]);

export function findStockForSku(
  rawSku: string,
  maps: MatchingMaps
): MatchResult {
  const cleaned = cleanSku(rawSku);
  if (!cleaned) {
    return { stockItem: null, notes: 'SKU Kosong' };
  }

  const upper = cleaned.toUpperCase();

  // 1. Direct exact match in STOCK LIST by Code (Column 1)
  if (maps.stockListByCode.has(upper)) {
    return {
      stockItem: maps.stockListByCode.get(upper)!,
      matchedBy: 'code',
      notes: 'Cocok tepat dengan Kode SKU STOCK LIST (Kolom 1)',
    };
  }

  // 2. Exact 5-digit match (e.g. "07945" or "7945" -> "07945")
  if (/^\d{1,5}$/.test(cleaned)) {
    const code5 = cleaned.padStart(5, '0');
    if (maps.stockListBy5Digits.has(code5)) {
      const matched = maps.stockListBy5Digits.get(code5)!;
      return {
        stockItem: matched,
        matchedBy: '5digits_sku',
        notes: `Cocok tepat SKU 5 Digit (${code5}) dengan Kolom 1 STOCK LIST (${matched.code})`,
      };
    }
    if (maps.stockListByCode.has(code5)) {
      const matched = maps.stockListByCode.get(code5)!;
      return {
        stockItem: matched,
        matchedBy: '5digits_sku',
        notes: `Cocok tepat SKU 5 Digit (${code5}) dengan Kolom 1 STOCK LIST (${matched.code})`,
      };
    }
  }

  // 3. Extract exact 5-digit tokens from SKU string (e.g. "ABC-07945-XL" -> "07945")
  const sku5Tokens = extract5DigitTokens(cleaned);
  for (const token of sku5Tokens) {
    if (maps.stockListBy5Digits.has(token)) {
      const matched = maps.stockListBy5Digits.get(token)!;
      return {
        stockItem: matched,
        matchedBy: '5digits_sku',
        notes: `Cocok tepat 5 digit (${token}) dengan SKU Kolom 1 STOCK LIST (${matched.code})`,
      };
    }
    if (maps.stockListByCode.has(token)) {
      const matched = maps.stockListByCode.get(token)!;
      return {
        stockItem: matched,
        matchedBy: '5digits_sku',
        notes: `Cocok tepat 5 digit (${token}) dengan SKU Kolom 1 STOCK LIST (${matched.code})`,
      };
    }
  }

  // 4. Match by Barcode in STOCK LIST (Column 2)
  if (maps.stockListByBarcode.has(upper)) {
    return {
      stockItem: maps.stockListByBarcode.get(upper)!,
      matchedBy: 'barcode',
      notes: 'Cocok dengan Barcode di STOCK LIST (Kolom 2)',
    };
  }

  // 5. Normalized match for minor formatting (without aggressive truncation)
  const norm = normalizeSkuKey(cleaned);
  if (norm && maps.stockListByNormalizedCode.has(norm)) {
    return {
      stockItem: maps.stockListByNormalizedCode.get(norm)!,
      matchedBy: 'code',
      notes: 'Cocok kode STOCK LIST (normalisasi format)',
    };
  }

  const noZero = stripLeadingZeros(cleaned).toUpperCase();
  if (noZero && maps.stockListByNormalizedCode.has(noZero)) {
    return {
      stockItem: maps.stockListByNormalizedCode.get(noZero)!,
      matchedBy: 'code',
      notes: 'Cocok kode STOCK LIST (tanpa awalan 0)',
    };
  }

  return { stockItem: null, notes: 'SKU tidak ditemukan di database STOCK LIST' };
}

export function calculateSummary(
  fileName: string,
  matches: ShopeeRowMatch[]
): ProcessSummary {
  let matchedCount = 0;
  let unmatchedCount = 0;
  let stockChangedCount = 0;
  let outOfStockCount = 0;
  let inStockCount = 0;

  for (const m of matches) {
    if (m.matchStatus === 'matched') {
      matchedCount++;
      if (m.newStock !== null) {
        if (m.newStock === 0) {
          outOfStockCount++;
        } else {
          inStockCount++;
        }
        if (parseNumber(m.originalStock) !== m.newStock) {
          stockChangedCount++;
        }
      }
    } else {
      unmatchedCount++;
    }
  }

  return {
    fileName,
    totalRows: matches.length,
    matchedCount,
    unmatchedCount,
    stockChangedCount,
    outOfStockCount,
    inStockCount,
  };
}

/**
 * Matches a single Shopee row to STOCK LIST database following user's rule:
 * Kolom 5 (E): Kode SKU
 * Kolom 6 (F): Kode Integrasi / SKU Variasi
 *
 * Rule:
 * "kode di kolom 6 dan kolom 5 sama jadi ambil salah satunya yang ada pada kolom,
 * jika di kolom ke 5 tidak ada namun di kolom ke 6 ada, berarti SKU nya ambil dari kolom ke 6."
 */
export function getEffectiveShopeeSku(
  skuCol5: string,
  skuCol6: string
): { sku: string; source: 'col5' | 'col6' | 'none' } {
  const s5 = cleanSku(skuCol5);
  const s6 = cleanSku(skuCol6);

  if (s5) {
    return { sku: s5, source: 'col5' };
  }
  if (s6) {
    return { sku: s6, source: 'col6' };
  }
  return { sku: '', source: 'none' };
}

export function matchShopeeRowToStock(
  skuCol5: string,
  skuCol6: string,
  rawRow: any[] | undefined,
  maps: MatchingMaps
): {
  matchedStock: StockListItem | null;
  matchedBy?: 'col5' | 'col6' | '5digits_sku' | 'parent_sku' | 'code' | 'barcode' | undefined;
  notes: string;
} {
  const s5 = cleanSku(skuCol5);
  const s6 = cleanSku(skuCol6);

  // 1. Primary candidate: Take SKU from Kolom 5 (E) if present; if Kolom 5 is empty, take from Kolom 6 (F)
  const primary = getEffectiveShopeeSku(s5, s6);

  if (primary.sku) {
    const resPrimary = findStockForSku(primary.sku, maps);
    if (resPrimary.stockItem) {
      const colLabel = primary.source === 'col5' ? 'Kolom 5 (E)' : 'Kolom 6 (F)';
      return {
        matchedStock: resPrimary.stockItem,
        matchedBy: primary.source === 'col5' ? 'col5' : 'col6',
        notes: resPrimary.notes || `Cocok tepat via ${colLabel} (${primary.sku}) dengan Kolom 1 STOCK LIST (${resPrimary.stockItem.code})`,
      };
    }
  }

  // 2. If Kolom 5 was taken but didn't match, and Kolom 6 has a distinct code, try Kolom 6 as fallback
  if (s6 && s6 !== primary.sku) {
    const res6 = findStockForSku(s6, maps);
    if (res6.stockItem) {
      return {
        matchedStock: res6.stockItem,
        matchedBy: 'col6',
        notes: res6.notes || `Cocok tepat via Kolom 6 (F) (${s6}) dengan Kolom 1 STOCK LIST (${res6.stockItem.code})`,
      };
    }
  }

  // 3. Fallback: Check Column 10 (SKU Induk / Parent SKU in index 9)
  if (rawRow && rawRow[9]) {
    const parentSku = cleanSku(rawRow[9]);
    if (parentSku && parentSku !== s5 && parentSku !== s6) {
      const resParent = findStockForSku(parentSku, maps);
      if (resParent.stockItem) {
        return {
          matchedStock: resParent.stockItem,
          matchedBy: 'parent_sku',
          notes: resParent.notes || `Cocok tepat via SKU Induk Kolom 10 (${parentSku}) dengan Kolom 1 STOCK LIST (${resParent.stockItem.code})`,
        };
      }
    }
  }

  // 4. Fallback: Search all other columns in the row (Kode Produk, Kode Variasi, Nama Produk) for 5-digit codes
  if (rawRow && Array.isArray(rawRow)) {
    for (let c = 0; c < Math.min(rawRow.length, 10); c++) {
      if (c === 4 || c === 5 || c === 9) continue; // Already checked
      const cellVal = cleanSku(rawRow[c]);
      if (!cellVal) continue;

      const tokens = extract5DigitTokens(cellVal);
      for (const tok of tokens) {
        if (tok === s5 || tok === s6) continue;
        const resTok = findStockForSku(tok, maps);
        if (resTok.stockItem) {
          return {
            matchedStock: resTok.stockItem,
            matchedBy: '5digits_sku',
            notes: resTok.notes || `Cocok via SKU 5 Digit (${tok}) dari baris data dengan STOCK LIST (${resTok.stockItem.code})`,
          };
        }
      }
    }
  }

  return {
    matchedStock: null,
    matchedBy: undefined,
    notes: s5 || s6
      ? `SKU (${s5 || s6}) tidak ditemukan di database STOCK LIST`
      : 'Baris tanpa kode SKU di Kolom 5 maupun Kolom 6',
  };
}

export function compareBalistWithStockList(
  balistList: BalistShopeeItem[],
  maps: MatchingMaps
): {
  items: import('../types').BalistComparisonItem[];
  summary: import('../types').BalistComparisonSummary;
} {
  const items: import('../types').BalistComparisonItem[] = [];
  let matchedCount = 0;
  let unmatchedCount = 0;
  let inStockCount = 0;
  let outOfStockCount = 0;
  let totalStockQuantity = 0;

  for (const balistItem of balistList) {
    const matchResult = matchShopeeRowToStock(
      balistItem.skuCol5,
      balistItem.skuCol6,
      balistItem.rawRow,
      maps
    );

    const matchedStock = matchResult.matchedStock;
    const productName = (balistItem.rawRow && balistItem.rawRow[1] !== undefined && balistItem.rawRow[1] !== null && String(balistItem.rawRow[1]).trim() !== '')
      ? String(balistItem.rawRow[1]).trim()
      : (balistItem.productName || (balistItem.rawRow && balistItem.rawRow[2] ? String(balistItem.rawRow[2]).trim() : ''));
    const variationName = balistItem.variationName || (balistItem.rawRow && balistItem.rawRow[3] ? String(balistItem.rawRow[3]).trim() : (balistItem.rawRow && balistItem.rawRow[4] ? String(balistItem.rawRow[4]).trim() : ''));
    const parentSku = balistItem.parentSku || (balistItem.rawRow && balistItem.rawRow[9] ? cleanSku(balistItem.rawRow[9]) : '');

    if (matchedStock) {
      matchedCount++;
      const qty = matchedStock.qty;
      totalStockQuantity += qty;
      if (qty > 0) {
        inStockCount++;
      } else {
        outOfStockCount++;
      }

      items.push({
        rowIndex: balistItem.rowIndex,
        skuCol5: balistItem.skuCol5,
        skuCol6: balistItem.skuCol6,
        productName,
        variationName,
        parentSku,
        matchedStockItem: matchedStock,
        stockQty: qty,
        matchStatus: 'matched',
        matchedBy: matchResult.matchedBy as any,
        notes: matchResult.notes,
        rawRow: balistItem.rawRow,
      });
    } else {
      unmatchedCount++;
      items.push({
        rowIndex: balistItem.rowIndex,
        skuCol5: balistItem.skuCol5,
        skuCol6: balistItem.skuCol6,
        productName,
        variationName,
        parentSku,
        matchedStockItem: null,
        stockQty: null,
        matchStatus: 'unmatched',
        notes: matchResult.notes,
        rawRow: balistItem.rawRow,
      });
    }
  }

  return {
    items,
    summary: {
      totalRows: balistList.length,
      matchedCount,
      unmatchedCount,
      inStockCount,
      outOfStockCount,
      totalStockQuantity,
    },
  };
}

/**
 * Updates the designated "Stok Masuk" / "Stok" column in rows using matched stock quantities from STOCK LIST
 */
export function updateBalistRowsWithStock(
  rows: any[][],
  maps: MatchingMaps,
  stockColIndex: number,
  options?: {
    unmatchedAction?: 'zero' | 'keep';
    col5Index?: number;
    col6Index?: number;
  }
): {
  updatedRows: any[][];
  stats: {
    total: number;
    matchedCount: number;
    unmatchedCount: number;
    inStockCount: number;
    outOfStockCount: number;
  };
} {
  const col5 = options?.col5Index ?? 4; // Column 5 (E)
  const col6 = options?.col6Index ?? 5; // Column 6 (F)
  const unmatchedAction = options?.unmatchedAction ?? 'zero';

  let matchedCount = 0;
  let unmatchedCount = 0;
  let inStockCount = 0;
  let outOfStockCount = 0;

  const updatedRows = rows.map((row) => {
    const newRow = [...row];
    // Ensure array is long enough to include stockColIndex
    while (newRow.length <= stockColIndex) {
      newRow.push('');
    }

    const sku5 = cleanSku(newRow[col5]);
    const sku6 = cleanSku(newRow[col6]);

    let matchedStock: StockListItem | null = null;
    if (sku5) {
      const res = findStockForSku(sku5, maps);
      if (res.stockItem) matchedStock = res.stockItem;
    }
    if (!matchedStock && sku6) {
      const res = findStockForSku(sku6, maps);
      if (res.stockItem) matchedStock = res.stockItem;
    }

    if (matchedStock) {
      matchedCount++;
      const qty = matchedStock.qty;
      if (qty > 0) {
        inStockCount++;
      } else {
        outOfStockCount++;
      }
      newRow[stockColIndex] = qty;
    } else {
      unmatchedCount++;
      if (unmatchedAction === 'zero') {
        newRow[stockColIndex] = 0;
      }
      // if 'keep', do not modify newRow[stockColIndex]
    }

    return newRow;
  });

  return {
    updatedRows,
    stats: {
      total: rows.length,
      matchedCount,
      unmatchedCount,
      inStockCount,
      outOfStockCount,
    },
  };
}
