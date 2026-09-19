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

// Normalizes an SKU for flexible matching: lowercase, strip all quotes, dashes, spaces, and leading zeros
export const normalizeSkuKey = (sku: string): string => {
  const cleaned = cleanSku(sku).toLowerCase();
  // Strip non-alphanumeric except maybe keep alphanumeric core
  return cleaned.replace(/[^a-z0-9]/gi, '');
};

export const stripLeadingZeros = (sku: string): string => {
  const cleaned = cleanSku(sku);
  return cleaned.replace(/^0+/, '');
};

// Extract all 5-digit numerical or alphanumeric sequence tokens from an SKU string
export const extract5DigitTokens = (sku: string): string[] => {
  const cleaned = cleanSku(sku);
  if (!cleaned) return [];
  const tokens = new Set<string>();

  // 1. Any 5 consecutive digits in the raw string (e.g. "12345" from "ABC-12345-XL")
  const digitMatches = cleaned.match(/\d{5}/g);
  if (digitMatches) {
    for (const d of digitMatches) {
      tokens.add(d);
    }
  }

  // 2. Split by common delimiters (-, _, /, ., space) and take 5-char parts
  const parts = cleaned.split(/[-_/. ,|#]+/);
  for (const part of parts) {
    const pTrim = part.trim();
    if (pTrim.length === 5) {
      tokens.add(pTrim.toUpperCase());
    }
    const pNoZero = stripLeadingZeros(pTrim);
    if (pNoZero.length === 5) {
      tokens.add(pNoZero.toUpperCase());
    }
    // If 4 digits, also try 5-digit padded with leading zero
    if (/^\d{4}$/.test(pTrim)) {
      tokens.add(('0' + pTrim).toUpperCase());
    }
  }

  // 3. Normalized string first 5 chars
  const norm = normalizeSkuKey(cleaned);
  if (norm.length >= 5) {
    tokens.add(norm.slice(0, 5).toUpperCase());
  }

  return Array.from(tokens);
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
      stockListByCode.set(codeUpper, item);

      const norm = normalizeSkuKey(item.code);
      if (norm) {
        stockListByNormalizedCode.set(norm, item);
        if (norm.length >= 5) {
          const pref5 = norm.slice(0, 5).toUpperCase();
          if (!stockListByPrefix5.has(pref5)) {
            stockListByPrefix5.set(pref5, item);
          }
        }
      }

      const noZero = stripLeadingZeros(item.code);
      if (noZero) stockListByNormalizedCode.set(noZero.toUpperCase(), item);

      // Index 5-digit tokens from Stock List item code
      const tokens = extract5DigitTokens(item.code);
      for (const tok of tokens) {
        if (!stockListBy5Digits.has(tok)) {
          stockListBy5Digits.set(tok, item);
        }
        if (/^\d{5}$/.test(tok)) {
          fiveDigitStockEntries.push({ code5: tok, item });
        }
      }

      // If code itself is 5 characters
      if (codeClean.length === 5) {
        stockListBy5Digits.set(codeUpper, item);
      }
      // If code is 4 digits numeric, index both 4 and 0-padded 5 digits
      if (/^\d{4}$/.test(codeClean)) {
        stockListBy5Digits.set(('0' + codeClean).toUpperCase(), item);
      }
    }

    if (item.barcode) {
      const barcodeUpper = cleanSku(item.barcode).toUpperCase();
      stockListByBarcode.set(barcodeUpper, item);
      const normBarcode = normalizeSkuKey(item.barcode);
      if (normBarcode) stockListByNormalizedCode.set(normBarcode, item);

      // Extract 5-digit tokens from barcode if applicable
      const barcodeTokens = extract5DigitTokens(item.barcode);
      for (const tok of barcodeTokens) {
        if (!stockListBy5Digits.has(tok)) {
          stockListBy5Digits.set(tok, item);
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
  'PCS', 'BOX', 'LUSIN', 'SET', 'SATUAN', 'KECIL', 'SEDANG', 'BESAR', 'S', 'M', 'L', 'XL'
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
  const norm = normalizeSkuKey(cleaned);
  const noZero = stripLeadingZeros(cleaned).toUpperCase();

  // 1. Direct match in STOCK LIST by Code (Column 1)
  if (maps.stockListByCode.has(upper)) {
    return {
      stockItem: maps.stockListByCode.get(upper)!,
      matchedBy: 'code',
      notes: 'Cocok langsung dengan Kode SKU STOCK LIST (Kolom 1)',
    };
  }

  // 2. Normalized match in STOCK LIST (e.g. leading zero difference "07945" vs "7945")
  if (maps.stockListByNormalizedCode.has(upper)) {
    return {
      stockItem: maps.stockListByNormalizedCode.get(upper)!,
      matchedBy: 'code',
      notes: 'Cocok kode STOCK LIST (format variasi)',
    };
  }
  if (norm && maps.stockListByNormalizedCode.has(norm)) {
    return {
      stockItem: maps.stockListByNormalizedCode.get(norm)!,
      matchedBy: 'code',
      notes: 'Cocok kode STOCK LIST (normalisasi format)',
    };
  }
  if (noZero && maps.stockListByNormalizedCode.has(noZero)) {
    return {
      stockItem: maps.stockListByNormalizedCode.get(noZero)!,
      matchedBy: 'code',
      notes: 'Cocok kode STOCK LIST (tanpa awalan 0)',
    };
  }

  // 3. Sub-token matching (split by delimiter -, _, /, space, etc.)
  const subTokens = cleaned.split(/[-_/ ,.:|#()]+/).map(t => t.trim().toUpperCase()).filter(Boolean);
  for (const token of subTokens) {
    if (GENERIC_VARIATION_WORDS.has(token) || token.length < 3) continue;
    if (maps.stockListByCode.has(token)) {
      return {
        stockItem: maps.stockListByCode.get(token)!,
        matchedBy: '5digits_sku',
        notes: `Cocok token SKU (${token}) dengan Kode STOCK LIST`,
      };
    }
    const tokenNorm = normalizeSkuKey(token);
    if (tokenNorm && maps.stockListByNormalizedCode.has(tokenNorm)) {
      return {
        stockItem: maps.stockListByNormalizedCode.get(tokenNorm)!,
        matchedBy: '5digits_sku',
        notes: `Cocok token SKU (${token}) dengan STOCK LIST`,
      };
    }
  }

  // 4. Analisa Persamaan Kode SKU 5 Digit dengan sheet STOCK LIST
  const sku5Tokens = extract5DigitTokens(cleaned);
  for (const token of sku5Tokens) {
    if (maps.stockListBy5Digits.has(token)) {
      const matched = maps.stockListBy5Digits.get(token)!;
      return {
        stockItem: matched,
        matchedBy: '5digits_sku',
        notes: `Cocok analisa kode 5 digit SKU (${token}) dengan STOCK LIST (${matched.code})`,
      };
    }
    if (maps.stockListByCode.has(token)) {
      const matched = maps.stockListByCode.get(token)!;
      return {
        stockItem: matched,
        matchedBy: '5digits_sku',
        notes: `Cocok analisa kode 5 digit SKU (${token}) dengan STOCK LIST (${matched.code})`,
      };
    }
    if (maps.stockListByNormalizedCode.has(token)) {
      const matched = maps.stockListByNormalizedCode.get(token)!;
      return {
        stockItem: matched,
        matchedBy: '5digits_sku',
        notes: `Cocok analisa kode 5 digit SKU (${token}) dengan STOCK LIST (${matched.code})`,
      };
    }
  }

  // 5. Containment check: Is there a 5-digit code in STOCK LIST contained in the SKU?
  if (maps.fiveDigitStockEntries && maps.fiveDigitStockEntries.length > 0) {
    for (const entry of maps.fiveDigitStockEntries) {
      if (cleaned.includes(entry.code5) || (norm && norm.includes(entry.code5))) {
        return {
          stockItem: entry.item,
          matchedBy: '5digits_sku',
          notes: `Cocok analisa pola 5 digit (${entry.code5}) di dalam SKU dengan STOCK LIST`,
        };
      }
    }
  }

  // 6. Prefix 5 chars on normalized SKU
  if (norm && norm.length >= 5) {
    const pref5 = norm.slice(0, 5).toUpperCase();
    if (maps.stockListByPrefix5.has(pref5)) {
      const matched = maps.stockListByPrefix5.get(pref5)!;
      return {
        stockItem: matched,
        matchedBy: '5digits_sku',
        notes: `Cocok awalan 5 digit SKU (${pref5}) dengan STOCK LIST (${matched.code})`,
      };
    }
  }

  // 7. Match by Barcode in STOCK LIST (Column 2)
  if (maps.stockListByBarcode.has(upper)) {
    return {
      stockItem: maps.stockListByBarcode.get(upper)!,
      matchedBy: 'barcode',
      notes: 'Cocok dengan Barcode di STOCK LIST',
    };
  }

  return { stockItem: null, notes: 'SKU tidak ditemukan di database stok' };
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

export function matchShopeeRowToStock(
  skuCol5: string,
  skuCol6: string,
  rawRow: any[] | undefined,
  maps: MatchingMaps
): {
  matchedStock: StockListItem | null;
  matchedBy?: 'col5' | 'col6' | '5digits_sku' | 'code' | 'barcode' | undefined;
  notes: string;
} {
  const s5 = cleanSku(skuCol5);
  const s6 = cleanSku(skuCol6);

  // 1. Try SKU from Column 6 (Kode Integrasi / SKU Variasi) - Primary check
  if (s6) {
    const res6 = findStockForSku(s6, maps);
    if (res6.stockItem) {
      return {
        matchedStock: res6.stockItem,
        matchedBy: 'col6',
        notes: res6.notes || `Cocok via Kolom 6 (${s6}) dengan STOCK LIST (${res6.stockItem.code})`,
      };
    }
  }

  // 2. Try SKU from Column 5 (Kode Variasi / SKU Variasi)
  if (s5 && !GENERIC_VARIATION_WORDS.has(s5.toUpperCase())) {
    const res5 = findStockForSku(s5, maps);
    if (res5.stockItem) {
      return {
        matchedStock: res5.stockItem,
        matchedBy: 'col5',
        notes: res5.notes || `Cocok via Kolom 5 (${s5}) dengan STOCK LIST (${res5.stockItem.code})`,
      };
    }
  }

  // 3. Check other columns in raw row (Col 10 SKU Induk, Col 4 Kode Variasi, Col 1 Kode Produk, Col 3 No Integrasi)
  if (rawRow && rawRow.length > 0) {
    const candidateCols = [9, 3, 0, 2];
    for (const colIdx of candidateCols) {
      if (rawRow[colIdx]) {
        const candSku = cleanSku(rawRow[colIdx]);
        if (candSku && candSku !== s5 && candSku !== s6 && !GENERIC_VARIATION_WORDS.has(candSku.toUpperCase())) {
          const resCand = findStockForSku(candSku, maps);
          if (resCand.stockItem) {
            return {
              matchedStock: resCand.stockItem,
              matchedBy: '5digits_sku',
              notes: resCand.notes || `Cocok via Kolom ${colIdx + 1} (${candSku}) analisa 5 digit dengan STOCK LIST (${resCand.stockItem.code})`,
            };
          }
        }
      }
    }
  }

  return {
    matchedStock: null,
    matchedBy: undefined,
    notes: s5 || s6 ? 'SKU tidak ditemukan di STOCK LIST' : 'Baris tanpa SKU',
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
