import { StockListItem, BalistShopeeItem, ShopeeRowMatch, ProcessSummary } from '../types';
import { cleanSku, parseNumber } from './sheets';

export interface MatchingMaps {
  stockListByCode: Map<string, StockListItem>;
  stockListByNormalizedCode: Map<string, StockListItem>;
  stockListByBarcode: Map<string, StockListItem>;
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

export function buildMatchingMaps(
  stockList: StockListItem[],
  balistList: BalistShopeeItem[]
): MatchingMaps {
  const stockListByCode = new Map<string, StockListItem>();
  const stockListByNormalizedCode = new Map<string, StockListItem>();
  const stockListByBarcode = new Map<string, StockListItem>();
  const balistByCol5 = new Map<string, BalistShopeeItem>();
  const balistByCol6 = new Map<string, BalistShopeeItem>();

  // Index Stock List
  for (const item of stockList) {
    if (item.code) {
      const codeClean = cleanSku(item.code);
      const codeUpper = codeClean.toUpperCase();
      stockListByCode.set(codeUpper, item);

      const norm = normalizeSkuKey(item.code);
      if (norm) stockListByNormalizedCode.set(norm, item);

      const noZero = stripLeadingZeros(item.code);
      if (noZero) stockListByNormalizedCode.set(noZero.toUpperCase(), item);
    }

    if (item.barcode) {
      const barcodeUpper = cleanSku(item.barcode).toUpperCase();
      stockListByBarcode.set(barcodeUpper, item);
      const normBarcode = normalizeSkuKey(item.barcode);
      if (normBarcode) stockListByNormalizedCode.set(normBarcode, item);
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
    balistByCol5,
    balistByCol6,
  };
}

export interface MatchResult {
  stockItem: StockListItem | null;
  matchedBy?: 'code' | 'barcode' | 'balist_col5' | 'balist_col6';
  notes?: string;
}

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

  // 3. Match via Balistshopee: Check if SKU exists in Balistshopee Col 5 -> Cross-reference Col 6 to STOCK LIST
  const balistFromCol5 = maps.balistByCol5.get(upper) || (norm ? maps.balistByCol5.get(norm) : undefined);
  if (balistFromCol5 && balistFromCol5.skuCol6) {
    const targetSku = balistFromCol5.skuCol6;
    const fromCol6 =
      maps.stockListByCode.get(targetSku.toUpperCase()) ||
      maps.stockListByNormalizedCode.get(normalizeSkuKey(targetSku)) ||
      maps.stockListByNormalizedCode.get(stripLeadingZeros(targetSku).toUpperCase());

    if (fromCol6) {
      return {
        stockItem: fromCol6,
        matchedBy: 'balist_col5',
        notes: `Tercatat di Balistshopee Kolom 5 (${balistFromCol5.skuCol5}) -> Direferensikan ke Kolom 6 (${balistFromCol5.skuCol6})`,
      };
    }
  }

  // 4. Match via Balistshopee: Check if SKU exists in Balistshopee Col 6 -> Cross-reference Col 5 to STOCK LIST
  const balistFromCol6 = maps.balistByCol6.get(upper) || (norm ? maps.balistByCol6.get(norm) : undefined);
  if (balistFromCol6 && balistFromCol6.skuCol5) {
    const targetSku = balistFromCol5?.skuCol5 || balistFromCol6.skuCol5;
    const fromCol5 =
      maps.stockListByCode.get(targetSku.toUpperCase()) ||
      maps.stockListByNormalizedCode.get(normalizeSkuKey(targetSku)) ||
      maps.stockListByNormalizedCode.get(stripLeadingZeros(targetSku).toUpperCase());

    if (fromCol5) {
      return {
        stockItem: fromCol5,
        matchedBy: 'balist_col6',
        notes: `Tercatat di Balistshopee Kolom 6 (${balistFromCol6.skuCol6}) -> Direferensikan ke Kolom 5 (${balistFromCol6.skuCol5})`,
      };
    }
  }

  // 5. Match by Barcode in STOCK LIST (Column 2)
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
    const sku5 = cleanSku(balistItem.skuCol5);
    const sku6 = cleanSku(balistItem.skuCol6);

    let matchedStock: StockListItem | null = null;
    let matchedBy: 'col5' | 'col6' | undefined = undefined;
    let notes = '';

    // First try SKU from Column 5
    if (sku5) {
      const matchResult = findStockForSku(sku5, maps);
      if (matchResult.stockItem) {
        matchedStock = matchResult.stockItem;
        matchedBy = 'col5';
        notes = `Cocok via Kolom 5 (${sku5}) dengan STOCK LIST (${matchedStock.code})`;
      }
    }

    // If not matched, try SKU from Column 6
    if (!matchedStock && sku6) {
      const matchResult = findStockForSku(sku6, maps);
      if (matchResult.stockItem) {
        matchedStock = matchResult.stockItem;
        matchedBy = 'col6';
        notes = `Cocok via Kolom 6 (${sku6}) dengan STOCK LIST (${matchedStock.code})`;
      }
    }

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
        matchedBy,
        notes,
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
        notes: sku5 || sku6 ? 'SKU tidak ditemukan di STOCK LIST' : 'Baris tanpa SKU',
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
