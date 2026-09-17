export interface StockListItem {
  code: string; // Col 1 (Code / SKU)
  barcode?: string; // Col 2
  description?: string; // Col 3
  unit?: string;
  category?: string;
  qty: number; // Col 15 (Stock Qty)
  rawRow: (string | number)[];
  rowIndex: number;
}

export interface BalistShopeeItem {
  skuCol5: string; // Column 5 (E)
  skuCol6: string; // Column 6 (F)
  rawRow: (string | number)[];
  rowIndex: number;
}

export interface ShopeeRowMatch {
  rowIndex: number; // 0-based index in the worksheet
  displayRow: number; // 1-based row number for user display
  rawSku: string;
  cleanedSku: string;
  productName?: string;
  variationName?: string;
  originalStock: number | string;
  newStock: number | null;
  matchStatus: 'matched' | 'unmatched';
  matchedBy?: 'code' | 'barcode' | 'balist_col5' | 'balist_col6';
  stockDiff?: number;
  notes?: string;
}

export interface BalistComparisonItem {
  rowIndex: number; // row in Balistshopee sheet (1-based)
  skuCol5: string; // Column 5 (E)
  skuCol6: string; // Column 6 (F)
  matchedStockItem: StockListItem | null;
  stockQty: number | null; // Quantity from STOCK LIST (Col 15)
  matchStatus: 'matched' | 'unmatched';
  matchedBy?: 'col5' | 'col6';
  notes?: string;
  rawRow: (string | number)[];
}

export interface BalistComparisonSummary {
  totalRows: number;
  matchedCount: number;
  unmatchedCount: number;
  inStockCount: number;
  outOfStockCount: number;
  totalStockQuantity: number;
}

export interface ProcessSummary {
  fileName: string;
  totalRows: number;
  matchedCount: number;
  unmatchedCount: number;
  stockChangedCount: number;
  outOfStockCount: number;
  inStockCount: number;
}

export interface SheetsConfig {
  balistSpreadsheetId: string;
  balistSheetName: string;
  stockSpreadsheetId: string;
  stockSheetName: string;
}
