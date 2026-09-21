export interface StockListItem {
  code: string; // Col 1 (Code / SKU)
  barcode?: string; // Col 2
  description?: string; // Col 3
  unit?: string;
  category?: string;
  brand?: string; // Merk / Brand
  qty: number; // Col 15 (Stock Qty)
  rawRow: (string | number)[];
  rowIndex: number;
}

export interface BalistShopeeItem {
  skuCol5: string; // Column 5 (E) / Variation Name or SKU
  skuCol6: string; // Column 6 (F) / Integration Code or SKU Variasi
  productName?: string; // Column 2 (B) / Nama Produk Shopee
  variationName?: string; // Column 5 (E) / Nama Variasi Shopee
  parentSku?: string; // Column 10 (J) / SKU Induk
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
  matchedBy?: 'code' | 'barcode' | 'balist_col5' | 'balist_col6' | '5digits_sku' | 'parent_sku';
  stockDiff?: number;
  notes?: string;
}

export interface BalistComparisonItem {
  rowIndex: number; // row in Balistshopee sheet (1-based)
  skuCol5: string; // Column 5 (E)
  skuCol6: string; // Column 6 (F)
  productName?: string; // Nama Produk Shopee (Kolom B)
  variationName?: string; // Nama Variasi (Kolom E)
  parentSku?: string; // SKU Induk (Kolom J)
  matchedStockItem: StockListItem | null;
  stockQty: number | null; // Quantity from STOCK LIST (Col 15)
  matchStatus: 'matched' | 'unmatched';
  matchedBy?: 'col5' | 'col6' | '5digits_sku' | 'parent_sku' | 'code' | 'barcode';
  notes?: string;
  rawRow: (string | number)[];
}

export type ExportFilterMode = 'all' | 'category' | 'brand' | 'custom';

export interface ExportFilterOptions {
  mode: ExportFilterMode;
  selectedCategories: string[];
  selectedBrands: string[];
  selectedRowIndices: number[]; // 1-based rowIndex or raw index
  storePrefix?: string;
  unmatchedStockAction?: 'keep' | 'zero';
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

export type LogType =
  | 'sync'
  | 'sheet_update'
  | 'upload'
  | 'download'
  | 'auth'
  | 'settings'
  | 'error'
  | 'info';

export type LogStatus = 'success' | 'error' | 'loading' | 'info';

export interface ActivityLogItem {
  id: string;
  timestamp: string; // ISO string
  formattedTime?: string;
  type: LogType;
  title: string;
  description: string;
  status: LogStatus;
  target?: string;
  details?: string;
  errorMessage?: string;
  rowCount?: number;
}
