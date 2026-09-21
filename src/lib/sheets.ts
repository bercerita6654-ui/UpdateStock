import * as XLSX from 'xlsx';
import { StockListItem, BalistShopeeItem } from '../types';

export const cleanSku = (val: unknown): string => {
  if (val === null || val === undefined) return '';
  let str = String(val).trim();
  // Remove leading and trailing double or single quotes and backticks
  str = str.replace(/^["'`]+|["'`]+$/g, '').trim();
  return str;
};

/**
 * Safely parses numbers from various formats:
 * - Indonesian thousand dots (e.g. "99.000" -> 99000, "1.000.000" -> 1000000, "25.500" -> 25500)
 * - English thousand commas (e.g. "99,000" -> 99000, "1,000,000" -> 1000000)
 * - Indonesian decimal commas (e.g. "99,5" -> 99.5, "99.000,50" -> 99000.5)
 * - English decimal dots (e.g. "99.5" -> 99.5, "99,000.50" -> 99000.5)
 * - Zero and accounting formats (e.g. 0, "0", "0.0", "(50)" -> -50)
 * - Text with currency or units (e.g. "Rp 99.000", "99.000 pcs" -> 99000)
 */
export const parseNumber = (val: unknown): number => {
  if (val === null || val === undefined) return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;

  let str = String(val).trim();
  if (!str) return 0;

  // Handle accounting parentheses negative format: e.g. "(50)" -> -50
  const isParenthesesNegative = /^\(.*\)$/.test(str);

  // Strip currency prefixes/suffixes like "Rp", "IDR", "pcs", ",-", etc.
  str = str.replace(/,\s*-$/, '');
  str = str.replace(/[^0-9.,\-+]/g, '').trim();
  if (!str) return 0;

  const isNegative = str.startsWith('-') || isParenthesesNegative;
  str = str.replace(/^[+-]/, '');

  const hasDot = str.includes('.');
  const hasComma = str.includes(',');

  let numStr = str;

  if (hasDot && hasComma) {
    const lastDotIndex = str.lastIndexOf('.');
    const lastCommaIndex = str.lastIndexOf(',');
    if (lastDotIndex > lastCommaIndex) {
      // e.g. "1,250,000.50" -> comma is thousand separator, dot is decimal
      numStr = str.replace(/,/g, '');
    } else {
      // e.g. "1.250.000,50" -> dot is thousand separator, comma is decimal
      numStr = str.replace(/\./g, '').replace(/,/g, '.');
    }
  } else if (hasDot && !hasComma) {
    // Only dots present: e.g. "99.000", "1.000.000", "99.5", "12.34"
    const dotParts = str.split('.');
    if (dotParts.length > 2) {
      // Multiple dots: definitely thousand separators (e.g. "1.000.000")
      numStr = str.replace(/\./g, '');
    } else {
      // Single dot: e.g. "99.000" vs "99.5"
      const decimalPart = dotParts[1];
      // In Indonesian inventory/stock and price context, a dot followed by 3 digits
      // (like .000, .500, .250, .100) is a thousand separator (e.g. "99.000" = 99000).
      if (decimalPart.length === 3) {
        numStr = str.replace(/\./g, '');
      } else {
        // 1 or 2 digits -> standard decimal
        numStr = str;
      }
    }
  } else if (hasComma && !hasDot) {
    // Only commas present: e.g. "99,000", "1,000,000", "99,5"
    const commaParts = str.split(',');
    if (commaParts.length > 2) {
      // Multiple commas -> definitely thousand separators (e.g. "1,000,000")
      numStr = str.replace(/,/g, '');
    } else {
      const decimalPart = commaParts[1];
      if (decimalPart.length === 3) {
        // "99,000" -> thousand separator
        numStr = str.replace(/,/g, '');
      } else {
        // "99,5" -> Indonesian decimal comma
        numStr = str.replace(/,/g, '.');
      }
    }
  }

  const result = parseFloat(numStr);
  if (isNaN(result)) return 0;
  return isNegative ? -result : result;
};

export const isOfficeFileError = (err: any): boolean => {
  const msg = (typeof err === 'string' ? err : err?.message || '').toLowerCase();
  return (
    msg.includes('not be an office file') ||
    msg.includes('office file') ||
    msg.includes('not supported for this document')
  );
};

export const isAuthError = (err: any): boolean => {
  if (!err) return false;
  if (typeof err === 'object' && err.isAuthError) return true;
  const msg = (typeof err === 'string' ? err : err?.message || err?.error?.message || '').toLowerCase();
  const status = err?.status || err?.code || 0;
  return (
    status === 401 ||
    status === 403 && (msg.includes('credential') || msg.includes('oauth') || msg.includes('auth') || msg.includes('permission')) ||
    msg.includes('invalid authentication credentials') ||
    msg.includes('invalid credentials') ||
    msg.includes('unauthenticated') ||
    msg.includes('oauth 2 access token') ||
    msg.includes('expected oauth') ||
    msg.includes('token expired') ||
    msg.includes('login cookie') ||
    msg.includes('auth/invalid-credential') ||
    msg.includes('auth/user-token-expired')
  );
};

export class GoogleAuthError extends Error {
  isAuthError = true;
  constructor(message: string = 'Sesi akun Google telah berakhir atau token akses tidak valid. Silakan Masuk Kembali dengan Google.') {
    super(message);
    this.name = 'GoogleAuthError';
  }
}

/**
 * Fetch raw file content from Google Drive API for Office files (.xlsx / .xls)
 */
export async function fetchOfficeFileRows(
  spreadsheetId: string,
  sheetName: string,
  accessToken: string
): Promise<any[][]> {
  const url = `https://www.googleapis.com/drive/v3/files/${spreadsheetId}?alt=media`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    const message = errorData?.error?.message || response.statusText;
    if (response.status === 401 || isAuthError(message)) {
      throw new GoogleAuthError();
    }
    throw new Error(
      `Dokumen berformat Excel (.xlsx) di Google Drive. Tidak dapat mengunduh via Drive API (${response.status}): ${message}`
    );
  }

  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(new Uint8Array(arrayBuffer), {
    type: 'array',
    cellDates: true,
  });

  let targetSheet = workbook.SheetNames[0];
  if (sheetName) {
    const found = workbook.SheetNames.find(
      (s) => s.trim().toLowerCase() === sheetName.trim().toLowerCase()
    );
    if (found) targetSheet = found;
  }

  const worksheet = workbook.Sheets[targetSheet];
  if (!worksheet) return [];

  const rows: any[][] = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: '',
  });

  return rows;
}

/**
 * Find exact sheet title from Google Spreadsheet metadata (case-insensitive)
 */
export async function getExactSheetTitle(
  spreadsheetId: string,
  targetName: string,
  accessToken: string
): Promise<string> {
  if (!targetName) return targetName;
  try {
    const metaUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets.properties.title`;
    const res = await fetch(metaUrl, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    if (res.ok) {
      const meta = await res.json();
      const sheets = meta.sheets || [];
      const match = sheets.find(
        (s: any) =>
          s.properties?.title?.trim().toLowerCase() === targetName.trim().toLowerCase()
      );
      if (match && match.properties?.title) {
        return match.properties.title;
      }
      if (sheets.length > 0 && sheets[0].properties?.title) {
        return sheets[0].properties.title;
      }
    }
  } catch (err) {
    console.warn('Could not query exact sheet title, using targetName:', err);
  }
  return targetName;
}

/**
 * Convert 0-based column index to Excel column letters (0 -> A, 26 -> AA, etc.)
 */
export function getColumnLetter(colIndex: number): string {
  let temp = colIndex;
  let letter = '';
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

/**
 * Automatically detects the "Stok Masuk" / "Stok" column index in Balistshopee or Shopee sheets
 */
export function detectBalistStockColumn(rows: any[][]): {
  stockColIndex: number;
  headerName: string;
} {
  if (!rows || rows.length === 0) {
    return { stockColIndex: 6, headerName: 'Kolom 7 (G)' };
  }

  // Scan rows 0 to 6 (the header rows)
  const maxSearchRows = Math.min(rows.length, 7);
  for (let r = 0; r < maxSearchRows; r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      const text = String(row[c] || '').trim().toLowerCase();
      if (!text) continue;
      // Exact or includes keywords for stock
      if (
        text.includes('stok masuk') ||
        text === 'stok' ||
        text === 'stock' ||
        text.includes('jumlah stok') ||
        text.includes('stok saat ini') ||
        text.includes('current stock') ||
        text.includes('stok gudang')
      ) {
        return {
          stockColIndex: c,
          headerName: String(row[c]).trim() || `Kolom ${c + 1} (${getColumnLetter(c)})`,
        };
      }
    }
  }

  // Secondary search for 'qty' or 'kuantitas'
  for (let r = 0; r < maxSearchRows; r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      const text = String(row[c] || '').trim().toLowerCase();
      if (text.includes('qty') || text.includes('kuantitas')) {
        return {
          stockColIndex: c,
          headerName: String(row[c]).trim() || `Kolom ${c + 1} (${getColumnLetter(c)})`,
        };
      }
    }
  }

  // Fallback: Default to column 7 (index 6, Col G in standard Shopee template)
  const sampleRow = rows[Math.min(rows.length - 1, 6)] || rows[0] || [];
  const fallbackIndex = sampleRow.length > 6 ? 6 : Math.max(0, sampleRow.length - 1);
  return {
    stockColIndex: fallbackIndex,
    headerName: `Kolom ${fallbackIndex + 1} (${getColumnLetter(fallbackIndex)})`,
  };
}

/**
 * Convert an Office file (.xlsx) stored on Google Drive to a true native Google Spreadsheet.
 * Uses Google Drive API files.copy with mimeType: 'application/vnd.google-apps.spreadsheet'.
 */
export async function convertOfficeFileToGoogleSheet(
  fileId: string,
  newTitle: string,
  accessToken: string
): Promise<{ id: string; name: string; webViewLink?: string }> {
  const url = `https://www.googleapis.com/drive/v3/files/${fileId}/copy?supportsAllDrives=true&fields=id,name,webViewLink`;
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      name: newTitle || 'Balistshopee (Google Spreadsheet)',
      mimeType: 'application/vnd.google-apps.spreadsheet',
    }),
  });

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    const message = err?.error?.message || response.statusText;
    if (response.status === 401 || isAuthError(message)) {
      throw new GoogleAuthError();
    }
    throw new Error(
      `Gagal mengonversi file Excel ke Google Spreadsheet (${response.status}): ${message}`
    );
  }

  return response.json();
}

/**
 * Directly download an array of rows as an Excel (.xlsx) file to the user's browser.
 * Preserves original template formatting, column widths, headers, styles and formulas if rawWorkbookData is provided.
 */
export function downloadBalistRowsAsXlsx(
  originalRows: any[][],
  fileName: string = 'Balistshopee_Updated.xlsx',
  options?: {
    rawWorkbookData?: Uint8Array;
    sheetName?: string;
    startRowIndex?: number;
    updatedRowValues?: any[][];
  }
) {
  const ws = XLSX.utils.aoa_to_sheet(originalRows);

  // Set standardized column widths for Shopee Mass Update template
  const colWidths = [
    { wch: 18 }, // Col 1: Kode Produk
    { wch: 42 }, // Col 2: Nama Produk
    { wch: 22 }, // Col 3: No. Integrasi Produk
    { wch: 18 }, // Col 4: Kode Variasi
    { wch: 26 }, // Col 5: Nama Variasi
    { wch: 20 }, // Col 6: Kode Integrasi
    { wch: 14 }, // Col 7: Stok
    { wch: 16 }, // Col 8: Harga
    { wch: 16 }, // Col 9: Status Produk
    { wch: 20 }, // Col 10: SKU Induk
  ];
  ws['!cols'] = colWidths;

  const wb = XLSX.utils.book_new();
  const targetSheetName = options?.sheetName || 'Template';
  XLSX.utils.book_append_sheet(wb, ws, targetSheetName);

  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([out], {
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

/**
 * Update an Office file (.xlsx) directly on Google Drive using Drive upload API
 */
export async function updateOfficeFileOnDrive(
  spreadsheetId: string,
  sheetName: string,
  values: any[][],
  accessToken: string,
  startCell: string = 'A1'
): Promise<any> {
  // Try to download existing workbook first to preserve other sheets/formatting
  let wb: XLSX.WorkBook;
  try {
    const getUrl = `https://www.googleapis.com/drive/v3/files/${spreadsheetId}?alt=media&supportsAllDrives=true`;
    const getResp = await fetch(getUrl, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    if (getResp.ok) {
      const arrayBuffer = await getResp.arrayBuffer();
      wb = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
    } else {
      wb = XLSX.utils.book_new();
    }
  } catch (e) {
    wb = XLSX.utils.book_new();
  }

  // Find sheet case-insensitively, or default to first sheet
  let targetSheet = wb.SheetNames[0] || 'Balistshopee';
  if (sheetName) {
    const found = wb.SheetNames.find(
      (s) => s.trim().toLowerCase() === sheetName.trim().toLowerCase()
    );
    if (found) {
      targetSheet = found;
    }
  }

  let ws = wb.Sheets[targetSheet];

  if (!ws) {
    ws = XLSX.utils.aoa_to_sheet([]);
    wb.Sheets[targetSheet] = ws;
    if (!wb.SheetNames.includes(targetSheet)) {
      wb.SheetNames.push(targetSheet);
    }
  }

  // Parse startCell row (e.g. 'A7' -> row 7, 0-based index 6)
  const cellMatch = startCell.match(/([A-Za-z]+)([0-9]+)/);
  const startRow1Based = cellMatch ? parseInt(cellMatch[2], 10) : 1;
  const startRow0Based = startRow1Based - 1;

  if (startRow0Based > 0) {
    // Clean all existing data from startRow0Based downwards to preserve header rows 0..(startRow0Based - 1)
    for (const key of Object.keys(ws)) {
      if (key.startsWith('!')) continue;
      try {
        const coord = XLSX.utils.decode_cell(key);
        if (coord.r >= startRow0Based) {
          delete ws[key];
        }
      } catch {
        // ignore
      }
    }
  }

  // Add new values starting at startCell
  XLSX.utils.sheet_add_aoa(ws, values, { origin: startCell });

  // Recalculate range !ref so Excel viewer shows all rows
  let minR = 0;
  let minC = 0;
  let maxR = 0;
  let maxC = 0;
  let hasCells = false;
  for (const key of Object.keys(ws)) {
    if (key.startsWith('!')) continue;
    try {
      const coord = XLSX.utils.decode_cell(key);
      if (!hasCells) {
        minR = maxR = coord.r;
        minC = maxC = coord.c;
        hasCells = true;
      } else {
        if (coord.r < minR) minR = coord.r;
        if (coord.r > maxR) maxR = coord.r;
        if (coord.c < minC) minC = coord.c;
        if (coord.c > maxC) maxC = coord.c;
      }
    } catch {
      // ignore
    }
  }
  if (hasCells) {
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: minR, c: minC }, e: { r: maxR, c: maxC } });
  }

  // Ensure active sheet view points to the updated target sheet
  if (!wb.Workbook) {
    wb.Workbook = {};
  }
  const sheetIdx = wb.SheetNames.indexOf(targetSheet);
  (wb.Workbook as any).Views = [{ activeTab: sheetIdx >= 0 ? sheetIdx : 0 }];

  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

  const uploadUrl = `https://www.googleapis.com/upload/drive/v3/files/${spreadsheetId}?uploadType=media&supportsAllDrives=true`;
  const uploadResp = await fetch(uploadUrl, {
    method: 'PATCH',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    },
    body: new Uint8Array(out),
  });

  if (!uploadResp.ok) {
    const err = await uploadResp.json().catch(() => ({}));
    throw new Error(
      err?.error?.message ||
        `Gagal memperbarui file Excel di Google Drive (${uploadResp.status}: ${uploadResp.statusText})`
    );
  }

  return uploadResp.json();
}

export async function fetchSheetValues(
  spreadsheetId: string,
  range: string,
  accessToken: string
): Promise<any[][]> {
  const sheetNameMatch = range.match(/^([^!]+)!/);
  const sheetName = sheetNameMatch ? sheetNameMatch[1] : '';

  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
    range
  )}`;

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    const message = errorData?.error?.message || `Error ${response.status}`;

    if (response.status === 401 || isAuthError(message)) {
      throw new GoogleAuthError();
    }

    // If it's an Office file error, attempt fallback to Google Drive API
    if (isOfficeFileError(message)) {
      try {
        console.warn(
          `Document ${spreadsheetId} is an Office file. Attempting Google Drive API fallback...`
        );
        return await fetchOfficeFileRows(spreadsheetId, sheetName, accessToken);
      } catch (driveErr: any) {
        if (isAuthError(driveErr)) {
          throw new GoogleAuthError();
        }
        console.error('Drive fallback failed:', driveErr);
        throw new Error(
          `Dokumen "${spreadsheetId}" adalah file Excel (.xlsx) di Google Drive, bukan Google Spreadsheet. ` +
          `Silakan buka file di Google Drive lalu pilih: File > Simpan sebagai Google Spreadsheet (Save as Google Sheets), ` +
          `kemudian salin ID Spreadsheet baru ke menu Pengaturan.`
        );
      }
    }

    throw new Error(message || `Gagal memuat Google Sheet (${response.status})`);
  }

  const data = await response.json();
  return data.values || [];
}

export async function updateSheetValues(
  spreadsheetId: string,
  range: string,
  values: any[][],
  accessToken: string
): Promise<any> {
  const sheetNameMatch = range.match(/^([^!]+)!/);
  const rawSheetName = sheetNameMatch ? sheetNameMatch[1].replace(/^['"]|['"]$/g, '') : '';
  const cellMatch = range.match(/!([A-Za-z]+[0-9]+.*)$/);
  const cellPart = cellMatch ? cellMatch[1] : 'A1';
  const startCellOnlyMatch = cellPart.match(/^[A-Za-z]+[0-9]+/);
  const startCell = startCellOnlyMatch ? startCellOnlyMatch[0] : 'A1';

  // Resolve exact sheet title
  const exactSheetName = rawSheetName
    ? await getExactSheetTitle(spreadsheetId, rawSheetName, accessToken)
    : rawSheetName;
  const safeRange = exactSheetName ? `'${exactSheetName}'!${cellPart}` : range;

  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
    safeRange
  )}?valueInputOption=USER_ENTERED`;

  const response = await fetch(url, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      range: safeRange,
      majorDimension: 'ROWS',
      values,
    }),
  });

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    const message = err?.error?.message || '';

    if (response.status === 401 || isAuthError(message)) {
      throw new GoogleAuthError();
    }

    if (isOfficeFileError(message)) {
      try {
        console.warn(
          `Document ${spreadsheetId} is an Office file. Updating via Google Drive upload API starting at ${startCell}...`
        );
        return await updateOfficeFileOnDrive(
          spreadsheetId,
          exactSheetName || rawSheetName,
          values,
          accessToken,
          startCell
        );
      } catch (driveErr: any) {
        if (isAuthError(driveErr)) {
          throw new GoogleAuthError();
        }
        console.error('Drive update failed:', driveErr);
        throw new Error(
          `Dokumen "${spreadsheetId}" adalah file Excel (.xlsx) di Google Drive, bukan Google Spreadsheet. ` +
          `Silakan buka file di Google Drive lalu pilih: File > Simpan sebagai Google Spreadsheet (Save as Google Sheets), ` +
          `lalu gunakan ID Spreadsheet baru.`
        );
      }
    }

    throw new Error(message || 'Gagal memperbarui Google Sheet');
  }

  return response.json();
}

export async function clearSheetValues(
  spreadsheetId: string,
  range: string,
  accessToken: string
): Promise<any> {
  const sheetNameMatch = range.match(/^([^!]+)!/);
  const rawSheetName = sheetNameMatch ? sheetNameMatch[1].replace(/^['"]|['"]$/g, '') : '';
  const cellMatch = range.match(/!(.*)$/);
  const cellPart = cellMatch ? cellMatch[1] : 'A7:ZZ';

  const exactSheetName = rawSheetName
    ? await getExactSheetTitle(spreadsheetId, rawSheetName, accessToken)
    : rawSheetName;
  const safeRange = exactSheetName ? `'${exactSheetName}'!${cellPart}` : range;

  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
    safeRange
  )}:clear`;

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
  });

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    const message = err?.error?.message || '';

    if (response.status === 401 || isAuthError(message)) {
      throw new GoogleAuthError();
    }

    // If it's an office file, clearing will be handled during updateOfficeFileOnDrive
    if (isOfficeFileError(message)) {
      console.warn(`Clear sheet skipped for Office file ${spreadsheetId}. Will overwrite during update.`);
      return { cleared: true };
    }

    throw new Error(message || 'Gagal mengosongkan rentang Google Sheet');
  }

  return response.json();
}

export function parseStockListFromRows(rows: any[][]): StockListItem[] {
  if (!rows || rows.length <= 1) return [];

  // Default column indices based on standard inventory/STOCK LIST structure:
  // Col 1 (index 0): Kode Barang / SKU
  // Col 2 (index 1): Barcode
  // Col 3 (index 2): Nama Barang / Deskripsi
  // Col 4 (index 3): Satuan / Unit
  // Col 5 (index 4): Kategori
  // Col 6 (index 5): Merk / Brand
  // Col 15 (index 14): Saldo Akhir / Qty / Stok
  let codeIdx = 0;
  let barcodeIdx = 1;
  let descIdx = 2;
  let unitIdx = 3;
  let catIdx = 4;
  let brandIdx = 5;
  let qtyIdx = 14;
  let headerRowIndex = 0;

  // Search rows 0..5 to see if any row contains header keywords
  for (let r = 0; r < Math.min(rows.length, 6); r++) {
    const rRow = rows[r] || [];
    let matchCount = 0;
    for (let c = 0; c < rRow.length; c++) {
      const h = String(rRow[c] || '').trim().toLowerCase();
      if (!h) continue;

      if (h.includes('barcode') || h.includes('bar code') || h.includes('bar_code')) {
        barcodeIdx = c;
        matchCount++;
      } else if (
        h === 'code' ||
        h === 'kode' ||
        h === 'sku' ||
        h.includes('kode barang') ||
        h.includes('kd barang') ||
        h.includes('kd_barang') ||
        h.includes('kode_barang') ||
        h.includes('item code') ||
        h.includes('item no') ||
        h.includes('sku 5') ||
        h.includes('kode sku') ||
        (h.includes('kode') && !h.includes('barcode')) ||
        (h.includes('code') && !h.includes('barcode'))
      ) {
        codeIdx = c;
        matchCount++;
      } else if (
        h.includes('nama') ||
        h.includes('deskripsi') ||
        h.includes('description') ||
        h.includes('item') ||
        h.includes('produk')
      ) {
        descIdx = c;
        matchCount++;
      } else if (h.includes('satuan') || h.includes('unit') || h.includes('kemasan')) {
        unitIdx = c;
        matchCount++;
      } else if (
        h.includes('kategori') ||
        h.includes('category') ||
        h.includes('kelompok') ||
        h.includes('group') ||
        h.includes('kat') ||
        h.includes('jenis') ||
        h.includes('golongan') ||
        h.includes('tipe') ||
        h.includes('type') ||
        h.includes('dept') ||
        h.includes('departemen') ||
        h.includes('klasifikasi')
      ) {
        catIdx = c;
        matchCount++;
      } else if (
        h.includes('merk') ||
        h.includes('brand') ||
        h.includes('merek') ||
        h.includes('brand/merk') ||
        h.includes('merk/brand') ||
        h.includes('pabrik') ||
        h.includes('produsen') ||
        h.includes('principal') ||
        h.includes('vendor') ||
        h.includes('supplier')
      ) {
        brandIdx = c;
        matchCount++;
      } else if (
        h.includes('qty') ||
        h.includes('stok') ||
        h.includes('stock') ||
        h.includes('kuantitas') ||
        h.includes('saldo') ||
        h.includes('sisa')
      ) {
        qtyIdx = c;
        matchCount++;
      }
    }

    if (matchCount >= 2) {
      headerRowIndex = r;
      break;
    }
  }

  const items: StockListItem[] = [];
  // Parse rows starting after the header row
  for (let i = headerRowIndex + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;

    const rawCode = row[codeIdx] !== undefined ? row[codeIdx] : row[0];
    const rawCleaned = cleanSku(rawCode);
    if (!rawCleaned) continue;

    // Standardize 5-digit numeric SKU code (e.g. 7945 -> 07945, 123 -> 00123)
    const code = /^\d{1,5}$/.test(rawCleaned) ? rawCleaned.padStart(5, '0') : rawCleaned;

    const barcode = cleanSku(row[barcodeIdx] !== undefined ? row[barcodeIdx] : row[1]);
    const description = row[descIdx] !== undefined ? String(row[descIdx]).trim() : (row[2] ? String(row[2]).trim() : '');
    const unit = row[unitIdx] !== undefined ? String(row[unitIdx]).trim() : (row[3] ? String(row[3]).trim() : '');
    
    // Category (Kategori) - Col 5 or detected
    let category = '';
    if (row[catIdx] !== undefined && row[catIdx] !== null && String(row[catIdx]).trim() !== '') {
      category = String(row[catIdx]).trim();
    } else if (row[4] !== undefined && row[4] !== null && String(row[4]).trim() !== '') {
      category = String(row[4]).trim();
    }

    // Brand / Merk (Merek) - Col 6 or detected
    let brand = '';
    if (row[brandIdx] !== undefined && row[brandIdx] !== null && String(row[brandIdx]).trim() !== '') {
      brand = String(row[brandIdx]).trim();
    } else if (row[5] !== undefined && row[5] !== null && String(row[5]).trim() !== '') {
      brand = String(row[5]).trim();
    }

    // Qty / Stok - Col 15 or detected
    const qtyVal = row[qtyIdx] !== undefined ? row[qtyIdx] : (row[14] !== undefined ? row[14] : 0);
    const qty = parseNumber(qtyVal);

    items.push({
      code,
      barcode,
      description,
      unit,
      category,
      brand,
      qty,
      rawRow: row,
      rowIndex: i + 1,
    });
  }

  return items;
}

export function parseBalistShopeeFromRows(
  rows: any[][],
  startRowIndex: number = 7
): BalistShopeeItem[] {
  if (!rows || rows.length === 0) return [];

  // Default indices according to Shopee Mass Update format:
  // Col 2 (index 1 / B): Nama Produk
  // Col 5 (index 4 / E): Kode SKU / Nama Variasi (Kolom 5)
  // Col 6 (index 5 / F): Kode Integrasi / SKU Variasi (Kolom 6)
  // Col 10 (index 9 / J): SKU Induk
  let nameIdx = 1;
  let col5Idx = 4;
  let col6Idx = 5;
  let skuParentIdx = 9;

  // Optional dynamic scan from header rows (0..5)
  for (let r = 0; r < Math.min(rows.length, 6); r++) {
    const rRow = rows[r] || [];
    for (let c = 0; c < rRow.length; c++) {
      const h = String(rRow[c] || '').trim().toLowerCase();
      if (!h) continue;
      if (h === 'nama produk' || h === 'product name' || (h.includes('nama') && h.includes('produk'))) {
        nameIdx = c;
      } else if (
        h === 'kode variasi' ||
        h === 'nama variasi' ||
        h === 'kode sku' ||
        h === 'variation sku'
      ) {
        if (c <= 4) col5Idx = c;
      } else if (
        h.includes('kode integrasi') ||
        h.includes('nomor integrasi') ||
        h.includes('sku variasi') ||
        h.includes('integration code')
      ) {
        col6Idx = c;
      } else if (h.includes('sku induk') || h.includes('parent sku')) {
        skuParentIdx = c;
      }
    }
  }

  // If the sheet has fewer rows than startRowIndex, fallback to row 2
  const effectiveStart = rows.length >= startRowIndex ? startRowIndex : 2;

  const items: BalistShopeeItem[] = [];
  for (let i = effectiveStart - 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;
    
    // Check if row has any non-empty data
    const hasData = row.some((cell) => cell !== undefined && cell !== null && String(cell).trim() !== '');
    if (!hasData) continue;

    const skuCol5 = cleanSku(row[col5Idx] !== undefined ? row[col5Idx] : row[4]); // Column 5 (E) / Kode SKU
    const skuCol6 = cleanSku(row[col6Idx] !== undefined ? row[col6Idx] : row[5]); // Column 6 (F) / Kode Integrasi / SKU Variasi
    const productName = row[nameIdx] !== undefined ? String(row[nameIdx]).trim() : (row[1] ? String(row[1]).trim() : '');
    const variationName = row[col5Idx] !== undefined ? String(row[col5Idx]).trim() : (row[4] ? String(row[4]).trim() : '');
    const parentSku = row[skuParentIdx] !== undefined ? cleanSku(row[skuParentIdx]) : (row[9] ? cleanSku(row[9]) : '');

    items.push({
      skuCol5,
      skuCol6,
      productName,
      variationName,
      parentSku,
      rawRow: row,
      rowIndex: i + 1,
    });
  }

  return items;
}

export async function loadStockListData(
  spreadsheetId: string,
  sheetName: string,
  accessToken: string
): Promise<StockListItem[]> {
  const rows = await fetchSheetValues(spreadsheetId, `${sheetName}!A:Z`, accessToken);
  return parseStockListFromRows(rows);
}

export async function loadBalistShopeeData(
  spreadsheetId: string,
  sheetName: string,
  accessToken: string
): Promise<BalistShopeeItem[]> {
  // Read Balistshopee data. Structure starts product rows at row 7.
  const rows = await fetchSheetValues(spreadsheetId, `${sheetName}!A:Z`, accessToken);
  return parseBalistShopeeFromRows(rows, 7);
}
