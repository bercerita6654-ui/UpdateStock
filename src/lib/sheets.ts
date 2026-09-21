import * as XLSX from 'xlsx';
import { StockListItem, BalistShopeeItem } from '../types';

export const cleanSku = (val: unknown): string => {
  if (val === null || val === undefined) return '';
  let str = String(val);
  // Replace zero-width spaces, non-breaking spaces, BOM, control characters, tabs, linebreaks with space
  str = str.replace(/[\u200B-\u200D\uFEFF\u00A0\u180E\u2000-\u200A\u202F\u205F\u3000\t\r\n]/g, ' ');
  str = str.trim();
  // Remove leading Excel formulas '=', quotes, single/double backticks, typographic quotes
  str = str.replace(/^[="'`‘“′]+|[="'`’”″]+$/g, '').trim();
  // Remove trailing .0 or .00 or ,0 or ,00 from numeric cell exports (e.g. "10510.0" -> "10510")
  str = str.replace(/[.,]0+$/, '').trim();
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

export interface DetectedStockListColumns {
  headerRowIndex: number;
  codeColIndex: number;
  codeHeader: string;
  barcodeColIndex: number;
  barcodeHeader: string;
  descColIndex: number;
  descHeader: string;
  unitColIndex: number;
  unitHeader: string;
  catColIndex: number;
  catHeader: string;
  brandColIndex: number;
  brandHeader: string;
  qtyColIndex: number;
  qtyHeader: string;
}

/**
 * Intelligently scans rows to find the header row and dynamically maps columns
 * by prioritizing the exact header name "Qty" (or its standard variants),
 * allowing any custom column layout (e.g. Column I, L, O, or anywhere) to work seamlessly.
 */
export function detectStockListColumns(rows: any[][]): DetectedStockListColumns {
  if (!rows || rows.length === 0) {
    return {
      headerRowIndex: 0,
      codeColIndex: 0,
      codeHeader: 'Kode Barang',
      barcodeColIndex: 1,
      barcodeHeader: 'Barcode',
      descColIndex: 2,
      descHeader: 'Nama Barang',
      unitColIndex: 3,
      unitHeader: 'Satuan',
      catColIndex: 4,
      catHeader: 'Kategori',
      brandColIndex: 5,
      brandHeader: 'Merk',
      qtyColIndex: 14,
      qtyHeader: 'Qty',
    };
  }

  let bestHeaderRow = 0;
  let highestScore = -1;
  let detected = {
    codeColIndex: 0,
    codeHeader: 'Kolom 1 (A)',
    barcodeColIndex: 1,
    barcodeHeader: 'Kolom 2 (B)',
    descColIndex: 2,
    descHeader: 'Kolom 3 (C)',
    unitColIndex: 3,
    unitHeader: 'Kolom 4 (D)',
    catColIndex: 4,
    catHeader: 'Kolom 5 (E)',
    brandColIndex: 5,
    brandHeader: 'Kolom 6 (F)',
    qtyColIndex: -1,
    qtyHeader: '',
  };

  const maxScanRows = Math.min(rows.length, 15);

  for (let r = 0; r < maxScanRows; r++) {
    const row = rows[r] || [];
    if (row.length === 0) continue;

    let rowScore = 0;
    let rCode = -1, rCodeH = '';
    let rBarcode = -1, rBarcodeH = '';
    let rDesc = -1, rDescH = '';
    let rUnit = -1, rUnitH = '';
    let rCat = -1, rCatH = '';
    let rBrand = -1, rBrandH = '';
    let rQty = -1, rQtyH = '', rQtyScore = 0;

    for (let c = 0; c < row.length; c++) {
      const rawVal = String(row[c] || '').trim();
      if (!rawVal) continue;
      const h = rawVal.toLowerCase();
      const normH = h.replace(/[[\]().,:;_/\\-]/g, ' ').replace(/\s+/g, ' ').trim();

      // 1. QTY / STOK COLUMN DETECTION (HIGHEST PRIORITY INTELLIGENT SCANNER)
      // Disambiguate against irrelevant min/max/po/buffer columns
      const isExcludedStock =
        normH.includes('min') ||
        normH.includes('max') ||
        normH.includes('maks') ||
        normH.includes('safety') ||
        normH.includes('buffer') ||
        normH.includes('pesan') ||
        normH.includes('order') ||
        normH.includes('beli') ||
        normH.includes('rusak') ||
        normH.includes('reject') ||
        normH.includes('retur') ||
        normH.includes('pending') ||
        normH.includes('harga');

      if (!isExcludedStock) {
        let currentQtyScore = 0;
        // Exact "qty", "quantity", "kuantitas" (Priority #1)
        if (normH === 'qty' || normH === 'quantity' || normH === 'kuantitas') {
          currentQtyScore = 100;
        } else if (
          normH === 'qty akhir' ||
          normH === 'saldo akhir' ||
          normH === 'stok akhir' ||
          normH === 'total qty' ||
          normH === 'qty total' ||
          normH === 'stock qty' ||
          normH === 'qty stock' ||
          normH === 'stok fisik' ||
          normH === 'qty fisik' ||
          normH === 'qty on hand' ||
          normH === 'on hand' ||
          normH === 'saldo qty' ||
          normH === 'qty saldo' ||
          normH === 'stok saat ini' ||
          normH === 'current stock' ||
          normH === 'sisa stok' ||
          normH === 'saldo stok' ||
          normH === 'stok gudang' ||
          normH === 'qty gudang'
        ) {
          currentQtyScore = 90;
        } else if (normH === 'stok' || normH === 'stock' || normH === 'jumlah' || normH === 'saldo' || normH === 'sisa') {
          currentQtyScore = 80;
        } else if (normH.startsWith('qty') || normH.endsWith('qty') || normH.includes(' qty ') || normH.includes('kuantitas')) {
          currentQtyScore = 70;
        } else if (normH.includes('stok') || normH.includes('stock') || normH.includes('saldo')) {
          currentQtyScore = 60;
        }

        if (currentQtyScore > rQtyScore) {
          rQty = c;
          rQtyH = rawVal;
          rQtyScore = currentQtyScore;
        }
      }

      // 2. KODE / SKU
      if (
        normH === 'kode barang' ||
        normH === 'kode item' ||
        normH === 'kode produk' ||
        normH === 'item code' ||
        normH === 'product code' ||
        normH === 'sku' ||
        normH === 'kode sku' ||
        normH === 'kode' ||
        normH === 'part number' ||
        normH === 'id barang' ||
        normH === 'no item'
      ) {
        rCode = c;
        rCodeH = rawVal;
        rowScore += 20;
      }

      // 3. BARCODE
      if (
        normH.includes('barcode') ||
        normH.includes('bar code') ||
        normH === 'upc' ||
        normH === 'ean'
      ) {
        rBarcode = c;
        rBarcodeH = rawVal;
        rowScore += 15;
      }

      // 4. NAMA BARANG / DESKRIPSI
      if (
        normH === 'nama barang' ||
        normH === 'nama item' ||
        normH === 'nama produk' ||
        normH === 'deskripsi' ||
        normH === 'description' ||
        normH === 'item name' ||
        normH === 'product name' ||
        normH === 'nama'
      ) {
        rDesc = c;
        rDescH = rawVal;
        rowScore += 20;
      }

      // 5. SATUAN / UNIT
      if (
        normH === 'satuan' ||
        normH === 'unit' ||
        normH === 'uom' ||
        normH === 'kemasan' ||
        normH === 'satuan barang'
      ) {
        rUnit = c;
        rUnitH = rawVal;
        rowScore += 10;
      }

      // 6. KATEGORI
      if (
        normH.includes('kategori') ||
        normH.includes('category') ||
        normH.includes('kelompok') ||
        normH.includes('group') ||
        normH === 'jenis' ||
        normH === 'kat' ||
        normH.includes('dept') ||
        normH.includes('departemen') ||
        normH.includes('golongan') ||
        normH.includes('klasifikasi')
      ) {
        rCat = c;
        rCatH = rawVal;
        rowScore += 10;
      }

      // 7. MERK / BRAND
      if (
        normH.includes('merk') ||
        normH.includes('brand') ||
        normH.includes('merek') ||
        normH.includes('pabrik') ||
        normH.includes('produsen') ||
        normH.includes('principal') ||
        normH.includes('vendor') ||
        normH.includes('supplier')
      ) {
        rBrand = c;
        rBrandH = rawVal;
        rowScore += 10;
      }
    }

    if (rQtyScore > 0) {
      rowScore += rQtyScore;
    }

    if (rowScore > highestScore && (rowScore >= 40 || (rQty !== -1 && rowScore >= 30))) {
      highestScore = rowScore;
      bestHeaderRow = r;
      detected = {
        codeColIndex: rCode !== -1 ? rCode : 0,
        codeHeader: rCodeH || 'Kolom 1 (A)',
        barcodeColIndex: rBarcode !== -1 ? rBarcode : 1,
        barcodeHeader: rBarcodeH || 'Kolom 2 (B)',
        descColIndex: rDesc !== -1 ? rDesc : 2,
        descHeader: rDescH || 'Kolom 3 (C)',
        unitColIndex: rUnit !== -1 ? rUnit : 3,
        unitHeader: rUnitH || 'Kolom 4 (D)',
        catColIndex: rCat !== -1 ? rCat : 4,
        catHeader: rCatH || 'Kolom 5 (E)',
        brandColIndex: rBrand !== -1 ? rBrand : 5,
        brandHeader: rBrandH || 'Kolom 6 (F)',
        qtyColIndex: rQty,
        qtyHeader: rQtyH || (rQty !== -1 ? `Kolom ${rQty + 1} (${getColumnLetter(rQty)})` : ''),
      };
    }
  }

  // Fallback for Qty if no explicit Qty header found in scanned rows:
  // Check the last numeric columns or standard default position
  if (detected.qtyColIndex === -1) {
    const sampleRow = rows[Math.min(rows.length - 1, bestHeaderRow + 1)] || rows[0] || [];
    // If standard inventory layout has 15 columns, use 14, else check last column
    detected.qtyColIndex = sampleRow.length > 14 ? 14 : Math.max(0, sampleRow.length - 1);
    detected.qtyHeader = `Kolom ${detected.qtyColIndex + 1} (${getColumnLetter(detected.qtyColIndex)})`;
  }

  return {
    headerRowIndex: bestHeaderRow,
    ...detected,
  };
}

export function parseStockListFromRows(rows: any[][]): StockListItem[] {
  if (!rows || rows.length <= 1) return [];

  // Intelligently detect header locations and column indices dynamically
  const colMap = detectStockListColumns(rows);
  const {
    headerRowIndex,
    codeColIndex,
    barcodeColIndex,
    descColIndex,
    unitColIndex,
    catColIndex,
    brandColIndex,
    qtyColIndex,
  } = colMap;

  const items: StockListItem[] = [];

  // Parse rows starting after the detected header row
  for (let i = headerRowIndex + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;

    // Check if the row has any non-empty cell
    const hasAnyCell = row.some((c) => c !== undefined && c !== null && String(c).trim() !== '');
    if (!hasAnyCell) continue;

    // Kode SKU is retrieved from detected code column, falling back to Kolom 1 (index 0 / A)
    let rawCode =
      row[codeColIndex] !== undefined && row[codeColIndex] !== null && String(row[codeColIndex]).trim() !== ''
        ? row[codeColIndex]
        : (row[0] !== undefined && row[0] !== null ? row[0] : '');

    let rawCleaned = cleanSku(rawCode);
    if (!rawCleaned) continue;

    // Standardize numeric SKU code (e.g. 7945 -> 07945, 10510 -> 10510)
    const code = /^\d{1,5}$/.test(rawCleaned) ? rawCleaned.padStart(5, '0') : rawCleaned;

    const barcode = cleanSku(
      row[barcodeColIndex] !== undefined && row[barcodeColIndex] !== null
        ? row[barcodeColIndex]
        : (row[1] !== undefined ? row[1] : '')
    );

    const description =
      row[descColIndex] !== undefined && row[descColIndex] !== null
        ? String(row[descColIndex]).trim()
        : (row[2] ? String(row[2]).trim() : '');

    const unit =
      row[unitColIndex] !== undefined && row[unitColIndex] !== null
        ? String(row[unitColIndex]).trim()
        : (row[3] ? String(row[3]).trim() : '');

    // Category (Kategori)
    let category = '';
    if (row[catColIndex] !== undefined && row[catColIndex] !== null && String(row[catColIndex]).trim() !== '') {
      category = String(row[catColIndex]).trim();
    } else if (row[4] !== undefined && row[4] !== null && String(row[4]).trim() !== '') {
      category = String(row[4]).trim();
    }

    // Brand / Merk (Merek)
    let brand = '';
    if (row[brandColIndex] !== undefined && row[brandColIndex] !== null && String(row[brandColIndex]).trim() !== '') {
      brand = String(row[brandColIndex]).trim();
    } else if (row[5] !== undefined && row[5] !== null && String(row[5]).trim() !== '') {
      brand = String(row[5]).trim();
    }

    // Qty / Stok - Dynamically read from the detected Qty column!
    let qtyVal = 0;
    if (qtyColIndex >= 0 && row[qtyColIndex] !== undefined && row[qtyColIndex] !== null) {
      qtyVal = row[qtyColIndex];
    } else if (row[14] !== undefined && row[14] !== null) {
      qtyVal = row[14];
    }

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

  // Exact Shopee Mass Update column positions:
  // Kolom 1 (index 0 / A): ID Produk
  // Kolom 2 (index 1 / B): ID Variasi
  // Kolom 3 (index 2 / C): Nama Produk
  // Kolom 4 (index 3 / D): Nama Variasi
  // Kolom 5 (index 4 / E): Kode SKU
  // Kolom 6 (index 5 / F): Kode SKU / SKU Variasi
  // Kolom 10 (index 9 / J): SKU Induk

  // If the sheet has fewer rows than startRowIndex, fallback to row 2
  const effectiveStart = rows.length >= startRowIndex ? startRowIndex : 2;

  const items: BalistShopeeItem[] = [];
  for (let i = effectiveStart - 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;
    
    // Check if row has any non-empty data
    const hasData = row.some((cell) => cell !== undefined && cell !== null && String(cell).trim() !== '');
    if (!hasData) continue;

    // Kolom 5 (E, index 4): Kode SKU
    const skuCol5 = cleanSku(row[4]);
    // Kolom 6 (F, index 5): Kode SKU / SKU Variasi
    const skuCol6 = cleanSku(row[5]);
    // Kolom 2 (B, index 1): Nama Produk dari Kolom ke-2 file BALISTSHOPEE
    const productName = row[1] !== undefined && row[1] !== null && String(row[1]).trim() !== ''
      ? String(row[1]).trim()
      : (row[2] !== undefined && row[2] !== null ? String(row[2]).trim() : '');
    // Kolom 4 (D, index 3): Nama Variasi
    const variationName = row[3] ? String(row[3]).trim() : (row[4] ? String(row[4]).trim() : '');
    // Kolom 10 (J, index 9): SKU Induk
    const parentSku = cleanSku(row[9]);

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
  const rows = await fetchSheetValues(spreadsheetId, `${sheetName}!A:ZZ`, accessToken);
  return parseStockListFromRows(rows);
}

export async function loadBalistShopeeData(
  spreadsheetId: string,
  sheetName: string,
  accessToken: string
): Promise<BalistShopeeItem[]> {
  // Read Balistshopee data. Structure starts product rows at row 7.
  const rows = await fetchSheetValues(spreadsheetId, `${sheetName}!A:ZZ`, accessToken);
  return parseBalistShopeeFromRows(rows, 7);
}
