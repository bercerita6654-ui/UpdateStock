import * as XLSX from 'xlsx';
import { StockListItem, BalistShopeeItem } from '../types';

export const cleanSku = (val: unknown): string => {
  if (val === null || val === undefined) return '';
  let str = String(val).trim();
  // Remove leading and trailing double or single quotes and backticks
  str = str.replace(/^["'`]+|["'`]+$/g, '').trim();
  return str;
};

export const parseNumber = (val: unknown): number => {
  if (val === null || val === undefined) return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  const clean = String(val).replace(/[^0-9.-]+/g, '');
  const num = parseFloat(clean);
  return isNaN(num) ? 0 : num;
};

export const isOfficeFileError = (err: any): boolean => {
  const msg = (typeof err === 'string' ? err : err?.message || '').toLowerCase();
  return (
    msg.includes('not be an office file') ||
    msg.includes('office file') ||
    msg.includes('not supported for this document')
  );
};

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
  if (options?.rawWorkbookData) {
    try {
      const wb = XLSX.read(options.rawWorkbookData, {
        type: 'array',
        cellStyles: true,
        cellDates: true,
      });

      let targetSheetName = wb.SheetNames[0] || 'Balistshopee';
      if (options.sheetName) {
        const found = wb.SheetNames.find(
          (s) => s.trim().toLowerCase() === options.sheetName!.trim().toLowerCase()
        );
        if (found) targetSheetName = found;
      }

      let ws = wb.Sheets[targetSheetName];
      if (!ws) {
        ws = XLSX.utils.aoa_to_sheet(originalRows);
        wb.Sheets[targetSheetName] = ws;
      } else {
        const startRow = (options.startRowIndex && options.startRowIndex > 0) ? options.startRowIndex : 7;
        const startRow0Based = startRow - 1;

        // Clear existing data rows from startRow downwards
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

        const dataToAdd = options.updatedRowValues || (originalRows.length >= startRow ? originalRows.slice(startRow0Based) : originalRows);
        XLSX.utils.sheet_add_aoa(ws, dataToAdd, { origin: `A${startRow}` });

        // Update range ref
        const totalRows = startRow0Based + dataToAdd.length;
        const maxCol = dataToAdd.reduce((max, r) => Math.max(max, r?.length || 0), 26);
        ws['!ref'] = `A1:${getColumnLetter(Math.max(0, maxCol - 1))}${Math.max(1, totalRows)}`;
      }

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
      return;
    } catch (err) {
      console.warn('Could not preserve raw workbook styles, falling back to clean aoa sheet:', err);
    }
  }

  const ws = XLSX.utils.aoa_to_sheet(originalRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, options?.sheetName || 'Balistshopee');
  XLSX.writeFile(wb, fileName);
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

    // If it's an Office file error, attempt fallback to Google Drive API
    if (isOfficeFileError(message)) {
      try {
        console.warn(
          `Document ${spreadsheetId} is an Office file. Attempting Google Drive API fallback...`
        );
        return await fetchOfficeFileRows(spreadsheetId, sheetName, accessToken);
      } catch (driveErr: any) {
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

  const items: StockListItem[] = [];
  // Skip header (row 0)
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const code = cleanSku(row[0]);
    if (!code) continue;

    const barcode = cleanSku(row[1]);
    const description = row[2] ? String(row[2]).trim() : '';
    const unit = row[3] ? String(row[3]).trim() : '';
    const category = row[4] ? String(row[4]).trim() : '';
    // Col 15 is index 14
    const qtyVal = row[14];
    const qty = parseNumber(qtyVal);

    items.push({
      code,
      barcode,
      description,
      unit,
      category,
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

  // If the sheet has fewer rows than startRowIndex, fallback to row 2
  const effectiveStart = rows.length >= startRowIndex ? startRowIndex : 2;

  const items: BalistShopeeItem[] = [];
  for (let i = effectiveStart - 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const skuCol5 = cleanSku(row[4]); // Column 5 (E)
    const skuCol6 = cleanSku(row[5]); // Column 6 (F)

    if (!skuCol5 && !skuCol6) continue;

    items.push({
      skuCol5,
      skuCol6,
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
  const rows = await fetchSheetValues(spreadsheetId, `${sheetName}!A:O`, accessToken);
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
