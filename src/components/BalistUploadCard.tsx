import React, { useRef, useMemo, useState } from 'react';
import {
  UploadCloud,
  FileSpreadsheet,
  CheckCircle2,
  Send,
  X,
  Layers,
  Info,
  Sliders,
  Sparkles,
  Download,
  AlertTriangle,
  ExternalLink,
  Database,
  RefreshCw,
  HardDrive,
  Cloud,
} from 'lucide-react';
import { ParsedGenericXlsx } from '../lib/excelProcessor';
import { getColumnLetter } from '../lib/sheets';

interface BalistUploadCardProps {
  parsedFile: ParsedGenericXlsx | null;
  uploadedFileName: string | null;
  sourceStartRow: number;
  onSourceStartRowChange: (row: number) => void;
  onFileUpload: (file: File) => void;
  onClearFile: () => void;
  onSelectSheetName: (sheetName: string) => void;
  onUpdateSheet: () => void;
  isUpdating: boolean;
  isAuthenticated: boolean;
  spreadsheetId: string;
  sheetName: string;
  onPromptSignIn: () => void;
  // Controls for "Stok Masuk" column update
  selectedStockColIndex: number;
  onSelectedStockColIndexChange: (col: number) => void;
  updateStockFromStockList: boolean;
  onUpdateStockFromStockListChange: (val: boolean) => void;
  unmatchedStockAction: 'zero' | 'keep';
  onUnmatchedStockActionChange: (action: 'zero' | 'keep') => void;
  matchedStockCount: number;
  // Conversion & direct download
  onConvertOfficeToGoogleSheet?: () => void;
  isConverting?: boolean;
  onDownloadUpdatedBalistXlsx?: () => void;
  isOfficeFile?: boolean;
  // Google Sheets Direct Sync mode support
  dataSourceMode?: 'upload' | 'sheets';
  onDataSourceModeChange?: (mode: 'upload' | 'sheets') => void;
  balistRowCount?: number;
  stockRowCount?: number;
  isLoadingSheets?: boolean;
  onSyncFromSheets?: () => void;
  onDirectUpdateSheetsStock?: () => void;
  isDirectUpdatingStock?: boolean;
  onDownloadBalistFromSheetsXlsx?: () => void;
}

export const BalistUploadCard: React.FC<BalistUploadCardProps> = ({
  parsedFile,
  uploadedFileName,
  sourceStartRow,
  onSourceStartRowChange,
  onFileUpload,
  onClearFile,
  onSelectSheetName,
  onUpdateSheet,
  isUpdating,
  isAuthenticated,
  spreadsheetId,
  sheetName,
  onPromptSignIn,
  selectedStockColIndex,
  onSelectedStockColIndexChange,
  updateStockFromStockList,
  onUpdateStockFromStockListChange,
  unmatchedStockAction,
  onUnmatchedStockActionChange,
  matchedStockCount,
  onConvertOfficeToGoogleSheet,
  isConverting = false,
  onDownloadUpdatedBalistXlsx,
  isOfficeFile = false,
  dataSourceMode = 'upload',
  onDataSourceModeChange,
  balistRowCount = 0,
  stockRowCount = 0,
  isLoadingSheets = false,
  onSyncFromSheets,
  onDirectUpdateSheetsStock,
  isDirectUpdatingStock = false,
  onDownloadBalistFromSheetsXlsx,
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [internalMode, setInternalMode] = useState<'upload' | 'sheets'>(dataSourceMode);

  const activeMode = onDataSourceModeChange ? dataSourceMode : internalMode;
  const setMode = (mode: 'upload' | 'sheets') => {
    setInternalMode(mode);
    onDataSourceModeChange?.(mode);
  };

  const availableColumns = useMemo(() => {
    if (!parsedFile || parsedFile.rows.length === 0) {
      // Fallback columns if parsedFile is not present
      return Array.from({ length: 15 }, (_, i) => {
        const letter = getColumnLetter(i);
        let name = '';
        if (i === 4) name = ' - SKU Induk (E)';
        if (i === 5) name = ' - Kode SKU Variasi (F)';
        if (i === 6) name = ' - Stok Masuk (G)';
        return {
          index: i,
          label: `Kolom ${i + 1} (${letter})${name}`,
        };
      });
    }
    const sampleRow = parsedFile.rows[Math.min(parsedFile.rows.length - 1, 6)] || parsedFile.rows[0] || [];
    const maxCols = Math.min(sampleRow.length, 30);

    const colNames: string[] = [];
    for (let c = 0; c < maxCols; c++) {
      let name = '';
      for (let r = 0; r < Math.min(parsedFile.rows.length, 6); r++) {
        const val = String(parsedFile.rows[r]?.[c] || '').trim();
        if (val && !val.match(/^[0-9]+$/)) {
          name = val;
          break;
        }
      }
      colNames.push(name);
    }

    return Array.from({ length: maxCols }, (_, i) => {
      const letter = getColumnLetter(i);
      const name = colNames[i] ? ` - ${colNames[i]}` : '';
      return {
        index: i,
        label: `Kolom ${i + 1} (${letter})${name}`,
      };
    });
  }, [parsedFile]);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        onFileUpload(file);
      } else {
        alert('Mohon unggah file dengan format Excel (.xlsx atau .xls)');
      }
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      onFileUpload(e.target.files[0]);
    }
  };

  const calculateRowsToSend = () => {
    if (!parsedFile) return 0;
    if (sourceStartRow > parsedFile.rows.length) return 0;
    return parsedFile.rows.length - (sourceStartRow - 1);
  };

  return (
    <div className="bg-white rounded-xl border border-stone-200 p-5 shadow-xs">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-stone-100">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded-lg bg-emerald-600 text-white flex items-center justify-center font-bold text-xs shrink-0 shadow-xs">
            1
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h2 className="text-sm sm:text-base font-semibold text-stone-900">
                Pembaruan &amp; Sinkronisasi Data BALISTSHOPEE
              </h2>
              <span className="px-2 py-0.5 rounded text-[11px] font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200">
                Langkah 1
              </span>
            </div>
            <p className="text-xs text-stone-500 mt-0.5">
              Pilih sumber data Balistshopee: unggah file Excel lokal atau sinkronkan langsung dari Google Sheet online.
            </p>
          </div>
        </div>

        {/* Protection Badge */}
        <div className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-emerald-50 text-emerald-800 text-xs border border-emerald-200 font-medium">
          <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
          <span>Mulai Baris ke-7 (Header 1-6 Utuh)</span>
        </div>
      </div>

      {/* Mode Selector Tabs: Upload Excel File vs Sync from Google Sheets */}
      <div className="mt-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 p-1.5 bg-stone-100/90 border border-stone-200 rounded-xl">
        <div className="flex items-center gap-1.5 w-full sm:w-auto">
          <button
            type="button"
            id="tab-mode-upload-excel"
            onClick={() => setMode('upload')}
            className={`flex-1 sm:flex-initial inline-flex items-center justify-center gap-2 px-3.5 py-2 rounded-lg text-xs font-semibold transition-all cursor-pointer ${
              activeMode === 'upload'
                ? 'bg-white text-emerald-800 shadow-xs border border-emerald-200'
                : 'text-stone-600 hover:text-stone-900 hover:bg-stone-200/60'
            }`}
          >
            <HardDrive className="w-4 h-4 text-emerald-600 shrink-0" />
            <span>Opsi 1: Upload File Excel (.xlsx)</span>
            {parsedFile && (
              <span className="w-2 h-2 rounded-full bg-emerald-500 inline-block" />
            )}
          </button>

          <button
            type="button"
            id="tab-mode-sync-sheets"
            onClick={() => setMode('sheets')}
            className={`flex-1 sm:flex-initial inline-flex items-center justify-center gap-2 px-3.5 py-2 rounded-lg text-xs font-semibold transition-all cursor-pointer ${
              activeMode === 'sheets'
                ? 'bg-white text-emerald-800 shadow-xs border border-emerald-200'
                : 'text-stone-600 hover:text-stone-900 hover:bg-stone-200/60'
            }`}
          >
            <Cloud className="w-4 h-4 text-emerald-600 shrink-0" />
            <span>Opsi 2: Sync dari Google Sheets</span>
            {balistRowCount > 0 && (
              <span className="px-1.5 py-0.2 rounded-full bg-emerald-100 text-emerald-800 text-[10px] font-bold">
                {balistRowCount.toLocaleString('id-ID')} baris
              </span>
            )}
          </button>
        </div>

        <div className="text-[11px] text-stone-500 px-2 flex items-center gap-1.5">
          <Info className="w-3.5 h-3.5 text-stone-400 shrink-0" />
          <span>
            {activeMode === 'upload'
              ? 'Unggah file Excel dari komputer'
              : 'Ambil langsung dari spreadsheet Google Sheets online'}
          </span>
        </div>
      </div>

      {/* MODE 1: UPLOAD FILE EXCEL */}
      {activeMode === 'upload' && (
        <>
          {!parsedFile ? (
            <div
              onDragOver={handleDragOver}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current?.click()}
              className="mt-4 border-2 border-dashed border-emerald-300 hover:border-emerald-500 hover:bg-emerald-50/30 rounded-xl p-7 text-center cursor-pointer transition-colors"
            >
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx, .xls"
                onChange={handleFileChange}
                className="hidden"
                id="file-upload-balist-input"
              />
              <div className="w-12 h-12 rounded-full bg-emerald-50 text-emerald-600 flex items-center justify-center mx-auto mb-3 border border-emerald-200">
                <UploadCloud className="w-6 h-6" />
              </div>
              <p className="text-sm font-semibold text-stone-800">
                Klik atau Tarik File Excel (.xlsx / .xls) untuk Sheet Balistshopee
              </p>
              <p className="text-xs text-stone-500 mt-1 max-w-md mx-auto">
                File ini akan dibaca baris datanya, lalu diperbarui ke Google Sheet Balistshopee mulai baris ke-7 agar format kolom 1-6 tetap utuh.
              </p>
            </div>
          ) : (
            <div className="mt-4 space-y-4">
              <div className="flex items-center justify-between p-3.5 bg-emerald-50/70 border border-emerald-200 rounded-lg">
                <div className="flex items-center gap-3 min-w-0">
                  <div className="w-10 h-10 rounded-lg bg-emerald-600 text-white flex items-center justify-center shrink-0">
                    <FileSpreadsheet className="w-5 h-5" />
                  </div>
                  <div className="min-w-0">
                    <p className="text-sm font-semibold text-stone-900 truncate">
                      {uploadedFileName}
                    </p>
                    <div className="flex flex-wrap items-center gap-2 text-xs text-stone-600 mt-0.5">
                      <span className="font-medium text-emerald-800">
                        {parsedFile.rows.length.toLocaleString('id-ID')} total baris terbaca
                      </span>
                      <span>•</span>
                      <span>{parsedFile.headers.length} kolom</span>
                      {parsedFile.sheetNames.length > 1 && (
                        <>
                          <span>•</span>
                          <span className="text-stone-500">Sheet:</span>
                          <select
                            value={parsedFile.selectedSheetName}
                            onChange={(e) => onSelectSheetName(e.target.value)}
                            className="bg-white border border-emerald-300 rounded px-1.5 py-0.5 text-xs text-stone-800 font-medium"
                          >
                            {parsedFile.sheetNames.map((s) => (
                              <option key={s} value={s}>
                                {s}
                              </option>
                            ))}
                          </select>
                        </>
                      )}
                    </div>
                  </div>
                </div>

                <button
                  type="button"
                  onClick={onClearFile}
                  className="p-1.5 text-stone-400 hover:text-rose-600 hover:bg-white rounded-lg transition-colors"
                  title="Ganti file"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              {/* Row 7 Protection Policy & Source Selector */}
              <div className="p-4 bg-emerald-50/50 border border-emerald-200 rounded-xl space-y-3">
                <div className="flex items-start gap-2.5">
                  <Info className="w-4 h-4 text-emerald-700 shrink-0 mt-0.5" />
                  <div className="text-xs">
                    <p className="font-bold text-emerald-900">
                      Target Penyimpanan: Mulai dari Baris ke-7 (<code className="bg-emerald-100/80 px-1 py-0.5 rounded font-mono">A7</code>)
                    </p>
                    <p className="text-emerald-800 mt-0.5 leading-relaxed">
                      Baris 1 s/d 6 pada Google Sheet <strong>"{sheetName}"</strong> diproteksi dan tidak akan dihapus ataupun ditimpa, sehingga nama kolom, format, dan formula struktur tetap utuh.
                    </p>
                  </div>
                </div>

                <div className="pt-2 border-t border-emerald-200/60">
                  <label className="text-xs font-semibold text-stone-700 block mb-2">
                    Pilih baris awal data dari file Excel yang diunggah:
                  </label>
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 text-xs">
                    <label
                      className={`flex items-start gap-2 p-2.5 rounded-lg border cursor-pointer transition-colors ${
                        sourceStartRow === 7
                          ? 'bg-white border-emerald-500 shadow-xs'
                          : 'bg-stone-50/80 border-stone-200 hover:bg-white'
                      }`}
                    >
                      <input
                        type="radio"
                        name="sourceRowStart"
                        value={7}
                        checked={sourceStartRow === 7}
                        onChange={() => onSourceStartRowChange(7)}
                        className="mt-0.5 text-emerald-600 focus:ring-emerald-500"
                      />
                      <div>
                        <span className="font-semibold text-stone-900 block">
                          Mulai Baris ke-7 (Default)
                        </span>
                        <span className="text-[11px] text-stone-500 leading-tight block mt-0.5">
                          Jika file Excel memiliki 6 baris judul/header.
                        </span>
                      </div>
                    </label>

                    <label
                      className={`flex items-start gap-2 p-2.5 rounded-lg border cursor-pointer transition-colors ${
                        sourceStartRow === 2
                          ? 'bg-white border-emerald-500 shadow-xs'
                          : 'bg-stone-50/80 border-stone-200 hover:bg-white'
                      }`}
                    >
                      <input
                        type="radio"
                        name="sourceRowStart"
                        value={2}
                        checked={sourceStartRow === 2}
                        onChange={() => onSourceStartRowChange(2)}
                        className="mt-0.5 text-emerald-600 focus:ring-emerald-500"
                      />
                      <div>
                        <span className="font-semibold text-stone-900 block">
                          Mulai Baris ke-2
                        </span>
                        <span className="text-[11px] text-stone-500 leading-tight block mt-0.5">
                          Jika file hanya memiliki 1 baris header kolom.
                        </span>
                      </div>
                    </label>

                    <label
                      className={`flex items-start gap-2 p-2.5 rounded-lg border cursor-pointer transition-colors ${
                        sourceStartRow === 1
                          ? 'bg-white border-emerald-500 shadow-xs'
                          : 'bg-stone-50/80 border-stone-200 hover:bg-white'
                      }`}
                    >
                      <input
                        type="radio"
                        name="sourceRowStart"
                        value={1}
                        checked={sourceStartRow === 1}
                        onChange={() => onSourceStartRowChange(1)}
                        className="mt-0.5 text-emerald-600 focus:ring-emerald-500"
                      />
                      <div>
                        <span className="font-semibold text-stone-900 block">
                          Mulai Baris ke-1 (Semua)
                        </span>
                        <span className="text-[11px] text-stone-500 leading-tight block mt-0.5">
                          Semua baris ditulis langsung ke baris 7.
                        </span>
                      </div>
                    </label>
                  </div>
                </div>
              </div>

              {/* Pengaturan Pembaruan Kolom Stok Masuk dari STOCK LIST */}
              <div className="p-4 bg-amber-50/70 border border-amber-200 rounded-xl space-y-3">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                  <div>
                    <h3 className="text-xs font-bold text-amber-950 flex items-center gap-1.5">
                      <Sliders className="w-4 h-4 text-amber-700" />
                      Pembaruan Kolom "Stok Masuk" dari Sheet STOCK LIST
                    </h3>
                    <p className="text-[11px] text-amber-800 mt-0.5">
                      Mencocokkan SKU Kolom 5 &amp; Kolom 6 ke Kolom 1 STOCK LIST, lalu otomatis mengisikan stok gudang ke kolom Stok Masuk.
                    </p>
                  </div>
                  <label className="flex items-center gap-2 cursor-pointer text-xs font-semibold text-amber-900 bg-white px-3 py-1.5 rounded-lg border border-amber-300 shadow-2xs self-start sm:self-auto">
                    <input
                      type="checkbox"
                      checked={updateStockFromStockList}
                      onChange={(e) => onUpdateStockFromStockListChange(e.target.checked)}
                      className="rounded text-emerald-600 focus:ring-emerald-500"
                    />
                    <span>Isi Stok Masuk dari STOCK LIST</span>
                  </label>
                </div>

                {updateStockFromStockList && (
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2 border-t border-amber-200/80 text-xs">
                    <div>
                      <label className="block text-[11px] font-semibold text-stone-700 mb-1">
                        Kolom Target "Stok Masuk" yang akan diisi nilainya:
                      </label>
                      <select
                        value={selectedStockColIndex}
                        onChange={(e) => onSelectedStockColIndexChange(Number(e.target.value))}
                        className="w-full bg-white border border-amber-300 rounded-lg p-2 text-xs font-medium text-stone-800 focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
                      >
                        {availableColumns.map((col) => (
                          <option key={col.index} value={col.index}>
                            {col.label}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <label className="block text-[11px] font-semibold text-stone-700 mb-1">
                        Jika SKU tidak ditemukan di STOCK LIST:
                      </label>
                      <select
                        value={unmatchedStockAction}
                        onChange={(e) => onUnmatchedStockActionChange(e.target.value as 'zero' | 'keep')}
                        className="w-full bg-white border border-amber-300 rounded-lg p-2 text-xs font-medium text-stone-800 focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
                      >
                        <option value="keep">Pertahankan Stok Asal dari File Excel (Default)</option>
                        <option value="zero">Isi Stok = 0 (Habis)</option>
                      </select>
                    </div>

                    {matchedStockCount > 0 && (
                      <div className="sm:col-span-2 flex items-center gap-2 p-2.5 bg-white/90 rounded-lg border border-amber-200 text-xs text-amber-900">
                        <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
                        <span>
                          <strong>{matchedStockCount.toLocaleString('id-ID')} produk</strong> terdeteksi cocok dengan STOCK LIST. Jumlah stoknya akan diisikan ke kolom <strong>{getColumnLetter(selectedStockColIndex)}</strong> saat memperbarui Google Sheets.
                        </span>
                      </div>
                    )}
                  </div>
                )}
              </div>

              {/* Office File Warning & 1-Click Convert to Google Spreadsheet */}
              {(isOfficeFile || spreadsheetId === '1wTchgk4-YRyQv-Sk10SZUrOooGMrC08S') && (
                <div className="p-4 bg-gradient-to-r from-amber-50 to-orange-50 border border-amber-300 rounded-xl space-y-2.5 text-xs">
                  <div className="flex items-start gap-2.5">
                    <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                    <div className="flex-1">
                      <div className="flex items-center gap-2">
                        <h4 className="font-bold text-stone-900">
                          Setup Awal 1 Kali: File Terdeteksi Sebagai Microsoft Excel (.xlsx) di Drive
                        </h4>
                        <span className="px-2 py-0.5 text-[10px] font-bold bg-amber-200 text-amber-900 rounded-full">
                          Cukup 1x Saja
                        </span>
                      </div>
                      <p className="text-stone-600 mt-1 leading-relaxed text-[11px]">
                        File spreadsheet lama (<code className="bg-white px-1 py-0.2 rounded border border-amber-200">1wTchgk4-YRyQv-Sk10SZUrOooGMrC08S</code>) berformat file Excel mentah. Konversi ini <strong>hanya dilakukan 1 kali saja seumur hidup</strong> agar menjadi Google Spreadsheet asli.
                      </p>
                      <p className="text-emerald-800 font-medium mt-1 text-[11px]">
                        ✓ <strong>Setelah 1x konversi</strong>, setiap upload berikutnya <strong>TIDAK AKAN</strong> membuat file baru. Data akan selalu menimpa sheet yang sama secara otomatis!
                      </p>
                    </div>
                  </div>

                  {onConvertOfficeToGoogleSheet && (
                    <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-amber-200/80">
                      <div className="text-[11px] text-amber-900">
                        Klik tombol ini sekali untuk mengubah file menjadi Google Spreadsheet asli:
                      </div>
                      <button
                        type="button"
                        onClick={onConvertOfficeToGoogleSheet}
                        disabled={isConverting}
                        className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold text-white bg-amber-600 hover:bg-amber-700 shadow-xs transition-colors disabled:opacity-50"
                      >
                        <Sparkles className="w-3.5 h-3.5" />
                        {isConverting ? 'Mengonversi ke Google Spreadsheet...' : '⚡ Konversi 1x ke Google Spreadsheet Resmi'}
                      </button>
                    </div>
                  )}
                </div>
              )}

              {/* Action to update Balistshopee Sheet */}
              <div className="p-4 bg-stone-50 border border-stone-200 rounded-lg space-y-3">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs">
                  <div>
                    <span className="font-semibold text-stone-800 block">
                      Perbarui Sheet "{sheetName}" (Mulai Baris 7)
                    </span>
                    <span className="text-stone-500">
                      {calculateRowsToSend().toLocaleString('id-ID')} baris data akan ditulis ke rentang{' '}
                      <code className="text-emerald-800 bg-emerald-50 px-1 py-0.5 rounded font-mono font-bold">
                        {sheetName}!A7
                      </code>
                      {updateStockFromStockList && (
                        <span className="text-emerald-700 font-semibold ml-1">
                          (dengan Stok Masuk dari STOCK LIST di Kolom {getColumnLetter(selectedStockColIndex)})
                        </span>
                      )}
                    </span>
                  </div>

                  <div className="flex flex-wrap items-center gap-2">
                    {onDownloadUpdatedBalistXlsx && (
                      <button
                        type="button"
                        onClick={onDownloadUpdatedBalistXlsx}
                        className="inline-flex items-center justify-center gap-1.5 px-3 py-2 rounded-lg text-xs font-semibold text-stone-700 bg-white border border-stone-300 hover:bg-stone-100 transition-colors shadow-2xs"
                        title="Unduh file Excel (.xlsx) langsung ke komputer dengan baris 7+ dan kolom Stok Masuk terisi"
                      >
                        <Download className="w-3.5 h-3.5 text-stone-600" />
                        Unduh Excel (.xlsx)
                      </button>
                    )}

                    {isAuthenticated ? (
                      <button
                        type="button"
                        id="btn-update-balist-to-sheets"
                        onClick={onUpdateSheet}
                        disabled={isUpdating}
                        className="inline-flex items-center justify-center gap-1.5 px-4 py-2 rounded-lg text-xs font-semibold text-white bg-emerald-700 hover:bg-emerald-800 shadow-xs transition-colors disabled:opacity-50 shrink-0"
                      >
                        <Send className="w-3.5 h-3.5" />
                        {isUpdating
                          ? 'Memperbarui Baris 7 Google Sheets...'
                          : 'Perbarui Sheet Balistshopee & Isi Stok Masuk'}
                      </button>
                    ) : (
                      <button
                        type="button"
                        onClick={onPromptSignIn}
                        className="inline-flex items-center justify-center gap-1.5 px-4 py-2 rounded-lg text-xs font-semibold text-blue-700 bg-blue-50 border border-blue-200 hover:bg-blue-100 transition-colors shrink-0"
                      >
                        Masuk Google untuk Memperbarui Sheet
                      </button>
                    )}
                  </div>
                </div>
              </div>
            </div>
          )}
        </>
      )}

      {/* MODE 2: SYNC LANGSUNG DARI GOOGLE SHEETS */}
      {activeMode === 'sheets' && (
        <div className="mt-4 space-y-4">
          {!isAuthenticated ? (
            <div className="p-6 bg-stone-50 border border-stone-200 rounded-xl text-center space-y-3">
              <div className="w-12 h-12 rounded-full bg-blue-50 text-blue-600 flex items-center justify-center mx-auto border border-blue-200">
                <Database className="w-6 h-6" />
              </div>
              <div>
                <h3 className="text-sm font-bold text-stone-900">
                  Sinkronisasi Langsung dengan Google Sheets
                </h3>
                <p className="text-xs text-stone-500 max-w-md mx-auto mt-1">
                  Hubungkan akun Google Anda untuk membaca dan menyinkronkan data langsung dari spreadsheet <strong>{sheetName}</strong> secara real-time tanpa perlu mengunggah file manual.
                </p>
              </div>
              <button
                type="button"
                onClick={onPromptSignIn}
                className="inline-flex items-center gap-2 px-5 py-2.5 rounded-lg text-xs font-semibold text-white bg-blue-600 hover:bg-blue-700 shadow-xs transition-colors"
              >
                <Cloud className="w-4 h-4" />
                <span>Masuk dengan Akun Google</span>
              </button>
            </div>
          ) : (
            <div className="space-y-4">
              {/* Google Sheets Connection Info */}
              <div className="p-4 bg-emerald-50/70 border border-emerald-200 rounded-xl">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div className="flex items-start gap-3">
                    <div className="w-10 h-10 rounded-lg bg-emerald-600 text-white flex items-center justify-center shrink-0 shadow-2xs">
                      <Database className="w-5 h-5" />
                    </div>
                    <div>
                      <div className="flex items-center gap-2">
                        <h3 className="text-sm font-bold text-stone-900">
                          Google Sheet "{sheetName}"
                        </h3>
                        <a
                          href={`https://docs.google.com/spreadsheets/d/${spreadsheetId}`}
                          target="_blank"
                          rel="noreferrer"
                          className="inline-flex items-center gap-1 text-[11px] text-emerald-700 hover:text-emerald-900 hover:underline font-medium"
                        >
                          <span>Buka di Google Sheets</span>
                          <ExternalLink className="w-3 h-3" />
                        </a>
                      </div>
                      <div className="flex flex-wrap items-center gap-2 text-xs text-stone-600 mt-1 font-mono">
                        <span className="text-stone-500">ID:</span>
                        <span className="bg-white px-1.5 py-0.5 rounded border border-emerald-200 text-stone-800 text-[11px] truncate max-w-[200px]">
                          {spreadsheetId}
                        </span>
                        <span>•</span>
                        <span className="font-semibold text-emerald-900">
                          {balistRowCount > 0
                            ? `${balistRowCount.toLocaleString('id-ID')} baris data termuat`
                            : 'Belum ada data termuat'}
                        </span>
                      </div>
                    </div>
                  </div>

                  <div className="flex items-center gap-2">
                    {onSyncFromSheets && (
                      <button
                        type="button"
                        id="btn-sync-balist-from-sheets"
                        onClick={onSyncFromSheets}
                        disabled={isLoadingSheets}
                        className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-xs font-semibold text-emerald-900 bg-white border border-emerald-300 hover:bg-emerald-100/60 shadow-2xs transition-colors disabled:opacity-50"
                      >
                        <RefreshCw className={`w-3.5 h-3.5 text-emerald-700 ${isLoadingSheets ? 'animate-spin' : ''}`} />
                        <span>{isLoadingSheets ? 'Menyinkronkan...' : 'Tarik Data dari Sheets'}</span>
                      </button>
                    )}
                  </div>
                </div>
              </div>

              {/* Stock Update Mapping Settings in Sheets Mode */}
              <div className="p-4 bg-amber-50/70 border border-amber-200 rounded-xl space-y-3">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                  <div>
                    <h3 className="text-xs font-bold text-amber-950 flex items-center gap-1.5">
                      <Sliders className="w-4 h-4 text-amber-700" />
                      Pembaruan Kolom "Stok Masuk" dari Sheet STOCK LIST
                    </h3>
                    <p className="text-[11px] text-amber-800 mt-0.5">
                      Mencocokkan SKU Kolom 5 &amp; 6 dari Google Sheet Balistshopee ke STOCK LIST ({stockRowCount.toLocaleString('id-ID')} produk gudang).
                    </p>
                  </div>
                  <label className="flex items-center gap-2 cursor-pointer text-xs font-semibold text-amber-900 bg-white px-3 py-1.5 rounded-lg border border-amber-300 shadow-2xs self-start sm:self-auto">
                    <input
                      type="checkbox"
                      checked={updateStockFromStockList}
                      onChange={(e) => onUpdateStockFromStockListChange(e.target.checked)}
                      className="rounded text-emerald-600 focus:ring-emerald-500"
                    />
                    <span>Isi Stok Masuk dari STOCK LIST</span>
                  </label>
                </div>

                {updateStockFromStockList && (
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2 border-t border-amber-200/80 text-xs">
                    <div>
                      <label className="block text-[11px] font-semibold text-stone-700 mb-1">
                        Kolom Target "Stok Masuk" pada Sheet Balistshopee:
                      </label>
                      <select
                        value={selectedStockColIndex}
                        onChange={(e) => onSelectedStockColIndexChange(Number(e.target.value))}
                        className="w-full bg-white border border-amber-300 rounded-lg p-2 text-xs font-medium text-stone-800 focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
                      >
                        {availableColumns.map((col) => (
                          <option key={col.index} value={col.index}>
                            {col.label}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <label className="block text-[11px] font-semibold text-stone-700 mb-1">
                        Jika SKU tidak ditemukan di STOCK LIST:
                      </label>
                      <select
                        value={unmatchedStockAction}
                        onChange={(e) => onUnmatchedStockActionChange(e.target.value as 'zero' | 'keep')}
                        className="w-full bg-white border border-amber-300 rounded-lg p-2 text-xs font-medium text-stone-800 focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
                      >
                        <option value="keep">Pertahankan Nilai Asal dari Sheet (Default)</option>
                        <option value="zero">Isi Stok = 0 (Habis)</option>
                      </select>
                    </div>

                    {matchedStockCount > 0 && (
                      <div className="sm:col-span-2 flex items-center gap-2 p-2.5 bg-white/90 rounded-lg border border-amber-200 text-xs text-amber-900">
                        <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
                        <span>
                          <strong>{matchedStockCount.toLocaleString('id-ID')} produk</strong> cocok dengan STOCK LIST. Jumlah stoknya akan diisikan ke kolom <strong>{getColumnLetter(selectedStockColIndex)}</strong> (Stok Masuk).
                        </span>
                      </div>
                    )}
                  </div>
                )}
              </div>

              {/* Action Bar for Sheets Mode */}
              <div className="p-4 bg-stone-50 border border-stone-200 rounded-lg space-y-3">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs">
                  <div>
                    <span className="font-semibold text-stone-800 block">
                      Aksi Data Sinkronisasi Google Sheets
                    </span>
                    <span className="text-stone-500">
                      {balistRowCount.toLocaleString('id-ID')} baris data siap diproses atau diekspor ke Excel dengan format resmi Shopee.
                    </span>
                  </div>

                  <div className="flex flex-wrap items-center gap-2">
                    {onDownloadBalistFromSheetsXlsx && (
                      <button
                        type="button"
                        id="btn-download-balist-sheets-xlsx"
                        onClick={onDownloadBalistFromSheetsXlsx}
                        disabled={balistRowCount === 0}
                        className="inline-flex items-center justify-center gap-1.5 px-3.5 py-2 rounded-lg text-xs font-semibold text-stone-700 bg-white border border-stone-300 hover:bg-stone-100 transition-colors shadow-2xs disabled:opacity-50 cursor-pointer"
                        title="Download file Excel (.xlsx) dengan format Shopee dan nama file otomatis berdasarkan tanggal"
                      >
                        <Download className="w-3.5 h-3.5 text-stone-600" />
                        Download XLSX Format Shopee
                      </button>
                    )}

                    {onDirectUpdateSheetsStock && (
                      <button
                        type="button"
                        id="btn-direct-update-sheets-stock"
                        onClick={onDirectUpdateSheetsStock}
                        disabled={isDirectUpdatingStock || balistRowCount === 0}
                        className="inline-flex items-center justify-center gap-1.5 px-4 py-2 rounded-lg text-xs font-semibold text-white bg-emerald-700 hover:bg-emerald-800 shadow-xs transition-colors disabled:opacity-50 shrink-0 cursor-pointer"
                      >
                        <Send className="w-3.5 h-3.5" />
                        {isDirectUpdatingStock
                          ? 'Memperbarui Stok Masuk...'
                          : 'Perbarui Kolom Stok ke Google Sheets'}
                      </button>
                    )}
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
};


