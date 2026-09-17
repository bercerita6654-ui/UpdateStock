import React, { useRef, useMemo } from 'react';
import { UploadCloud, FileSpreadsheet, Download, RefreshCw, X, SlidersHorizontal, Check, EyeOff, Store } from 'lucide-react';
import { ParsedShopeeSheet, detectStoreFromFilename } from '../lib/excelProcessor';

interface UploadSectionProps {
  parsedFile: ParsedShopeeSheet | null;
  uploadedFileName?: string | null;
  isLoading: boolean;
  selectedSkuCol: number;
  selectedStockCol: number;
  unmatchedAction: 'keep' | 'zero';
  onFileUpload: (file: File) => void;
  onClearFile: () => void;
  onChangeSkuCol: (colIndex: number) => void;
  onChangeStockCol: (colIndex: number) => void;
  onChangeUnmatchedAction: (action: 'keep' | 'zero') => void;
  onDownloadSample: () => void;
  onHide?: () => void;
}

export const UploadSection: React.FC<UploadSectionProps> = ({
  parsedFile,
  uploadedFileName,
  isLoading,
  selectedSkuCol,
  selectedStockCol,
  unmatchedAction,
  onFileUpload,
  onClearFile,
  onChangeSkuCol,
  onChangeStockCol,
  onChangeUnmatchedAction,
  onDownloadSample,
  onHide,
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);

  const detectedStore = useMemo(() => {
    const raw = uploadedFileName || parsedFile?.fileName || '';
    return detectStoreFromFilename(raw);
  }, [uploadedFileName, parsedFile?.fileName]);

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

  return (
    <div className="bg-white rounded-xl border border-stone-200 p-5 shadow-xs">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-stone-100">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded-lg bg-orange-600 text-white flex items-center justify-center font-bold text-xs shrink-0 shadow-xs">
            2
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h2 className="text-sm sm:text-base font-semibold text-stone-900">
                Tombol 2: Upload File XLSX Shopee (Mass Update Seller Centre)
              </h2>
              <span className="px-2 py-0.5 rounded text-[11px] font-semibold bg-orange-50 text-orange-700 border border-orange-200">
                Langkah 2
              </span>
            </div>
            <p className="text-xs text-stone-500 mt-0.5">
              Unggah file mass update stok dari Shopee Seller Centre (.xlsx) untuk dicocokkan stoknya
            </p>
          </div>
        </div>
        <div className="flex items-center gap-2 w-fit shrink-0">
          <button
            type="button"
            id="btn-download-sample"
            onClick={onDownloadSample}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium text-stone-700 bg-stone-100 hover:bg-stone-200 transition-colors"
          >
            <Download className="w-3.5 h-3.5" />
            Unduh Format Contoh
          </button>
          {onHide && (
            <button
              type="button"
              id="btn-hide-shopee-section"
              onClick={onHide}
              title="Sembunyikan menu ini"
              className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium text-stone-500 hover:text-stone-800 hover:bg-stone-100 border border-stone-200 transition-colors"
            >
              <EyeOff className="w-3.5 h-3.5" />
              <span>Sembunyikan</span>
            </button>
          )}
        </div>
      </div>

      {!parsedFile ? (
        <div
          onDragOver={handleDragOver}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
          className="mt-4 border-2 border-dashed border-stone-300 hover:border-orange-500 hover:bg-orange-50/20 rounded-xl p-8 text-center cursor-pointer transition-colors"
        >
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx, .xls"
            onChange={handleFileChange}
            className="hidden"
            id="file-upload-input"
          />
          <div className="w-12 h-12 rounded-full bg-orange-50 text-orange-600 flex items-center justify-center mx-auto mb-3 border border-orange-200">
            <UploadCloud className="w-6 h-6" />
          </div>
          <p className="text-sm font-medium text-stone-800">
            Tarik dan lepas file Excel Shopee di sini, atau{' '}
            <span className="text-orange-600 font-semibold underline">pilih file</span>
          </p>
          <p className="text-xs text-stone-500 mt-1">
            Mendukung format file .xlsx dari Mass Update Shopee Seller Centre
          </p>
        </div>
      ) : (
        <div className="mt-4 space-y-4">
          {/* File Card info */}
          <div className="flex items-center justify-between p-3.5 bg-orange-50/60 border border-orange-200 rounded-lg">
            <div className="flex items-center gap-3 min-w-0">
              <div className="w-10 h-10 rounded-lg bg-orange-600 text-white flex items-center justify-center shrink-0">
                <FileSpreadsheet className="w-5 h-5" />
              </div>
              <div className="min-w-0">
                <p className="text-sm font-semibold text-stone-900 truncate">
                  {uploadedFileName || parsedFile.fileName || parsedFile.sheetName} ({parsedFile.rows.length - 1} baris produk)
                </p>
                <p className="text-xs text-stone-600 mt-0.5">
                  Header terdeteksi di baris ke-{parsedFile.headerRowIndex + 1}
                </p>
                {detectedStore.storeName && (
                  <div className="flex flex-wrap items-center gap-1.5 mt-1.5">
                    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-[11px] font-bold bg-orange-600 text-white shadow-2xs">
                      <Store className="w-3 h-3" />
                      Toko: {detectedStore.storeName}
                    </span>
                    {detectedStore.storeCode && (
                      <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-mono bg-orange-100 text-orange-800 border border-orange-200">
                        Kode: {detectedStore.storeCode}
                      </span>
                    )}
                    <span className="text-[11px] text-orange-700 font-medium">
                      ➔ Format unduh otomatis: <strong>{detectedStore.storePrefix}</strong>
                    </span>
                  </div>
                )}
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

          {/* Mapping settings */}
          <div className="p-4 bg-stone-50 border border-stone-200 rounded-lg text-xs space-y-3">
            <div className="flex items-center gap-1.5 font-medium text-stone-800 pb-2 border-b border-stone-200">
              <SlidersHorizontal className="w-3.5 h-3.5 text-stone-500" />
              <span>Pemetaan Kolom File Shopee</span>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
              <div>
                <label className="block text-stone-600 font-medium mb-1">
                  Kolom SKU Acuan di Shopee:
                </label>
                <select
                  value={selectedSkuCol}
                  onChange={(e) => onChangeSkuCol(Number(e.target.value))}
                  className="w-full bg-white border border-stone-300 rounded-md px-2.5 py-1.5 text-stone-800 focus:outline-hidden focus:ring-1 focus:ring-orange-500"
                >
                  {parsedFile.headers.map((h, i) => (
                    <option key={i} value={i}>
                      Kolom {i + 1}: {h || `(Tanpa Judul)`}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-stone-600 font-medium mb-1">
                  Kolom Stok di Shopee:
                </label>
                <select
                  value={selectedStockCol}
                  onChange={(e) => onChangeStockCol(Number(e.target.value))}
                  className="w-full bg-white border border-stone-300 rounded-md px-2.5 py-1.5 text-stone-800 focus:outline-hidden focus:ring-1 focus:ring-orange-500"
                >
                  {parsedFile.headers.map((h, i) => (
                    <option key={i} value={i}>
                      Kolom {i + 1}: {h || `(Tanpa Judul)`}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-stone-600 font-medium mb-1">
                  Jika SKU Tidak Ditemukan:
                </label>
                <select
                  value={unmatchedAction}
                  onChange={(e) => onChangeUnmatchedAction(e.target.value as 'keep' | 'zero')}
                  className="w-full bg-white border border-stone-300 rounded-md px-2.5 py-1.5 text-stone-800 focus:outline-hidden focus:ring-1 focus:ring-orange-500"
                >
                  <option value="keep">Biarkan Stok Asli (Tidak Diubah)</option>
                  <option value="zero">Ubah Stok Menjadi 0 (Kosongkan)</option>
                </select>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
