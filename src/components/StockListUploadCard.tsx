import React, { useRef, useState } from 'react';
import {
  UploadCloud,
  FileSpreadsheet,
  CheckCircle2,
  X,
  Layers,
  Info,
  Download,
  Database,
  RefreshCw,
  Eye,
  Check,
  PackageCheck,
  Tag,
  FolderOpen,
} from 'lucide-react';
import { StockListItem } from '../types';
import { downloadBlob, createSampleStockListFile } from '../lib/excelProcessor';

interface StockListUploadCardProps {
  stockList: StockListItem[];
  uploadedStockFileName: string | null;
  uploadedStockFileSize?: number | null;
  dataSource: 'file' | 'sheets';
  lastUpdated: Date | null;
  onFileUpload: (file: File) => void;
  onClearFile: () => void;
  onRefreshGoogleSheets?: () => void;
  isLoadingSheets?: boolean;
  isAuthenticated?: boolean;
  stockSheetName?: string;
}

export const StockListUploadCard: React.FC<StockListUploadCardProps> = ({
  stockList,
  uploadedStockFileName,
  uploadedStockFileSize,
  dataSource,
  lastUpdated,
  onFileUpload,
  onClearFile,
  onRefreshGoogleSheets,
  isLoadingSheets = false,
  isAuthenticated = false,
  stockSheetName = 'STOCK LIST',
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [showPreviewModal, setShowPreviewModal] = useState(false);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      if (
        file.name.endsWith('.xlsx') ||
        file.name.endsWith('.xls') ||
        file.name.endsWith('.csv')
      ) {
        onFileUpload(file);
      } else {
        alert('Mohon unggah file dengan format Excel (.xlsx / .xls) atau .csv');
      }
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      onFileUpload(e.target.files[0]);
    }
  };

  const handleDownloadSampleStock = () => {
    try {
      const data = createSampleStockListFile();
      downloadBlob(data, 'STOCK_LIST_Sample.xlsx');
    } catch (e) {
      console.error(e);
    }
  };

  const formatFileSize = (bytes?: number | null) => {
    if (!bytes) return '';
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  return (
    <div className="bg-white rounded-xl border border-stone-200 p-5 shadow-xs">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-stone-100">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded-lg bg-blue-600 text-white flex items-center justify-center font-bold text-xs shrink-0 shadow-xs">
            <Database className="w-4 h-4" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h2 className="text-sm sm:text-base font-semibold text-stone-900">
                Database STOCK LIST (Stok Gudang)
              </h2>
              <span
                className={`px-2 py-0.5 rounded text-[11px] font-semibold border ${
                  dataSource === 'file'
                    ? 'bg-blue-50 text-blue-700 border-blue-200'
                    : 'bg-emerald-50 text-emerald-700 border-emerald-200'
                }`}
              >
                {dataSource === 'file' ? 'File Excel Terunggah' : 'Google Sheets Cloud'}
              </span>
            </div>
            <p className="text-xs text-stone-500 mt-0.5">
              Sumber referensi stok gudang untuk mencocokkan SKU produk Shopee &amp; Balistshopee.
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={handleDownloadSampleStock}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium text-stone-700 bg-stone-100 hover:bg-stone-200 transition-colors"
            title="Unduh file contoh template STOCK LIST Excel"
          >
            <Download className="w-3.5 h-3.5" />
            <span>Contoh Format .xlsx</span>
          </button>
        </div>
      </div>

      {stockList.length === 0 ? (
        <div
          onDragOver={handleDragOver}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
          className="mt-4 border-2 border-dashed border-blue-300 hover:border-blue-500 hover:bg-blue-50/30 rounded-xl p-7 text-center cursor-pointer transition-colors"
        >
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx, .xls, .csv"
            onChange={handleFileChange}
            className="hidden"
            id="file-upload-stock-input"
          />
          <div className="w-12 h-12 rounded-full bg-blue-50 text-blue-600 flex items-center justify-center mx-auto mb-3 border border-blue-200">
            <UploadCloud className="w-6 h-6" />
          </div>
          <p className="text-sm font-semibold text-stone-800">
            Klik atau Tarik File Excel STOCK LIST (.xlsx / .xls / .csv) di Sini
          </p>
          <p className="text-xs text-stone-500 mt-1 max-w-md mx-auto">
            Sistem otomatis mendeteksi kolom <strong>Kode SKU</strong>, <strong>Nama Produk</strong>, dan <strong>Stok/Saldo Akhir</strong> tanpa perlu login Google.
          </p>
        </div>
      ) : (
        <div className="mt-4 space-y-3">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between p-3.5 bg-blue-50/70 border border-blue-200 rounded-lg gap-3">
            <div className="flex items-center gap-3 min-w-0">
              <div className="w-10 h-10 rounded-lg bg-blue-600 text-white flex items-center justify-center shrink-0 shadow-2xs">
                <FileSpreadsheet className="w-5 h-5" />
              </div>
              <div className="min-w-0">
                <div className="flex items-center gap-2">
                  <p className="text-sm font-bold text-stone-900 truncate">
                    {dataSource === 'file'
                      ? uploadedStockFileName || 'STOCK_LIST_Uploaded.xlsx'
                      : `Google Sheet: ${stockSheetName}`}
                  </p>
                  {uploadedStockFileSize ? (
                    <span className="text-[11px] text-stone-500 font-mono">
                      ({formatFileSize(uploadedStockFileSize)})
                    </span>
                  ) : null}
                </div>
                <div className="flex flex-wrap items-center gap-2 text-xs text-stone-600 mt-0.5">
                  <span className="font-semibold text-blue-800 inline-flex items-center gap-1">
                    <CheckCircle2 className="w-3.5 h-3.5 text-blue-600" />
                    {stockList.length.toLocaleString('id-ID')} produk aktif siap dicocokkan
                  </span>
                  {lastUpdated && (
                    <>
                      <span>•</span>
                      <span className="text-stone-500">
                        Dimuat: {lastUpdated.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })}
                      </span>
                    </>
                  )}
                </div>
              </div>
            </div>

            <div className="flex items-center gap-2 self-end sm:self-auto">
              <button
                type="button"
                onClick={() => setShowPreviewModal(true)}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium text-blue-800 bg-white hover:bg-blue-100 border border-blue-200 transition-colors shadow-2xs"
                title="Lihat sampel data produk yang terbaca"
              >
                <Eye className="w-3.5 h-3.5 text-blue-600" />
                <span>Pratinjau Data</span>
              </button>

              <button
                type="button"
                onClick={() => fileInputRef.current?.click()}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium text-stone-700 bg-white hover:bg-stone-100 border border-stone-200 transition-colors shadow-2xs"
                title="Ganti file STOCK LIST dengan file Excel lain"
              >
                <FolderOpen className="w-3.5 h-3.5 text-stone-600" />
                <span>Ganti File</span>
              </button>

              {dataSource === 'file' && (
                <button
                  type="button"
                  onClick={onClearFile}
                  className="p-1.5 text-stone-400 hover:text-rose-600 hover:bg-white rounded-lg transition-colors"
                  title="Hapus file unggahan ini"
                >
                  <X className="w-4 h-4" />
                </button>
              )}

              {dataSource === 'file' && isAuthenticated && onRefreshGoogleSheets && (
                <button
                  type="button"
                  onClick={onRefreshGoogleSheets}
                  disabled={isLoadingSheets}
                  className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium text-emerald-800 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 transition-colors"
                  title="Beralih dan muat data terbaru dari Google Sheets"
                >
                  <RefreshCw className={`w-3 h-3 ${isLoadingSheets ? 'animate-spin' : ''}`} />
                  <span>Pakai Google Sheets</span>
                </button>
              )}
            </div>
          </div>

          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx, .xls, .csv"
            onChange={handleFileChange}
            className="hidden"
            id="file-upload-stock-input-replace"
          />
        </div>
      )}

      {/* Pratinjau Modal STOCK LIST */}
      {showPreviewModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-xs animate-in fade-in duration-150">
          <div className="bg-white rounded-2xl border border-stone-200 shadow-2xl max-w-3xl w-full max-h-[85vh] flex flex-col overflow-hidden">
            <div className="p-4 border-b border-stone-100 flex items-center justify-between bg-stone-50">
              <div className="flex items-center gap-2">
                <FileSpreadsheet className="w-5 h-5 text-blue-600" />
                <div>
                  <h3 className="text-sm font-bold text-stone-900">
                    Pratinjau Data STOCK LIST ({stockList.length.toLocaleString('id-ID')} Produk)
                  </h3>
                  <p className="text-[11px] text-stone-500">
                    Menampilkan 10 baris pertama data yang berhasil diparsing
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setShowPreviewModal(false)}
                className="p-1.5 text-stone-400 hover:text-stone-700 hover:bg-stone-200/60 rounded-lg transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-4 overflow-auto flex-1">
              <div className="border border-stone-200 rounded-lg overflow-hidden">
                <table className="w-full text-xs text-left">
                  <thead className="bg-stone-100 text-stone-700 font-semibold border-b border-stone-200">
                    <tr>
                      <th className="py-2 px-3">No</th>
                      <th className="py-2 px-3">Kode SKU</th>
                      <th className="py-2 px-3">Barcode</th>
                      <th className="py-2 px-3">Nama Produk</th>
                      <th className="py-2 px-3">Kategori</th>
                      <th className="py-2 px-3">Merk</th>
                      <th className="py-2 px-3 text-right">Stok Gudang</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-stone-100">
                    {stockList.slice(0, 10).map((item, idx) => (
                      <tr key={idx} className="hover:bg-blue-50/40">
                        <td className="py-2 px-3 text-stone-400">{idx + 1}</td>
                        <td className="py-2 px-3 font-mono font-semibold text-blue-900">{item.code}</td>
                        <td className="py-2 px-3 font-mono text-stone-600">{item.barcode || '-'}</td>
                        <td className="py-2 px-3 font-medium text-stone-800">{item.description || '-'}</td>
                        <td className="py-2 px-3 text-stone-600">{item.category || '-'}</td>
                        <td className="py-2 px-3 text-stone-600">{item.brand || '-'}</td>
                        <td className="py-2 px-3 text-right font-mono font-bold text-emerald-700">
                          {item.qty.toLocaleString('id-ID')}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="p-3 border-t border-stone-100 bg-stone-50 flex justify-end">
              <button
                type="button"
                onClick={() => setShowPreviewModal(false)}
                className="px-4 py-1.5 text-xs font-semibold text-stone-700 bg-white hover:bg-stone-100 border border-stone-200 rounded-lg transition-colors"
              >
                Tutup
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
