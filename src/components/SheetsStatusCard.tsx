import React, { useRef, useState } from 'react';
import {
  Database,
  RefreshCw,
  CheckCircle2,
  AlertTriangle,
  ExternalLink,
  ShieldCheck,
  FileSpreadsheet,
  Upload,
  HelpCircle,
  Check,
  Copy,
  Sparkles,
  EyeOff,
  History,
} from 'lucide-react';
import { SheetsConfig } from '../types';
import { isAuthError } from '../lib/sheets';

interface SheetsStatusCardProps {
  config: SheetsConfig;
  isLoading: boolean;
  stockCount: number;
  balistCount: number;
  lastLoaded: Date | null;
  error: string | null;
  isAuthenticated: boolean;
  onRefresh: () => void;
  onPromptSignIn: () => void;
  onRefreshStockList?: () => void;
  isRefreshingStockList?: boolean;
  onUploadStockListFile?: (file: File) => void;
  onOpenSettings?: () => void;
  onConvertBalistToGoogleSheet?: () => void;
  isConvertingBalist?: boolean;
  onHide?: () => void;
  onToggleActivityLogs?: () => void;
  showActivityLogs?: boolean;
}

export const SheetsStatusCard: React.FC<SheetsStatusCardProps> = ({
  config,
  isLoading,
  stockCount,
  balistCount,
  lastLoaded,
  error,
  isAuthenticated,
  onRefresh,
  onPromptSignIn,
  onRefreshStockList,
  isRefreshingStockList = false,
  onUploadStockListFile,
  onOpenSettings,
  onConvertBalistToGoogleSheet,
  isConvertingBalist = false,
  onHide,
  onToggleActivityLogs,
  showActivityLogs = false,
}) => {
  const stockFileInputRef = useRef<HTMLInputElement>(null);
  const [showOfficeGuide, setShowOfficeGuide] = useState(true);
  const [copiedId, setCopiedId] = useState<string | null>(null);

  const isOfficeError =
    error &&
    (error.toLowerCase().includes('office file') ||
      error.toLowerCase().includes('not supported for this document') ||
      error.toLowerCase().includes('excel (.xlsx)'));

  const handleCopy = (text: string, id: string) => {
    navigator.clipboard.writeText(text);
    setCopiedId(id);
    setTimeout(() => setCopiedId(null), 2000);
  };

  const handleStockFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0] && onUploadStockListFile) {
      onUploadStockListFile(e.target.files[0]);
    }
  };

  return (
    <div className="bg-white rounded-xl border border-stone-200 p-5 shadow-xs space-y-4">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-stone-100">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 rounded-lg bg-emerald-50 text-emerald-700 flex items-center justify-center border border-emerald-200 shrink-0">
            <Database className="w-5 h-5" />
          </div>
          <div>
            <h2 className="text-sm sm:text-base font-semibold text-stone-900">
              Koneksi Sumber Data Google Sheets &amp; Excel
            </h2>
            <p className="text-xs text-stone-500">
              Sinkronisasi data stok gudang &amp; referensi SKU dengan Google Drive / Google Sheets
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          {isAuthenticated ? (
            <button
              type="button"
              id="btn-refresh-sheets"
              onClick={onRefresh}
              disabled={isLoading}
              className="inline-flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-medium text-emerald-800 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 transition-colors disabled:opacity-50"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isLoading ? 'animate-spin' : ''}`} />
              {isLoading ? 'Memuat Data...' : 'Muat Ulang Sheets'}
            </button>
          ) : (
            <button
              type="button"
              id="btn-connect-sheets-auth"
              onClick={onPromptSignIn}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium text-blue-700 bg-blue-50 hover:bg-blue-100 border border-blue-200 transition-colors"
            >
              <ShieldCheck className="w-3.5 h-3.5" />
              Hubungkan Akun Google
            </button>
          )}

          {onToggleActivityLogs && (
            <button
              type="button"
              id="btn-toggle-activity-logs-card"
              onClick={onToggleActivityLogs}
              title={showActivityLogs ? 'Sembunyikan panel log riwayat' : 'Tampilkan panel log riwayat'}
              className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-medium border transition-colors ${
                showActivityLogs
                  ? 'bg-stone-800 text-white border-stone-800 hover:bg-stone-900'
                  : 'text-stone-700 bg-stone-100 hover:bg-stone-200 border-stone-200'
              }`}
            >
              <History className="w-3.5 h-3.5" />
              <span>{showActivityLogs ? 'Tutup Log' : 'Log Riwayat'}</span>
            </button>
          )}

          {onHide && (
            <button
              type="button"
              id="btn-hide-sheets-status-card"
              onClick={onHide}
              title="Sembunyikan kartu status koneksi ini"
              className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium text-stone-500 hover:text-stone-800 hover:bg-stone-100 border border-stone-200 transition-colors"
            >
              <EyeOff className="w-3.5 h-3.5" />
              <span>Sembunyikan</span>
            </button>
          )}
        </div>
      </div>

      {/* OFFICE FILE NOTICE & FIX GUIDE */}
      {isOfficeError && showOfficeGuide && (
        <div className="p-4 bg-amber-50/90 border border-amber-300 rounded-xl space-y-3 text-xs">
          <div className="flex items-start justify-between gap-3">
            <div className="flex items-start gap-2.5">
              <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />
              <div>
                <h3 className="font-bold text-amber-900 text-sm">
                  File Terdeteksi Berformat Microsoft Excel (.xlsx) di Google Drive
                </h3>
                <p className="text-amber-800 mt-1 leading-relaxed">
                  Google Sheets API hanya dapat membaca dan menulis dokumen bertipe{' '}
                  <strong>Google Spreadsheet asli</strong>. Dokumen yang Anda gunakan saat ini diunggah sebagai file Excel mentah (.xlsx).
                </p>
              </div>
            </div>
            <button
              type="button"
              onClick={() => setShowOfficeGuide(false)}
              className="text-amber-500 hover:text-amber-800 text-xs px-2 py-1"
            >
              Tutup
            </button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-3 pt-2">
            {/* Solution 1: 1-Click Automated Conversion or Manual in Google Drive */}
            <div className="p-3 bg-white border border-amber-200 rounded-lg space-y-2">
              <span className="font-bold text-stone-900 block flex items-center gap-1.5">
                <span className="w-4 h-4 rounded-full bg-emerald-600 text-white flex items-center justify-center text-[10px]">
                  1
                </span>
                Solusi Cepat (Cukup 1x Saja): Konversi ke Google Spreadsheet Asli
              </span>
              <p className="text-stone-600 text-[11px] leading-relaxed">
                Hanya perlu dijalankan <strong>1 kali di awal</strong>. Setelah itu, untuk setiap upload rutin berikutnya, data akan otomatis <strong>menimpa file yang sama</strong> tanpa membuat file baru.
              </p>

              {onConvertBalistToGoogleSheet && (
                <button
                  type="button"
                  onClick={onConvertBalistToGoogleSheet}
                  disabled={isConvertingBalist}
                  className="w-full inline-flex items-center justify-center gap-1.5 px-3 py-2 rounded-lg text-xs font-bold text-white bg-emerald-700 hover:bg-emerald-800 shadow-xs transition-colors disabled:opacity-50"
                >
                  <Sparkles className="w-3.5 h-3.5" />
                  {isConvertingBalist
                    ? 'Sedang Mengonversi ke Google Spreadsheet...'
                    : '⚡ Konversi 1x ke Google Spreadsheet (Hanya Sekali)'}
                </button>
              )}

              <div className="flex flex-wrap gap-2 pt-1">
                <a
                  href={`https://docs.google.com/spreadsheets/d/${config.balistSpreadsheetId}`}
                  target="_blank"
                  rel="noreferrer"
                  className="inline-flex items-center gap-1 px-2.5 py-1 rounded bg-stone-100 hover:bg-stone-200 text-stone-700 font-medium text-[11px]"
                >
                  Buka Balistshopee <ExternalLink className="w-3 h-3" />
                </a>
                <a
                  href={`https://docs.google.com/spreadsheets/d/${config.stockSpreadsheetId}`}
                  target="_blank"
                  rel="noreferrer"
                  className="inline-flex items-center gap-1 px-2.5 py-1 rounded bg-stone-100 hover:bg-stone-200 text-stone-700 font-medium text-[11px]"
                >
                  Buka STOCK LIST <ExternalLink className="w-3 h-3" />
                </a>
                {onOpenSettings && (
                  <button
                    type="button"
                    onClick={onOpenSettings}
                    className="inline-flex items-center gap-1 px-2.5 py-1 rounded bg-emerald-50 hover:bg-emerald-100 text-emerald-800 font-semibold text-[11px] border border-emerald-300"
                  >
                    Ubah ID di Pengaturan
                  </button>
                )}
              </div>
            </div>

            {/* Solution 2: Local Upload */}
            <div className="p-3 bg-white border border-amber-200 rounded-lg space-y-2">
              <span className="font-bold text-stone-900 block flex items-center gap-1.5">
                <span className="w-4 h-4 rounded-full bg-blue-600 text-white flex items-center justify-center text-[10px]">
                  2
                </span>
                Solusi Instan: Gunakan File Excel (.xlsx) Langsung
              </span>
              <p className="text-stone-600 text-[11px] leading-relaxed">
                Anda tidak harus memakai Google Drive jika memiliki file Excel di komputer. Cukup upload file:
              </p>
              <div className="flex flex-col gap-1.5 pt-1">
                <button
                  type="button"
                  onClick={() => stockFileInputRef.current?.click()}
                  className="inline-flex items-center justify-center gap-1.5 px-3 py-1.5 rounded bg-blue-50 hover:bg-blue-100 text-blue-800 font-medium text-[11px] border border-blue-200"
                >
                  <Upload className="w-3 h-3" />
                  Upload STOCK LIST (.xlsx) Sekarang
                </button>
                <span className="text-[10px] text-stone-500">
                  Untuk Balistshopee, gunakan tombol <strong>"Tombol 1: Upload Excel Balistshopee"</strong> di bawah.
                </span>
              </div>
            </div>
          </div>
        </div>
      )}

      {error && !isOfficeError && (
        <div className="p-3.5 bg-amber-50 border border-amber-200 rounded-lg flex items-start gap-2.5 text-xs text-amber-900">
          <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
          <div className="flex-1 space-y-2">
            <div>
              <p className="font-semibold text-stone-900">Pemberitahuan Akses Google Sheets:</p>
              <p className="mt-0.5 text-stone-700 whitespace-pre-line leading-relaxed">{error}</p>
            </div>
            {(isAuthError(error) || !isAuthenticated) && (
              <div className="pt-1">
                <button
                  type="button"
                  id="btn-card-reconnect-google"
                  onClick={onPromptSignIn}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold text-white bg-blue-600 hover:bg-blue-700 shadow-xs transition-colors"
                >
                  <ShieldCheck className="w-3.5 h-3.5" />
                  Masuk Kembali dengan Google
                </button>
              </div>
            )}
          </div>
        </div>
      )}

      {/* Detail Cards */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        {/* Card 1: Balistshopee */}
        <div className="p-3.5 rounded-lg bg-stone-50 border border-stone-200/80">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2">
              <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
              <span className="text-xs font-semibold text-stone-800 uppercase tracking-wide">
                Sheet: {config.balistSheetName}
              </span>
            </div>
            <a
              href={`https://docs.google.com/spreadsheets/d/${config.balistSpreadsheetId}`}
              target="_blank"
              rel="noreferrer"
              className="text-stone-400 hover:text-stone-700 text-xs inline-flex items-center gap-1"
            >
              Buka <ExternalLink className="w-3 h-3" />
            </a>
          </div>
          <p className="text-xs text-stone-600 mt-2 font-mono truncate">
            ID: {config.balistSpreadsheetId}
          </p>
          <div className="flex items-center justify-between mt-3 pt-2.5 border-t border-stone-200/60 text-xs">
            <span className="text-stone-500">SKU Acuan:</span>
            <span className="font-medium text-stone-800 bg-white px-2 py-0.5 rounded border border-stone-200">
              Kolom 5 (E) &amp; Kolom 6 (F)
            </span>
          </div>
          <div className="flex items-center justify-between mt-1 text-xs">
            <span className="text-stone-500">Baris Termuat:</span>
            <span className="font-semibold text-emerald-700">
              {balistCount > 0 ? `${balistCount.toLocaleString('id-ID')} baris` : 'Belum dimuat'}
            </span>
          </div>
        </div>

        {/* Card 2: STOCK LIST */}
        <div className="p-3.5 rounded-lg bg-stone-50 border border-stone-200/80 flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <span className="w-2 h-2 rounded-full bg-blue-500"></span>
                <span className="text-xs font-semibold text-stone-800 uppercase tracking-wide">
                  Sheet: {config.stockSheetName}
                </span>
              </div>
              <div className="flex items-center gap-2">
                {isAuthenticated && onRefreshStockList && (
                  <button
                    type="button"
                    onClick={onRefreshStockList}
                    disabled={isRefreshingStockList || isLoading}
                    className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-semibold text-blue-700 hover:text-blue-900 bg-blue-50 hover:bg-blue-100 border border-blue-200 transition-colors disabled:opacity-50 cursor-pointer"
                    title="Segarkan data terbaru dari sheet STOCK LIST"
                  >
                    <RefreshCw className={`w-3 h-3 ${isRefreshingStockList ? 'animate-spin' : ''}`} />
                    <span>{isRefreshingStockList ? 'Memperbarui...' : 'Refresh'}</span>
                  </button>
                )}
                <a
                  href={`https://docs.google.com/spreadsheets/d/${config.stockSpreadsheetId}`}
                  target="_blank"
                  rel="noreferrer"
                  className="text-stone-400 hover:text-stone-700 text-xs inline-flex items-center gap-1"
                >
                  Buka <ExternalLink className="w-3 h-3" />
                </a>
              </div>
            </div>
            <p className="text-xs text-stone-600 mt-2 font-mono truncate">
              ID: {config.stockSpreadsheetId}
            </p>
            <div className="flex items-center justify-between mt-3 pt-2.5 border-t border-stone-200/60 text-xs">
              <span className="text-stone-500">Kode SKU &amp; Stok:</span>
              <span className="font-medium text-stone-800 bg-white px-2 py-0.5 rounded border border-stone-200">
                Kolom 1 (SKU) &rarr; Kolom "Qty" (Deteksi Otomatis)
              </span>
            </div>
            <div className="flex items-center justify-between mt-1 text-xs">
              <span className="text-stone-500">Total Stok Gudang:</span>
              <span className="font-semibold text-blue-700">
                {stockCount > 0 ? `${stockCount.toLocaleString('id-ID')} produk` : 'Belum dimuat'}
              </span>
            </div>
          </div>
        </div>
      </div>

      {lastLoaded && (
        <div className="pt-2 flex flex-col sm:flex-row items-center justify-between text-[11px] text-stone-400 gap-1 border-t border-stone-100">
          <div className="flex items-center gap-1 text-emerald-700">
            <CheckCircle2 className="w-3.5 h-3.5 text-emerald-500" />
            <span>Data siap digunakan untuk komparasi &amp; sinkronisasi stok Shopee</span>
          </div>
          <span>Terakhir dimuat: {lastLoaded.toLocaleTimeString('id-ID')}</span>
        </div>
      )}
    </div>
  );
};
