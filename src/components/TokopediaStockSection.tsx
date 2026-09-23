import React, { useState, useRef, useMemo, useEffect } from 'react';
import {
  UploadCloud,
  FileSpreadsheet,
  Download,
  CheckCircle2,
  AlertCircle,
  TrendingDown,
  RefreshCw,
  Search,
  X,
  SlidersHorizontal,
  ChevronLeft,
  ChevronRight,
  Package,
  Layers,
  Sparkles,
  ArrowRight,
  Filter,
  Check,
  Eye,
  FileDown,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { TokopediaRowMatch, StockListItem, ExportFilterOptions } from '../types';
import {
  ParsedTokopediaSheet,
  generateUpdatedTokopediaWorkbook,
  generateFilteredTokopediaWorkbook,
  generateTokopediaFilename,
  createSampleTokopediaFile,
} from '../lib/tokopediaProcessor';
import { downloadBlob } from '../lib/excelProcessor';
import { getColumnLetter } from '../lib/sheets';
import { TokopediaDownloadModal } from './TokopediaDownloadModal';

interface TokopediaStockSectionProps {
  parsedFile: ParsedTokopediaSheet | null;
  uploadedFileName?: string | null;
  matches: TokopediaRowMatch[];
  stockCount: number;
  stockSheetName?: string;
  selectedSkuCol: number;
  selectedStockCol: number;
  dataStartRow: number;
  unmatchedAction: 'keep' | 'zero';
  onFileUpload: (file: File) => void;
  onClearFile: () => void;
  onChangeSkuCol: (colIndex: number) => void;
  onChangeStockCol: (colIndex: number) => void;
  onChangeDataStartRow: (startRow: number) => void;
  onChangeUnmatchedAction: (action: 'keep' | 'zero') => void;
  onRefreshStockList?: () => void;
  isRefreshingStockList?: boolean;
}

export const TokopediaStockSection: React.FC<TokopediaStockSectionProps> = ({
  parsedFile,
  uploadedFileName,
  matches,
  stockCount,
  stockSheetName = 'STOCK LIST',
  selectedSkuCol,
  selectedStockCol,
  dataStartRow,
  unmatchedAction,
  onFileUpload,
  onClearFile,
  onChangeSkuCol,
  onChangeStockCol,
  onChangeDataStartRow,
  onChangeUnmatchedAction,
  onRefreshStockList,
  isRefreshingStockList,
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterType, setFilterType] = useState<'all' | 'matched' | 'unmatched' | 'inStock' | 'outOfStock'>('all');
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(25);
  const [showConfigPanel, setShowConfigPanel] = useState(false);
  const [isDownloadModalOpen, setIsDownloadModalOpen] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);

  // Keyboard shortcut listener for '/' search and 'Esc'
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (
        (e.target instanceof HTMLInputElement || e.target instanceof HTMLTextAreaElement) &&
        e.target.id !== 'input-search-tokopedia'
      ) {
        return;
      }

      if (e.key === '/' && e.target !== document.getElementById('input-search-tokopedia')) {
        e.preventDefault();
        const el = document.getElementById('input-search-tokopedia');
        if (el) (el as HTMLInputElement).focus();
      } else if (e.key === 'Escape' && searchTerm) {
        setSearchTerm('');
        setCurrentPage(1);
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [searchTerm]);

  const summary = useMemo(() => {
    let matchedCount = 0;
    let unmatchedCount = 0;
    let inStockCount = 0;
    let outOfStockCount = 0;
    let stockChangedCount = 0;

    for (const m of matches) {
      if (m.matchStatus === 'matched') {
        matchedCount++;
        if (m.newStock !== null && m.newStock > 0) {
          inStockCount++;
        } else if (m.newStock === 0) {
          outOfStockCount++;
        }
        if (m.stockDiff !== undefined && m.stockDiff !== 0) {
          stockChangedCount++;
        }
      } else {
        unmatchedCount++;
        if (unmatchedAction === 'zero') {
          outOfStockCount++;
        }
      }
    }

    return {
      total: matches.length,
      matchedCount,
      unmatchedCount,
      inStockCount,
      outOfStockCount,
      stockChangedCount,
    };
  }, [matches, unmatchedAction]);

  const filteredMatches = useMemo(() => {
    return matches.filter((m) => {
      if (filterType === 'matched' && m.matchStatus !== 'matched') return false;
      if (filterType === 'unmatched' && m.matchStatus !== 'unmatched') return false;
      if (filterType === 'inStock' && (m.newStock === null || m.newStock <= 0)) return false;
      if (filterType === 'outOfStock' && (m.newStock === null || m.newStock > 0)) return false;

      if (searchTerm.trim()) {
        const q = searchTerm.trim().toLowerCase();
        const digitsQ = q.replace(/\D/g, '');

        const skuMatch = m.skuCol4.toLowerCase().includes(q);
        const nameMatch = m.productName?.toLowerCase().includes(q);
        const varMatch = m.variationName?.toLowerCase().includes(q);
        const codeMatch = m.matchedStockItem?.code?.toLowerCase().includes(q);
        const barcodeMatch = m.matchedStockItem?.barcode?.toLowerCase().includes(q);
        const brandMatch = m.matchedStockItem?.brand?.toLowerCase().includes(q);
        const descMatch = m.matchedStockItem?.description?.toLowerCase().includes(q);
        const rowMatch = String(m.displayRow) === q || `baris ${m.displayRow}`.includes(q);

        let digitsMatch = false;
        if (digitsQ.length >= 3) {
          const skuDigits = m.skuCol4.replace(/\D/g, '');
          const codeDigits = m.matchedStockItem?.code ? m.matchedStockItem.code.replace(/\D/g, '') : '';
          digitsMatch = skuDigits.includes(digitsQ) || codeDigits.includes(digitsQ);
        }

        return (
          skuMatch ||
          nameMatch ||
          varMatch ||
          codeMatch ||
          barcodeMatch ||
          brandMatch ||
          descMatch ||
          rowMatch ||
          digitsMatch
        );
      }

      return true;
    });
  }, [matches, filterType, searchTerm]);

  const totalPages = Math.ceil(filteredMatches.length / pageSize) || 1;
  const paginatedMatches = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredMatches.slice(start, start + pageSize);
  }, [filteredMatches, currentPage, pageSize]);

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

  const handleDownloadSample = () => {
    try {
      const sample = createSampleTokopediaFile();
      downloadBlob(sample, 'Template_Mass_Update_Stok_Tokopedia_Contoh.xlsx');
    } catch (err) {
      console.error('Failed to download sample Tokopedia file:', err);
    }
  };

  const handleOpenDownloadModal = () => {
    if (!parsedFile || matches.length === 0) {
      alert('Tidak ada data Tokopedia untuk diunduh.');
      return;
    }
    setIsDownloadModalOpen(true);
  };

  const handleConfirmFilteredDownload = (
    options: ExportFilterOptions,
    dateFormat: 'dash' | 'readable' = 'dash'
  ) => {
    if (!parsedFile || matches.length === 0) return;
    setIsDownloading(true);

    try {
      const data = generateFilteredTokopediaWorkbook(
        parsedFile,
        matches,
        selectedStockCol ?? 8,
        {
          mode: options.mode,
          selectedCategories: options.selectedCategories,
          selectedBrands: options.selectedBrands,
          selectedRowIndices: options.selectedRowIndices,
          unmatchedAction,
        }
      );

      const prefix = options.storePrefix?.trim() || 'tokopedia';
      const downloadName = generateTokopediaFilename(prefix, new Date(), dateFormat);
      downloadBlob(data, downloadName);
    } catch (err: any) {
      console.error('Download error:', err);
      alert(`Gagal mengunduh file Tokopedia: ${err?.message || 'Error tidak diketahui'}`);
    } finally {
      setIsDownloading(false);
    }
  };

  const handleDownloadRestockList = () => {
    const zeroItems = matches.filter((m) => m.newStock !== null && m.newStock === 0);
    if (zeroItems.length === 0) {
      alert('Tidak ada produk dengan stok 0 / habis untuk diunduh.');
      return;
    }

    const headers = [
      'No',
      'Baris Sheet',
      'SKU Tokopedia (Kolom 4)',
      'Kode STOCK LIST (Kolom 1)',
      'Barcode Gudang',
      'Nama Produk Tokopedia',
      'Deskripsi Barang Gudang',
      'Kategori',
      'Merk',
      'Stok Asli Tokopedia',
      'Stok Baru Gudang',
      'Status',
      'Catatan Pencocokan',
    ];

    const rows = zeroItems.map((m, idx) => [
      idx + 1,
      m.displayRow,
      m.skuCol4 || '-',
      m.matchedStockItem?.code || 'Tidak Terhubung',
      m.matchedStockItem?.barcode || '-',
      m.productName || '-',
      m.matchedStockItem?.description || '-',
      m.matchedStockItem?.category || '-',
      m.matchedStockItem?.brand || '-',
      m.originalStock,
      0,
      'Perlu Restock Segera (Stok 0)',
      m.notes || '',
    ]);

    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    ws['!cols'] = [
      { wch: 6 },
      { wch: 12 },
      { wch: 24 },
      { wch: 24 },
      { wch: 18 },
      { wch: 38 },
      { wch: 38 },
      { wch: 18 },
      { wch: 16 },
      { wch: 18 },
      { wch: 16 },
      { wch: 28 },
      { wch: 32 },
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Daftar Restock Tokopedia');
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const fileName = generateTokopediaFilename('DAFTAR_RESTOCK_TOKOPEDIA_STOK_0');
    downloadBlob(new Uint8Array(out), fileName);
  };

  return (
    <div className="space-y-4">
      {/* Tokopedia Main Upload Card */}
      <div className="bg-white rounded-xl border border-stone-200 p-5 shadow-xs">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-stone-100">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-xl bg-emerald-600 text-white flex items-center justify-center font-bold text-sm shrink-0 shadow-xs">
              <Package className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center gap-2 flex-wrap">
                <h2 className="text-base font-bold text-stone-900 tracking-tight">
                  Pembaruan Stok Tokopedia (Mass Update)
                </h2>
                <span className="px-2 py-0.5 rounded text-[11px] font-bold bg-emerald-50 text-emerald-800 border border-emerald-300">
                  Tokopedia Seller Center
                </span>
                <span className="text-[11px] font-semibold text-stone-600 bg-stone-100 px-2 py-0.5 rounded">
                  SKU: Kolom 4 (D) • Stok: Kolom 9 (I)
                </span>
              </div>
              <p className="text-xs text-stone-500 mt-0.5">
                Unggah file template ubah stok massal dari Tokopedia untuk disinkronkan secara instan dengan database{' '}
                <strong className="text-stone-700">{stockSheetName}</strong>.
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <button
              type="button"
              id="btn-download-tokopedia-sample"
              onClick={handleDownloadSample}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold text-stone-700 bg-stone-100 hover:bg-stone-200 border border-stone-200 transition-colors cursor-pointer"
              title="Unduh file contoh format Tokopedia (SKU Kolom 4, Stok Kolom 9)"
            >
              <Download className="w-3.5 h-3.5 text-stone-600" />
              <span>Unduh Format Contoh</span>
            </button>
            {onRefreshStockList && (
              <button
                type="button"
                onClick={onRefreshStockList}
                disabled={isRefreshingStockList}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold text-emerald-800 bg-emerald-50 hover:bg-emerald-100 border border-emerald-300 transition-colors disabled:opacity-50 cursor-pointer"
                title="Segarkan data terbaru dari STOCK LIST"
              >
                <RefreshCw className={`w-3.5 h-3.5 text-emerald-600 ${isRefreshingStockList ? 'animate-spin' : ''}`} />
                <span>{isRefreshingStockList ? 'Memperbarui...' : 'Refresh Gudang'}</span>
              </button>
            )}
          </div>
        </div>

        {/* Upload Zone */}
        {!parsedFile ? (
          <div
            onDragOver={handleDragOver}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            className="mt-4 border-2 border-dashed border-emerald-300 hover:border-emerald-500 hover:bg-emerald-50/30 rounded-xl p-8 text-center cursor-pointer transition-all group"
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx, .xls"
              onChange={handleFileChange}
              className="hidden"
              id="file-upload-input-tokopedia"
            />
            <div className="w-12 h-12 rounded-full bg-emerald-50 text-emerald-600 group-hover:scale-110 flex items-center justify-center mx-auto mb-3 border border-emerald-200 transition-transform">
              <UploadCloud className="w-6 h-6" />
            </div>
            <p className="text-sm font-semibold text-stone-800">
              Tarik dan lepas file Excel Tokopedia di sini, atau{' '}
              <span className="text-emerald-700 font-bold underline underline-offset-2">pilih file dari perangkat</span>
            </p>
            <p className="text-xs text-stone-500 mt-1 max-w-lg mx-auto">
              Mendukung file <code className="text-emerald-700 font-mono font-semibold">.xlsx</code> hasil ekspor Ubah Stok Massal Tokopedia Seller Center (SKU di Kolom ke-4, Stok di Kolom ke-9, baris data mulai baris 4).
            </p>
          </div>
        ) : (
          <div className="mt-4 space-y-4">
            {/* Uploaded File Banner */}
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 p-3.5 bg-emerald-50/70 border border-emerald-200 rounded-xl">
              <div className="flex items-center gap-3 min-w-0">
                <div className="w-10 h-10 rounded-xl bg-emerald-600 text-white flex items-center justify-center shrink-0 shadow-xs">
                  <FileSpreadsheet className="w-5 h-5" />
                </div>
                <div className="min-w-0">
                  <div className="flex items-center gap-2">
                    <p className="text-sm font-bold text-stone-900 truncate">
                      {uploadedFileName || parsedFile.fileName}
                    </p>
                    <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300">
                      {parsedFile.rows.length - (dataStartRow - 1)} baris produk
                    </span>
                  </div>
                  <p className="text-xs text-stone-600 mt-0.5">
                    Sheet: <strong className="font-semibold text-stone-800">"{parsedFile.sheetName}"</strong> • Kolom SKU:{' '}
                    <strong className="text-emerald-800">{getColumnLetter(selectedSkuCol)} (Kolom {selectedSkuCol + 1})</strong> • Kolom Stok:{' '}
                    <strong className="text-emerald-800">{getColumnLetter(selectedStockCol)} (Kolom {selectedStockCol + 1})</strong> • Mulai Baris:{' '}
                    <strong className="text-stone-800">Baris {dataStartRow}</strong>
                  </p>
                </div>
              </div>

              <div className="flex items-center gap-2 shrink-0 self-end sm:self-center">
                <button
                  type="button"
                  onClick={() => setShowConfigPanel((prev) => !prev)}
                  className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold border transition-colors cursor-pointer ${
                    showConfigPanel
                      ? 'bg-emerald-700 text-white border-emerald-800'
                      : 'bg-white text-stone-700 hover:bg-stone-100 border-stone-200'
                  }`}
                  title="Atur kolom SKU, kolom stok, atau baris mulai data"
                >
                  <SlidersHorizontal className="w-3.5 h-3.5" />
                  <span>{showConfigPanel ? 'Tutup Pengaturan' : 'Atur Kolom'}</span>
                </button>
                <button
                  type="button"
                  onClick={() => fileInputRef.current?.click()}
                  className="inline-flex items-center gap-1 px-3 py-1.5 bg-white hover:bg-stone-100 text-stone-700 rounded-lg text-xs font-semibold border border-stone-200 transition-colors cursor-pointer"
                  title="Ganti dengan file Tokopedia lainnya"
                >
                  <RefreshCw className="w-3.5 h-3.5 text-stone-500" />
                  <span>Ganti File</span>
                </button>
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".xlsx, .xls"
                  onChange={handleFileChange}
                  className="hidden"
                />
                <button
                  type="button"
                  onClick={onClearFile}
                  className="p-1.5 text-stone-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-colors cursor-pointer"
                  title="Hapus file ini"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>
            </div>

            {/* Optional Configuration Panel */}
            {showConfigPanel && (
              <div className="p-4 bg-stone-50 border border-stone-200 rounded-xl space-y-3 animate-fadeIn">
                <div className="flex items-center justify-between pb-2 border-b border-stone-200">
                  <div className="flex items-center gap-2">
                    <SlidersHorizontal className="w-4 h-4 text-emerald-700" />
                    <span className="text-xs font-bold text-stone-800">
                      Penyesuaian Posisi Kolom &amp; Baris Tokopedia
                    </span>
                  </div>
                  <span className="text-[11px] text-stone-500">
                    Standar Tokopedia: SKU = Kolom 4 (D), Stok = Kolom 9 (I), Baris Mulai = Baris 4
                  </span>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-4 gap-3">
                  {/* Kolom SKU */}
                  <div>
                    <label className="block text-[11px] font-bold text-stone-700 mb-1">
                      Kolom SKU Tokopedia:
                    </label>
                    <select
                      value={selectedSkuCol}
                      onChange={(e) => onChangeSkuCol(Number(e.target.value))}
                      className="w-full text-xs bg-white border border-stone-300 rounded-lg px-2.5 py-1.5 focus:ring-2 focus:ring-emerald-500 font-medium"
                    >
                      {parsedFile.headers.map((h, i) => (
                        <option key={i} value={i}>
                          Kolom {i + 1} ({getColumnLetter(i)}) - {h || `Kolom ${i + 1}`}
                        </option>
                      ))}
                    </select>
                  </div>

                  {/* Kolom Stok */}
                  <div>
                    <label className="block text-[11px] font-bold text-stone-700 mb-1">
                      Kolom Stok Tokopedia:
                    </label>
                    <select
                      value={selectedStockCol}
                      onChange={(e) => onChangeStockCol(Number(e.target.value))}
                      className="w-full text-xs bg-white border border-stone-300 rounded-lg px-2.5 py-1.5 focus:ring-2 focus:ring-emerald-500 font-medium"
                    >
                      {parsedFile.headers.map((h, i) => (
                        <option key={i} value={i}>
                          Kolom {i + 1} ({getColumnLetter(i)}) - {h || `Kolom ${i + 1}`}
                        </option>
                      ))}
                    </select>
                  </div>

                  {/* Baris Mulai Data */}
                  <div>
                    <label className="block text-[11px] font-bold text-stone-700 mb-1">
                      Baris Mulai Data:
                    </label>
                    <select
                      value={dataStartRow}
                      onChange={(e) => onChangeDataStartRow(Number(e.target.value))}
                      className="w-full text-xs bg-white border border-stone-300 rounded-lg px-2.5 py-1.5 focus:ring-2 focus:ring-emerald-500 font-medium"
                    >
                      {[1, 2, 3, 4, 5, 6, 7, 8].map((r) => (
                        <option key={r} value={r}>
                          Mulai Baris ke-{r} {r === 4 ? '(Standar Tokopedia)' : ''}
                        </option>
                      ))}
                    </select>
                  </div>

                  {/* Aksi SKU Tidak Cocok */}
                  <div>
                    <label className="block text-[11px] font-bold text-stone-700 mb-1">
                      Jika SKU Tidak Ditemukan:
                    </label>
                    <select
                      value={unmatchedAction}
                      onChange={(e) => onChangeUnmatchedAction(e.target.value as 'keep' | 'zero')}
                      className="w-full text-xs bg-white border border-stone-300 rounded-lg px-2.5 py-1.5 focus:ring-2 focus:ring-emerald-500 font-medium"
                    >
                      <option value="keep">Pertahankan Stok Asli Tokopedia</option>
                      <option value="zero">Ubah Menjadi Stok 0 (Kosongkan)</option>
                    </select>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}
      </div>

      {/* When File is Loaded: Real-time Search, Metric Cards, and Synchronized Table */}
      {parsedFile && matches.length > 0 && (
        <div className="space-y-4">
          {/* Top Real-time Search Bar */}
          <div className="bg-white p-3.5 sm:p-4 rounded-xl border border-stone-200 shadow-xs">
            <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-3">
              <div className="relative flex-1">
                <Search className="w-4 h-4 text-stone-400 absolute left-3.5 top-3 pointer-events-none" />
                <input
                  type="text"
                  id="input-search-tokopedia"
                  placeholder="Cari cepat: Ketik SKU Tokopedia (Kolom 4), 5-digit angka, Nama Produk, Barcode, atau Brand..."
                  value={searchTerm}
                  onChange={(e) => {
                    setSearchTerm(e.target.value);
                    setCurrentPage(1);
                  }}
                  className="w-full pl-10 pr-28 py-2.5 text-xs sm:text-sm bg-stone-50 hover:bg-white focus:bg-white border border-stone-300 rounded-lg text-stone-900 placeholder-stone-400 focus:outline-hidden focus:ring-2 focus:ring-emerald-500/25 focus:border-emerald-600 shadow-2xs transition-all font-medium"
                />
                <div className="absolute right-2.5 top-2.5 flex items-center gap-1.5">
                  {searchTerm && (
                    <button
                      type="button"
                      onClick={() => {
                        setSearchTerm('');
                        setCurrentPage(1);
                      }}
                      className="p-1 text-stone-400 hover:text-stone-700 hover:bg-stone-200/60 rounded-md transition-colors cursor-pointer"
                      title="Hapus pencarian (Esc)"
                    >
                      <X className="w-3.5 h-3.5" />
                    </button>
                  )}
                  <span className="text-[11px] font-semibold text-stone-600 bg-stone-200/70 px-2 py-0.5 rounded border border-stone-300/80">
                    {filteredMatches.length.toLocaleString('id-ID')} / {matches.length.toLocaleString('id-ID')}
                  </span>
                </div>
              </div>

              <div className="flex items-center gap-2 shrink-0">
                {searchTerm && (
                  <button
                    type="button"
                    onClick={() => {
                      setSearchTerm('');
                      setCurrentPage(1);
                    }}
                    className="inline-flex items-center gap-1 px-3 py-2 bg-stone-100 hover:bg-stone-200 text-stone-700 rounded-lg text-xs font-semibold border border-stone-200 transition-colors cursor-pointer"
                  >
                    <X className="w-3.5 h-3.5 text-stone-500" />
                    <span>Reset Cari</span>
                  </button>
                )}

                <button
                  type="button"
                  id="btn-download-updated-tokopedia"
                  onClick={handleOpenDownloadModal}
                  className="inline-flex items-center gap-2 px-4 py-2.5 bg-emerald-700 hover:bg-emerald-800 text-white rounded-lg text-xs font-bold shadow-xs hover:shadow-md transition-all cursor-pointer"
                  title="Unduh file Excel Tokopedia dengan opsi filter kategori, merk, atau semua produk (Kolom 9 otomatis terupdate)"
                >
                  <Download className="w-4 h-4" />
                  <span>Unduh File Tokopedia Terupdate (.xlsx)</span>
                </button>
              </div>
            </div>

            <div className="flex items-center justify-between gap-2 mt-2 pt-2 border-t border-stone-100 text-[11px] text-stone-500">
              <div className="flex items-center gap-1.5 flex-wrap">
                <span className="font-semibold text-emerald-800 bg-emerald-50 px-1.5 py-0.5 rounded border border-emerald-200">
                  Sinkronisasi Otomatis:
                </span>
                <span>
                  SKU Kolom 4 Tokopedia dicocokkan ke Kolom 1 database {stockSheetName}, dan jumlah stok gudang ditulis ke Kolom 9 Tokopedia.
                </span>
              </div>
              <div className="hidden sm:flex items-center gap-1 text-stone-400">
                <span>Shortcut:</span>
                <kbd className="px-1.5 py-0.5 bg-stone-100 border border-stone-300 rounded font-mono text-[10px] text-stone-600">/</kbd>
                <span>fokus cari,</span>
                <kbd className="px-1.5 py-0.5 bg-stone-100 border border-stone-300 rounded font-mono text-[10px] text-stone-600">Esc</kbd>
                <span>reset</span>
              </div>
            </div>
          </div>

          {/* Metric Summary Cards */}
          <div className="grid grid-cols-2 sm:grid-cols-5 gap-3">
            {/* Card 1: Total Item Tokopedia */}
            <button
              type="button"
              onClick={() => {
                setFilterType('all');
                setCurrentPage(1);
              }}
              className={`p-3.5 rounded-xl border shadow-xs transition-all text-left cursor-pointer ${
                filterType === 'all'
                  ? 'bg-emerald-50/90 border-emerald-400 ring-2 ring-emerald-400/50'
                  : 'bg-white border-stone-200 hover:border-emerald-300'
              }`}
            >
              <div className="flex items-center justify-between text-xs font-semibold text-stone-600">
                <span>Total Produk</span>
                <Package className="w-4 h-4 text-stone-400" />
              </div>
              <p className="text-xl font-bold text-stone-900 mt-1">
                {summary.total.toLocaleString('id-ID')}
              </p>
              <div className="mt-2 pt-2 border-t border-stone-100 text-[11px] text-stone-500">
                Mulai baris {dataStartRow}
              </div>
            </button>

            {/* Card 2: Cocok Terhubung */}
            <button
              type="button"
              onClick={() => {
                setFilterType('matched');
                setCurrentPage(1);
              }}
              className={`p-3.5 rounded-xl border shadow-xs transition-all text-left cursor-pointer ${
                filterType === 'matched'
                  ? 'bg-emerald-50/90 border-emerald-400 ring-2 ring-emerald-400/50'
                  : 'bg-white border-emerald-200 hover:border-emerald-400'
              }`}
            >
              <div className="flex items-center justify-between text-xs font-semibold text-emerald-800">
                <span>Cocok Gudang</span>
                <CheckCircle2 className="w-4 h-4 text-emerald-600" />
              </div>
              <p className="text-xl font-bold text-emerald-700 mt-1">
                {summary.matchedCount.toLocaleString('id-ID')}
              </p>
              <div className="mt-2 pt-2 border-t border-emerald-100 text-[11px] text-emerald-700 font-medium">
                {summary.total > 0
                  ? `${((summary.matchedCount / summary.total) * 100).toFixed(1)}% terhubung`
                  : '0%'}
              </div>
            </button>

            {/* Card 3: Belum Cocok (Unmatched) */}
            <button
              type="button"
              onClick={() => {
                setFilterType('unmatched');
                setCurrentPage(1);
              }}
              className={`p-3.5 rounded-xl border shadow-xs transition-all text-left cursor-pointer ${
                filterType === 'unmatched'
                  ? 'bg-amber-50/90 border-amber-400 ring-2 ring-amber-400/50'
                  : 'bg-white border-amber-200 hover:border-amber-400'
              }`}
            >
              <div className="flex items-center justify-between text-xs font-semibold text-amber-800">
                <span>Belum Terhubung</span>
                <AlertCircle className="w-4 h-4 text-amber-600" />
              </div>
              <p className="text-xl font-bold text-amber-700 mt-1">
                {summary.unmatchedCount.toLocaleString('id-ID')}
              </p>
              <div className="mt-2 pt-2 border-t border-amber-100 text-[11px] text-amber-700 font-medium">
                {unmatchedAction === 'keep' ? 'Stok asli tetap' : 'Diubah jadi 0'}
              </div>
            </button>

            {/* Card 4: Stok Tersedia (> 0) */}
            <button
              type="button"
              onClick={() => {
                setFilterType('inStock');
                setCurrentPage(1);
              }}
              className={`p-3.5 rounded-xl border shadow-xs transition-all text-left cursor-pointer ${
                filterType === 'inStock'
                  ? 'bg-blue-50/90 border-blue-400 ring-2 ring-blue-400/50'
                  : 'bg-white border-blue-200 hover:border-blue-400'
              }`}
            >
              <div className="flex items-center justify-between text-xs font-semibold text-blue-800">
                <span>Stok Tersedia</span>
                <Sparkles className="w-4 h-4 text-blue-600" />
              </div>
              <p className="text-xl font-bold text-blue-700 mt-1">
                {summary.inStockCount.toLocaleString('id-ID')}
              </p>
              <div className="mt-2 pt-2 border-t border-blue-100 text-[11px] text-blue-700 font-medium">
                Stok Gudang &gt; 0
              </div>
            </button>

            {/* Card 5: Stok Habis / Butuh Restock (0) */}
            <button
              type="button"
              onClick={() => {
                setFilterType('outOfStock');
                setCurrentPage(1);
              }}
              className={`p-3.5 rounded-xl border shadow-xs transition-all text-left cursor-pointer ${
                filterType === 'outOfStock'
                  ? 'bg-rose-50/90 border-rose-400 ring-2 ring-rose-400/50'
                  : 'bg-white border-rose-200 hover:border-rose-400'
              }`}
            >
              <div className="flex items-center justify-between text-xs font-semibold text-rose-800">
                <span className="flex items-center gap-1">
                  <span className="w-2 h-2 rounded-full bg-rose-500 animate-pulse"></span>
                  Habis (Stok 0)
                </span>
                <TrendingDown className="w-4 h-4 text-rose-600" />
              </div>
              <p className="text-xl font-bold text-rose-700 mt-1">
                {summary.outOfStockCount.toLocaleString('id-ID')}
              </p>
              <div className="mt-2 pt-2 border-t border-rose-100 text-[11px] text-rose-700 font-medium">
                Butuh Restock Segera
              </div>
            </button>
          </div>

          {/* Dedicated Restock Alert Banner when outOfStock filter is active */}
          {filterType === 'outOfStock' && summary.outOfStockCount > 0 && (
            <div className="px-4 py-3 bg-rose-50/95 border border-rose-200 rounded-xl text-rose-950 flex flex-col md:flex-row md:items-center justify-between gap-3 animate-fadeIn shadow-xs">
              <div className="flex items-start gap-2.5">
                <div className="p-2 bg-rose-200 text-rose-800 rounded-lg shrink-0 mt-0.5">
                  <TrendingDown className="w-4 h-4 text-rose-700" />
                </div>
                <div>
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-bold uppercase tracking-wider text-rose-800 bg-rose-200/80 px-1.5 py-0.5 rounded text-[10px]">
                      Mode Fokus Restock Tokopedia
                    </span>
                    <span className="text-xs font-bold text-rose-900">
                      {summary.outOfStockCount.toLocaleString('id-ID')} Produk Memerlukan Pengadaan Ulang (Stok = 0)
                    </span>
                  </div>
                  <p className="text-[11px] text-rose-700 mt-0.5">
                    Menampilkan produk Tokopedia yang stoknya kosong di gudang {stockSheetName}. Gunakan data ini untuk PO ke supplier.
                  </p>
                </div>
              </div>

              <div className="flex items-center gap-2 shrink-0 self-end md:self-center">
                <button
                  type="button"
                  onClick={handleDownloadRestockList}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-rose-700 hover:bg-rose-800 text-white rounded-lg text-xs font-semibold shadow-xs transition-colors cursor-pointer"
                  title="Unduh daftar produk stok 0 ke format Excel"
                >
                  <FileDown className="w-3.5 h-3.5" />
                  <span>Unduh List Restock (.xlsx)</span>
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setFilterType('all');
                    setCurrentPage(1);
                  }}
                  className="inline-flex items-center gap-1 px-2.5 py-1.5 bg-white hover:bg-stone-100 text-stone-700 rounded-lg text-xs font-medium border border-rose-200 transition-colors cursor-pointer"
                >
                  <X className="w-3 h-3" />
                  <span>Reset Filter</span>
                </button>
              </div>
            </div>
          )}

          {/* Main Table Container */}
          <div className="bg-white rounded-xl border border-stone-200 shadow-xs overflow-hidden">
            {/* Table Action Bar */}
            <div className="p-3.5 sm:p-4 border-b border-stone-200 bg-stone-50/70 flex flex-col md:flex-row md:items-center justify-between gap-3">
              {/* Filter Tabs */}
              <div className="flex items-center gap-1.5 flex-wrap">
                <button
                  type="button"
                  onClick={() => {
                    setFilterType('all');
                    setCurrentPage(1);
                  }}
                  className={`px-3 py-1 rounded-md font-semibold text-xs transition-colors cursor-pointer ${
                    filterType === 'all'
                      ? 'bg-stone-800 text-white shadow-2xs'
                      : 'bg-white text-stone-700 hover:bg-stone-200 border border-stone-200'
                  }`}
                >
                  Semua ({summary.total.toLocaleString('id-ID')})
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setFilterType('matched');
                    setCurrentPage(1);
                  }}
                  className={`px-3 py-1 rounded-md font-semibold text-xs transition-colors cursor-pointer ${
                    filterType === 'matched'
                      ? 'bg-emerald-700 text-white shadow-2xs'
                      : 'bg-emerald-50 text-emerald-800 hover:bg-emerald-100 border border-emerald-300'
                  }`}
                >
                  Cocok ({summary.matchedCount.toLocaleString('id-ID')})
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setFilterType('unmatched');
                    setCurrentPage(1);
                  }}
                  className={`px-3 py-1 rounded-md font-semibold text-xs transition-colors cursor-pointer ${
                    filterType === 'unmatched'
                      ? 'bg-amber-700 text-white shadow-2xs'
                      : 'bg-amber-50 text-amber-800 hover:bg-amber-100 border border-amber-300'
                  }`}
                >
                  Belum Cocok ({summary.unmatchedCount.toLocaleString('id-ID')})
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setFilterType('inStock');
                    setCurrentPage(1);
                  }}
                  className={`px-3 py-1 rounded-md font-semibold text-xs transition-colors cursor-pointer ${
                    filterType === 'inStock'
                      ? 'bg-blue-700 text-white shadow-2xs'
                      : 'bg-blue-50 text-blue-800 hover:bg-blue-100 border border-blue-300'
                  }`}
                >
                  Stok Tersedia ({summary.inStockCount.toLocaleString('id-ID')})
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setFilterType('outOfStock');
                    setCurrentPage(1);
                  }}
                  className={`inline-flex items-center gap-1.5 px-3 py-1 rounded-md font-semibold text-xs transition-colors cursor-pointer ${
                    filterType === 'outOfStock'
                      ? 'bg-rose-700 text-white shadow-2xs ring-2 ring-rose-300'
                      : 'bg-rose-50 text-rose-800 hover:bg-rose-100 border border-rose-300'
                  }`}
                >
                  <span className={`w-2 h-2 rounded-full ${filterType === 'outOfStock' ? 'bg-white' : 'bg-rose-500 animate-pulse'}`}></span>
                  <span>Stok 0 / Butuh Restock</span>
                  <span className={`px-1.5 py-0.2 rounded-full text-[10px] font-bold ${
                    filterType === 'outOfStock' ? 'bg-rose-900/40 text-white' : 'bg-rose-200 text-rose-900'
                  }`}>
                    {summary.outOfStockCount.toLocaleString('id-ID')}
                  </span>
                </button>
              </div>

              {/* Page size & Table Stats */}
              <div className="flex items-center gap-2 text-xs text-stone-500 self-end md:self-center">
                <span>Tampilkan:</span>
                <select
                  value={pageSize}
                  onChange={(e) => {
                    setPageSize(Number(e.target.value));
                    setCurrentPage(1);
                  }}
                  className="bg-white border border-stone-200 rounded-lg px-2 py-1 text-xs font-semibold text-stone-700"
                >
                  <option value={15}>15 baris</option>
                  <option value={25}>25 baris</option>
                  <option value={50}>50 baris</option>
                  <option value={100}>100 baris</option>
                </select>
              </div>
            </div>

            {/* Table Content */}
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-stone-700">
                <thead className="bg-stone-100/90 text-stone-700 uppercase font-bold text-[10px] tracking-wider border-b border-stone-200">
                  <tr>
                    <th scope="col" className="py-3 px-3 w-14 text-center">Baris</th>
                    <th scope="col" className="py-3 px-3 min-w-[140px]">
                      <div className="flex flex-col">
                        <span className="text-emerald-800">SKU Tokopedia</span>
                        <span className="text-[9px] font-normal text-stone-400 lowercase">Kolom 4 (D)</span>
                      </div>
                    </th>
                    <th scope="col" className="py-3 px-3 min-w-[200px]">Nama Produk Tokopedia</th>
                    <th scope="col" className="py-3 px-3 min-w-[140px]">
                      <div className="flex flex-col">
                        <span className="text-stone-800">Kode STOCK LIST</span>
                        <span className="text-[9px] font-normal text-stone-400 lowercase">Kolom 1 Gudang</span>
                      </div>
                    </th>
                    <th scope="col" className="py-3 px-3 min-w-[180px]">Nama Barang Gudang</th>
                    <th scope="col" className="py-3 px-3 text-center w-28">
                      <div className="flex flex-col items-center">
                        <span className="text-stone-600">Stok Asli</span>
                        <span className="text-[9px] font-normal text-stone-400 lowercase">Kolom 9 (I)</span>
                      </div>
                    </th>
                    <th scope="col" className="py-3 px-3 text-center w-28">
                      <div className="flex flex-col items-center">
                        <span className="text-emerald-800">Stok Gudang</span>
                        <span className="text-[9px] font-normal text-stone-400 lowercase">Kolom 15 Gudang</span>
                      </div>
                    </th>
                    <th scope="col" className="py-3 px-3 text-center w-24">Selisih</th>
                    <th scope="col" className="py-3 px-3 text-center w-28">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-stone-100">
                  {paginatedMatches.length === 0 ? (
                    <tr>
                      <td colSpan={9} className="py-12 text-center text-stone-500">
                        <div className="flex flex-col items-center justify-center gap-2">
                          <Package className="w-8 h-8 text-stone-300" />
                          <p className="text-sm font-semibold text-stone-700">
                            {searchTerm
                              ? `Tidak ditemukan produk Tokopedia yang cocok dengan "${searchTerm}"`
                              : 'Tidak ada data produk pada kategori filter ini.'}
                          </p>
                          <p className="text-xs text-stone-400">
                            Coba periksa kata kunci pencarian atau reset filter untuk melihat semua baris.
                          </p>
                          {(searchTerm || filterType !== 'all') && (
                            <button
                              type="button"
                              onClick={() => {
                                setSearchTerm('');
                                setFilterType('all');
                                setCurrentPage(1);
                              }}
                              className="mt-2 px-3 py-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 rounded-lg text-xs font-semibold border border-emerald-300 transition-colors"
                            >
                              Reset Semua Filter &amp; Pencarian
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                  ) : (
                    paginatedMatches.map((m) => {
                      const isZeroStock = m.newStock === 0;
                      const isMatched = m.matchStatus === 'matched';
                      return (
                        <tr
                          key={m.rowIndex}
                          className={`hover:bg-stone-50/90 transition-colors ${
                            !isMatched
                              ? 'bg-amber-50/20'
                              : isZeroStock
                              ? 'bg-rose-50/30'
                              : ''
                          }`}
                        >
                          {/* Baris */}
                          <td className="py-2.5 px-3 text-center font-mono text-[11px] text-stone-500 font-semibold">
                            {m.displayRow}
                          </td>

                          {/* SKU Tokopedia (Kolom 4 / D) */}
                          <td className="py-2.5 px-3">
                            <span className="font-mono text-xs font-bold text-stone-900 bg-stone-100 px-2 py-0.5 rounded border border-stone-200">
                              {m.skuCol4 || '<Kosong>'}
                            </span>
                          </td>

                          {/* Nama Produk Tokopedia */}
                          <td className="py-2.5 px-3">
                            <div className="flex flex-col max-w-[240px]">
                              <span className="text-xs font-semibold text-stone-900 line-clamp-2">
                                {m.productName || '-'}
                              </span>
                              {m.variationName && (
                                <span className="text-[10px] text-stone-500">
                                  Varian: {m.variationName}
                                </span>
                              )}
                            </div>
                          </td>

                          {/* Kode Gudang STOCK LIST */}
                          <td className="py-2.5 px-3">
                            {m.matchedStockItem ? (
                              <span className="font-mono text-xs font-bold text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                                {m.matchedStockItem.code}
                              </span>
                            ) : (
                              <span className="text-stone-400 italic text-[11px]">-</span>
                            )}
                          </td>

                          {/* Nama Barang Gudang */}
                          <td className="py-2.5 px-3">
                            {m.matchedStockItem ? (
                              <div className="flex flex-col max-w-[220px]">
                                <span className="text-xs text-stone-800 font-medium line-clamp-2">
                                  {m.matchedStockItem.description || '-'}
                                </span>
                                {(m.matchedStockItem.brand || m.matchedStockItem.category) && (
                                  <span className="text-[10px] text-stone-400">
                                    {[m.matchedStockItem.brand, m.matchedStockItem.category].filter(Boolean).join(' • ')}
                                  </span>
                                )}
                              </div>
                            ) : (
                              <span className="text-stone-400 italic text-[11px]">-</span>
                            )}
                          </td>

                          {/* Stok Asli Tokopedia */}
                          <td className="py-2.5 px-3 text-center font-mono text-xs text-stone-600">
                            {m.originalStock !== null && m.originalStock !== undefined ? (
                              <span>{Number(m.originalStock).toLocaleString('id-ID')}</span>
                            ) : (
                              <span className="text-stone-400 italic">-</span>
                            )}
                          </td>

                          {/* Stok Baru Gudang */}
                          <td className="py-2.5 px-3 text-center">
                            {m.newStock !== null ? (
                              <div className="flex flex-col items-center gap-0.5">
                                <span
                                  className={`inline-flex items-center px-2 py-0.5 rounded-full font-bold font-mono text-xs ${
                                    m.newStock > 0
                                      ? 'bg-emerald-100 text-emerald-800 border border-emerald-200'
                                      : 'bg-rose-100 text-rose-800 border border-rose-300 shadow-2xs'
                                  }`}
                                >
                                  {m.newStock.toLocaleString('id-ID')}
                                </span>
                                {m.newStock === 0 && (
                                  <span className="text-[9px] font-bold text-rose-600 tracking-tight uppercase">
                                    Butuh Restock
                                  </span>
                                )}
                              </div>
                            ) : (
                              <span className="text-stone-300 italic">-</span>
                            )}
                          </td>

                          {/* Selisih */}
                          <td className="py-2.5 px-3 text-center font-mono text-xs font-semibold">
                            {m.stockDiff !== undefined && m.stockDiff !== 0 ? (
                              <span
                                className={
                                  m.stockDiff > 0
                                    ? 'text-emerald-700 bg-emerald-50 px-1.5 py-0.5 rounded border border-emerald-200'
                                    : 'text-rose-700 bg-rose-50 px-1.5 py-0.5 rounded border border-rose-200'
                                }
                              >
                                {m.stockDiff > 0 ? `+${m.stockDiff}` : m.stockDiff}
                              </span>
                            ) : (
                              <span className="text-stone-400">-</span>
                            )}
                          </td>

                          {/* Status */}
                          <td className="py-2.5 px-3 text-center">
                            {isMatched ? (
                              isZeroStock ? (
                                <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-rose-700 bg-rose-50 border border-rose-200 px-2 py-0.5 rounded-full">
                                  <TrendingDown className="w-3 h-3 text-rose-600" />
                                  Cocok (Stok 0)
                                </span>
                              ) : (
                                <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-full">
                                  <CheckCircle2 className="w-3 h-3" />
                                  Cocok
                                </span>
                              )
                            ) : (
                              <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-amber-800 bg-amber-50 border border-amber-200 px-2 py-0.5 rounded-full">
                                <AlertCircle className="w-3 h-3" />
                                Belum Cocok
                              </span>
                            )}
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>

            {/* Pagination Controls */}
            {totalPages > 1 && (
              <div className="p-3 bg-stone-50 border-t border-stone-200 flex flex-col sm:flex-row items-center justify-between gap-2 text-xs text-stone-600">
                <span>
                  Menampilkan{' '}
                  <strong className="text-stone-800">
                    {(currentPage - 1) * pageSize + 1} -{' '}
                    {Math.min(currentPage * pageSize, filteredMatches.length)}
                  </strong>{' '}
                  dari <strong className="text-stone-800">{filteredMatches.length}</strong> baris
                </span>

                <div className="flex items-center gap-1">
                  <button
                    type="button"
                    onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                    disabled={currentPage === 1}
                    className="p-1.5 rounded-lg border border-stone-200 bg-white text-stone-700 hover:bg-stone-100 disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
                  >
                    <ChevronLeft className="w-4 h-4" />
                  </button>

                  <div className="flex items-center gap-1 px-2 font-medium">
                    <span>Hal</span>
                    <input
                      type="number"
                      min={1}
                      max={totalPages}
                      value={currentPage}
                      onChange={(e) => {
                        const v = parseInt(e.target.value, 10);
                        if (!isNaN(v) && v >= 1 && v <= totalPages) {
                          setCurrentPage(v);
                        }
                      }}
                      className="w-12 text-center py-0.5 border border-stone-300 rounded font-bold text-xs"
                    />
                    <span>dari {totalPages}</span>
                  </div>

                  <button
                    type="button"
                    onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                    disabled={currentPage === totalPages}
                    className="p-1.5 rounded-lg border border-stone-200 bg-white text-stone-700 hover:bg-stone-100 disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
                  >
                    <ChevronRight className="w-4 h-4" />
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* Tokopedia Download Filter Modal: Filter by Category, Brand, Custom Checklist, or All */}
      {parsedFile && (
        <TokopediaDownloadModal
          isOpen={isDownloadModalOpen}
          onClose={() => setIsDownloadModalOpen(false)}
          onConfirmDownload={handleConfirmFilteredDownload}
          matches={matches}
          initialPrefix={uploadedFileName ? uploadedFileName.replace(/\.xlsx?$/i, '') : 'tokopedia'}
          defaultPrefix="tokopedia"
          stockSheetName={stockSheetName}
          isProcessing={isDownloading}
        />
      )}
    </div>
  );
};
