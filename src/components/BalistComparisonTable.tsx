import React, { useState, useMemo, useEffect } from 'react';
import {
  Search,
  CheckCircle2,
  AlertCircle,
  Package,
  Layers,
  Check,
  TrendingDown,
  TrendingUp,
  Download,
  Filter,
  Info,
  Send,
  ExternalLink,
  ChevronRight,
  Store,
  Lock,
  LogIn,
  RefreshCw,
  Eye,
  EyeOff,
  ChevronDown,
  ChevronUp,
  Table,
} from 'lucide-react';
import { BalistComparisonItem, BalistComparisonSummary } from '../types';
import * as XLSX from 'xlsx';
import { downloadBlob, generateShopeeBalistFilename, detectStoreFromFilename } from '../lib/excelProcessor';
import { parseNumber } from '../lib/sheets';
import { BalistProductListModal, BalistModalCategory } from './BalistProductListModal';

interface BalistComparisonTableProps {
  items: BalistComparisonItem[];
  summary: BalistComparisonSummary;
  balistSheetName: string;
  stockSheetName: string;
  uploadedFileName?: string | null;
  onUpdateBalistStockInSheet?: () => void;
  isUpdatingBalistStock?: boolean;
  isAuthenticated?: boolean;
  onPromptSignIn?: () => void;
  onRefreshStockList?: () => void;
  isRefreshingStockList?: boolean;
  onDownloadUpdatedBalistXlsx?: (prefix?: string) => void;
  stockColIndex?: number;
}

export const BalistComparisonTable: React.FC<BalistComparisonTableProps> = ({
  items,
  summary,
  balistSheetName,
  stockSheetName,
  uploadedFileName,
  onUpdateBalistStockInSheet,
  isUpdatingBalistStock,
  isAuthenticated = false,
  onPromptSignIn,
  onRefreshStockList,
  isRefreshingStockList = false,
  onDownloadUpdatedBalistXlsx,
  stockColIndex = 6,
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [filterType, setFilterType] = useState<
    'all' | 'matched' | 'unmatched' | 'inStock' | 'outOfStock'
  >('all');
  const [currentPage, setCurrentPage] = useState(1);
  const [modalCategory, setModalCategory] = useState<BalistModalCategory | null>(null);
  const pageSize = 50;

  const detectedStore = useMemo(() => {
    return detectStoreFromFilename(uploadedFileName || '');
  }, [uploadedFileName]);

  const [exportPrefix, setExportPrefix] = useState<'balist' | 'gomall'>(() => {
    if (uploadedFileName) {
      const d = detectStoreFromFilename(uploadedFileName);
      if (d.storePrefix) return d.storePrefix;
    }
    const saved = localStorage.getItem('balist_download_prefix_option');
    return saved === 'gomall' ? 'gomall' : 'balist';
  });

  // State to hide/collapse the large table list (defaults to false / hidden)
  const [showDetailedTable, setShowDetailedTable] = useState<boolean>(() => {
    const saved = localStorage.getItem('show_balist_comparison_table');
    return saved === 'true'; // Default is false (hidden)
  });

  const handleToggleDetailedTable = () => {
    setShowDetailedTable((prev) => {
      const next = !prev;
      localStorage.setItem('show_balist_comparison_table', String(next));
      return next;
    });
  };

  useEffect(() => {
    if (detectedStore.storePrefix) {
      setExportPrefix(detectedStore.storePrefix);
    }
  }, [detectedStore.storePrefix]);

  const handleOpenModal = (category: BalistModalCategory) => {
    setModalCategory(category);
  };

  const filteredItems = useMemo(() => {
    return items.filter((item) => {
      if (filterType === 'matched' && item.matchStatus !== 'matched') return false;
      if (filterType === 'unmatched' && item.matchStatus !== 'unmatched') return false;
      if (filterType === 'inStock' && (item.stockQty === null || item.stockQty <= 0))
        return false;
      if (filterType === 'outOfStock' && (item.stockQty === null || item.stockQty > 0))
        return false;

      if (searchTerm.trim()) {
        const q = searchTerm.toLowerCase();
        const nameMatch = item.productName?.toLowerCase().includes(q);
        const varMatch = item.variationName?.toLowerCase().includes(q);
        const parentMatch = item.parentSku?.toLowerCase().includes(q);
        const sku5Match = item.skuCol5.toLowerCase().includes(q);
        const sku6Match = item.skuCol6.toLowerCase().includes(q);
        const codeMatch = item.matchedStockItem?.code.toLowerCase().includes(q);
        const descMatch = item.matchedStockItem?.description?.toLowerCase().includes(q);
        const catMatch = item.matchedStockItem?.category?.toLowerCase().includes(q);
        const brandMatch = item.matchedStockItem?.brand?.toLowerCase().includes(q);
        const notesMatch = item.notes?.toLowerCase().includes(q);
        const rowMatch = String(item.rowIndex).includes(q);
        return (
          nameMatch ||
          varMatch ||
          parentMatch ||
          sku5Match ||
          sku6Match ||
          codeMatch ||
          descMatch ||
          catMatch ||
          brandMatch ||
          notesMatch ||
          rowMatch
        );
      }
      return true;
    });
  }, [items, filterType, searchTerm]);

  const totalPages = Math.ceil(filteredItems.length / pageSize) || 1;
  const paginatedItems = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredItems.slice(start, start + pageSize);
  }, [filteredItems, currentPage]);

  const handleExportComparison = (prefix: string = exportPrefix) => {
    // If the parent provided the full Shopee XLSX generator (which preserves the uploaded template exactly), use it
    if (onDownloadUpdatedBalistXlsx) {
      onDownloadUpdatedBalistXlsx(prefix);
      return;
    }

    // Otherwise, generate the exact 6-header Shopee Mass Update (Informasi Penjualan) template
    const shopeeHeaderRows = [
      ['Pusat Edukasi Penjual > Pelajari Lebih Lanjut Tentang Update Massal Informasi Penjualan', '', '', '', '', '', '', '', '', ''],
      ['Kategori', 'Informasi Penjualan', '', '', '', '', '', '', '', ''],
      ['Catatan: 1. Jangan ubah format baris header (Baris 1-6) | 2. Jangan ubah data pada kolom bertanda bintang (*) | 3. Pastikan format file tetap .xlsx', '', '', '', '', '', '', '', '', ''],
      ['Kode Produk', 'Nama Produk', 'No. Integrasi Produk', 'Kode Variasi', 'Nama Variasi', 'Kode Integrasi', 'Stok', 'Harga', 'Status Produk', 'SKU Induk'],
      ['Wajib', 'Hanya baca', 'Hanya baca', 'Wajib', 'Hanya baca', 'Opsional', 'Wajib', 'Opsional', 'Hanya baca', 'Opsional'],
      ['Contoh: 12345678', 'Contoh: Produk A', 'Contoh: P001', 'Contoh: 87654321', 'Contoh: Standar', 'Contoh: SKU001', 'Contoh: 100', 'Contoh: 50000', 'Contoh: Aktif', 'Contoh: SKU000'],
    ];

    const dataRows = items.map((it) => {
      // Matched stock quantity from STOCK LIST or fallback to raw
      const finalStock =
        it.stockQty !== null
          ? it.stockQty
          : it.rawRow && it.rawRow[stockColIndex] !== undefined && it.rawRow[stockColIndex] !== ''
          ? parseNumber(it.rawRow[stockColIndex])
          : 0;

      // If raw row is complete, clone and overwrite stock column
      if (it.rawRow && it.rawRow.length >= 6) {
        const cloned = [...it.rawRow];
        while (cloned.length <= stockColIndex) cloned.push('');
        cloned[stockColIndex] = finalStock;
        return cloned;
      }

      // Default fallback row structured precisely for Shopee Mass Update format
      const productId = it.rawRow && it.rawRow[0] ? it.rawRow[0] : it.rowIndex;
      const productName =
        it.matchedStockItem?.description ||
        (it.rawRow && it.rawRow[1] ? String(it.rawRow[1]) : `Produk SKU ${it.skuCol5 || it.skuCol6}`);
      const parentIntegration = it.rawRow && it.rawRow[2] ? it.rawRow[2] : '';
      const variationId = it.rawRow && it.rawRow[3] ? it.rawRow[3] : '';
      const variationName = it.skuCol5 || (it.rawRow && it.rawRow[4] ? it.rawRow[4] : 'Standar');
      const variationSku = it.skuCol6 || it.skuCol5 || (it.rawRow && it.rawRow[5] ? it.rawRow[5] : '');
      const price = it.rawRow && it.rawRow[7] !== undefined ? it.rawRow[7] : '';
      const status = it.rawRow && it.rawRow[8] ? it.rawRow[8] : 'Aktif';
      const parentSku = it.skuCol5 || '';

      return [
        productId,
        productName,
        parentIntegration,
        variationId,
        variationName,
        variationSku,
        finalStock,
        price,
        status,
        parentSku,
      ];
    });

    const fullRows = [...shopeeHeaderRows, ...dataRows];
    const ws = XLSX.utils.aoa_to_sheet(fullRows);

    // Standard column widths for Shopee Mass Update
    ws['!cols'] = [
      { wch: 16 }, // Kode Produk
      { wch: 40 }, // Nama Produk
      { wch: 20 }, // No Integrasi Produk
      { wch: 16 }, // Kode Variasi
      { wch: 22 }, // Nama Variasi (Kolom 5)
      { wch: 22 }, // Kode Integrasi / SKU Variasi (Kolom 6)
      { wch: 12 }, // Stok (Kolom 7)
      { wch: 14 }, // Harga
      { wch: 14 }, // Status Produk
      { wch: 20 }, // SKU Induk
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Mass Update Template');

    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const fileName = generateShopeeBalistFilename(prefix || 'balist');
    downloadBlob(new Uint8Array(out), fileName);
  };

  return (
    <div className="space-y-4">
      {/* Header Bar with Quick Action Buttons & Table Visibility Toggle */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 pb-1">
        <div className="flex items-center gap-2">
          <h3 className="text-sm font-bold text-stone-800 flex items-center gap-1.5">
            <Package className="w-4 h-4 text-emerald-600" />
            <span>Hasil Analisa Stok Balistshopee</span>
          </h3>
          {detectedStore.storeName && (
            <span className="text-[11px] font-semibold bg-emerald-100/80 text-emerald-900 border border-emerald-300/80 px-2 py-0.5 rounded-md flex items-center gap-1">
              <Store className="w-3 h-3 text-emerald-700" />
              {detectedStore.storeName}
            </span>
          )}
        </div>

        <div className="flex items-center gap-2 flex-wrap">
          {onRefreshStockList && (
            <button
              type="button"
              id="btn-refresh-stocklist-comparison-top"
              onClick={onRefreshStockList}
              disabled={isRefreshingStockList}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold text-emerald-800 hover:text-emerald-950 bg-emerald-50 hover:bg-emerald-100 border border-emerald-300 shadow-2xs transition-colors shrink-0 cursor-pointer disabled:opacity-50"
              title="Segarkan data terbaru dari Sheet STOCK LIST di Google Sheets"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isRefreshingStockList ? 'animate-spin' : ''}`} />
              <span>{isRefreshingStockList ? 'Memperbarui...' : 'Refresh STOCK LIST'}</span>
            </button>
          )}

          <button
            type="button"
            onClick={() => handleExportComparison(exportPrefix)}
            id="btn-export-comparison-xlsx-top"
            className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-2xs transition-colors shrink-0 cursor-pointer text-white bg-emerald-700 hover:bg-emerald-800"
            title={`Unduh file Excel format Mass Update Shopee (.xlsx) untuk toko ${exportPrefix}`}
          >
            <Download className="w-3.5 h-3.5" />
            <span>Unduh File Shopee (.xlsx)</span>
          </button>

          <button
            type="button"
            id="btn-toggle-comparison-table"
            onClick={handleToggleDetailedTable}
            className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium border transition-colors cursor-pointer shadow-2xs ${
              showDetailedTable
                ? 'bg-stone-100 text-stone-800 border-stone-300 hover:bg-stone-200'
                : 'bg-white text-stone-600 border-stone-200 hover:bg-stone-50 hover:text-stone-900'
            }`}
            title={showDetailedTable ? 'Sembunyikan tabel list rincian' : 'Tampilkan tabel list rincian'}
          >
            {showDetailedTable ? (
              <>
                <EyeOff className="w-3.5 h-3.5 text-stone-500" />
                <span>Sembunyikan Tabel</span>
                <ChevronUp className="w-3.5 h-3.5 text-stone-400" />
              </>
            ) : (
              <>
                <Eye className="w-3.5 h-3.5 text-stone-500" />
                <span>Tampilkan Tabel Rincian</span>
                <ChevronDown className="w-3.5 h-3.5 text-stone-400" />
              </>
            )}
          </button>
        </div>
      </div>

      {/* Metric Cards for Comparison */}
      <div className="grid grid-cols-2 sm:grid-cols-5 gap-3">
        {/* Card 1: Total Item Balist */}
        <button
          type="button"
          onClick={() => handleOpenModal('all')}
          className="p-3.5 bg-white rounded-xl border border-stone-200 shadow-xs hover:border-stone-400 hover:shadow-md transition-all text-left cursor-pointer group focus:outline-hidden focus:ring-2 focus:ring-stone-400"
          title="Klik untuk membuka popup daftar semua item Balist"
        >
          <div className="flex items-center justify-between text-stone-500 text-xs">
            <span className="font-semibold text-stone-700">Total Item Balist</span>
            <Layers className="w-4 h-4 text-stone-400 group-hover:text-stone-700 transition-colors" />
          </div>
          <p className="text-xl font-bold text-stone-900 mt-1">
            {summary.totalRows.toLocaleString('id-ID')}
          </p>
          <div className="flex items-center justify-between mt-2 pt-2 border-t border-stone-100 text-[11px]">
            <span className="text-stone-400">Sheet {balistSheetName}</span>
            <span className="text-stone-500 font-medium group-hover:text-stone-900 inline-flex items-center gap-0.5">
              Lihat list <ChevronRight className="w-3 h-3 group-hover:translate-x-0.5 transition-transform" />
            </span>
          </div>
        </button>

        {/* Card 2: Cocok di STOCK LIST */}
        <button
          type="button"
          onClick={() => handleOpenModal('matched')}
          className="p-3.5 bg-white rounded-xl border border-emerald-200 shadow-xs hover:border-emerald-400 hover:shadow-md transition-all text-left cursor-pointer group focus:outline-hidden focus:ring-2 focus:ring-emerald-500"
          title="Klik untuk membuka popup daftar produk cocok di STOCK LIST"
        >
          <div className="flex items-center justify-between text-emerald-600 text-xs">
            <span className="font-semibold text-emerald-800">Cocok di STOCK LIST</span>
            <CheckCircle2 className="w-4 h-4 text-emerald-500 group-hover:text-emerald-700 transition-colors" />
          </div>
          <p className="text-xl font-bold text-emerald-700 mt-1">
            {summary.matchedCount.toLocaleString('id-ID')}
          </p>
          <div className="flex items-center justify-between mt-2 pt-2 border-t border-emerald-100/70 text-[11px]">
            <span className="text-emerald-600 font-medium">
              {summary.totalRows > 0
                ? `${Math.round((summary.matchedCount / summary.totalRows) * 100)}% terhubung`
                : '0%'}
            </span>
            <span className="text-emerald-700 font-medium inline-flex items-center gap-0.5">
              Lihat list <ChevronRight className="w-3 h-3 group-hover:translate-x-0.5 transition-transform" />
            </span>
          </div>
        </button>

        {/* Card 3: Tidak Ditemukan */}
        <button
          type="button"
          onClick={() => handleOpenModal('unmatched')}
          className="p-3.5 bg-white rounded-xl border border-amber-200 shadow-xs hover:border-amber-400 hover:shadow-md transition-all text-left cursor-pointer group focus:outline-hidden focus:ring-2 focus:ring-amber-500"
          title="Klik untuk membuka popup daftar produk yang tidak ditemukan di STOCK LIST"
        >
          <div className="flex items-center justify-between text-amber-600 text-xs">
            <span className="font-semibold text-amber-800">Tidak Ditemukan</span>
            <AlertCircle className="w-4 h-4 text-amber-500 group-hover:text-amber-700 transition-colors" />
          </div>
          <p className="text-xl font-bold text-amber-700 mt-1">
            {summary.unmatchedCount.toLocaleString('id-ID')}
          </p>
          <div className="flex items-center justify-between mt-2 pt-2 border-t border-amber-100/70 text-[11px]">
            <span className="text-amber-600">Perlu cek SKU</span>
            <span className="text-amber-700 font-medium inline-flex items-center gap-0.5">
              Lihat list <ChevronRight className="w-3 h-3 group-hover:translate-x-0.5 transition-transform" />
            </span>
          </div>
        </button>

        {/* Card 4: Tersedia (Ready) */}
        <button
          type="button"
          onClick={() => handleOpenModal('inStock')}
          className="p-3.5 bg-white rounded-xl border border-blue-200 shadow-xs hover:border-blue-400 hover:shadow-md transition-all text-left cursor-pointer group focus:outline-hidden focus:ring-2 focus:ring-blue-500"
          title="Klik untuk membuka popup daftar produk stok tersedia (> 0)"
        >
          <div className="flex items-center justify-between text-blue-600 text-xs">
            <span className="font-semibold text-blue-800">Tersedia (Ready)</span>
            <TrendingUp className="w-4 h-4 text-blue-500 group-hover:text-blue-700 transition-colors" />
          </div>
          <p className="text-xl font-bold text-blue-700 mt-1">
            {summary.inStockCount.toLocaleString('id-ID')}
          </p>
          <div className="flex items-center justify-between mt-2 pt-2 border-t border-blue-100/70 text-[11px]">
            <span className="text-blue-600 font-medium">Stok &gt; 0</span>
            <span className="text-blue-700 font-medium inline-flex items-center gap-0.5">
              Lihat list <ChevronRight className="w-3 h-3 group-hover:translate-x-0.5 transition-transform" />
            </span>
          </div>
        </button>

        {/* Card 5: Habis (Kosong) */}
        <button
          type="button"
          onClick={() => handleOpenModal('outOfStock')}
          className="p-3.5 bg-white rounded-xl border border-rose-200 shadow-xs hover:border-rose-400 hover:shadow-md transition-all text-left cursor-pointer group focus:outline-hidden focus:ring-2 focus:ring-rose-500"
          title="Klik untuk membuka popup daftar produk stok kosong (= 0)"
        >
          <div className="flex items-center justify-between text-rose-600 text-xs">
            <span className="font-semibold text-rose-800">Habis (Kosong)</span>
            <TrendingDown className="w-4 h-4 text-rose-500 group-hover:text-rose-700 transition-colors" />
          </div>
          <p className="text-xl font-bold text-rose-700 mt-1">
            {summary.outOfStockCount.toLocaleString('id-ID')}
          </p>
          <div className="flex items-center justify-between mt-2 pt-2 border-t border-rose-100/70 text-[11px]">
            <span className="text-rose-600 font-medium">Stok = 0</span>
            <span className="text-rose-700 font-medium inline-flex items-center gap-0.5">
              Lihat list <ChevronRight className="w-3 h-3 group-hover:translate-x-0.5 transition-transform" />
            </span>
          </div>
        </button>
      </div>

      {/* Table Container - Collapsible (Hidden by default based on user preference) */}
      {showDetailedTable && (
        <div className="bg-white rounded-xl border border-stone-200 shadow-xs overflow-hidden transition-all">
          {/* Filter, Search, and Action */}
          <div className="p-4 border-b border-stone-100 flex flex-col lg:flex-row lg:items-center justify-between gap-3">
          <div className="flex items-center gap-2">
            <Filter className="w-4 h-4 text-stone-400" />
            <div className="flex flex-wrap gap-1.5 text-xs">
              <button
                type="button"
                onClick={() => {
                  setFilterType('all');
                  setCurrentPage(1);
                }}
                className={`px-3 py-1.5 rounded-lg font-medium transition-colors ${
                  filterType === 'all'
                    ? 'bg-stone-900 text-white'
                    : 'bg-stone-100 text-stone-600 hover:bg-stone-200'
                }`}
              >
                Semua ({items.length})
              </button>
              <button
                type="button"
                onClick={() => {
                  setFilterType('matched');
                  setCurrentPage(1);
                }}
                className={`px-3 py-1.5 rounded-lg font-medium transition-colors ${
                  filterType === 'matched'
                    ? 'bg-emerald-600 text-white'
                    : 'bg-emerald-50 text-emerald-700 hover:bg-emerald-100'
                }`}
              >
                Cocok ({summary.matchedCount})
              </button>
              <button
                type="button"
                onClick={() => {
                  setFilterType('inStock');
                  setCurrentPage(1);
                }}
                className={`px-3 py-1.5 rounded-lg font-medium transition-colors ${
                  filterType === 'inStock'
                    ? 'bg-blue-600 text-white'
                    : 'bg-blue-50 text-blue-700 hover:bg-blue-100'
                }`}
              >
                Stok Ada ({summary.inStockCount})
              </button>
              <button
                type="button"
                onClick={() => {
                  setFilterType('outOfStock');
                  setCurrentPage(1);
                }}
                className={`px-3 py-1.5 rounded-lg font-medium transition-colors ${
                  filterType === 'outOfStock'
                    ? 'bg-rose-600 text-white'
                    : 'bg-rose-50 text-rose-700 hover:bg-rose-100'
                }`}
              >
                Stok 0 ({summary.outOfStockCount})
              </button>
              <button
                type="button"
                onClick={() => {
                  setFilterType('unmatched');
                  setCurrentPage(1);
                }}
                className={`px-3 py-1.5 rounded-lg font-medium transition-colors ${
                  filterType === 'unmatched'
                    ? 'bg-amber-600 text-white'
                    : 'bg-amber-50 text-amber-700 hover:bg-amber-100'
                }`}
              >
                Tidak Cocok ({summary.unmatchedCount})
              </button>
            </div>
          </div>

          <div className="flex items-center gap-2 w-full lg:w-auto">
            <div className="relative flex-1 lg:w-64">
              <Search className="w-4 h-4 text-stone-400 absolute left-3 top-2.5" />
              <input
                type="text"
                placeholder="Cari SKU atau nama produk..."
                value={searchTerm}
                onChange={(e) => {
                  setSearchTerm(e.target.value);
                  setCurrentPage(1);
                }}
                className="w-full pl-9 pr-3 py-1.5 text-xs bg-stone-50 border border-stone-200 rounded-lg text-stone-800 placeholder-stone-400 focus:outline-hidden focus:bg-white focus:border-emerald-500"
              />
            </div>

            <div className="flex items-center gap-1.5 shrink-0">
              {onRefreshStockList && (
                <button
                  type="button"
                  id="btn-refresh-stocklist-comparison"
                  onClick={onRefreshStockList}
                  disabled={isRefreshingStockList}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold text-emerald-800 hover:text-emerald-950 bg-emerald-50 hover:bg-emerald-100 border border-emerald-300 shadow-2xs transition-colors shrink-0 cursor-pointer disabled:opacity-50"
                  title="Segarkan data terbaru dari Sheet STOCK LIST di Google Sheets"
                >
                  <RefreshCw className={`w-3.5 h-3.5 ${isRefreshingStockList ? 'animate-spin' : ''}`} />
                  <span>{isRefreshingStockList ? 'Memperbarui...' : 'Refresh STOCK LIST'}</span>
                </button>
              )}

              <button
                type="button"
                onClick={() => {
                  handleExportComparison(exportPrefix);
                }}
                id="btn-export-comparison-xlsx"
                className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-2xs transition-colors shrink-0 cursor-pointer text-white bg-emerald-700 hover:bg-emerald-800"
                title={`Unduh file Excel format Mass Update Shopee (.xlsx) untuk toko ${exportPrefix}`}
              >
                <Download className="w-3.5 h-3.5" />
                <span>Unduh File Shopee (.xlsx)</span>
              </button>
            </div>
          </div>
        </div>

        {/* Informative helper note about column mapping */}
        <div className="px-4 py-2.5 bg-stone-50 border-b border-stone-200/80 text-[11px] text-stone-600 flex flex-wrap items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <span className="font-semibold text-emerald-800 bg-emerald-100/70 border border-emerald-200 px-1.5 py-0.5 rounded text-[10px]">
              KOLOM SHOPEE
            </span>
            <span>SKU dari Kolom 5 (E) atau Kolom 6 (F), SKU Induk (J)</span>
            <span className="text-stone-300">→</span>
            <span className="font-semibold text-blue-800 bg-blue-100/70 border border-blue-200 px-1.5 py-0.5 rounded text-[10px]">
              STOCK LIST
            </span>
            <span>SKU 5 Digit (Kolom 1), Barcode (Kolom 2) & Nama Barang (Kolom 3)</span>
          </div>
          <span className="text-stone-400">
            Total {items.length} baris dianalisis
          </span>
        </div>

        {/* Table Content */}
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-stone-50/90 border-b border-stone-200 text-stone-600 font-semibold uppercase tracking-wider text-[11px]">
                <th className="py-2.5 px-3 w-12 text-center">Baris</th>
                <th className="py-2.5 px-3.5 min-w-[220px]">Produk & Variasi di Shopee</th>
                <th className="py-2.5 px-3 min-w-[170px]">SKU Shopee (Kolom 5 / 6)</th>
                <th className="py-2.5 px-3 min-w-[150px]">SKU 5 Digit (Kolom 1 STOCK LIST)</th>
                <th className="py-2.5 px-4 min-w-[220px]">Nama Barang di STOCK LIST</th>
                <th className="py-2.5 px-3 text-center min-w-[120px]">Stok Gudang</th>
                <th className="py-2.5 px-4 min-w-[200px]">Catatan Pencocokan</th>
                <th className="py-2.5 px-3 text-center min-w-[90px]">Status</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-stone-100">
              {paginatedItems.length === 0 ? (
                <tr>
                  <td colSpan={8} className="py-10 text-center text-stone-400">
                    Tidak ada data yang sesuai dengan pencarian atau filter.
                  </td>
                </tr>
              ) : (
                paginatedItems.map((item) => {
                  const isZeroStock = item.stockQty !== null && item.stockQty === 0;
                  const displayShopeeName =
                    (item.rawRow && item.rawRow[1] !== undefined && item.rawRow[1] !== null && String(item.rawRow[1]).trim() !== '')
                      ? String(item.rawRow[1]).trim()
                      : (item.productName || (item.rawRow && item.rawRow[2] ? String(item.rawRow[2]).trim() : ''));
                  const displayVariation =
                    item.variationName ||
                    (item.rawRow && item.rawRow[3] ? String(item.rawRow[3]).trim() : (item.rawRow && item.rawRow[4] ? String(item.rawRow[4]).trim() : ''));
                  const s5 = item.skuCol5 || (item.rawRow && item.rawRow[4] ? String(item.rawRow[4]).trim() : '');
                  const s6 = item.skuCol6 || (item.rawRow && item.rawRow[5] ? String(item.rawRow[5]).trim() : '');
                  const displayParentSku =
                    item.parentSku ||
                    (item.rawRow && item.rawRow[9] ? String(item.rawRow[9]).trim() : '');

                  const effectiveSku = s5 || s6;
                  const effectiveSource = s5 ? 'Kolom 5 (E)' : (s6 ? 'Kolom 6 (F)' : '');

                  return (
                    <tr
                      key={item.rowIndex}
                      className={`hover:bg-stone-50/70 transition-colors ${
                        item.matchStatus === 'unmatched'
                          ? 'bg-amber-50/25'
                          : isZeroStock
                          ? 'bg-rose-50/20'
                          : ''
                      }`}
                    >
                      <td className="py-2.5 px-3 text-center text-stone-400 font-mono text-[11px]">
                        {item.rowIndex}
                      </td>

                      {/* Produk di Shopee (Nama & Variasi) */}
                      <td className="py-2.5 px-3.5">
                        <div className="font-medium text-stone-900 line-clamp-2" title={displayShopeeName}>
                          {displayShopeeName || <span className="text-stone-400 italic">Baris {item.rowIndex}</span>}
                        </div>
                        {displayVariation && (
                          <div className="mt-0.5 flex items-center gap-1">
                            <span className="text-[10px] text-stone-400">Variasi:</span>
                            <span className="inline-block px-1.5 py-0.2 rounded bg-stone-100 text-stone-700 text-[10px] font-medium border border-stone-200">
                              {displayVariation}
                            </span>
                          </div>
                        )}
                      </td>

                      {/* SKU Shopee (Kolom 5 / Kolom 6) */}
                      <td className="py-2.5 px-3">
                        {effectiveSku ? (
                          <div className="flex flex-col gap-0.5">
                            <div className="flex items-center gap-1.5">
                              <span className="font-mono font-bold text-stone-900 bg-stone-100 px-1.5 py-0.5 rounded border border-stone-200/80 text-xs">
                                {effectiveSku}
                              </span>
                              <span className="text-[9px] px-1 py-0.2 rounded bg-stone-200/60 text-stone-600 font-medium">
                                {effectiveSource}
                              </span>
                            </div>
                            {s6 && s5 && s6 !== s5 && (
                              <div className="text-[10px] text-stone-400 font-mono">
                                Kolom 6: {s6}
                              </div>
                            )}
                          </div>
                        ) : (
                          <span className="text-stone-300 italic">-</span>
                        )}
                        {displayParentSku && displayParentSku !== effectiveSku && (
                          <div className="text-[10px] text-stone-400 mt-1 font-mono">
                            Induk (J): {displayParentSku}
                          </div>
                        )}
                      </td>

                      {/* SKU 5 Digit di STOCK LIST (Kolom 1) */}
                      <td className="py-2.5 px-3">
                        {item.matchedStockItem?.code ? (
                          <div className="flex flex-col gap-0.5">
                            <span className="font-mono font-bold text-emerald-950 bg-emerald-50 border border-emerald-300/90 px-2 py-0.5 rounded text-xs inline-block shadow-2xs">
                              {item.matchedStockItem.code}
                            </span>
                            {item.matchedStockItem.barcode && item.matchedStockItem.barcode !== item.matchedStockItem.code && (
                              <span className="text-[10px] text-stone-400 font-mono" title={`Barcode: ${item.matchedStockItem.barcode}`}>
                                Barcode: {item.matchedStockItem.barcode}
                              </span>
                            )}
                          </div>
                        ) : (
                          <span className="text-amber-600/90 text-[11px] font-medium italic">Belum terdaftar</span>
                        )}
                      </td>

                      {/* Nama Barang di STOCK LIST */}
                      <td className="py-2.5 px-4">
                        {item.matchedStockItem ? (
                          <div>
                            <div className="text-stone-800 font-medium line-clamp-2" title={item.matchedStockItem.description}>
                              {item.matchedStockItem.description || '-'}
                            </div>
                            {(item.matchedStockItem.category || item.matchedStockItem.brand) && (
                              <div className="text-[10px] text-stone-400 mt-0.5">
                                {[item.matchedStockItem.category, item.matchedStockItem.brand].filter(Boolean).join(' • ')}
                              </div>
                            )}
                          </div>
                        ) : (
                          <span className="text-stone-400 italic text-[11px]">-</span>
                        )}
                      </td>

                      {/* Stok Gudang */}
                      <td className="py-2.5 px-3 text-center">
                        {item.stockQty !== null ? (
                          <span
                            className={`inline-flex items-center px-2 py-0.5 rounded-full font-bold font-mono text-xs ${
                              item.stockQty > 0
                                ? 'bg-emerald-100 text-emerald-800 border border-emerald-200'
                                : 'bg-rose-100 text-rose-800 border border-rose-200'
                            }`}
                          >
                            {item.stockQty.toLocaleString('id-ID')}
                          </span>
                        ) : (
                          <span className="text-stone-300 italic">-</span>
                        )}
                      </td>

                      {/* Catatan Pencocokan */}
                      <td className="py-2.5 px-4 text-[11px] text-stone-500">
                        {item.notes}
                      </td>

                      {/* Status */}
                      <td className="py-2.5 px-3 text-center">
                        {item.matchStatus === 'matched' ? (
                          <span className="inline-flex items-center gap-1 text-[11px] font-medium text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-full">
                            <CheckCircle2 className="w-3 h-3" />
                            Cocok
                          </span>
                        ) : (
                          <span className="inline-flex items-center gap-1 text-[11px] font-medium text-amber-700 bg-amber-50 border border-amber-200 px-2 py-0.5 rounded-full">
                            <AlertCircle className="w-3 h-3" />
                            Belum Ada
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

        {/* Pagination */}
        {filteredItems.length > pageSize && (
          <div className="p-3 border-t border-stone-100 flex items-center justify-between text-xs text-stone-500">
            <span>
              Menampilkan {Math.min((currentPage - 1) * pageSize + 1, filteredItems.length)} -{' '}
              {Math.min(currentPage * pageSize, filteredItems.length)} dari {filteredItems.length} data
            </span>
            <div className="flex items-center gap-1.5">
              <button
                type="button"
                disabled={currentPage === 1}
                onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                className="px-2.5 py-1 rounded border border-stone-200 bg-white hover:bg-stone-50 disabled:opacity-40 cursor-pointer"
              >
                Sebelumnya
              </button>
              <span className="px-2 font-medium text-stone-700">
                {currentPage} / {totalPages}
              </span>
              <button
                type="button"
                disabled={currentPage === totalPages}
                onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                className="px-2.5 py-1 rounded border border-stone-200 bg-white hover:bg-stone-50 disabled:opacity-40 cursor-pointer"
              >
                Berikutnya
              </button>
            </div>
          </div>
        )}
      </div>
      )}

      {/* Product List Popup Modal */}
      {modalCategory && (
        <BalistProductListModal
          isOpen={modalCategory !== null}
          onClose={() => setModalCategory(null)}
          category={modalCategory}
          items={items}
          balistSheetName={balistSheetName}
          stockSheetName={stockSheetName}
        />
      )}
    </div>
  );
};
