import React, { useState, useMemo } from 'react';
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
} from 'lucide-react';
import { BalistComparisonItem, BalistComparisonSummary } from '../types';
import * as XLSX from 'xlsx';
import { downloadBlob, generateShopeeBalistFilename } from '../lib/excelProcessor';
import { parseNumber } from '../lib/sheets';
import { BalistProductListModal, BalistModalCategory } from './BalistProductListModal';

interface BalistComparisonTableProps {
  items: BalistComparisonItem[];
  summary: BalistComparisonSummary;
  balistSheetName: string;
  stockSheetName: string;
  onUpdateBalistStockInSheet?: () => void;
  isUpdatingBalistStock?: boolean;
}

export const BalistComparisonTable: React.FC<BalistComparisonTableProps> = ({
  items,
  summary,
  balistSheetName,
  stockSheetName,
  onUpdateBalistStockInSheet,
  isUpdatingBalistStock,
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [filterType, setFilterType] = useState<
    'all' | 'matched' | 'unmatched' | 'inStock' | 'outOfStock'
  >('all');
  const [currentPage, setCurrentPage] = useState(1);
  const [modalCategory, setModalCategory] = useState<BalistModalCategory | null>(null);
  const pageSize = 50;

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
        const sku5Match = item.skuCol5.toLowerCase().includes(q);
        const sku6Match = item.skuCol6.toLowerCase().includes(q);
        const codeMatch = item.matchedStockItem?.code.toLowerCase().includes(q);
        const descMatch = item.matchedStockItem?.description?.toLowerCase().includes(q);
        const notesMatch = item.notes?.toLowerCase().includes(q);
        return sku5Match || sku6Match || codeMatch || descMatch || notesMatch;
      }
      return true;
    });
  }, [items, filterType, searchTerm]);

  const totalPages = Math.ceil(filteredItems.length / pageSize) || 1;
  const paginatedItems = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredItems.slice(start, start + pageSize);
  }, [filteredItems, currentPage]);

  const handleExportComparison = () => {
    // Header standard Shopee Mass Update format
    const headers = [
      'No',
      'Nama Produk',
      'No. Variasi',
      'Nama Variasi',
      'Kode Variasi (SKU)',
      'SKU Induk / Referensi',
      'Stok',
      'Status Stok',
      'Keterangan',
    ];

    const rows = items.map((it, idx) => {
      const productName = it.matchedStockItem?.description || (it.rawRow && it.rawRow[1] ? String(it.rawRow[1]) : `Produk SKU ${it.skuCol5 || it.skuCol6}`);
      const variationSku = it.skuCol6 || it.skuCol5;
      const parentSku = it.skuCol5 || it.skuCol6;
      
      // Stock quantity from STOCK LIST if matched, otherwise 0 or original stock
      const finalStock = it.stockQty !== null ? it.stockQty : (it.rawRow && it.rawRow[6] !== undefined && it.rawRow[6] !== '' ? parseNumber(it.rawRow[6]) : 0);
      const stockStatus = finalStock > 0 ? 'Tersedia' : 'Habis';

      return [
        idx + 1,
        productName,
        it.rowIndex,
        it.skuCol6 ? `Varian (${it.skuCol6})` : 'Standar',
        variationSku,
        parentSku,
        finalStock,
        stockStatus,
        it.notes || (it.matchStatus === 'matched' ? 'Stok Terupdate' : 'SKU Belum Terdaftar di Gudang'),
      ];
    });

    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);

    // Set column widths for clean look
    ws['!cols'] = [
      { wch: 6 },
      { wch: 38 },
      { wch: 12 },
      { wch: 20 },
      { wch: 22 },
      { wch: 22 },
      { wch: 12 },
      { wch: 14 },
      { wch: 35 },
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Shopee Stock Update');

    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const fileName = generateShopeeBalistFilename('shopee_balist');
    downloadBlob(new Uint8Array(out), fileName);
  };

  return (
    <div className="space-y-4">
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

      {/* Table Container */}
      <div className="bg-white rounded-xl border border-stone-200 shadow-xs overflow-hidden">
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

            <button
              type="button"
              onClick={handleExportComparison}
              id="btn-export-comparison-xlsx"
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold text-white bg-emerald-700 hover:bg-emerald-800 shadow-2xs transition-colors shrink-0 cursor-pointer"
              title="Download file Excel (.xlsx) format Shopee"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Download XLSX</span>
            </button>
          </div>
        </div>

        {/* Table Content */}
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-stone-50/80 border-b border-stone-200 text-stone-600 font-semibold uppercase tracking-wider text-[11px]">
                <th className="py-2.5 px-3 w-12 text-center">Baris</th>
                <th className="py-2.5 px-3">SKU Kolom 5 (E)</th>
                <th className="py-2.5 px-3">SKU Kolom 6 (F)</th>
                <th className="py-2.5 px-3">Kode di STOCK LIST</th>
                <th className="py-2.5 px-4">Deskripsi Produk</th>
                <th className="py-2.5 px-4 text-center">Jumlah Stok (STOCK LIST)</th>
                <th className="py-2.5 px-4">Catatan Pencocokan</th>
                <th className="py-2.5 px-3 text-center">Status</th>
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
                  const hasStock = item.stockQty !== null && item.stockQty > 0;
                  const isZeroStock = item.stockQty !== null && item.stockQty === 0;

                  return (
                    <tr
                      key={item.rowIndex}
                      className={`hover:bg-stone-50/60 transition-colors ${
                        item.matchStatus === 'unmatched'
                          ? 'bg-amber-50/20'
                          : isZeroStock
                          ? 'bg-rose-50/20'
                          : ''
                      }`}
                    >
                      <td className="py-2.5 px-3 text-center text-stone-400 font-mono">
                        {item.rowIndex}
                      </td>
                      <td className="py-2.5 px-3 font-mono font-medium text-stone-900">
                        {item.skuCol5 || <span className="text-stone-300 italic">-</span>}
                      </td>
                      <td className="py-2.5 px-3 font-mono text-stone-700">
                        {item.skuCol6 || <span className="text-stone-300 italic">-</span>}
                      </td>
                      <td className="py-2.5 px-3 font-mono text-stone-800">
                        {item.matchedStockItem?.code || (
                          <span className="text-stone-400 italic">Tidak ditemukan</span>
                        )}
                      </td>
                      <td className="py-2.5 px-4 text-stone-700 max-w-xs truncate">
                        {item.matchedStockItem?.description || (
                          <span className="text-stone-400 italic">-</span>
                        )}
                      </td>
                      <td className="py-2.5 px-4 text-center">
                        {item.stockQty !== null ? (
                          <span
                            className={`inline-flex items-center px-2 py-0.5 rounded-full font-bold font-mono text-xs ${
                              item.stockQty > 0
                                ? 'bg-emerald-100 text-emerald-800'
                                : 'bg-rose-100 text-rose-800'
                            }`}
                          >
                            {item.stockQty.toLocaleString('id-ID')}
                          </span>
                        ) : (
                          <span className="text-stone-300 italic">-</span>
                        )}
                      </td>
                      <td className="py-2.5 px-4 text-[11px] text-stone-500">
                        {item.notes}
                      </td>
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
