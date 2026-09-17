import React, { useState, useMemo } from 'react';
import {
  X,
  Search,
  Download,
  CheckCircle2,
  AlertCircle,
  TrendingUp,
  TrendingDown,
  Layers,
  Package,
} from 'lucide-react';
import { BalistComparisonItem } from '../types';
import * as XLSX from 'xlsx';
import { downloadBlob, generateShopeeBalistFilename } from '../lib/excelProcessor';
import { parseNumber } from '../lib/sheets';

export type BalistModalCategory =
  | 'all'
  | 'matched'
  | 'unmatched'
  | 'inStock'
  | 'outOfStock';

interface BalistProductListModalProps {
  isOpen: boolean;
  onClose: () => void;
  category: BalistModalCategory;
  items: BalistComparisonItem[];
  balistSheetName: string;
  stockSheetName: string;
}

export const BalistProductListModal: React.FC<BalistProductListModalProps> = ({
  isOpen,
  onClose,
  category,
  items,
  balistSheetName,
  stockSheetName,
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const pageSize = 50;

  // Filter items based on category
  const categoryItems = useMemo(() => {
    return items.filter((item) => {
      if (category === 'matched') return item.matchStatus === 'matched';
      if (category === 'unmatched') return item.matchStatus === 'unmatched';
      if (category === 'inStock')
        return item.stockQty !== null && item.stockQty > 0;
      if (category === 'outOfStock')
        return item.stockQty !== null && item.stockQty === 0;
      return true; // 'all'
    });
  }, [items, category]);

  // Apply search query within category
  const filteredItems = useMemo(() => {
    if (!searchTerm.trim()) return categoryItems;
    const q = searchTerm.toLowerCase();
    return categoryItems.filter((item) => {
      const sku5Match = item.skuCol5.toLowerCase().includes(q);
      const sku6Match = item.skuCol6.toLowerCase().includes(q);
      const codeMatch = item.matchedStockItem?.code?.toLowerCase().includes(q);
      const descMatch = item.matchedStockItem?.description?.toLowerCase().includes(q);
      const notesMatch = item.notes?.toLowerCase().includes(q);
      const rowMatch = String(item.rowIndex).includes(q);
      return sku5Match || sku6Match || codeMatch || descMatch || notesMatch || rowMatch;
    });
  }, [categoryItems, searchTerm]);

  const totalPages = Math.ceil(filteredItems.length / pageSize) || 1;
  const paginatedItems = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredItems.slice(start, start + pageSize);
  }, [filteredItems, currentPage]);

  const categoryConfig = useMemo(() => {
    switch (category) {
      case 'matched':
        return {
          title: 'Daftar Produk: Cocok di STOCK LIST',
          subtitle: `Menampilkan semua produk dari sheet ${balistSheetName} yang SKU-nya cocok di sheet ${stockSheetName}`,
          icon: CheckCircle2,
          iconBg: 'bg-emerald-100 text-emerald-700 border-emerald-300',
          badgeBg: 'bg-emerald-100 text-emerald-800 border-emerald-200',
          exportNamePrefix: 'produk_cocok_stock_list',
        };
      case 'unmatched':
        return {
          title: 'Daftar Produk: Tidak Ditemukan di STOCK LIST',
          subtitle: `Produk dari sheet ${balistSheetName} yang kode SKU (Kolom 5 & 6)-nya belum ada di sheet ${stockSheetName}`,
          icon: AlertCircle,
          iconBg: 'bg-amber-100 text-amber-700 border-amber-300',
          badgeBg: 'bg-amber-100 text-amber-800 border-amber-200',
          exportNamePrefix: 'produk_tidak_ditemukan',
        };
      case 'inStock':
        return {
          title: 'Daftar Produk: Stok Tersedia (Ready > 0)',
          subtitle: `Produk yang memiliki saldo kuantitas stok lebih dari 0 pada Kolom 15 STOCK LIST`,
          icon: TrendingUp,
          iconBg: 'bg-blue-100 text-blue-700 border-blue-300',
          badgeBg: 'bg-blue-100 text-blue-800 border-blue-200',
          exportNamePrefix: 'produk_stok_tersedia',
        };
      case 'outOfStock':
        return {
          title: 'Daftar Produk: Stok Habis (Kosong = 0)',
          subtitle: `Produk yang kuantitas stoknya bernilai 0 di gudang STOCK LIST`,
          icon: TrendingDown,
          iconBg: 'bg-rose-100 text-rose-700 border-rose-300',
          badgeBg: 'bg-rose-100 text-rose-800 border-rose-200',
          exportNamePrefix: 'produk_stok_habis',
        };
      case 'all':
      default:
        return {
          title: `Daftar Semua Produk: Sheet ${balistSheetName}`,
          subtitle: `Seluruh baris item produk yang dimuat dari sheet ${balistSheetName}`,
          icon: Layers,
          iconBg: 'bg-stone-100 text-stone-700 border-stone-300',
          badgeBg: 'bg-stone-100 text-stone-800 border-stone-200',
          exportNamePrefix: 'semua_produk_balist',
        };
    }
  }, [category, balistSheetName, stockSheetName]);

  const IconComponent = categoryConfig.icon;

  const totalQty = useMemo(() => {
    return filteredItems.reduce((acc, it) => acc + (it.stockQty || 0), 0);
  }, [filteredItems]);

  const handleExportCategoryXlsx = () => {
    const headers = [
      'No',
      'Baris Sheet',
      'SKU Kolom 5 (E)',
      'SKU Kolom 6 (F)',
      'Kode STOCK LIST',
      'Nama / Deskripsi Produk',
      'Jumlah Stok',
      'Status Pencocokan',
      'Catatan',
    ];

    const rows = filteredItems.map((it, idx) => [
      idx + 1,
      it.rowIndex,
      it.skuCol5 || '-',
      it.skuCol6 || '-',
      it.matchedStockItem?.code || 'Tidak Ada',
      it.matchedStockItem?.description || (it.rawRow && it.rawRow[1] ? String(it.rawRow[1]) : '-'),
      it.stockQty !== null ? it.stockQty : 0,
      it.matchStatus === 'matched' ? 'Cocok' : 'Tidak Ditemukan',
      it.notes || '',
    ]);

    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    ws['!cols'] = [
      { wch: 6 },
      { wch: 12 },
      { wch: 22 },
      { wch: 22 },
      { wch: 22 },
      { wch: 38 },
      { wch: 14 },
      { wch: 18 },
      { wch: 30 },
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Daftar Produk');

    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const fileName = generateShopeeBalistFilename(categoryConfig.exportNamePrefix);
    downloadBlob(new Uint8Array(out), fileName);
  };

  if (!isOpen) return null;

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-5 bg-stone-900/60 backdrop-blur-xs animate-fadeIn"
      role="dialog"
      aria-modal="true"
      onClick={onClose}
    >
      <div
        className="bg-white rounded-2xl border border-stone-200 shadow-2xl w-full max-w-5xl max-h-[90vh] flex flex-col overflow-hidden animate-scaleUp"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header Modal */}
        <div className="p-4 sm:p-5 border-b border-stone-200 flex items-start justify-between gap-4 bg-stone-50/70">
          <div className="flex items-start gap-3">
            <div
              className={`w-10 h-10 rounded-xl flex items-center justify-center shrink-0 border ${categoryConfig.iconBg}`}
            >
              <IconComponent className="w-5 h-5" />
            </div>
            <div>
              <div className="flex flex-wrap items-center gap-2">
                <h2 className="text-base sm:text-lg font-bold text-stone-900">
                  {categoryConfig.title}
                </h2>
                <span
                  className={`px-2.5 py-0.5 rounded-full text-xs font-bold border ${categoryConfig.badgeBg}`}
                >
                  {categoryItems.length.toLocaleString('id-ID')} Produk
                </span>
              </div>
              <p className="text-xs text-stone-500 mt-1 max-w-2xl">
                {categoryConfig.subtitle}
              </p>
            </div>
          </div>

          <button
            type="button"
            onClick={onClose}
            className="p-1.5 rounded-lg text-stone-400 hover:text-stone-700 hover:bg-stone-200/60 transition-colors cursor-pointer"
            aria-label="Tutup"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Toolbar: Search & Export */}
        <div className="p-3 sm:p-4 border-b border-stone-100 flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-white">
          <div className="relative flex-1 max-w-md">
            <Search className="w-4 h-4 text-stone-400 absolute left-3 top-2.5" />
            <input
              type="text"
              placeholder="Cari berdasarkan SKU, Baris, Nama Produk, atau Kode..."
              value={searchTerm}
              onChange={(e) => {
                setSearchTerm(e.target.value);
                setCurrentPage(1);
              }}
              className="w-full pl-9 pr-3 py-1.5 text-xs bg-stone-50 border border-stone-200 rounded-lg text-stone-800 placeholder-stone-400 focus:outline-hidden focus:bg-white focus:border-emerald-500"
              autoFocus
            />
          </div>

          <div className="flex items-center gap-2">
            <span className="text-xs text-stone-500 font-medium hidden sm:inline">
              Ditemukan: <strong>{filteredItems.length.toLocaleString('id-ID')}</strong> item
            </span>
            <button
              type="button"
              onClick={handleExportCategoryXlsx}
              disabled={filteredItems.length === 0}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold text-white bg-emerald-700 hover:bg-emerald-800 shadow-2xs transition-colors cursor-pointer disabled:opacity-50 shrink-0"
              title="Download daftar produk ini ke file Excel (.xlsx)"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Download Excel (.xlsx)</span>
            </button>
          </div>
        </div>

        {/* Product Table List */}
        <div className="overflow-y-auto flex-1 max-h-[58vh]">
          <table className="w-full text-left border-collapse text-xs">
            <thead className="sticky top-0 bg-stone-100 z-10 border-b border-stone-200 text-stone-700 font-semibold uppercase tracking-wider text-[11px]">
              <tr>
                <th className="py-2.5 px-3 w-12 text-center">No</th>
                <th className="py-2.5 px-3 w-16 text-center">Baris</th>
                <th className="py-2.5 px-3">SKU Kolom 5 (E)</th>
                <th className="py-2.5 px-3">SKU Kolom 6 (F)</th>
                <th className="py-2.5 px-3">Kode STOCK LIST</th>
                <th className="py-2.5 px-4">Nama / Deskripsi Produk</th>
                <th className="py-2.5 px-3 text-center">Stok Gudang</th>
                <th className="py-2.5 px-3 text-center">Status</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-stone-100">
              {paginatedItems.length === 0 ? (
                <tr>
                  <td colSpan={8} className="py-12 text-center text-stone-400">
                    <Package className="w-8 h-8 text-stone-300 mx-auto mb-2" />
                    <p className="font-medium text-stone-600">Tidak ada produk ditemukan</p>
                    <p className="text-xs text-stone-400 mt-0.5">
                      Coba ganti kata kunci pencarian Anda.
                    </p>
                  </td>
                </tr>
              ) : (
                paginatedItems.map((item, idx) => {
                  const globalIdx = (currentPage - 1) * pageSize + idx + 1;
                  const isZeroStock = item.stockQty !== null && item.stockQty === 0;

                  return (
                    <tr
                      key={`${item.rowIndex}-${idx}`}
                      className={`hover:bg-stone-50/80 transition-colors ${
                        item.matchStatus === 'unmatched'
                          ? 'bg-amber-50/20'
                          : isZeroStock
                          ? 'bg-rose-50/20'
                          : ''
                      }`}
                    >
                      <td className="py-2.5 px-3 text-center text-stone-400 font-mono text-[11px]">
                        {globalIdx}
                      </td>
                      <td className="py-2.5 px-3 text-center text-stone-500 font-mono font-medium">
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
                          <span className="text-amber-600 italic">Tidak ditemukan</span>
                        )}
                      </td>
                      <td className="py-2.5 px-4 text-stone-700 max-w-sm truncate">
                        {item.matchedStockItem?.description ||
                          (item.rawRow && item.rawRow[1] ? String(item.rawRow[1]) : (
                            <span className="text-stone-300 italic">-</span>
                          ))}
                      </td>
                      <td className="py-2.5 px-3 text-center">
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
                          <span className="text-stone-400 italic">0</span>
                        )}
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

        {/* Modal Footer with Pagination and Summary */}
        <div className="p-3 sm:p-4 border-t border-stone-200 bg-stone-50 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs">
          <div className="flex flex-wrap items-center gap-3 text-stone-600">
            <span>
              Total: <strong>{filteredItems.length.toLocaleString('id-ID')}</strong> produk
            </span>
            <span>•</span>
            <span>
              Total Akumulasi Stok: <strong className="text-stone-900 font-mono">{totalQty.toLocaleString('id-ID')} pcs</strong>
            </span>
          </div>

          <div className="flex items-center justify-between sm:justify-end gap-3">
            {filteredItems.length > pageSize && (
              <div className="flex items-center gap-1.5">
                <button
                  type="button"
                  disabled={currentPage === 1}
                  onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                  className="px-2.5 py-1 rounded border border-stone-300 bg-white hover:bg-stone-100 disabled:opacity-40 cursor-pointer"
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
                  className="px-2.5 py-1 rounded border border-stone-300 bg-white hover:bg-stone-100 disabled:opacity-40 cursor-pointer"
                >
                  Berikutnya
                </button>
              </div>
            )}

            <button
              type="button"
              onClick={onClose}
              className="px-4 py-1.5 rounded-lg border border-stone-300 bg-white hover:bg-stone-100 text-stone-700 font-medium transition-colors cursor-pointer"
            >
              Tutup
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
