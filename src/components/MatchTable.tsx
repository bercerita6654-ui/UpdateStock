import React, { useState, useMemo } from 'react';
import { Search, CheckCircle2, AlertCircle, ArrowUpRight, ArrowDownRight, Minus, ChevronLeft, ChevronRight, Filter } from 'lucide-react';
import { ShopeeRowMatch } from '../types';

interface MatchTableProps {
  matches: ShopeeRowMatch[];
}

export const MatchTable: React.FC<MatchTableProps> = ({ matches }) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [filterType, setFilterType] = useState<'all' | 'matched' | 'unmatched' | 'changed'>('all');
  const [currentPage, setCurrentPage] = useState(1);
  const pageSize = 50;

  const filteredMatches = useMemo(() => {
    return matches.filter((item) => {
      // Filter by status
      if (filterType === 'matched' && item.matchStatus !== 'matched') return false;
      if (filterType === 'unmatched' && item.matchStatus !== 'unmatched') return false;
      if (filterType === 'changed' && (item.stockDiff === 0 || item.newStock === null)) return false;

      // Filter by search term
      if (searchTerm.trim()) {
        const q = searchTerm.toLowerCase();
        const skuMatch = item.rawSku.toLowerCase().includes(q);
        const nameMatch = item.productName?.toLowerCase().includes(q);
        const varMatch = item.variationName?.toLowerCase().includes(q);
        const notesMatch = item.notes?.toLowerCase().includes(q);
        return skuMatch || nameMatch || varMatch || notesMatch;
      }
      return true;
    });
  }, [matches, filterType, searchTerm]);

  const totalPages = Math.ceil(filteredMatches.length / pageSize) || 1;
  const paginatedItems = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredMatches.slice(start, start + pageSize);
  }, [filteredMatches, currentPage]);

  const handleFilterChange = (type: 'all' | 'matched' | 'unmatched' | 'changed') => {
    setFilterType(type);
    setCurrentPage(1);
  };

  return (
    <div className="bg-white rounded-xl border border-stone-200 shadow-xs overflow-hidden">
      {/* Table Header & Controls */}
      <div className="p-4 border-b border-stone-100 flex flex-col md:flex-row md:items-center justify-between gap-3">
        <div className="flex items-center gap-2">
          <Filter className="w-4 h-4 text-stone-400" />
          <div className="flex flex-wrap gap-1.5 text-xs">
            <button
              type="button"
              onClick={() => handleFilterChange('all')}
              className={`px-3 py-1.5 rounded-lg font-medium transition-colors ${
                filterType === 'all'
                  ? 'bg-stone-900 text-white'
                  : 'bg-stone-100 text-stone-600 hover:bg-stone-200'
              }`}
            >
              Semua ({matches.length})
            </button>
            <button
              type="button"
              onClick={() => handleFilterChange('matched')}
              className={`px-3 py-1.5 rounded-lg font-medium transition-colors ${
                filterType === 'matched'
                  ? 'bg-emerald-600 text-white'
                  : 'bg-emerald-50 text-emerald-700 hover:bg-emerald-100'
              }`}
            >
              Cocok ({matches.filter((m) => m.matchStatus === 'matched').length})
            </button>
            <button
              type="button"
              onClick={() => handleFilterChange('changed')}
              className={`px-3 py-1.5 rounded-lg font-medium transition-colors ${
                filterType === 'changed'
                  ? 'bg-blue-600 text-white'
                  : 'bg-blue-50 text-blue-700 hover:bg-blue-100'
              }`}
            >
              Stok Berubah ({matches.filter((m) => m.stockDiff !== 0 && m.newStock !== null).length})
            </button>
            <button
              type="button"
              onClick={() => handleFilterChange('unmatched')}
              className={`px-3 py-1.5 rounded-lg font-medium transition-colors ${
                filterType === 'unmatched'
                  ? 'bg-amber-600 text-white'
                  : 'bg-amber-50 text-amber-700 hover:bg-amber-100'
              }`}
            >
              Tidak Cocok ({matches.filter((m) => m.matchStatus === 'unmatched').length})
            </button>
          </div>
        </div>

        {/* Search input */}
        <div className="relative w-full md:w-72">
          <Search className="w-4 h-4 text-stone-400 absolute left-3 top-2.5" />
          <input
            type="text"
            placeholder="Cari SKU atau nama produk..."
            value={searchTerm}
            onChange={(e) => {
              setSearchTerm(e.target.value);
              setCurrentPage(1);
            }}
            className="w-full pl-9 pr-3 py-1.5 text-xs bg-stone-50 border border-stone-200 rounded-lg text-stone-800 placeholder-stone-400 focus:outline-hidden focus:bg-white focus:border-orange-500"
          />
        </div>
      </div>

      {/* Table Content */}
      <div className="overflow-x-auto">
        <table className="w-full text-left border-collapse text-xs">
          <thead>
            <tr className="bg-stone-50/80 border-b border-stone-200 text-stone-600 font-semibold uppercase tracking-wider text-[11px]">
              <th className="py-2.5 px-3 w-12 text-center">Baris</th>
              <th className="py-2.5 px-4">SKU Shopee</th>
              <th className="py-2.5 px-4">Nama Produk / Variasi</th>
              <th className="py-2.5 px-4 text-center">Stok Lama</th>
              <th className="py-2.5 px-4 text-center">Stok Baru (STOCK LIST)</th>
              <th className="py-2.5 px-4 text-center">Selisih</th>
              <th className="py-2.5 px-4">Keterangan / Jalur Acuan</th>
              <th className="py-2.5 px-4 text-center">Status</th>
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
                const isChanged = item.stockDiff !== 0 && item.newStock !== null;
                const isDiffPositive = (item.stockDiff || 0) > 0;

                return (
                  <tr
                    key={item.rowIndex}
                    className={`hover:bg-stone-50/60 transition-colors ${
                      item.matchStatus === 'unmatched'
                        ? 'bg-amber-50/20'
                        : isChanged
                        ? 'bg-blue-50/10'
                        : ''
                    }`}
                  >
                    <td className="py-2.5 px-3 text-center text-stone-400 font-mono">
                      {item.displayRow}
                    </td>
                    <td className="py-2.5 px-4 font-mono font-medium text-stone-900">
                      {item.rawSku || <span className="text-stone-300 italic">(Kosong)</span>}
                    </td>
                    <td className="py-2.5 px-4 text-stone-700 max-w-xs truncate">
                      {item.productName || item.variationName || '-'}
                      {item.variationName && item.productName && (
                        <span className="block text-[10px] text-stone-400">
                          Variasi: {item.variationName}
                        </span>
                      )}
                    </td>
                    <td className="py-2.5 px-4 text-center font-mono text-stone-600">
                      {item.originalStock}
                    </td>
                    <td className="py-2.5 px-4 text-center font-mono font-bold">
                      {item.newStock !== null ? (
                        <span
                          className={
                            item.newStock === 0
                              ? 'text-rose-600 bg-rose-50 px-2 py-0.5 rounded border border-rose-200'
                              : 'text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200'
                          }
                        >
                          {item.newStock}
                        </span>
                      ) : (
                        <span className="text-stone-400">-</span>
                      )}
                    </td>
                    <td className="py-2.5 px-4 text-center font-mono">
                      {item.newStock === null ? (
                        <span className="text-stone-400">-</span>
                      ) : item.stockDiff === 0 ? (
                        <span className="inline-flex items-center text-stone-400">
                          <Minus className="w-3 h-3" />
                        </span>
                      ) : isDiffPositive ? (
                        <span className="inline-flex items-center gap-0.5 text-emerald-600 font-semibold">
                          <ArrowUpRight className="w-3 h-3" />+{item.stockDiff}
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-0.5 text-rose-600 font-semibold">
                          <ArrowDownRight className="w-3 h-3" />
                          {item.stockDiff}
                        </span>
                      )}
                    </td>
                    <td className="py-2.5 px-4 text-stone-600">
                      <span className="line-clamp-1" title={item.notes}>
                        {item.notes || '-'}
                      </span>
                    </td>
                    <td className="py-2.5 px-4 text-center">
                      {item.matchStatus === 'matched' ? (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-medium bg-emerald-50 text-emerald-700 border border-emerald-200">
                          <CheckCircle2 className="w-3 h-3" /> Cocok
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-medium bg-amber-50 text-amber-700 border border-amber-200">
                          <AlertCircle className="w-3 h-3" /> Tidak Ada
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

      {/* Pagination Bar */}
      {totalPages > 1 && (
        <div className="p-3 border-t border-stone-100 flex items-center justify-between text-xs text-stone-500">
          <span>
            Menampilkan {(currentPage - 1) * pageSize + 1} -{' '}
            {Math.min(currentPage * pageSize, filteredMatches.length)} dari{' '}
            {filteredMatches.length} baris
          </span>

          <div className="flex items-center gap-1">
            <button
              type="button"
              disabled={currentPage <= 1}
              onClick={() => setCurrentPage((p) => Math.max(p - 1, 1))}
              className="p-1 rounded hover:bg-stone-100 disabled:opacity-30"
            >
              <ChevronLeft className="w-4 h-4" />
            </button>
            <span className="px-2 font-medium text-stone-700">
              Halaman {currentPage} dari {totalPages}
            </span>
            <button
              type="button"
              disabled={currentPage >= totalPages}
              onClick={() => setCurrentPage((p) => Math.min(p + 1, totalPages))}
              className="p-1 rounded hover:bg-stone-100 disabled:opacity-30"
            >
              <ChevronRight className="w-4 h-4" />
            </button>
          </div>
        </div>
      )}
    </div>
  );
};
