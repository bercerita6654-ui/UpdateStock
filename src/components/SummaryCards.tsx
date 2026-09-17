import React from 'react';
import { Download, CheckCircle, AlertCircle, ArrowUpDown, Package, FileCheck, Send } from 'lucide-react';
import { ProcessSummary } from '../types';

interface SummaryCardsProps {
  summary: ProcessSummary;
  onDownload: () => void;
  onSyncBalist: () => void;
  isSyncingBalist: boolean;
  canSyncBalist: boolean;
}

export const SummaryCards: React.FC<SummaryCardsProps> = ({
  summary,
  onDownload,
  onSyncBalist,
  isSyncingBalist,
  canSyncBalist,
}) => {
  return (
    <div className="space-y-4">
      {/* Metric Grid */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <div className="p-3.5 bg-white rounded-xl border border-stone-200 shadow-xs">
          <div className="flex items-center justify-between text-stone-500 text-xs">
            <span>Total Produk</span>
            <Package className="w-4 h-4 text-stone-400" />
          </div>
          <p className="text-xl font-bold text-stone-900 mt-1">
            {summary.totalRows.toLocaleString('id-ID')}
          </p>
          <span className="text-[11px] text-stone-400">Baris dari file Shopee</span>
        </div>

        <div className="p-3.5 bg-white rounded-xl border border-stone-200 shadow-xs">
          <div className="flex items-center justify-between text-emerald-600 text-xs">
            <span>SKU Cocok</span>
            <CheckCircle className="w-4 h-4 text-emerald-500" />
          </div>
          <p className="text-xl font-bold text-emerald-700 mt-1">
            {summary.matchedCount.toLocaleString('id-ID')}
          </p>
          <span className="text-[11px] text-emerald-600 font-medium">
            {summary.totalRows > 0
              ? `${Math.round((summary.matchedCount / summary.totalRows) * 100)}% teridentifikasi`
              : '0%'}
          </span>
        </div>

        <div className="p-3.5 bg-white rounded-xl border border-stone-200 shadow-xs">
          <div className="flex items-center justify-between text-amber-600 text-xs">
            <span>SKU Belum Cocok</span>
            <AlertCircle className="w-4 h-4 text-amber-500" />
          </div>
          <p className="text-xl font-bold text-amber-700 mt-1">
            {summary.unmatchedCount.toLocaleString('id-ID')}
          </p>
          <span className="text-[11px] text-amber-600">Perlu cek SKU toko</span>
        </div>

        <div className="p-3.5 bg-white rounded-xl border border-stone-200 shadow-xs">
          <div className="flex items-center justify-between text-blue-600 text-xs">
            <span>Perubahan Stok</span>
            <ArrowUpDown className="w-4 h-4 text-blue-500" />
          </div>
          <p className="text-xl font-bold text-blue-700 mt-1">
            {summary.stockChangedCount.toLocaleString('id-ID')}
          </p>
          <span className="text-[11px] text-blue-600 font-medium">
            Stok gudang berbeda
          </span>
        </div>
      </div>

      {/* Action Bar */}
      <div className="p-4 bg-orange-50/80 border border-orange-200 rounded-xl flex flex-col sm:flex-row items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-full bg-orange-600 text-white flex items-center justify-center shrink-0 shadow-xs">
            <FileCheck className="w-5 h-5" />
          </div>
          <div>
            <h3 className="text-sm font-semibold text-stone-900">
              Hasil Pencocokan Siap Diunduh
            </h3>
            <p className="text-xs text-stone-600">
              File baru diformat identik dengan format asli Shopee untuk diunggah langsung ke Seller Centre.
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2.5 w-full sm:w-auto justify-end">
          {canSyncBalist && (
            <button
              type="button"
              id="btn-sync-balistshopee"
              onClick={onSyncBalist}
              disabled={isSyncingBalist}
              className="px-3.5 py-2.5 rounded-lg text-xs font-semibold text-stone-700 bg-white hover:bg-stone-50 border border-stone-300 transition-colors shadow-xs flex items-center gap-1.5 disabled:opacity-50"
            >
              <Send className="w-3.5 h-3.5 text-orange-600" />
              {isSyncingBalist ? 'Menyinkronkan...' : 'Perbarui di Balistshopee'}
            </button>
          )}

          <button
            type="button"
            id="btn-download-xlsx"
            onClick={onDownload}
            className="w-full sm:w-auto px-5 py-2.5 rounded-lg text-xs font-semibold text-white bg-emerald-600 hover:bg-emerald-700 transition-colors shadow-sm flex items-center justify-center gap-2 cursor-pointer"
          >
            <Download className="w-4 h-4" />
            Download XLSX Terupdate
          </button>
        </div>
      </div>
    </div>
  );
};
