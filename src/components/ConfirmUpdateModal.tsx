import React from 'react';
import { AlertTriangle, CheckCircle, X, ShieldAlert } from 'lucide-react';

interface ConfirmUpdateModalProps {
  isOpen: boolean;
  onClose: () => void;
  onConfirm: () => void;
  isProcessing: boolean;
  spreadsheetName: string;
  spreadsheetId: string;
  matchedCount: number;
  title?: string;
  description?: string;
  actionText?: string;
}

export const ConfirmUpdateModal: React.FC<ConfirmUpdateModalProps> = ({
  isOpen,
  onClose,
  onConfirm,
  isProcessing,
  spreadsheetName,
  spreadsheetId,
  matchedCount,
  title = 'Konfirmasi Pembaruan Data Spreadsheet',
  description,
  actionText = 'Pembaruan nilai data pada sheet target',
}) => {
  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-stone-900/50 backdrop-blur-xs">
      <div className="bg-white rounded-2xl border border-stone-200 shadow-2xl max-w-md w-full overflow-hidden animate-in fade-in zoom-in-95 duration-150">
        <div className="p-5">
          <div className="w-12 h-12 rounded-full bg-amber-50 border border-amber-200 text-amber-600 flex items-center justify-center mb-4 mx-auto">
            <ShieldAlert className="w-6 h-6" />
          </div>

          <h3 className="text-base font-semibold text-stone-900 text-center">
            {title}
          </h3>

          <p className="text-xs text-stone-600 text-center mt-2 leading-relaxed">
            {description || (
              <>
                Apakah Anda yakin ingin memperbarui data pada spreadsheet{' '}
                <strong className="text-stone-800">"{spreadsheetName}"</strong>?
              </>
            )}
          </p>

          <div className="mt-4 p-3 bg-stone-50 border border-stone-200 rounded-lg text-xs space-y-1.5 font-mono">
            <div className="flex justify-between text-stone-600">
              <span>Spreadsheet ID:</span>
              <span className="text-stone-800 truncate max-w-[180px]">
                {spreadsheetId}
              </span>
            </div>
            <div className="flex justify-between text-stone-600">
              <span>Jumlah baris:</span>
              <span className="font-bold text-emerald-700">
                {matchedCount.toLocaleString('id-ID')} baris
              </span>
            </div>
            <div className="flex justify-between text-stone-600">
              <span>Aksi:</span>
              <span className="text-amber-700 font-medium">
                {actionText}
              </span>
            </div>
          </div>

          <p className="text-[11px] text-stone-400 mt-3 text-center">
            Tindakan ini akan memodifikasi sel di Google Sheets akun Anda sesuai hasil pencocokan.
          </p>

          <div className="mt-6 flex items-center justify-end gap-2.5">
            <button
              type="button"
              onClick={onClose}
              disabled={isProcessing}
              className="px-4 py-2 text-xs font-medium rounded-lg border border-stone-300 text-stone-700 hover:bg-stone-100 transition-colors"
            >
              Batal
            </button>
            <button
              type="button"
              id="btn-confirm-sheet-update"
              onClick={onConfirm}
              disabled={isProcessing}
              className="px-4 py-2 text-xs font-semibold rounded-lg bg-orange-600 hover:bg-orange-700 text-white shadow-xs transition-colors disabled:opacity-50 flex items-center gap-1.5"
            >
              {isProcessing ? 'Memperbarui...' : 'Ya, Perbarui Sekarang'}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
