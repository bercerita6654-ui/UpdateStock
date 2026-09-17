import React, { useState } from 'react';
import { X, Sliders, RotateCcw, ExternalLink } from 'lucide-react';
import { SheetsConfig } from '../types';

interface SettingsModalProps {
  isOpen: boolean;
  onClose: () => void;
  config: SheetsConfig;
  onSave: (newConfig: SheetsConfig) => void;
  onReset: () => void;
}

export const SettingsModal: React.FC<SettingsModalProps> = ({
  isOpen,
  onClose,
  config,
  onSave,
  onReset,
}) => {
  const [formData, setFormData] = useState<SheetsConfig>(config);

  if (!isOpen) return null;

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    onSave(formData);
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-stone-900/50 backdrop-blur-xs">
      <div className="bg-white rounded-2xl border border-stone-200 shadow-xl max-w-lg w-full overflow-hidden">
        <div className="flex items-center justify-between p-4 border-b border-stone-100">
          <div className="flex items-center gap-2">
            <Sliders className="w-4 h-4 text-orange-600" />
            <h3 className="font-semibold text-stone-900 text-sm">
              Pengaturan Spreadsheet Sumber Data
            </h3>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1 text-stone-400 hover:text-stone-700 rounded-lg hover:bg-stone-100"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-5 space-y-4 text-xs">
          {/* Section 1: Balistshopee */}
          <div className="p-3 bg-stone-50 border border-stone-200 rounded-lg space-y-2.5">
            <div className="flex items-center justify-between font-semibold text-stone-800">
              <span>Spreadsheet Balistshopee</span>
              <a
                href={`https://docs.google.com/spreadsheets/d/${formData.balistSpreadsheetId}`}
                target="_blank"
                rel="noreferrer"
                className="text-orange-600 hover:underline inline-flex items-center gap-1 font-normal text-[11px]"
              >
                Buka Link <ExternalLink className="w-3 h-3" />
              </a>
            </div>

            <div>
              <label className="block text-stone-600 font-medium mb-1">
                Spreadsheet ID:
              </label>
              <input
                type="text"
                value={formData.balistSpreadsheetId}
                onChange={(e) =>
                  setFormData({ ...formData, balistSpreadsheetId: e.target.value })
                }
                className="w-full bg-white border border-stone-300 rounded px-2.5 py-1.5 font-mono text-stone-800 focus:outline-hidden focus:ring-1 focus:ring-orange-500"
                required
              />
            </div>

            <div>
              <label className="block text-stone-600 font-medium mb-1">
                Nama Sheet:
              </label>
              <input
                type="text"
                value={formData.balistSheetName}
                onChange={(e) =>
                  setFormData({ ...formData, balistSheetName: e.target.value })
                }
                className="w-full bg-white border border-stone-300 rounded px-2.5 py-1.5 font-medium text-stone-800 focus:outline-hidden focus:ring-1 focus:ring-orange-500"
                required
              />
            </div>
            <p className="text-[11px] text-stone-500">
              * SKU acuan berada pada Kolom 5 dan Kolom 6
            </p>
          </div>

          {/* Section 2: STOCK LIST */}
          <div className="p-3 bg-stone-50 border border-stone-200 rounded-lg space-y-2.5">
            <div className="flex items-center justify-between font-semibold text-stone-800">
              <span>Spreadsheet STOCK LIST (Stok Gudang)</span>
              <a
                href={`https://docs.google.com/spreadsheets/d/${formData.stockSpreadsheetId}`}
                target="_blank"
                rel="noreferrer"
                className="text-orange-600 hover:underline inline-flex items-center gap-1 font-normal text-[11px]"
              >
                Buka Link <ExternalLink className="w-3 h-3" />
              </a>
            </div>

            <div>
              <label className="block text-stone-600 font-medium mb-1">
                Spreadsheet ID:
              </label>
              <input
                type="text"
                value={formData.stockSpreadsheetId}
                onChange={(e) =>
                  setFormData({ ...formData, stockSpreadsheetId: e.target.value })
                }
                className="w-full bg-white border border-stone-300 rounded px-2.5 py-1.5 font-mono text-stone-800 focus:outline-hidden focus:ring-1 focus:ring-orange-500"
                required
              />
            </div>

            <div>
              <label className="block text-stone-600 font-medium mb-1">
                Nama Sheet:
              </label>
              <input
                type="text"
                value={formData.stockSheetName}
                onChange={(e) =>
                  setFormData({ ...formData, stockSheetName: e.target.value })
                }
                className="w-full bg-white border border-stone-300 rounded px-2.5 py-1.5 font-medium text-stone-800 focus:outline-hidden focus:ring-1 focus:ring-orange-500"
                required
              />
            </div>
            <p className="text-[11px] text-stone-500">
              * Kode SKU acuan pada Kolom 1 (Code) dan jumlah stok pada Kolom 15 (Qty)
            </p>
          </div>

          <div className="flex items-center justify-between pt-3 border-t border-stone-100">
            <button
              type="button"
              onClick={() => {
                onReset();
                onClose();
              }}
              className="inline-flex items-center gap-1 text-stone-500 hover:text-stone-800 font-medium"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              Reset ke Default
            </button>

            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={onClose}
                className="px-3 py-1.5 rounded-lg border border-stone-300 hover:bg-stone-100 font-medium text-stone-700"
              >
                Batal
              </button>
              <button
                type="submit"
                className="px-4 py-1.5 rounded-lg bg-orange-600 hover:bg-orange-700 font-medium text-white shadow-xs"
              >
                Simpan Perubahan
              </button>
            </div>
          </div>
        </form>
      </div>
    </div>
  );
};
