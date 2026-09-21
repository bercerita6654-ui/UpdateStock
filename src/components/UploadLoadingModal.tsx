import React from 'react';
import {
  FileSpreadsheet,
  Loader2,
  CheckCircle2,
  Clock,
  Sparkles,
  Database,
  ShoppingBag,
  Layers,
  ArrowRight,
} from 'lucide-react';

export interface UploadProgressState {
  isOpen: boolean;
  fileName: string;
  fileSize?: number | null;
  uploadType: 'stock_list' | 'balist' | 'shopee' | 'generic';
  step: number; // 1 to 4
  stepTitle: string;
  stepDescription?: string;
  progressPercent?: number;
}

interface UploadLoadingModalProps {
  progress: UploadProgressState;
}

export const UploadLoadingModal: React.FC<UploadLoadingModalProps> = ({ progress }) => {
  if (!progress.isOpen) return null;

  const formatFileSize = (bytes?: number | null) => {
    if (!bytes) return '';
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  const getTypeDetails = () => {
    switch (progress.uploadType) {
      case 'stock_list':
        return {
          title: 'Memproses File STOCK LIST (Database Gudang)',
          badge: 'Database Stok Gudang',
          badgeColor: 'bg-blue-100 text-blue-800 border-blue-200',
          icon: Database,
          iconBg: 'bg-blue-600',
        };
      case 'balist':
        return {
          title: 'Memproses File Data Balistshopee',
          badge: 'Template Balistshopee',
          badgeColor: 'bg-indigo-100 text-indigo-800 border-indigo-200',
          icon: Layers,
          iconBg: 'bg-indigo-600',
        };
      case 'shopee':
        return {
          title: 'Memproses File Mass Update Shopee',
          badge: 'Shopee Mass Update',
          badgeColor: 'bg-amber-100 text-amber-900 border-amber-200',
          icon: ShoppingBag,
          iconBg: 'bg-amber-600',
        };
      default:
        return {
          title: 'Memproses File Spreadsheet Excel',
          badge: 'File Excel (.xlsx)',
          badgeColor: 'bg-stone-100 text-stone-800 border-stone-200',
          icon: FileSpreadsheet,
          iconBg: 'bg-stone-800',
        };
    }
  };

  const typeInfo = getTypeDetails();
  const IconComponent = typeInfo.icon;

  const steps = [
    {
      num: 1,
      title: 'Membaca & Dekode File Excel',
      desc: 'Membaca struktur workbook, sheet, dan data binary.',
    },
    {
      num: 2,
      title: 'Menganalisis Kolom & Format Baris',
      desc: 'Mendeteksi otomatis kolom SKU, nama produk, dan kolom stok.',
    },
    {
      num: 3,
      title: 'Sinkronisasi & Pencocokan Data',
      desc: 'Menghubungkan kode barang ke database stok yang relevan.',
    },
    {
      num: 4,
      title: 'Menyiapkan Tampilan & Hasil',
      desc: 'Menghitung ringkasan perbandingan & menyusun tabel data.',
    },
  ];

  const currentPercent = progress.progressPercent ?? Math.min(progress.step * 25, 95);

  return (
    <div
      id="upload-loading-modal-overlay"
      className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs animate-in fade-in duration-200"
    >
      <div className="bg-white rounded-2xl border border-stone-200 shadow-2xl max-w-lg w-full overflow-hidden animate-in zoom-in-95 duration-200">
        {/* Modal Header */}
        <div className="bg-stone-900 text-white p-5 relative overflow-hidden">
          {/* Subtle background glow */}
          <div className="absolute -top-12 -right-12 w-40 h-40 bg-blue-500/20 rounded-full blur-2xl pointer-events-none" />
          
          <div className="relative z-10 flex items-start gap-3.5">
            <div className={`w-11 h-11 rounded-xl ${typeInfo.iconBg} text-white flex items-center justify-center shrink-0 shadow-md`}>
              <Loader2 className="w-6 h-6 animate-spin text-white" />
            </div>
            <div className="flex-1 min-w-0">
              <div className="flex items-center gap-2 mb-1">
                <span className={`px-2 py-0.5 rounded text-[10px] font-bold border uppercase tracking-wider ${typeInfo.badgeColor}`}>
                  {typeInfo.badge}
                </span>
                <span className="inline-flex items-center gap-1 text-[11px] text-blue-300 font-medium">
                  <Sparkles className="w-3 h-3 text-blue-400" />
                  Sedang Diproses...
                </span>
              </div>
              <h3 className="text-base font-bold text-white tracking-tight leading-snug truncate">
                {typeInfo.title}
              </h3>
              <p className="text-xs text-stone-300 mt-1 truncate">
                File: <span className="font-semibold text-white">{progress.fileName}</span>
                {progress.fileSize ? ` (${formatFileSize(progress.fileSize)})` : ''}
              </p>
            </div>
          </div>
        </div>

        {/* Modal Body: Progress Bar & Steps */}
        <div className="p-5 space-y-5">
          {/* Progress Bar Container */}
          <div className="space-y-1.5">
            <div className="flex items-center justify-between text-xs">
              <span className="font-semibold text-stone-700 flex items-center gap-1.5">
                <Clock className="w-3.5 h-3.5 text-blue-600 animate-pulse" />
                {progress.stepTitle || 'Memproses data...'}
              </span>
              <span className="font-mono font-bold text-blue-600">
                {Math.round(currentPercent)}%
              </span>
            </div>
            <div className="w-full h-2.5 bg-stone-100 rounded-full overflow-hidden border border-stone-200/80 p-0.5">
              <div
                className="h-full bg-linear-to-r from-blue-600 via-indigo-600 to-emerald-500 rounded-full transition-all duration-300 ease-out"
                style={{ width: `${Math.max(5, currentPercent)}%` }}
              />
            </div>
            {progress.stepDescription && (
              <p className="text-[11px] text-stone-500 italic mt-1">
                {progress.stepDescription}
              </p>
            )}
          </div>

          {/* Sequential Step Indicators */}
          <div className="space-y-2.5 bg-stone-50/70 p-3.5 rounded-xl border border-stone-100">
            {steps.map((s) => {
              const isCompleted = progress.step > s.num;
              const isCurrent = progress.step === s.num;
              const isPending = progress.step < s.num;

              return (
                <div
                  key={s.num}
                  className={`flex items-start gap-3 p-2 rounded-lg transition-colors ${
                    isCurrent
                      ? 'bg-white border border-blue-200 shadow-2xs'
                      : isCompleted
                      ? 'opacity-85'
                      : 'opacity-40'
                  }`}
                >
                  <div className="mt-0.5 shrink-0">
                    {isCompleted ? (
                      <div className="w-5 h-5 rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center font-bold text-[10px]">
                        <CheckCircle2 className="w-4 h-4" />
                      </div>
                    ) : isCurrent ? (
                      <div className="w-5 h-5 rounded-full bg-blue-600 text-white flex items-center justify-center font-bold text-[10px] animate-pulse">
                        <Loader2 className="w-3 h-3 animate-spin" />
                      </div>
                    ) : (
                      <div className="w-5 h-5 rounded-full bg-stone-200 text-stone-500 flex items-center justify-center font-bold text-[10px]">
                        {s.num}
                      </div>
                    )}
                  </div>
                  <div className="flex-1 min-w-0">
                    <p
                      className={`text-xs font-semibold ${
                        isCurrent
                          ? 'text-blue-900 font-bold'
                          : isCompleted
                          ? 'text-stone-800 line-through decoration-stone-300'
                          : 'text-stone-500'
                      }`}
                    >
                      {s.title}
                    </p>
                    <p className="text-[11px] text-stone-500 mt-0.5 leading-relaxed">
                      {s.desc}
                    </p>
                  </div>
                </div>
              );
            })}
          </div>

          {/* Helpful Information Notice */}
          <div className="p-3 bg-blue-50/60 border border-blue-100 rounded-lg text-xs text-blue-950 flex items-center gap-2">
            <span className="w-2 h-2 rounded-full bg-blue-500 shrink-0 animate-ping" />
            <p className="text-[11px] text-stone-600 leading-normal">
              Mohon tunggu sebentar, sistem sedang memproses dan mengekstrak data agar akurat.
            </p>
          </div>
        </div>
      </div>
    </div>
  );
};
