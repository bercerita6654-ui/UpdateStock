import React, { useState, useMemo } from 'react';
import {
  History,
  RefreshCw,
  Send,
  UploadCloud,
  Download,
  ShieldCheck,
  AlertTriangle,
  CheckCircle2,
  XCircle,
  Clock,
  Trash2,
  Copy,
  Check,
  ChevronDown,
  ChevronUp,
  Search,
  SlidersHorizontal,
  Info,
  Terminal,
  FileSpreadsheet,
  EyeOff,
} from 'lucide-react';
import { ActivityLogItem, LogType, LogStatus } from '../types';

interface ActivityLogPanelProps {
  logs: ActivityLogItem[];
  onClearLogs: () => void;
  lastSyncTime: Date | null;
  lastSheetUpdateTime: Date | null;
  onRefreshSync?: () => void;
  isSyncing?: boolean;
  onHide?: () => void;
}

export const ActivityLogPanel: React.FC<ActivityLogPanelProps> = ({
  logs,
  onClearLogs,
  lastSyncTime,
  lastSheetUpdateTime,
  onRefreshSync,
  isSyncing = false,
  onHide,
}) => {
  const [isExpanded, setIsExpanded] = useState<boolean>(() => {
    return localStorage.getItem('activity_log_expanded') !== 'false';
  });
  const [activeTab, setActiveTab] = useState<'all' | 'sync' | 'sheet_update' | 'files' | 'error'>('all');
  const [searchTerm, setSearchTerm] = useState('');
  const [copiedLogId, setCopiedLogId] = useState<string | null>(null);
  const [copiedAll, setCopiedAll] = useState(false);
  const [expandedLogId, setExpandedLogId] = useState<string | null>(null);

  const toggleExpanded = () => {
    setIsExpanded((prev) => {
      const next = !prev;
      localStorage.setItem('activity_log_expanded', String(next));
      return next;
    });
  };

  const errorCount = useMemo(() => {
    return logs.filter((l) => l.status === 'error').length;
  }, [logs]);

  const filteredLogs = useMemo(() => {
    return logs.filter((log) => {
      // Tab filter
      if (activeTab === 'sync' && log.type !== 'sync') return false;
      if (activeTab === 'sheet_update' && log.type !== 'sheet_update') return false;
      if (activeTab === 'files' && log.type !== 'upload' && log.type !== 'download') return false;
      if (activeTab === 'error' && log.status !== 'error') return false;

      // Search filter
      if (searchTerm.trim()) {
        const q = searchTerm.toLowerCase();
        const titleMatch = log.title.toLowerCase().includes(q);
        const descMatch = log.description.toLowerCase().includes(q);
        const targetMatch = log.target?.toLowerCase().includes(q);
        const errorMatch = log.errorMessage?.toLowerCase().includes(q);
        const detailsMatch = log.details?.toLowerCase().includes(q);
        return titleMatch || descMatch || !!targetMatch || !!errorMatch || !!detailsMatch;
      }
      return true;
    });
  }, [logs, activeTab, searchTerm]);

  const handleCopySingleLog = (log: ActivityLogItem, e: React.MouseEvent) => {
    e.stopPropagation();
    const text = `[${log.formattedTime || log.timestamp}] [${log.status.toUpperCase()}] ${log.title}
Deskripsi: ${log.description}
${log.target ? `Target: ${log.target}\n` : ''}${log.errorMessage ? `Error: ${log.errorMessage}\n` : ''}${log.details ? `Detail: ${log.details}\n` : ''}`;
    navigator.clipboard.writeText(text);
    setCopiedLogId(log.id);
    setTimeout(() => setCopiedLogId(null), 2000);
  };

  const handleCopyAllLogs = () => {
    if (logs.length === 0) return;
    const text = logs
      .map(
        (log) =>
          `[${log.formattedTime || log.timestamp}] [${log.type.toUpperCase()}] [${log.status.toUpperCase()}] ${log.title} - ${log.description}${
            log.errorMessage ? ` | ERROR: ${log.errorMessage}` : ''
          }`
      )
      .join('\n');
    navigator.clipboard.writeText(text);
    setCopiedAll(true);
    setTimeout(() => setCopiedAll(false), 2500);
  };

  const formatRelativeTime = (isoString: string) => {
    try {
      const date = new Date(isoString);
      const now = new Date();
      const diffSeconds = Math.floor((now.getTime() - date.getTime()) / 1000);

      if (diffSeconds < 10) return 'Baru saja';
      if (diffSeconds < 60) return `${diffSeconds} detik lalu`;
      const diffMinutes = Math.floor(diffSeconds / 60);
      if (diffMinutes < 60) return `${diffMinutes} menit lalu`;
      const diffHours = Math.floor(diffMinutes / 60);
      if (diffHours < 24) return `${diffHours} jam lalu`;
      return date.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' });
    } catch {
      return isoString;
    }
  };

  const getLogIcon = (type: LogType, status: LogStatus) => {
    if (status === 'error') {
      return <XCircle className="w-4 h-4 text-rose-600" />;
    }
    switch (type) {
      case 'sync':
        return <RefreshCw className={`w-4 h-4 ${status === 'loading' ? 'animate-spin text-blue-600' : 'text-emerald-600'}`} />;
      case 'sheet_update':
        return <Send className="w-4 h-4 text-blue-600" />;
      case 'upload':
        return <UploadCloud className="w-4 h-4 text-indigo-600" />;
      case 'download':
        return <Download className="w-4 h-4 text-purple-600" />;
      case 'auth':
        return <ShieldCheck className="w-4 h-4 text-emerald-600" />;
      case 'settings':
        return <SlidersHorizontal className="w-4 h-4 text-stone-600" />;
      case 'error':
      default:
        return <AlertTriangle className="w-4 h-4 text-amber-600" />;
    }
  };

  const getStatusBadge = (status: LogStatus) => {
    switch (status) {
      case 'success':
        return (
          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-50 text-emerald-700 border border-emerald-200">
            <CheckCircle2 className="w-3 h-3" />
            Sukses
          </span>
        );
      case 'error':
        return (
          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-rose-50 text-rose-700 border border-rose-200">
            <XCircle className="w-3 h-3" />
            Gagal
          </span>
        );
      case 'loading':
        return (
          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-blue-50 text-blue-700 border border-blue-200">
            <RefreshCw className="w-3 h-3 animate-spin" />
            Memproses
          </span>
        );
      case 'info':
      default:
        return (
          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-stone-100 text-stone-700 border border-stone-200">
            <Info className="w-3 h-3 text-stone-400" />
            Info
          </span>
        );
    }
  };

  return (
    <div className="bg-white rounded-xl border border-stone-200 shadow-xs overflow-hidden transition-all">
      {/* Panel Header */}
      <div className="p-4 sm:p-4.5 bg-stone-50/80 border-b border-stone-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div className="flex items-center justify-between sm:justify-start gap-3 w-full sm:w-auto">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-stone-100 text-stone-700 flex items-center justify-center border border-stone-300/80 shrink-0">
              <History className="w-4 h-4 text-stone-700" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h3 className="text-sm font-bold text-stone-900">
                  Log Riwayat Aktivitas &amp; Pembaruan
                </h3>
                <span className="px-2 py-0.2 rounded-full text-[11px] font-semibold bg-stone-200 text-stone-800">
                  {logs.length} Log
                </span>
                {errorCount > 0 ? (
                  <span className="px-2 py-0.2 rounded-full text-[11px] font-bold bg-rose-100 text-rose-700 border border-rose-200 flex items-center gap-1">
                    <AlertTriangle className="w-3 h-3" />
                    {errorCount} Error Terdeteksi
                  </span>
                ) : (
                  <span className="hidden md:inline-flex items-center gap-1 px-2 py-0.2 rounded-full text-[11px] font-medium bg-emerald-50 text-emerald-700 border border-emerald-200">
                    <CheckCircle2 className="w-3 h-3" />
                    Semua Normal
                  </span>
                )}
              </div>
              <p className="text-[11px] text-stone-500 mt-0.5">
                Mencatat riwayat sinkronisasi Google Sheets, eksekusi pembaruan baris sheet, dan audit error.
              </p>
            </div>
          </div>

          <button
            type="button"
            onClick={toggleExpanded}
            className="sm:hidden p-1.5 rounded-lg border border-stone-200 text-stone-500 hover:bg-stone-200/60 transition-colors"
            aria-label="Toggle Panel"
          >
            {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
          </button>
        </div>

        {/* Quick Timestamps & Header Actions */}
        <div className="flex flex-wrap items-center gap-2 sm:justify-end text-xs">
          {/* Last Sync indicator */}
          <div className="flex items-center gap-1.5 px-2.5 py-1 bg-white border border-stone-200 rounded-lg text-[11px] text-stone-600">
            <Clock className="w-3.5 h-3.5 text-stone-400 shrink-0" />
            <span>Sinkron Terakhir:</span>
            <strong className="text-stone-800 font-mono">
              {lastSyncTime
                ? lastSyncTime.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' })
                : 'Belum Sinkron'}
            </strong>
          </div>

          {/* Last Sheet Update indicator */}
          {lastSheetUpdateTime && (
            <div className="flex items-center gap-1.5 px-2.5 py-1 bg-emerald-50 border border-emerald-200 rounded-lg text-[11px] text-emerald-800">
              <Send className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
              <span>Update Sheet:</span>
              <strong className="font-mono">
                {lastSheetUpdateTime.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' })}
              </strong>
            </div>
          )}

          {/* Action buttons */}
          <div className="flex items-center gap-1.5">
            {onRefreshSync && (
              <button
                type="button"
                onClick={onRefreshSync}
                disabled={isSyncing}
                className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg text-[11px] font-medium bg-white hover:bg-stone-100 text-stone-700 border border-stone-300 transition-colors shadow-2xs disabled:opacity-50 cursor-pointer"
                title="Sinkronkan ulang Google Sheets"
              >
                <RefreshCw className={`w-3 h-3 text-stone-500 ${isSyncing ? 'animate-spin' : ''}`} />
                <span className="hidden md:inline">Sinkronkan</span>
              </button>
            )}

            <button
              type="button"
              onClick={handleCopyAllLogs}
              disabled={logs.length === 0}
              className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg text-[11px] font-medium bg-white hover:bg-stone-100 text-stone-700 border border-stone-300 transition-colors shadow-2xs disabled:opacity-40 cursor-pointer"
              title="Salin seluruh riwayat log aktivitas untuk pelacakan error"
            >
              {copiedAll ? <Check className="w-3 h-3 text-emerald-600" /> : <Copy className="w-3 h-3 text-stone-500" />}
              <span>{copiedAll ? 'Tersalin' : 'Salin Log'}</span>
            </button>

            <button
              type="button"
              onClick={onClearLogs}
              disabled={logs.length === 0}
              className="inline-flex items-center gap-1 px-2 py-1 rounded-lg text-[11px] font-medium text-stone-500 hover:text-rose-600 hover:bg-rose-50 border border-transparent hover:border-rose-200 transition-colors cursor-pointer disabled:opacity-40"
              title="Bersihkan semua catatan log"
            >
              <Trash2 className="w-3 h-3" />
              <span className="hidden lg:inline">Bersihkan</span>
            </button>

            <button
              type="button"
              onClick={toggleExpanded}
              className="hidden sm:inline-flex items-center gap-1 px-2.5 py-1 rounded-lg text-[11px] font-medium bg-white hover:bg-stone-100 text-stone-700 border border-stone-300 transition-colors shadow-2xs cursor-pointer"
              title={isExpanded ? 'Sembunyikan detail log' : 'Tampilkan detail log'}
            >
              <span>{isExpanded ? 'Tutup' : 'Buka Log'}</span>
              {isExpanded ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
            </button>

            {onHide && (
              <button
                type="button"
                onClick={onHide}
                className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg text-[11px] font-medium text-stone-600 hover:text-stone-900 bg-white hover:bg-stone-100 border border-stone-300 transition-colors shadow-2xs cursor-pointer"
                title="Sembunyikan panel log riwayat ini"
              >
                <EyeOff className="w-3.5 h-3.5 text-stone-500" />
                <span>Sembunyikan</span>
              </button>
            )}
          </div>
        </div>
      </div>

      {/* Expanded Content Area */}
      {isExpanded ? (
        <div className="p-4 space-y-3">
          {/* Sub Toolbar: Filter tabs & search */}
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-2.5 pt-0.5">
            <div className="flex items-center gap-1 overflow-x-auto pb-1 md:pb-0 text-xs">
              <button
                type="button"
                onClick={() => setActiveTab('all')}
                className={`px-3 py-1 rounded-lg font-semibold transition-all shrink-0 cursor-pointer ${
                  activeTab === 'all'
                    ? 'bg-stone-900 text-white shadow-2xs'
                    : 'text-stone-600 hover:text-stone-900 hover:bg-stone-100'
                }`}
              >
                Semua ({logs.length})
              </button>

              <button
                type="button"
                onClick={() => setActiveTab('sync')}
                className={`px-3 py-1 rounded-lg font-semibold transition-all shrink-0 cursor-pointer ${
                  activeTab === 'sync'
                    ? 'bg-emerald-700 text-white shadow-2xs'
                    : 'text-stone-600 hover:text-stone-900 hover:bg-stone-100'
                }`}
              >
                Sinkronisasi ({logs.filter((l) => l.type === 'sync').length})
              </button>

              <button
                type="button"
                onClick={() => setActiveTab('sheet_update')}
                className={`px-3 py-1 rounded-lg font-semibold transition-all shrink-0 cursor-pointer ${
                  activeTab === 'sheet_update'
                    ? 'bg-blue-700 text-white shadow-2xs'
                    : 'text-stone-600 hover:text-stone-900 hover:bg-stone-100'
                }`}
              >
                Pembaruan Sheet ({logs.filter((l) => l.type === 'sheet_update').length})
              </button>

              <button
                type="button"
                onClick={() => setActiveTab('files')}
                className={`px-3 py-1 rounded-lg font-semibold transition-all shrink-0 cursor-pointer ${
                  activeTab === 'files'
                    ? 'bg-purple-700 text-white shadow-2xs'
                    : 'text-stone-600 hover:text-stone-900 hover:bg-stone-100'
                }`}
              >
                File &amp; Excel ({logs.filter((l) => l.type === 'upload' || l.type === 'download').length})
              </button>

              {errorCount > 0 && (
                <button
                  type="button"
                  onClick={() => setActiveTab('error')}
                  className={`px-3 py-1 rounded-lg font-semibold transition-all shrink-0 cursor-pointer ${
                    activeTab === 'error'
                      ? 'bg-rose-700 text-white shadow-2xs'
                      : 'text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-200'
                  }`}
                >
                  Error ({errorCount})
                </button>
              )}
            </div>

            <div className="relative w-full md:w-64">
              <Search className="w-3.5 h-3.5 text-stone-400 absolute left-2.5 top-2" />
              <input
                type="text"
                placeholder="Cari kata kunci log / error..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full pl-8 pr-3 py-1 text-xs bg-stone-50 border border-stone-200 rounded-lg text-stone-800 placeholder-stone-400 focus:outline-hidden focus:bg-white focus:border-stone-400"
              />
            </div>
          </div>

          {/* Logs List Container */}
          <div className="border border-stone-200 rounded-xl bg-stone-50/50 max-h-72 overflow-y-auto divide-y divide-stone-200/80">
            {filteredLogs.length === 0 ? (
              <div className="py-8 px-4 text-center text-stone-400">
                <Terminal className="w-7 h-7 text-stone-300 mx-auto mb-1.5" />
                <p className="text-xs font-semibold text-stone-600">Tidak ada log aktivitas ditemukan</p>
                <p className="text-[11px] text-stone-400 mt-0.5">
                  {searchTerm ? 'Coba ganti kata kunci pencarian Anda.' : 'Setiap aktivitas sinkronisasi dan pembaruan akan dicatat di sini.'}
                </p>
              </div>
            ) : (
              filteredLogs.map((log) => {
                const isDetailExpanded = expandedLogId === log.id;
                const hasExtraDetails = !!log.errorMessage || !!log.details || !!log.target;

                return (
                  <div
                    key={log.id}
                    className={`p-3 transition-colors ${
                      log.status === 'error'
                        ? 'bg-rose-50/40 hover:bg-rose-50/70'
                        : log.status === 'loading'
                        ? 'bg-blue-50/30'
                        : 'bg-white hover:bg-stone-50'
                    }`}
                  >
                    <div className="flex items-start justify-between gap-3">
                      <div className="flex items-start gap-2.5 min-w-0">
                        <div className="mt-0.5 shrink-0">{getLogIcon(log.type, log.status)}</div>
                        <div className="min-w-0 flex-1">
                          <div className="flex flex-wrap items-center gap-2">
                            <span className="text-xs font-bold text-stone-900">
                              {log.title}
                            </span>
                            {getStatusBadge(log.status)}
                            {log.rowCount !== undefined && log.rowCount > 0 && (
                              <span className="text-[10px] font-mono px-1.5 py-0.2 bg-stone-100 text-stone-700 rounded border border-stone-200">
                                {log.rowCount.toLocaleString('id-ID')} baris
                              </span>
                            )}
                          </div>

                          <p className="text-xs text-stone-600 mt-0.5 break-words">
                            {log.description}
                          </p>

                          {/* Target identifier badge */}
                          {log.target && (
                            <div className="mt-1 flex items-center gap-1.5 text-[11px] text-stone-500 font-mono">
                              <span className="text-stone-400">Target:</span>
                              <span className="bg-stone-100 px-1.5 py-0.2 rounded border border-stone-200 text-stone-700 truncate max-w-sm">
                                {log.target}
                              </span>
                            </div>
                          )}

                          {/* Error Callout */}
                          {log.errorMessage && (
                            <div className="mt-1.5 p-2 bg-rose-100/70 border border-rose-200 rounded-lg text-xs text-rose-900 font-mono break-all flex items-start gap-2">
                              <AlertTriangle className="w-3.5 h-3.5 text-rose-600 shrink-0 mt-0.5" />
                              <div className="flex-1 min-w-0">
                                <span className="font-bold block text-[11px] text-rose-800">Detail Error:</span>
                                <span>{log.errorMessage}</span>
                              </div>
                            </div>
                          )}

                          {/* Expandable Details Drawer */}
                          {isDetailExpanded && log.details && (
                            <div className="mt-2 p-2.5 bg-stone-900 text-stone-100 rounded-lg text-[11px] font-mono overflow-x-auto whitespace-pre-wrap max-h-36">
                              {log.details}
                            </div>
                          )}
                        </div>
                      </div>

                      {/* Right side: Timestamp & Quick Action */}
                      <div className="flex flex-col items-end shrink-0 gap-1 text-[11px] text-stone-400">
                        <span className="font-mono text-stone-600 font-medium">
                          {log.formattedTime || new Date(log.timestamp).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' })}
                        </span>
                        <span className="text-[10px] text-stone-400">
                          {formatRelativeTime(log.timestamp)}
                        </span>

                        <div className="flex items-center gap-1 mt-1">
                          {hasExtraDetails && log.details && (
                            <button
                              type="button"
                              onClick={() => setExpandedLogId(isDetailExpanded ? null : log.id)}
                              className="text-[10px] text-stone-500 hover:text-stone-800 underline cursor-pointer"
                            >
                              {isDetailExpanded ? 'Sembunyikan' : 'Lihat Detail'}
                            </button>
                          )}
                          <button
                            type="button"
                            onClick={(e) => handleCopySingleLog(log, e)}
                            className="p-1 text-stone-400 hover:text-stone-700 rounded hover:bg-stone-200/50 transition-colors cursor-pointer"
                            title="Salin rincian log ini"
                          >
                            {copiedLogId === log.id ? (
                              <Check className="w-3 h-3 text-emerald-600" />
                            ) : (
                              <Copy className="w-3 h-3" />
                            )}
                          </button>
                        </div>
                      </div>
                    </div>
                  </div>
                );
              })
            )}
          </div>
        </div>
      ) : (
        /* Collapsed Preview summary */
        <div
          onClick={toggleExpanded}
          className="px-4 py-2 bg-stone-50/50 hover:bg-stone-100/70 border-t border-stone-100 flex items-center justify-between text-xs text-stone-600 cursor-pointer transition-colors"
        >
          <div className="flex items-center gap-2 truncate">
            <span className="text-stone-400">Aktivitas Terakhir:</span>
            {logs.length > 0 ? (
              <span className="font-medium text-stone-800 truncate">
                [{logs[0].formattedTime || 'Baru saja'}] {logs[0].title} — {logs[0].description}
              </span>
            ) : (
              <span className="text-stone-400 italic">Belum ada aktivitas tercatat</span>
            )}
          </div>

          <div className="flex items-center gap-1 text-stone-500 font-medium text-[11px] shrink-0">
            <span>Buka Riwayat Lengkap</span>
            <ChevronDown className="w-3.5 h-3.5" />
          </div>
        </div>
      )}
    </div>
  );
};
