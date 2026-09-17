import React, { useState, useEffect, useMemo, useCallback } from 'react';
import { User } from 'firebase/auth';
import {
  initAuth,
  googleSignIn,
  logout,
  getAccessToken,
  setCachedToken,
} from './lib/firebase';
import {
  loadStockListData,
  loadBalistShopeeData,
  updateSheetValues,
  clearSheetValues,
  fetchSheetValues,
  parseStockListFromRows,
  parseBalistShopeeFromRows,
  detectBalistStockColumn,
  getColumnLetter,
  convertOfficeFileToGoogleSheet,
  downloadBalistRowsAsXlsx,
} from './lib/sheets';
import {
  buildMatchingMaps,
  MatchingMaps,
  calculateSummary,
  compareBalistWithStockList,
  updateBalistRowsWithStock,
} from './lib/stockMatcher';
import {
  parseShopeeXlsx,
  matchShopeeFile,
  generateUpdatedShopeeWorkbook,
  downloadBlob,
  createSampleShopeeFile,
  generateShopeeBalistFilename,
  detectStoreFromFilename,
  ParsedShopeeSheet,
  parseGenericXlsx,
  ParsedGenericXlsx,
} from './lib/excelProcessor';
import {
  StockListItem,
  BalistShopeeItem,
  ShopeeRowMatch,
  ProcessSummary,
  SheetsConfig,
  BalistComparisonItem,
  BalistComparisonSummary,
  ActivityLogItem,
  LogType,
  LogStatus,
} from './types';
import { Header } from './components/Header';
import { SheetsStatusCard } from './components/SheetsStatusCard';
import { ActivityLogPanel } from './components/ActivityLogPanel';
import { BalistUploadCard } from './components/BalistUploadCard';
import { BalistComparisonTable } from './components/BalistComparisonTable';
import { UploadSection } from './components/UploadSection';
import { SummaryCards } from './components/SummaryCards';
import { MatchTable } from './components/MatchTable';
import { SettingsModal } from './components/SettingsModal';
import { ConfirmUpdateModal } from './components/ConfirmUpdateModal';
import { ArrowRight, CheckCircle, Info, Sparkles, Database, ShoppingBag, GitCompare, ExternalLink, Eye, EyeOff, RefreshCw, History } from 'lucide-react';

const DEFAULT_SHEETS_CONFIG: SheetsConfig = {
  balistSpreadsheetId: '1wTchgk4-YRyQv-Sk10SZUrOooGMrC08S',
  balistSheetName: 'Balistshopee',
  stockSpreadsheetId: '1mrD9sQK_Sffa1X1fzlCDmaJXs1Yj2q-XTNdi2sRGPos',
  stockSheetName: 'STOCK LIST',
};

export default function App() {
  // Navigation tab: 'all' | 'balist_sync' | 'shopee_sync'
  const [activeTab, setActiveTab] = useState<'workflow' | 'balist_comparison' | 'shopee_match'>('workflow');
  const [showShopeeSection, setShowShopeeSection] = useState<boolean>(() => {
    const saved = localStorage.getItem('show_shopee_section');
    return saved ? JSON.parse(saved) : false;
  });
  const [showSheetsStatusCard, setShowSheetsStatusCard] = useState<boolean>(() => {
    const saved = localStorage.getItem('show_sheets_status_card');
    return saved ? JSON.parse(saved) : false;
  });
  const [showActivityLogPanel, setShowActivityLogPanel] = useState<boolean>(() => {
    const saved = localStorage.getItem('show_activity_log_panel');
    return saved === 'true';
  });

  // Auth state
  const [user, setUser] = useState<User | null>(null);
  const [accessToken, setAccessToken] = useState<string | null>(null);
  const [isLoadingAuth, setIsLoadingAuth] = useState(false);

  // Sheets Config
  const [config, setConfig] = useState<SheetsConfig>(() => {
    const saved = localStorage.getItem('shopee_sheets_config');
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        return DEFAULT_SHEETS_CONFIG;
      }
    }
    return DEFAULT_SHEETS_CONFIG;
  });

  // Sheets Data
  const [stockList, setStockList] = useState<StockListItem[]>([]);
  const [balistList, setBalistList] = useState<BalistShopeeItem[]>([]);
  const [isLoadingSheets, setIsLoadingSheets] = useState(false);
  const [sheetsError, setSheetsError] = useState<string | null>(null);
  const [lastLoaded, setLastLoaded] = useState<Date | null>(null);

  // Office Conversion State
  const [isConvertingBalist, setIsConvertingBalist] = useState(false);
  const [convertedSpreadsheetInfo, setConvertedSpreadsheetInfo] = useState<{
    id: string;
    name: string;
    url?: string;
  } | null>(null);

  // Upload 1: Balistshopee XLSX update
  const [balistUploadedFile, setBalistUploadedFile] = useState<File | null>(null);
  const [parsedBalistXlsx, setParsedBalistXlsx] = useState<ParsedGenericXlsx | null>(null);
  const [balistSourceStartRow, setBalistSourceStartRow] = useState<number>(7);
  const [balistStockColIndex, setBalistStockColIndex] = useState<number>(6);
  const [updateStockFromStockList, setUpdateStockFromStockList] = useState<boolean>(true);
  const [balistUnmatchedStockAction, setBalistUnmatchedStockAction] = useState<'zero' | 'keep'>('keep');
  const [isUpdatingBalistSheet, setIsUpdatingBalistSheet] = useState(false);
  const [isConfirmBalistUploadModalOpen, setIsConfirmBalistUploadModalOpen] = useState(false);
  const [isDirectUpdatingBalistStock, setIsDirectUpdatingBalistStock] = useState(false);
  const [isConfirmDirectUpdateModalOpen, setIsConfirmDirectUpdateModalOpen] = useState(false);

  // Upload 2: Shopee Uploaded File & Matching
  const [shopeeUploadedFile, setShopeeUploadedFile] = useState<File | null>(null);
  const [parsedShopeeSheet, setParsedShopeeSheet] = useState<ParsedShopeeSheet | null>(null);
  const [selectedSkuCol, setSelectedSkuCol] = useState<number>(0);
  const [selectedStockCol, setSelectedStockCol] = useState<number>(1);
  const [unmatchedAction, setUnmatchedAction] = useState<'keep' | 'zero'>('keep');
  const [matches, setMatches] = useState<ShopeeRowMatch[]>([]);
  const [summary, setSummary] = useState<ProcessSummary | null>(null);

  // Modals & Feedback
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isConfirmModalOpen, setIsConfirmModalOpen] = useState(false);
  const [isSyncingBalist, setIsSyncingBalist] = useState(false);
  const [toastMessage, setToastMessage] = useState<{ text: string; type: 'success' | 'error' } | null>(null);

  // Activity Log State
  const [activityLogs, setActivityLogs] = useState<ActivityLogItem[]>(() => {
    try {
      const saved = localStorage.getItem('shopee_activity_logs');
      if (saved) {
        return JSON.parse(saved);
      }
    } catch (e) {
      console.error('Failed to parse activity logs:', e);
    }
    return [
      {
        id: 'init-ready',
        timestamp: new Date().toISOString(),
        formattedTime: new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' }),
        type: 'info',
        title: 'Sistem Dimuat',
        description: 'Aplikasi siap untuk sinkronisasi Google Sheets dan pembaruan stok.',
        status: 'info',
      },
    ];
  });

  const [lastSheetUpdateTime, setLastSheetUpdateTime] = useState<Date | null>(() => {
    const saved = localStorage.getItem('last_sheet_update_time');
    return saved ? new Date(saved) : null;
  });

  const addLog = useCallback(
    (
      type: LogType,
      title: string,
      description: string,
      status: LogStatus = 'info',
      extra?: {
        target?: string;
        details?: string;
        errorMessage?: string;
        rowCount?: number;
      }
    ) => {
      const now = new Date();
      const newEntry: ActivityLogItem = {
        id: `${Date.now()}-${Math.random().toString(36).substring(2, 7)}`,
        timestamp: now.toISOString(),
        formattedTime: now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' }),
        type,
        title,
        description,
        status,
        ...extra,
      };

      setActivityLogs((prev) => {
        const updated = [newEntry, ...prev.slice(0, 99)];
        try {
          localStorage.setItem('shopee_activity_logs', JSON.stringify(updated));
        } catch (e) {
          console.error('Failed to persist logs:', e);
        }
        return updated;
      });
    },
    []
  );

  const handleClearLogs = () => {
    const now = new Date();
    const resetEntry: ActivityLogItem = {
      id: `${Date.now()}-reset`,
      timestamp: now.toISOString(),
      formattedTime: now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' }),
      type: 'info',
      title: 'Log Dibersihkan',
      description: 'Riwayat aktivitas sebelumnya telah dibersihkan oleh pengguna.',
      status: 'info',
    };
    setActivityLogs([resetEntry]);
    localStorage.setItem('shopee_activity_logs', JSON.stringify([resetEntry]));
    showToast('Semua catatan log riwayat berhasil dibersihkan.');
  };

  const showToast = (text: string, type: 'success' | 'error' = 'success') => {
    setToastMessage({ text, type });
    setTimeout(() => {
      setToastMessage(null);
    }, 4000);
  };

  // Build lookup maps for fast matching
  const matchingMaps = useMemo<MatchingMaps>(() => {
    return buildMatchingMaps(stockList, balistList);
  }, [stockList, balistList]);

  // Balistshopee vs STOCK LIST comparison calculation
  const balistComparison = useMemo(() => {
    if (balistList.length === 0) {
      return {
        items: [] as BalistComparisonItem[],
        summary: {
          totalRows: 0,
          matchedCount: 0,
          unmatchedCount: 0,
          inStockCount: 0,
          outOfStockCount: 0,
          totalStockQuantity: 0,
        } as BalistComparisonSummary,
      };
    }
    return compareBalistWithStockList(balistList, matchingMaps);
  }, [balistList, matchingMaps]);

  // Handle Google Auth
  useEffect(() => {
    const unsubscribe = initAuth(
      (currentUser, token) => {
        setUser(currentUser);
        setAccessToken(token);
      },
      () => {
        setUser(null);
        setAccessToken(null);
      }
    );
    return () => unsubscribe();
  }, []);

  const handleSignIn = async () => {
    if (isLoadingAuth) return;
    setIsLoadingAuth(true);
    setSheetsError(null);
    try {
      const result = await googleSignIn();
      if (result) {
        setUser(result.user);
        setAccessToken(result.accessToken);
        addLog(
          'auth',
          'Login Google Berhasil',
          `Berhasil terhubung sebagai ${result.user.displayName || result.user.email}.`,
          'success',
          { target: result.user.email || undefined }
        );
        showToast(`Berhasil masuk sebagai ${result.user.displayName || result.user.email}`);
        loadSheets(result.accessToken);
      }
    } catch (err: any) {
      if (
        err?.code === 'auth/cancelled-popup-request' ||
        err?.code === 'auth/popup-closed-by-user' ||
        err?.message?.includes('cancelled-popup-request') ||
        err?.message?.includes('popup-closed-by-user')
      ) {
        // User closed the popup or duplicate request was cancelled - no toast needed
        return;
      }
      console.error('Sign in failed:', err);
      const errMsg = err?.message || 'Gagal menghubungkan akun Google';
      addLog('auth', 'Gagal Login Google', errMsg, 'error', { errorMessage: errMsg });
      showToast(errMsg, 'error');
    } finally {
      setIsLoadingAuth(false);
    }
  };

  const handleSignOut = async () => {
    try {
      await logout();
      setUser(null);
      setAccessToken(null);
      addLog('auth', 'Logout Akun Google', 'Pengguna telah keluar dari akun Google.', 'info');
      showToast('Berhasil keluar dari akun Google');
    } catch (err: any) {
      console.error('Logout error:', err);
    }
  };

  // Load Google Sheets data
  const loadSheets = useCallback(
    async (tokenToUse?: string) => {
      const token = tokenToUse || accessToken;
      if (!token) {
        setSheetsError('Silakan Masuk dengan Google untuk membaca data spreadsheet.');
        addLog(
          'sync',
          'Sinkronisasi Ditunda',
          'Akses akun Google diperlukan sebelum membaca spreadsheet.',
          'info'
        );
        return;
      }

      setIsLoadingSheets(true);
      setSheetsError(null);
      addLog(
        'sync',
        'Memulai Sinkronisasi Data',
        `Membaca sheet "${config.stockSheetName}" & "${config.balistSheetName}" dari Google Drive...`,
        'loading',
        { target: `${config.stockSheetName}, ${config.balistSheetName}` }
      );

      try {
        const stockPromise = loadStockListData(config.stockSpreadsheetId, config.stockSheetName, token);
        const balistPromise = loadBalistShopeeData(config.balistSpreadsheetId, config.balistSheetName, token);

        const [stockRes, balistRes] = await Promise.allSettled([stockPromise, balistPromise]);

        let successCount = 0;
        const errMessages: string[] = [];

        let stockRowCount = 0;
        let balistRowCount = 0;

        if (stockRes.status === 'fulfilled') {
          setStockList(stockRes.value);
          stockRowCount = stockRes.value.length;
          successCount++;
        } else {
          console.error('Stock list load error:', stockRes.reason);
          errMessages.push(`STOCK LIST: ${stockRes.reason?.message || 'Gagal memuat'}`);
        }

        if (balistRes.status === 'fulfilled') {
          setBalistList(balistRes.value);
          balistRowCount = balistRes.value.length;
          successCount++;
        } else {
          console.error('Balist load error:', balistRes.reason);
          errMessages.push(`Balistshopee: ${balistRes.reason?.message || 'Gagal memuat'}`);
        }

        if (errMessages.length > 0) {
          setSheetsError(errMessages.join(' \n• '));
        } else {
          setSheetsError(null);
        }

        if (successCount > 0) {
          const syncTime = new Date();
          setLastLoaded(syncTime);
          addLog(
            'sync',
            'Sinkronisasi Selesai',
            `Berhasil memuat ${stockRowCount.toLocaleString('id-ID')} baris STOCK LIST dan ${balistRowCount.toLocaleString('id-ID')} baris Balistshopee (${successCount}/2 sheet).`,
            errMessages.length > 0 ? 'error' : 'success',
            {
              rowCount: stockRowCount + balistRowCount,
              target: `${config.stockSheetName} (${stockRowCount}), ${config.balistSheetName} (${balistRowCount})`,
              errorMessage: errMessages.length > 0 ? errMessages.join(' | ') : undefined,
            }
          );
          showToast(`Berhasil menyinkronkan data (${successCount}/2 spreadsheet).`);
        } else {
          const errMsg = errMessages.join('\n') || 'Gagal memuat data spreadsheet';
          addLog('sync', 'Gagal Sinkronisasi Google Sheets', errMsg, 'error', {
            errorMessage: errMsg,
            target: `${config.stockSheetName}, ${config.balistSheetName}`,
          });
          showToast('Gagal memuat data spreadsheet. Periksa pesan di bawah.', 'error');
        }
      } catch (err: any) {
        console.error('Failed to load sheets data:', err);
        const errMsg = err?.message || 'Gagal memuat data dari Google Sheets';
        setSheetsError(errMsg);
        addLog('sync', 'Error Sinkronisasi', errMsg, 'error', {
          errorMessage: errMsg,
        });
        showToast(errMsg, 'error');
      } finally {
        setIsLoadingSheets(false);
      }
    },
    [accessToken, config, addLog]
  );

  // Upload Manual STOCK LIST XLSX File
  const handleStockListFileUpload = async (file: File) => {
    try {
      const parsed = await parseGenericXlsx(file, config.stockSheetName);
      const items = parseStockListFromRows(parsed.rows);
      if (items.length === 0) {
        showToast('Tidak ada data produk yang terbaca dari file ini.', 'error');
        addLog(
          'upload',
          'Gagal Membaca File STOCK LIST',
          `Tidak ada data produk yang valid ditemukan pada file "${file.name}".`,
          'error',
          { target: file.name }
        );
        return;
      }
      setStockList(items);
      const loadTime = new Date();
      setLastLoaded(loadTime);
      addLog(
        'upload',
        'Upload STOCK LIST Manual Sukses',
        `Berhasil memuat ${items.length.toLocaleString('id-ID')} produk dari file "${file.name}".`,
        'success',
        { rowCount: items.length, target: file.name }
      );
      showToast(`Berhasil memuat ${items.length.toLocaleString('id-ID')} produk STOCK LIST dari file "${file.name}"!`);
    } catch (err: any) {
      console.error('Stock file parse error:', err);
      const errMsg = err?.message || 'Gagal membaca file STOCK LIST';
      addLog('upload', 'Error Baca File STOCK LIST', errMsg, 'error', {
        errorMessage: errMsg,
        target: file.name,
      });
      showToast(errMsg, 'error');
    }
  };

  // UPLOAD 1: Handle Balist XLSX File
  const handleBalistFileUpload = async (file: File) => {
    try {
      setBalistUploadedFile(file);
      const parsed = await parseGenericXlsx(file, config.balistSheetName);
      setParsedBalistXlsx(parsed);

      // Auto-detect default start row: if >= 7 rows, default to row 7; otherwise row 2
      const defaultStartRow = parsed.rows.length >= 7 ? 7 : 2;
      setBalistSourceStartRow(defaultStartRow);

      // Auto-detect stock column index
      const detected = detectBalistStockColumn(parsed.rows);
      setBalistStockColIndex(detected.stockColIndex);

      // Auto-detect store from filename (31475604 -> balist, 56977507 -> Gomall)
      const detectedStore = detectStoreFromFilename(file.name);

      // Instantly parse into balistList starting at defaultStartRow
      const items = parseBalistShopeeFromRows(parsed.rows, defaultStartRow);
      if (items.length > 0) {
        setBalistList(items);
      }

      const storeDetectionMsg = detectedStore.storeName
        ? ` [Toko Terdeteksi: ${detectedStore.storeName} (${detectedStore.storeCode || detectedStore.storePrefix})] ➔ Format Unduh Otomatis: ${detectedStore.storePrefix}`
        : '';

      addLog(
        'upload',
        'Upload File Balistshopee Sukses',
        `File "${file.name}" (${parsed.rows.length.toLocaleString('id-ID')} baris total) terbaca. Kolom Stok Masuk terdeteksi pada Kolom ${
          detected.stockColIndex + 1
        } (${getColumnLetter(detected.stockColIndex)}).${storeDetectionMsg}`,
        'success',
        { rowCount: parsed.rows.length, target: file.name }
      );

      showToast(
        `File "${file.name}" terbaca (${parsed.rows.length} baris).${
          detectedStore.storeName ? ` Toko: ${detectedStore.storeName} (${detectedStore.storePrefix})` : ''
        }`
      );
    } catch (err: any) {
      console.error('Balist file parse error:', err);
      const errMsg = err?.message || 'Gagal membaca file Excel Balist';
      addLog('upload', 'Gagal Baca File Balist', errMsg, 'error', {
        errorMessage: errMsg,
        target: file.name,
      });
      showToast(errMsg, 'error');
    }
  };

  const handleBalistSourceStartRowChange = (row: number) => {
    setBalistSourceStartRow(row);
    if (parsedBalistXlsx) {
      const items = parseBalistShopeeFromRows(parsedBalistXlsx.rows, row);
      if (items.length > 0) {
        setBalistList(items);
      }
    }
  };

  const handleClearBalistFile = () => {
    setBalistUploadedFile(null);
    setParsedBalistXlsx(null);
    addLog('upload', 'File Balist Dihapus', 'File unggahan Balistshopee telah dibersihkan.', 'info');
  };

  const handleSelectBalistSheetName = async (sheetName: string) => {
    if (!balistUploadedFile) return;
    try {
      const parsed = await parseGenericXlsx(balistUploadedFile, sheetName);
      setParsedBalistXlsx(parsed);
      const detected = detectBalistStockColumn(parsed.rows);
      setBalistStockColIndex(detected.stockColIndex);
      const items = parseBalistShopeeFromRows(parsed.rows, balistSourceStartRow);
      if (items.length > 0) {
        setBalistList(items);
      }
      addLog(
        'upload',
        'Ganti Sheet Balist',
        `Sheet "${sheetName}" dipilih dari file "${balistUploadedFile.name}".`,
        'info',
        { target: sheetName }
      );
    } catch (err) {
      console.error(err);
    }
  };

  const balistMatchedStockCount = useMemo(() => {
    if (!parsedBalistXlsx || !updateStockFromStockList) return 0;
    const baseRows =
      balistSourceStartRow > 1
        ? parsedBalistXlsx.rows.slice(balistSourceStartRow - 1)
        : parsedBalistXlsx.rows;
    return updateBalistRowsWithStock(
      baseRows,
      matchingMaps,
      balistStockColIndex,
      { unmatchedAction: balistUnmatchedStockAction }
    ).stats.matchedCount;
  }, [parsedBalistXlsx, updateStockFromStockList, balistSourceStartRow, matchingMaps, balistStockColIndex, balistUnmatchedStockAction]);

  // Perform Google Sheet Update for Balistshopee starting from ROW 7
  const handleExecuteBalistSheetUpdate = async () => {
    if (!accessToken) {
      showToast('Akses akun Google diperlukan untuk memperbarui Google Sheet.', 'error');
      addLog('sheet_update', 'Pembaruan Sheet Dibatalkan', 'Belum terautentikasi dengan Google.', 'error');
      return;
    }
    if (!parsedBalistXlsx || parsedBalistXlsx.rows.length === 0) {
      showToast('Tidak ada data baris dari file Excel untuk diunggah.', 'error');
      return;
    }

    setIsUpdatingBalistSheet(true);
    addLog(
      'sheet_update',
      'Memulai Pembaruan Google Sheet',
      `Menyiapkan pembaruan data sheet "${config.balistSheetName}" mulai baris ke-7...`,
      'loading',
      { target: config.balistSheetName }
    );

    try {
      showToast(`Mengosongkan & memperbarui data sheet ${config.balistSheetName} mulai Baris ke-7...`);

      // Determine rows to write from uploaded file
      const baseRows =
        balistSourceStartRow > 1
          ? parsedBalistXlsx.rows.slice(balistSourceStartRow - 1)
          : parsedBalistXlsx.rows;

      let rowsToSend = baseRows;
      let stockStats: { matchedCount: number; inStockCount: number } | null = null;

      if (updateStockFromStockList) {
        const res = updateBalistRowsWithStock(
          baseRows,
          matchingMaps,
          balistStockColIndex,
          { unmatchedAction: balistUnmatchedStockAction }
        );
        rowsToSend = res.updatedRows;
        stockStats = res.stats;
      }

      // 1. Clear existing values starting ONLY from row 7 downwards (A7:ZZ)
      // Baris 1 s/d 6 (header, format, struktur kolom) tetap utuh terlindungi!
      await clearSheetValues(
        config.balistSpreadsheetId,
        `${config.balistSheetName}!A7:ZZ`,
        accessToken
      );

      // 2. Write rows to Balistshopee starting at cell A7
      await updateSheetValues(
        config.balistSpreadsheetId,
        `${config.balistSheetName}!A7`,
        rowsToSend,
        accessToken
      );

      const updateTime = new Date();
      setLastSheetUpdateTime(updateTime);
      localStorage.setItem('last_sheet_update_time', updateTime.toISOString());

      const stockMsg = stockStats
        ? ` & kolom ${getColumnLetter(balistStockColIndex)} (Stok Masuk) berhasil diisi untuk ${stockStats.matchedCount} produk cocok`
        : '';

      addLog(
        'sheet_update',
        'Pembaruan Sheet Balistshopee Berhasil',
        `Sheet "${config.balistSheetName}" berhasil diperbarui mulai baris 7 (${rowsToSend.length.toLocaleString(
          'id-ID'
        )} baris data)${stockMsg}. Format & baris 1-6 aman terlindungi.`,
        'success',
        {
          rowCount: rowsToSend.length,
          target: `${config.balistSheetName}!A7:ZZ`,
          details: `Diperbarui pada: ${updateTime.toLocaleString('id-ID')}\nTotal Baris: ${rowsToSend.length}\nKolom Stok: ${getColumnLetter(
            balistStockColIndex
          )}\nStok Terhubung: ${stockStats ? stockStats.matchedCount : 'N/A'}`,
        }
      );

      showToast(
        `Sheet "${config.balistSheetName}" berhasil diperbarui mulai baris 7 (${rowsToSend.length.toLocaleString(
          'id-ID'
        )} baris data)${stockMsg}! Silakan tekan Refresh / F5 di tab spreadsheet untuk melihat data terbaru.`
      );
      setIsConfirmBalistUploadModalOpen(false);

      // Ensure local state reflects newest uploaded data starting from row 7
      const items = parseBalistShopeeFromRows(rowsToSend, 1);
      if (items.length > 0) {
        setBalistList(items);
      }
    } catch (err: any) {
      console.error('Failed to update Balist sheet:', err);
      const errMsg = err?.message || 'Gagal memperbarui Google Sheet Balistshopee';
      setSheetsError(errMsg);
      addLog('sheet_update', 'Gagal Memperbarui Sheet Balistshopee', errMsg, 'error', {
        errorMessage: errMsg,
        target: `${config.balistSheetName}!A7`,
        details: `Spreadsheet ID: ${config.balistSpreadsheetId}\nError: ${errMsg}`,
      });
      showToast(errMsg, 'error');
    } finally {
      setIsUpdatingBalistSheet(false);
    }
  };

  // Perform Direct Stock Update on existing Balistshopee sheet without needing to re-upload Excel
  const handleExecuteDirectBalistStockUpdate = async () => {
    if (!accessToken) {
      showToast('Akses akun Google diperlukan untuk memperbarui Google Sheet.', 'error');
      addLog('sheet_update', 'Update Stok Gagal', 'Akun Google belum terhubung.', 'error');
      return;
    }
    if (balistList.length === 0) {
      showToast('Tidak ada data baris pada sheet Balistshopee untuk diperbarui.', 'error');
      return;
    }

    setIsDirectUpdatingBalistStock(true);
    addLog(
      'sheet_update',
      'Memulai Update Stok Balist Langsung',
      `Mencocokkan stok dari sheet "${config.stockSheetName}" ke "${config.balistSheetName}"...`,
      'loading',
      { target: config.balistSheetName }
    );

    try {
      showToast(`Mencocokkan stok dari sheet STOCK LIST ke ${config.balistSheetName}...`);

      const rawRows = balistList.map((item) => [...item.rawRow]);
      const detected = detectBalistStockColumn(rawRows);
      const targetCol = balistStockColIndex !== undefined ? balistStockColIndex : detected.stockColIndex;

      const { updatedRows, stats } = updateBalistRowsWithStock(
        rawRows,
        matchingMaps,
        targetCol,
        { unmatchedAction: balistUnmatchedStockAction }
      );

      // Write updated rows to A7
      await updateSheetValues(
        config.balistSpreadsheetId,
        `${config.balistSheetName}!A7`,
        updatedRows,
        accessToken
      );

      const updateTime = new Date();
      setLastSheetUpdateTime(updateTime);
      localStorage.setItem('last_sheet_update_time', updateTime.toISOString());

      addLog(
        'sheet_update',
        'Update Stok Balistshopee Berhasil',
        `Kolom Stok Masuk (${getColumnLetter(targetCol)}) pada sheet "${config.balistSheetName}" berhasil diisi (${stats.matchedCount} produk cocok).`,
        'success',
        {
          rowCount: updatedRows.length,
          target: `${config.balistSheetName}!A7`,
          details: `Stok Cocok: ${stats.matchedCount} baris\nStok Ready: ${stats.inStockCount} baris\nKolom Target: Kolom ${targetCol + 1} (${getColumnLetter(targetCol)})`,
        }
      );

      showToast(
        `Berhasil memperbarui kolom Stok Masuk (${getColumnLetter(targetCol)}) pada sheet "${config.balistSheetName}"! ` +
        `${stats.matchedCount} produk cocok dicocokkan dengan STOCK LIST. (Silakan tekan Refresh / F5 pada tab spreadsheet untuk melihat data terbaru).`
      );

      setIsConfirmDirectUpdateModalOpen(false);
      await loadSheets();
    } catch (err: any) {
      console.error('Direct stock update error:', err);
      const errMsg = err?.message || 'Gagal memperbarui stok pada Google Sheet';
      setSheetsError(errMsg);
      addLog('sheet_update', 'Gagal Update Stok Balist', errMsg, 'error', {
        errorMessage: errMsg,
        target: `${config.balistSheetName}!A7`,
      });
      showToast(errMsg, 'error');
    } finally {
      setIsDirectUpdatingBalistStock(false);
    }
  };

  // Convert Office File (.xlsx) to native Google Spreadsheet
  const handleConvertBalistToGoogleSheet = async () => {
    if (!accessToken) {
      showToast('Akses akun Google diperlukan untuk mengonversi spreadsheet.', 'error');
      return;
    }

    setIsConvertingBalist(true);
    addLog(
      'settings',
      'Mengonversi Format Spreadsheet',
      `Mengonversi file Excel ${config.balistSpreadsheetId} ke Google Spreadsheet resmi...`,
      'loading',
      { target: config.balistSpreadsheetId }
    );

    try {
      showToast(`Mengonversi file Excel ${config.balistSpreadsheetId} ke Google Spreadsheet resmi...`);

      const converted = await convertOfficeFileToGoogleSheet(
        config.balistSpreadsheetId,
        'Balistshopee (Resmi Google Sheets)',
        accessToken
      );

      const newId = converted.id;
      const newConfig: SheetsConfig = {
        ...config,
        balistSpreadsheetId: newId,
        balistSheetName: 'Balistshopee',
      };

      setConfig(newConfig);
      localStorage.setItem('shopee_sheets_config', JSON.stringify(newConfig));

      const sheetUrl = converted.webViewLink || `https://docs.google.com/spreadsheets/d/${newId}`;
      setConvertedSpreadsheetInfo({
        id: newId,
        name: converted.name || 'Balistshopee (Resmi Google Sheets)',
        url: sheetUrl,
      });

      // If user uploaded a file, immediately write it starting from row 7 to cell A7
      if (parsedBalistXlsx && parsedBalistXlsx.rows.length > 0) {
        showToast('Menyalin data ke Google Spreadsheet baru mulai baris ke-7...');
        const baseRows =
          balistSourceStartRow > 1
            ? parsedBalistXlsx.rows.slice(balistSourceStartRow - 1)
            : parsedBalistXlsx.rows;

        let rowsToSend = baseRows;
        if (updateStockFromStockList) {
          const res = updateBalistRowsWithStock(
            baseRows,
            matchingMaps,
            balistStockColIndex,
            { unmatchedAction: balistUnmatchedStockAction }
          );
          rowsToSend = res.updatedRows;
        }

        await updateSheetValues(newId, 'Balistshopee!A7', rowsToSend, accessToken);
      }

      addLog(
        'settings',
        'Konversi Google Spreadsheet Sukses',
        `File berhasil dikonversi menjadi Google Spreadsheet resmi (ID: ${newId}). Sinkronisasi kini berfungsi real-time.`,
        'success',
        { target: newId, details: `URL: ${sheetUrl}` }
      );

      showToast(
        `Sukses! File berhasil dikonversi menjadi Google Spreadsheet resmi. Sinkronisasi kini berfungsi real-time!`
      );

      await loadSheets(accessToken);
    } catch (err: any) {
      console.error('Conversion error:', err);
      const msg = err?.message || 'Gagal mengonversi file ke Google Spreadsheet';
      setSheetsError(msg);
      addLog('settings', 'Konversi Google Spreadsheet Gagal', msg, 'error', {
        errorMessage: msg,
      });
      showToast(msg, 'error');
    } finally {
      setIsConvertingBalist(false);
    }
  };

  // Download directly as Excel (.xlsx) file with custom/preset prefix (e.g. 'balist', 'gomall')
  const handleDownloadUpdatedBalistXlsx = (prefix: string = 'balist') => {
    if (!parsedBalistXlsx || parsedBalistXlsx.rows.length === 0) {
      showToast('Unggah file Excel Balistshopee terlebih dahulu.', 'error');
      return;
    }
    try {
      const baseRows =
        balistSourceStartRow > 1
          ? parsedBalistXlsx.rows.slice(balistSourceStartRow - 1)
          : parsedBalistXlsx.rows;

      let rowsToSend = baseRows;
      if (updateStockFromStockList) {
        const res = updateBalistRowsWithStock(
          baseRows,
          matchingMaps,
          balistStockColIndex,
          { unmatchedAction: balistUnmatchedStockAction }
        );
        rowsToSend = res.updatedRows;
      }

      const headerRows = parsedBalistXlsx.rows.slice(0, Math.min(6, balistSourceStartRow - 1));
      const fullRows = [...headerRows, ...rowsToSend];
      const balistFileName = generateShopeeBalistFilename(prefix || 'balist');

      downloadBalistRowsAsXlsx(fullRows, balistFileName, {
        rawWorkbookData: parsedBalistXlsx.rawWorkbookData,
        sheetName: parsedBalistXlsx.selectedSheetName || 'Balistshopee',
        startRowIndex: balistSourceStartRow,
        updatedRowValues: rowsToSend,
      });

      addLog(
        'download',
        `Unduh Excel ${prefix} Siap Pakai`,
        `File "${balistFileName}" (${fullRows.length} baris total) berhasil diunduh dengan kolom Stok Masuk terisi.`,
        'success',
        { rowCount: fullRows.length, target: balistFileName }
      );

      showToast(`File "${balistFileName}" berhasil diunduh dengan format asli persis!`);
    } catch (err: any) {
      console.error('Download Balist error:', err);
      const errMsg = err?.message || 'Gagal mengunduh file Excel';
      addLog('download', 'Gagal Unduh Excel Balist', errMsg, 'error', { errorMessage: errMsg });
      showToast('Gagal mengunduh file Excel: ' + errMsg, 'error');
    }
  };

  // Download Balistshopee data directly from Google Sheets as .xlsx with updated stock
  const handleDownloadBalistFromSheetsXlsx = async (prefix: string = 'balist') => {
    if (!accessToken) {
      showToast('Silakan masuk dengan Google untuk mengunduh data langsung dari Google Sheets.', 'error');
      return;
    }
    if (balistList.length === 0) {
      showToast('Belum ada data Google Sheets yang termuat. Silakan tarik/sinkronkan data terlebih dahulu.', 'error');
      return;
    }

    try {
      showToast('Mengambil data dari Google Sheets...');
      let fullRows: any[][] = [];
      try {
        fullRows = await fetchSheetValues(
          config.balistSpreadsheetId,
          `${config.balistSheetName}!A:Z`,
          accessToken
        );
      } catch (e) {
        console.warn('Could not fetch full A:Z, using loaded balist items:', e);
      }

      let headerRows: any[][] = [];
      let dataRows: any[][] = [];

      if (fullRows && fullRows.length >= 7) {
        headerRows = fullRows.slice(0, 6);
        dataRows = fullRows.slice(6);
      } else if (fullRows && fullRows.length > 0) {
        headerRows = fullRows.slice(0, 1);
        dataRows = fullRows.slice(1);
      } else {
        dataRows = balistList.map((item) => [...item.rawRow]);
      }

      let rowsToSend = dataRows;
      if (updateStockFromStockList) {
        const res = updateBalistRowsWithStock(
          dataRows,
          matchingMaps,
          balistStockColIndex,
          { unmatchedAction: balistUnmatchedStockAction }
        );
        rowsToSend = res.updatedRows;
      }

      const rowsToExport = [...headerRows, ...rowsToSend];
      const balistFileName = generateShopeeBalistFilename(prefix || 'balist');

      downloadBalistRowsAsXlsx(rowsToExport, balistFileName, {
        sheetName: config.balistSheetName || 'Balistshopee',
        startRowIndex: headerRows.length + 1,
        updatedRowValues: rowsToSend,
      });

      addLog(
        'download',
        `Unduh Excel ${prefix} dari Google Sheets`,
        `File "${balistFileName}" (${rowsToExport.length} baris) berhasil diunduh langsung dari Google Sheets.`,
        'success',
        { rowCount: rowsToExport.length, target: balistFileName }
      );

      showToast(`File "${balistFileName}" berhasil diunduh dengan stok terbaru!`);
    } catch (err: any) {
      console.error('Download Balist from Sheets error:', err);
      const errMsg = err?.message || '';
      addLog('download', 'Gagal Unduh Balist Sheets', errMsg, 'error', { errorMessage: errMsg });
      showToast('Gagal mengunduh file Excel dari Sheets: ' + errMsg, 'error');
    }
  };

  // UPLOAD 2: Handle Shopee XLSX Mass Update File
  const handleShopeeFileUpload = async (file: File) => {
    try {
      setShopeeUploadedFile(file);
      const parsed = await parseShopeeXlsx(file);
      setParsedShopeeSheet(parsed);
      setSelectedSkuCol(parsed.skuColIndex);
      setSelectedStockCol(parsed.stockColIndex);

      // Auto-detect store from filename (e.g. 31475604 -> balist, 56977507 -> Gomall)
      const detectedStore = detectStoreFromFilename(file.name);
      const storeDetectionMsg = detectedStore.storeName
        ? ` [Toko Terdeteksi: ${detectedStore.storeName} (${detectedStore.storeCode || detectedStore.storePrefix})] ➔ Format Unduh Otomatis: ${detectedStore.storePrefix}`
        : '';

      addLog(
        'upload',
        'Upload File Shopee Mass Update',
        `File "${file.name}" (${parsed.rows.length.toLocaleString('id-ID')} baris) berhasil dibaca. Kolom SKU: Kolom ${
          parsed.skuColIndex + 1
        }, Kolom Stok: Kolom ${parsed.stockColIndex + 1}.${storeDetectionMsg}`,
        'success',
        { rowCount: parsed.rows.length, target: file.name }
      );
      showToast(
        `File "${file.name}" berhasil dibaca.${
          detectedStore.storeName ? ` Terdeteksi: ${detectedStore.storeName} (${detectedStore.storePrefix})` : ''
        }`
      );
    } catch (err: any) {
      console.error('File parsing error:', err);
      const errMsg = err?.message || 'Gagal membaca file Excel Shopee';
      addLog('upload', 'Gagal Baca Shopee File', errMsg, 'error', {
        errorMessage: errMsg,
        target: file.name,
      });
      showToast(errMsg, 'error');
    }
  };

  const handleClearShopeeFile = () => {
    setShopeeUploadedFile(null);
    setParsedShopeeSheet(null);
    setMatches([]);
    setSummary(null);
    addLog('upload', 'File Shopee Dikosongkan', 'File unggahan mass update Shopee telah dibersihkan.', 'info');
  };

  // Re-run matching whenever parsed Shopee sheet, maps, column choices, or action changes
  useEffect(() => {
    if (!parsedShopeeSheet) {
      setMatches([]);
      setSummary(null);
      return;
    }

    const calculatedMatches = matchShopeeFile(parsedShopeeSheet, matchingMaps, {
      customSkuCol: selectedSkuCol,
      customStockCol: selectedStockCol,
      unmatchedAction,
    });

    setMatches(calculatedMatches);
    const calculatedSummary = calculateSummary(
      shopeeUploadedFile?.name || 'Shopee_Update.xlsx',
      calculatedMatches
    );
    setSummary(calculatedSummary);
  }, [parsedShopeeSheet, matchingMaps, selectedSkuCol, selectedStockCol, unmatchedAction, shopeeUploadedFile]);

  // Download Updated Shopee File
  const handleDownloadUpdatedXlsx = () => {
    if (!parsedShopeeSheet || matches.length === 0) {
      showToast('Tidak ada data untuk diunduh.', 'error');
      return;
    }

    try {
      const updatedData = generateUpdatedShopeeWorkbook(
        parsedShopeeSheet,
        matches,
        selectedStockCol,
        { unmatchedAction }
      );

      // Check if store code is present in filename
      const detectedStore = detectStoreFromFilename(shopeeUploadedFile?.name || '');
      const downloadName = detectedStore.storePrefix
        ? generateShopeeBalistFilename(detectedStore.storePrefix)
        : `${(shopeeUploadedFile?.name || 'shopee_stock').replace(/\.xlsx?$/i, '')}_STOK_TERUPDATE.xlsx`;

      downloadBlob(updatedData, downloadName);
      addLog(
        'download',
        'Unduh Hasil Shopee Mass Update',
        `File "${downloadName}" (${matches.length.toLocaleString('id-ID')} baris) berhasil diunduh.${
          detectedStore.storeName ? ` Format Toko: ${detectedStore.storeName}.` : ''
        } Siap diunggah ke Seller Centre Shopee.`,
        'success',
        { rowCount: matches.length, target: downloadName }
      );
      showToast(`File "${downloadName}" berhasil diunduh! Siap upload ke Shopee.`);
    } catch (err: any) {
      console.error('Download error:', err);
      const errMsg = err?.message || 'Gagal membuat file XLSX terupdate';
      addLog('download', 'Gagal Unduh Shopee Update', errMsg, 'error', { errorMessage: errMsg });
      showToast(errMsg, 'error');
    }
  };

  // Download Sample Shopee Template
  const handleDownloadSample = () => {
    try {
      const sample = createSampleShopeeFile();
      downloadBlob(sample, 'Shopee_Mass_Update_Sample.xlsx');
      addLog('download', 'Unduh Format Contoh Shopee', 'Template contoh Shopee Mass Update berhasil diunduh.', 'info');
      showToast('File template contoh berhasil diunduh.');
    } catch (err: any) {
      showToast('Gagal mengunduh sampel', 'error');
    }
  };

  const handleSaveConfig = (newConfig: SheetsConfig) => {
    setConfig(newConfig);
    localStorage.setItem('shopee_sheets_config', JSON.stringify(newConfig));
    addLog(
      'settings',
      'Pengaturan Spreadsheet Disimpan',
      `ID STOCK LIST: ${newConfig.stockSpreadsheetId.substring(0, 10)}..., ID Balist: ${newConfig.balistSpreadsheetId.substring(0, 10)}...`,
      'info',
      { target: `${newConfig.stockSheetName}, ${newConfig.balistSheetName}` }
    );
    showToast('Pengaturan spreadsheet disimpan.');
    if (accessToken) {
      loadSheets(accessToken);
    }
  };

  const handleResetConfig = () => {
    setConfig(DEFAULT_SHEETS_CONFIG);
    localStorage.removeItem('shopee_sheets_config');
    addLog('settings', 'Pengaturan Direset ke Default', 'Konfigurasi ID spreadsheet dikembalikan ke bawaan.', 'info');
    showToast('Pengaturan dikembalikan ke default.');
    if (accessToken) {
      loadSheets(accessToken);
    }
  };

  return (
    <div className="min-h-screen bg-stone-100/70 text-stone-900 flex flex-col font-sans">
      {/* Toast Notification */}
      {toastMessage && (
        <div
          className={`fixed bottom-6 right-6 z-50 px-4 py-3 rounded-xl shadow-lg border text-xs font-medium flex items-center gap-2 max-w-md animate-in slide-in-from-bottom-5 ${
            toastMessage.type === 'success'
              ? 'bg-emerald-900 text-white border-emerald-800'
              : 'bg-rose-900 text-white border-rose-800'
          }`}
        >
          {toastMessage.type === 'success' ? (
            <CheckCircle className="w-4 h-4 text-emerald-400 shrink-0" />
          ) : (
            <Info className="w-4 h-4 text-rose-400 shrink-0" />
          )}
          <span>{toastMessage.text}</span>
        </div>
      )}

      {/* Header */}
      <Header
        user={user}
        onSignIn={handleSignIn}
        onSignOut={handleSignOut}
        onOpenSettings={() => setIsSettingsOpen(true)}
        isLoadingAuth={isLoadingAuth}
        sheetsConfig={config}
      />

      {/* Main Container */}
      <main className="max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 flex-1 space-y-6">
        {/* Live Status Google Sheets Connection */}
        {showSheetsStatusCard ? (
          <SheetsStatusCard
            config={config}
            isLoading={isLoadingSheets}
            stockCount={stockList.length}
            balistCount={balistList.length}
            lastLoaded={lastLoaded}
            error={sheetsError}
            isAuthenticated={!!user}
            onRefresh={() => loadSheets()}
            onPromptSignIn={handleSignIn}
            onUploadStockListFile={handleStockListFileUpload}
            onOpenSettings={() => setIsSettingsOpen(true)}
            onConvertBalistToGoogleSheet={handleConvertBalistToGoogleSheet}
            isConvertingBalist={isConvertingBalist}
            onHide={() => {
              setShowSheetsStatusCard(false);
              localStorage.setItem('show_sheets_status_card', 'false');
              showToast('Kartu status koneksi Google Sheets disembunyikan.');
            }}
            showActivityLogs={showActivityLogPanel}
            onToggleActivityLogs={() => {
              setShowActivityLogPanel((prev) => {
                const next = !prev;
                localStorage.setItem('show_activity_log_panel', String(next));
                showToast(next ? 'Panel Log Riwayat ditampilkan.' : 'Panel Log Riwayat disembunyikan.');
                return next;
              });
            }}
          />
        ) : (
          <div className="flex items-center justify-between p-3 bg-stone-50 border border-stone-200 rounded-xl text-xs text-stone-600">
            <div className="flex items-center gap-2">
              <Database className="w-4 h-4 text-emerald-600 shrink-0" />
              <span>
                Koneksi Google Sheets:{' '}
                <strong className="text-stone-800">
                  {stockList.length > 0 ? `${stockList.length.toLocaleString('id-ID')} baris STOCK LIST termuat` : 'Terkoneksi'}
                </strong>
              </span>
            </div>
            <div className="flex items-center gap-2">
              {!!user && (
                <button
                  type="button"
                  onClick={() => loadSheets()}
                  disabled={isLoadingSheets}
                  className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-white hover:bg-stone-100 text-stone-700 border border-stone-200 rounded-lg text-xs font-medium transition-colors"
                >
                  <RefreshCw className={`w-3 h-3 ${isLoadingSheets ? 'animate-spin' : ''}`} />
                  <span>Muat Ulang</span>
                </button>
              )}
              <button
                type="button"
                id="btn-show-sheets-status"
                onClick={() => {
                  setShowSheetsStatusCard(true);
                  localStorage.setItem('show_sheets_status_card', 'true');
                }}
                className="inline-flex items-center gap-1.5 px-3 py-1 bg-white hover:bg-stone-100 text-stone-700 border border-stone-200 rounded-lg text-xs font-medium shadow-2xs transition-colors"
              >
                <Eye className="w-3.5 h-3.5 text-stone-500" />
                <span>Detail Koneksi</span>
              </button>
              {!showActivityLogPanel && (
                <button
                  type="button"
                  id="btn-show-activity-logs"
                  onClick={() => {
                    setShowActivityLogPanel(true);
                    localStorage.setItem('show_activity_log_panel', 'true');
                    showToast('Panel Log Riwayat ditampilkan.');
                  }}
                  className="inline-flex items-center gap-1.5 px-3 py-1 bg-white hover:bg-stone-100 text-stone-700 border border-stone-200 rounded-lg text-xs font-medium shadow-2xs transition-colors"
                  title="Tampilkan Panel Log Riwayat Aktivitas & Pembaruan"
                >
                  <History className="w-3.5 h-3.5 text-stone-500" />
                  <span>Log Riwayat ({activityLogs.length})</span>
                </button>
              )}
            </div>
          </div>
        )}

        {/* Panel Log Riwayat Aktivitas & Status Pembaruan (Hanya tampil jika diaktifkan) */}
        {showActivityLogPanel && (
          <ActivityLogPanel
            logs={activityLogs}
            onClearLogs={handleClearLogs}
            lastSyncTime={lastLoaded}
            lastSheetUpdateTime={lastSheetUpdateTime}
            isSyncing={isLoadingSheets}
            onHide={() => {
              setShowActivityLogPanel(false);
              localStorage.setItem('show_activity_log_panel', 'false');
              showToast('Panel Log Riwayat disembunyikan.');
            }}
          />
        )}

        {/* SECTION 1: TOMBOL UPLOAD 1 (PERBARUI BALISTSHOPEE & PERBANDINGAN STOCK LIST) */}
        {(activeTab === 'workflow' || activeTab === 'balist_comparison') && (
          <div className="space-y-4">
            <BalistUploadCard
              parsedFile={parsedBalistXlsx}
              uploadedFileName={balistUploadedFile?.name || null}
              sourceStartRow={balistSourceStartRow}
              onSourceStartRowChange={handleBalistSourceStartRowChange}
              onFileUpload={handleBalistFileUpload}
              onClearFile={handleClearBalistFile}
              onSelectSheetName={handleSelectBalistSheetName}
              onUpdateSheet={() => setIsConfirmBalistUploadModalOpen(true)}
              isUpdating={isUpdatingBalistSheet}
              isAuthenticated={!!user}
              spreadsheetId={config.balistSpreadsheetId}
              sheetName={config.balistSheetName}
              onPromptSignIn={handleSignIn}
              selectedStockColIndex={balistStockColIndex}
              onSelectedStockColIndexChange={setBalistStockColIndex}
              updateStockFromStockList={updateStockFromStockList}
              onUpdateStockFromStockListChange={setUpdateStockFromStockList}
              unmatchedStockAction={balistUnmatchedStockAction}
              onUnmatchedStockActionChange={setBalistUnmatchedStockAction}
              matchedStockCount={balistMatchedStockCount}
              onConvertOfficeToGoogleSheet={handleConvertBalistToGoogleSheet}
              isConverting={isConvertingBalist}
              onDownloadUpdatedBalistXlsx={handleDownloadUpdatedBalistXlsx}
              isOfficeFile={config.balistSpreadsheetId === '1wTchgk4-YRyQv-Sk10SZUrOooGMrC08S'}
            />

            {/* Live Comparison: Balistshopee vs STOCK LIST */}
            {balistList.length > 0 && (
              <BalistComparisonTable
                items={balistComparison.items}
                summary={balistComparison.summary}
                balistSheetName={config.balistSheetName}
                stockSheetName={config.stockSheetName}
                uploadedFileName={balistUploadedFile?.name || parsedBalistXlsx?.fileName}
                onUpdateBalistStockInSheet={() => setIsConfirmDirectUpdateModalOpen(true)}
                isUpdatingBalistStock={isDirectUpdatingBalistStock}
              />
            )}
          </div>
        )}

        {/* SECTION 2: TOMBOL UPLOAD 2 (UPLOAD SHOPEE MASS UPDATE & UPDATE STOK) */}
        {(activeTab === 'shopee_match' || (activeTab === 'workflow' && showShopeeSection)) ? (
          <div className="space-y-4">
            <UploadSection
              parsedFile={parsedShopeeSheet}
              uploadedFileName={shopeeUploadedFile?.name || parsedShopeeSheet?.fileName}
              isLoading={false}
              selectedSkuCol={selectedSkuCol}
              selectedStockCol={selectedStockCol}
              unmatchedAction={unmatchedAction}
              onFileUpload={handleShopeeFileUpload}
              onClearFile={handleClearShopeeFile}
              onChangeSkuCol={setSelectedSkuCol}
              onChangeStockCol={setSelectedStockCol}
              onChangeUnmatchedAction={setUnmatchedAction}
              onDownloadSample={handleDownloadSample}
              onHide={() => {
                setShowShopeeSection(false);
                localStorage.setItem('show_shopee_section', 'false');
                showToast('Menu Upload File Shopee berhasil disembunyikan.');
              }}
            />

            {/* Summary & Download Action for Shopee File */}
            {summary && (
              <SummaryCards
                summary={summary}
                onDownload={handleDownloadUpdatedXlsx}
                onSyncBalist={() => setIsConfirmModalOpen(true)}
                isSyncingBalist={isSyncingBalist}
                canSyncBalist={false}
              />
            )}

            {/* Match Table for Shopee File */}
            {matches.length > 0 && <MatchTable matches={matches} />}
          </div>
        ) : activeTab === 'workflow' && !showShopeeSection ? (
          <div className="flex items-center justify-between p-3.5 bg-stone-50 border border-dashed border-stone-300 rounded-xl text-xs text-stone-600">
            <div className="flex items-center gap-2">
              <EyeOff className="w-4 h-4 text-stone-400" />
              <span>Menu <strong>Tombol 2: Upload File XLSX Shopee (Mass Update)</strong> disembunyikan.</span>
            </div>
            <button
              type="button"
              id="btn-show-shopee-section"
              onClick={() => {
                setShowShopeeSection(true);
                localStorage.setItem('show_shopee_section', 'true');
              }}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-white hover:bg-stone-100 text-stone-700 border border-stone-300 rounded-lg font-medium shadow-2xs transition-colors"
            >
              <Eye className="w-3.5 h-3.5 text-stone-500" />
              <span>Tampilkan Menu Shopee</span>
            </button>
          </div>
        ) : null}
      </main>

      {/* Footer */}
      <footer className="mt-auto border-t border-stone-200 bg-white py-4 text-center text-xs text-stone-500">
        <div className="max-w-7xl mx-auto px-4 flex flex-col sm:flex-row items-center justify-between gap-2">
          <span>Shopee Stock Synchronizer — 2 Tombol Upload Excel Terintegrasi</span>
          <span className="text-stone-400 text-[11px]">
            Balistshopee (Kolom 5 &amp; 6) • STOCK LIST (Kolom 1 &amp; 15) • Shopee Mass Update
          </span>
        </div>
      </footer>

      {/* Settings Modal */}
      <SettingsModal
        isOpen={isSettingsOpen}
        onClose={() => setIsSettingsOpen(false)}
        config={config}
        onSave={handleSaveConfig}
        onReset={handleResetConfig}
      />

      {/* Confirmation Modal 1: For Uploading Excel to Balistshopee starting at Row 7 */}
      <ConfirmUpdateModal
        isOpen={isConfirmBalistUploadModalOpen}
        onClose={() => setIsConfirmBalistUploadModalOpen(false)}
        onConfirm={handleExecuteBalistSheetUpdate}
        isProcessing={isUpdatingBalistSheet}
        spreadsheetName={config.balistSheetName}
        spreadsheetId={config.balistSpreadsheetId}
        matchedCount={
          parsedBalistXlsx
            ? balistSourceStartRow > 1
              ? Math.max(0, parsedBalistXlsx.rows.length - (balistSourceStartRow - 1))
              : parsedBalistXlsx.rows.length
            : 0
        }
        title={`Konfirmasi Perbarui Sheet ${config.balistSheetName} (Mulai Baris 7)`}
        description={`Pembaruan akan diterapkan mulai dari Baris ke-7 (A7). Baris 1 s/d 6 pada Google Sheet "${config.balistSheetName}" akan tetap utuh dipertahankan agar struktur kolom tidak berubah.${
          updateStockFromStockList
            ? ` Nilai kolom Stok Masuk (${getColumnLetter(balistStockColIndex)}) juga akan otomatis diisi dengan data stok gudang dari sheet STOCK LIST (${balistMatchedStockCount} produk cocok).`
            : ''
        }`}
        actionText={`Mengosongkan baris 7+ lalu menulis data baru ke ${config.balistSheetName}!A7 (Baris 1-6 aman)`}
      />

      {/* Confirmation Modal 2: For Direct Stock Update from STOCK LIST to Balistshopee Sheet */}
      <ConfirmUpdateModal
        isOpen={isConfirmDirectUpdateModalOpen}
        onClose={() => setIsConfirmDirectUpdateModalOpen(false)}
        onConfirm={handleExecuteDirectBalistStockUpdate}
        isProcessing={isDirectUpdatingBalistStock}
        spreadsheetName={config.balistSheetName}
        spreadsheetId={config.balistSpreadsheetId}
        matchedCount={balistComparison.summary.matchedCount}
        title={`Perbarui Nilai Stok di Sheet ${config.balistSheetName}`}
        description={`Sistem akan mencocokkan SKU (Kolom 5 & 6) pada sheet "${config.balistSheetName}" dengan Kolom 1 sheet "${config.stockSheetName}", lalu mengisi jumlah stok gudang ke kolom Stok Masuk (${getColumnLetter(balistStockColIndex)}) mulai baris ke-7. Baris 1-6 tetap aman tidak diubah.`}
        actionText={`Tulis nilai stok ke kolom ${getColumnLetter(balistStockColIndex)} sheet ${config.balistSheetName}!A7`}
      />

      {/* Modal: Google Spreadsheet Berhasil Dikonversi */}
      {convertedSpreadsheetInfo && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-xs animate-in fade-in duration-200">
          <div className="bg-white rounded-2xl border border-emerald-200 shadow-2xl max-w-lg w-full p-6 space-y-4">
            <div className="flex items-start gap-3">
              <div className="w-10 h-10 rounded-xl bg-emerald-100 text-emerald-700 flex items-center justify-center shrink-0">
                <Sparkles className="w-6 h-6" />
              </div>
              <div className="flex-1">
                <h3 className="text-base font-bold text-stone-900">
                  Google Spreadsheet Baru Berhasil Dibuat!
                </h3>
                <p className="text-xs text-stone-600 mt-1 leading-relaxed">
                  Dokumen lama (<code className="bg-stone-100 px-1 py-0.5 rounded font-mono text-[11px]">1wTchgk4-YRyQv-Sk10SZUrOooGMrC08S</code>) yang berformat Excel telah sukses dikonversi menjadi Google Spreadsheet resmi.
                </p>
              </div>
            </div>

            <div className="p-3 bg-stone-50 border border-stone-200 rounded-xl space-y-2 text-xs">
              <div className="flex items-center justify-between">
                <span className="text-stone-500">Nama Spreadsheet Baru:</span>
                <span className="font-semibold text-stone-800">{convertedSpreadsheetInfo.name}</span>
              </div>
              <div className="flex items-center justify-between">
                <span className="text-stone-500">ID Baru:</span>
                <span className="font-mono text-stone-700 text-[11px] truncate max-w-[240px]">
                  {convertedSpreadsheetInfo.id}
                </span>
              </div>
              <p className="text-emerald-700 font-medium pt-1 text-[11px]">
                ID baru telah otomatis disimpan ke Pengaturan aplikasi. Sinkronisasi stok kini selalu ter-refresh seketika tanpa masalah cache!
              </p>
            </div>

            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setConvertedSpreadsheetInfo(null)}
                className="px-4 py-2 text-xs font-semibold text-stone-600 hover:bg-stone-100 rounded-lg transition-colors"
              >
                Tutup
              </button>
              <a
                href={convertedSpreadsheetInfo.url}
                target="_blank"
                rel="noreferrer"
                onClick={() => setConvertedSpreadsheetInfo(null)}
                className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-bold text-white bg-emerald-700 hover:bg-emerald-800 rounded-lg shadow-xs transition-colors"
              >
                <span>Buka Google Spreadsheet Baru</span>
                <ExternalLink className="w-3.5 h-3.5" />
              </a>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
