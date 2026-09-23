import React, { useState, useEffect, useMemo, useCallback } from 'react';
import { User } from 'firebase/auth';
import {
  initAuth,
  googleSignIn,
  logout,
  getAccessToken,
  setCachedToken,
  handleAuthExpiry,
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
  cleanSku,
  isAuthError,
} from './lib/sheets';
import {
  buildMatchingMaps,
  MatchingMaps,
  calculateSummary,
  compareBalistWithStockList,
  updateBalistRowsWithStock,
  findStockForSku,
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
  parseTokopediaXlsx,
  matchTokopediaFile,
  generateUpdatedTokopediaWorkbook,
  generateTokopediaFilename,
  createSampleTokopediaFile,
  ParsedTokopediaSheet,
} from './lib/tokopediaProcessor';
import {
  StockListItem,
  BalistShopeeItem,
  ShopeeRowMatch,
  TokopediaRowMatch,
  MarketplacePlatform,
  ProcessSummary,
  SheetsConfig,
  BalistComparisonItem,
  BalistComparisonSummary,
  ActivityLogItem,
  LogType,
  LogStatus,
  ExportFilterOptions,
} from './types';
import { Header } from './components/Header';
import { SheetsStatusCard } from './components/SheetsStatusCard';
import { ActivityLogPanel } from './components/ActivityLogPanel';
import { StockListUploadCard } from './components/StockListUploadCard';
import { BalistUploadCard } from './components/BalistUploadCard';
import { BalistComparisonTable } from './components/BalistComparisonTable';
import { UploadSection } from './components/UploadSection';
import { TokopediaStockSection } from './components/TokopediaStockSection';
import { SummaryCards } from './components/SummaryCards';
import { MatchTable } from './components/MatchTable';
import { SettingsModal } from './components/SettingsModal';
import { ConfirmUpdateModal } from './components/ConfirmUpdateModal';
import { ShopeeDownloadModal } from './components/ShopeeDownloadModal';
import { UploadLoadingModal, UploadProgressState } from './components/UploadLoadingModal';
import { ArrowRight, CheckCircle, Info, Sparkles, Database, ShoppingBag, GitCompare, ExternalLink, Eye, EyeOff, RefreshCw, History, FileSpreadsheet, Package } from 'lucide-react';

const DEFAULT_SHEETS_CONFIG: SheetsConfig = {
  balistSpreadsheetId: '1wTchgk4-YRyQv-Sk10SZUrOooGMrC08S',
  balistSheetName: 'Balistshopee',
  stockSpreadsheetId: '1mrD9sQK_Sffa1X1fzlCDmaJXs1Yj2q-XTNdi2sRGPos',
  stockSheetName: 'STOCK LIST',
};

export default function App() {
  // Active Platform: 'shopee' | 'tokopedia' | 'stocklist'
  const [activePlatform, setActivePlatform] = useState<MarketplacePlatform>(() => {
    const saved = localStorage.getItem('active_marketplace_platform');
    return (saved as MarketplacePlatform) || 'shopee';
  });

  // Navigation tab for Shopee view: 'all' | 'balist_sync' | 'shopee_sync'
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

  // STOCK LIST Data & Upload State
  const [stockList, setStockList] = useState<StockListItem[]>([]);
  const [uploadedStockFileName, setUploadedStockFileName] = useState<string | null>(null);
  const [uploadedStockFileSize, setUploadedStockFileSize] = useState<number | null>(null);
  const [stockDataSource, setStockDataSource] = useState<'file' | 'sheets'>('sheets');
  const [balistList, setBalistList] = useState<BalistShopeeItem[]>([]);
  const [isLoadingSheets, setIsLoadingSheets] = useState(false);
  const [isRefreshingStockList, setIsRefreshingStockList] = useState(false);
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

  // Upload 3: Tokopedia Uploaded File & Matching (SKU Kolom 4 / index 3, Stok Kolom 9 / index 8)
  const [tokopediaUploadedFile, setTokopediaUploadedFile] = useState<File | null>(null);
  const [parsedTokopediaSheet, setParsedTokopediaSheet] = useState<ParsedTokopediaSheet | null>(null);
  const [tokopediaSkuCol, setTokopediaSkuCol] = useState<number>(3); // Kolom 4 (D)
  const [tokopediaStockCol, setTokopediaStockCol] = useState<number>(8); // Kolom 9 (I)
  const [tokopediaStartRow, setTokopediaStartRow] = useState<number>(4); // Baris ke-4
  const [tokopediaUnmatchedAction, setTokopediaUnmatchedAction] = useState<'keep' | 'zero'>('keep');
  const [tokopediaMatches, setTokopediaMatches] = useState<TokopediaRowMatch[]>([]);

  // Modals & Feedback
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isConfirmModalOpen, setIsConfirmModalOpen] = useState(false);
  const [isSyncingBalist, setIsSyncingBalist] = useState(false);
  const [toastMessage, setToastMessage] = useState<{ text: string; type: 'success' | 'error' } | null>(null);

  // File Upload Loading Progress Modal State
  const [uploadProgress, setUploadProgress] = useState<UploadProgressState>({
    isOpen: false,
    fileName: '',
    fileSize: null,
    uploadType: 'generic',
    step: 1,
    stepTitle: '',
    stepDescription: '',
    progressPercent: 0,
  });

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

  // Modal download state & items for filtering by Category/Brand/Checklist
  const [isShopeeDownloadModalOpen, setIsShopeeDownloadModalOpen] = useState(false);
  const [modalDownloadPrefix, setModalDownloadPrefix] = useState<string>('balist');

  const modalComparisonItems = useMemo<BalistComparisonItem[]>(() => {
    if (parsedBalistXlsx && parsedBalistXlsx.rows.length > 0) {
      const rawItems = parseBalistShopeeFromRows(parsedBalistXlsx.rows, balistSourceStartRow);
      const res = compareBalistWithStockList(rawItems, matchingMaps);
      return res.items;
    }
    return balistComparison.items;
  }, [parsedBalistXlsx, balistSourceStartRow, matchingMaps, balistComparison.items]);

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

        let hasAuthError = false;

        if (stockRes.status === 'fulfilled') {
          setStockList(stockRes.value);
          setUploadedStockFileName(null);
          setUploadedStockFileSize(null);
          setStockDataSource('sheets');
          stockRowCount = stockRes.value.length;
          successCount++;
        } else {
          console.error('Stock list load error:', stockRes.reason);
          if (isAuthError(stockRes.reason)) {
            hasAuthError = true;
          }
          errMessages.push(`STOCK LIST: ${stockRes.reason?.message || 'Gagal memuat'}`);
        }

        if (balistRes.status === 'fulfilled') {
          setBalistList(balistRes.value);
          balistRowCount = balistRes.value.length;
          successCount++;
        } else {
          console.error('Balist load error:', balistRes.reason);
          if (isAuthError(balistRes.reason)) {
            hasAuthError = true;
          }
          errMessages.push(`Balistshopee: ${balistRes.reason?.message || 'Gagal memuat'}`);
        }

        if (hasAuthError) {
          handleAuthExpiry();
          setAccessToken(null);
          setSheetsError('Sesi akun Google Anda telah berakhir atau kredensial tidak valid. Silakan klik tombol "Masuk dengan Google" untuk menghubungkan kembali.');
          addLog('auth', 'Sesi Google Kedaluwarsa', 'Token Google kedaluwarsa atau tidak valid.', 'error');
          showToast('Sesi Google kedaluwarsa. Silakan Masuk dengan Google kembali.', 'error');
          return;
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

  // Handle Google Auth and Auto-login
  useEffect(() => {
    let isInitialMount = true;
    const unsubscribe = initAuth(
      (currentUser, token) => {
        setUser(currentUser);
        setAccessToken(token);
        if (isInitialMount) {
          loadSheets(token);
        }
      },
      () => {
        setUser(null);
        setAccessToken(null);
      }
    );
    isInitialMount = false;
    return () => unsubscribe();
  }, [loadSheets]);

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

  // Refresh only the STOCK LIST sheet on demand
  const handleRefreshStockList = useCallback(async () => {
    if (!accessToken) {
      handleSignIn();
      return;
    }

    setIsRefreshingStockList(true);
    addLog(
      'sync',
      'Memperbarui Data STOCK LIST',
      `Membaca ulang sheet "${config.stockSheetName}" dari Google Drive...`,
      'loading',
      { target: config.stockSheetName }
    );

    try {
      const items = await loadStockListData(config.stockSpreadsheetId, config.stockSheetName, accessToken);
      setStockList(items);
      setUploadedStockFileName(null);
      setUploadedStockFileSize(null);
      setStockDataSource('sheets');
      const syncTime = new Date();
      setLastLoaded(syncTime);
      setSheetsError(null);
      addLog(
        'sync',
        'STOCK LIST Berhasil Diperbarui',
        `Berhasil menyegarkan ${items.length.toLocaleString('id-ID')} baris produk dari sheet "${config.stockSheetName}".`,
        'success',
        { rowCount: items.length, target: config.stockSheetName }
      );
      showToast(`Berhasil memperbarui data STOCK LIST (${items.length.toLocaleString('id-ID')} produk)!`);
    } catch (err: any) {
      console.error('Failed to refresh STOCK LIST:', err);
      if (isAuthError(err)) {
        handleAuthExpiry();
        setAccessToken(null);
        setSheetsError('Sesi akun Google Anda telah berakhir atau token akses tidak valid. Silakan Masuk Kembali dengan Google.');
        addLog(
          'auth',
          'Sesi Google Kedaluwarsa',
          'Token akses Google kedaluwarsa saat memperbarui STOCK LIST. Silakan hubungkan kembali akun Google.',
          'error'
        );
        showToast('Sesi Google kedaluwarsa. Silakan Masuk Kembali dengan Google.', 'error');
        handleSignIn();
        return;
      }
      const errMsg = err?.message || 'Gagal memuat ulang data STOCK LIST dari Google Sheets';
      addLog('sync', 'Gagal Refresh STOCK LIST', errMsg, 'error', {
        errorMessage: errMsg,
        target: config.stockSheetName,
      });
      showToast(errMsg, 'error');
    } finally {
      setIsRefreshingStockList(false);
    }
  }, [accessToken, config, addLog, handleSignIn]);

  // Sleep utility for smooth visual progress feedback
  const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

  // Upload Manual STOCK LIST XLSX File
  const handleStockListFileUpload = async (file: File) => {
    setUploadProgress({
      isOpen: true,
      fileName: file.name,
      fileSize: file.size,
      uploadType: 'stock_list',
      step: 1,
      stepTitle: 'Membaca File STOCK LIST...',
      stepDescription: 'Mengekstrak workbook & lembar kerja spreadsheet...',
      progressPercent: 20,
    });

    try {
      await sleep(200);
      const parsed = await parseGenericXlsx(file, config.stockSheetName);

      setUploadProgress((prev) => ({
        ...prev,
        step: 2,
        stepTitle: 'Menganalisis Kolom & Format Produk...',
        stepDescription: `Memproses ${parsed.rows.length.toLocaleString('id-ID')} baris data produk...`,
        progressPercent: 55,
      }));

      await sleep(200);
      const items = parseStockListFromRows(parsed.rows);

      if (items.length === 0) {
        setUploadProgress((prev) => ({ ...prev, isOpen: false }));
        showToast('Tidak ada data produk yang terbaca dari file ini. Pastikan file memiliki kolom SKU dan Stok.', 'error');
        addLog(
          'upload',
          'Gagal Membaca File STOCK LIST',
          `Tidak ada data produk yang valid ditemukan pada file "${file.name}".`,
          'error',
          { target: file.name }
        );
        return;
      }

      setUploadProgress((prev) => ({
        ...prev,
        step: 3,
        stepTitle: 'Menyusun Database Stok Gudang...',
        stepDescription: `Menghubungkan ${items.length.toLocaleString('id-ID')} produk ke indeks pencocokan SKU...`,
        progressPercent: 85,
      }));

      await sleep(200);
      setStockList(items);
      setUploadedStockFileName(file.name);
      setUploadedStockFileSize(file.size);
      setStockDataSource('file');
      const loadTime = new Date();
      setLastLoaded(loadTime);

      setUploadProgress((prev) => ({
        ...prev,
        step: 4,
        stepTitle: 'Database Stok Siap Digunakan!',
        stepDescription: `Berhasil memuat ${items.length.toLocaleString('id-ID')} produk aktif.`,
        progressPercent: 100,
      }));

      await sleep(350);

      addLog(
        'upload',
        'Upload STOCK LIST Berhasil',
        `Berhasil memuat ${items.length.toLocaleString('id-ID')} produk dari file lokal "${file.name}". Pencocokan stok diperbarui seketika.`,
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
    } finally {
      setUploadProgress((prev) => ({ ...prev, isOpen: false }));
    }
  };

  const handleClearStockListFile = () => {
    setStockList([]);
    setUploadedStockFileName(null);
    setUploadedStockFileSize(null);
    setStockDataSource('sheets');
    addLog('upload', 'File STOCK LIST Dikosongkan', 'Data STOCK LIST lokal telah dibersihkan.', 'info');
    showToast('File STOCK LIST telah dibersihkan.');
  };

  // UPLOAD 1: Handle Balist XLSX File
  const handleBalistFileUpload = async (file: File) => {
    setUploadProgress({
      isOpen: true,
      fileName: file.name,
      fileSize: file.size,
      uploadType: 'balist',
      step: 1,
      stepTitle: 'Membaca File Data Balistshopee...',
      stepDescription: 'Mengekstrak workbook & membaca format template Shopee...',
      progressPercent: 20,
    });

    try {
      await sleep(200);
      setBalistUploadedFile(file);
      const parsed = await parseGenericXlsx(file, config.balistSheetName);
      setParsedBalistXlsx(parsed);

      setUploadProgress((prev) => ({
        ...prev,
        step: 2,
        stepTitle: 'Mendeteksi Kolom Stok & Identitas Toko...',
        stepDescription: `Menganalisis ${parsed.rows.length.toLocaleString('id-ID')} baris data & memeriksa kode toko...`,
        progressPercent: 55,
      }));

      await sleep(200);

      // Auto-detect default start row: if >= 7 rows, default to row 7; otherwise row 2
      const defaultStartRow = parsed.rows.length >= 7 ? 7 : 2;
      setBalistSourceStartRow(defaultStartRow);

      // Auto-detect stock column index
      const detected = detectBalistStockColumn(parsed.rows);
      setBalistStockColIndex(detected.stockColIndex);

      // Auto-detect store from filename (31475604 -> balist, 56977507 -> Gomall)
      const detectedStore = detectStoreFromFilename(file.name);

      setUploadProgress((prev) => ({
        ...prev,
        step: 3,
        stepTitle: 'Sinkronisasi dengan Database STOCK LIST...',
        stepDescription: 'Mencocokkan SKU variasi produk & menghitung perubahan stok...',
        progressPercent: 85,
      }));

      await sleep(200);

      // Instantly parse into balistList starting at defaultStartRow
      const items = parseBalistShopeeFromRows(parsed.rows, defaultStartRow);
      if (items.length > 0) {
        setBalistList(items);
      }

      setUploadProgress((prev) => ({
        ...prev,
        step: 4,
        stepTitle: 'Data Berhasil Diproses!',
        stepDescription: `${parsed.rows.length.toLocaleString('id-ID')} baris siap dibandingkan.`,
        progressPercent: 100,
      }));

      await sleep(350);

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
    } finally {
      setUploadProgress((prev) => ({ ...prev, isOpen: false }));
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

    setUploadProgress({
      isOpen: true,
      fileName: `${balistUploadedFile.name} [Sheet: ${sheetName}]`,
      fileSize: balistUploadedFile.size,
      uploadType: 'balist',
      step: 2,
      stepTitle: `Membaca Sheet "${sheetName}"...`,
      stepDescription: 'Mengekstrak baris & konfigurasi kolom...',
      progressPercent: 50,
    });

    try {
      await sleep(150);
      const parsed = await parseGenericXlsx(balistUploadedFile, sheetName);
      setParsedBalistXlsx(parsed);
      const detected = detectBalistStockColumn(parsed.rows);
      setBalistStockColIndex(detected.stockColIndex);
      const items = parseBalistShopeeFromRows(parsed.rows, balistSourceStartRow);
      if (items.length > 0) {
        setBalistList(items);
      }

      setUploadProgress((prev) => ({
        ...prev,
        step: 4,
        stepTitle: 'Sheet Berhasil Dimuat!',
        stepDescription: `Menampilkan data dari sheet "${sheetName}".`,
        progressPercent: 100,
      }));
      await sleep(250);

      addLog(
        'upload',
        'Ganti Sheet Balist',
        `Sheet "${sheetName}" dipilih dari file "${balistUploadedFile.name}".`,
        'info',
        { target: sheetName }
      );
    } catch (err) {
      console.error(err);
    } finally {
      setUploadProgress((prev) => ({ ...prev, isOpen: false }));
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
      if (isAuthError(err)) {
        handleAuthExpiry();
        setAccessToken(null);
        setSheetsError('Sesi akun Google Anda telah berakhir. Silakan Masuk Kembali dengan Google.');
        showToast('Sesi Google kedaluwarsa. Membuka login Google...', 'error');
        handleSignIn();
        return;
      }
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
      if (isAuthError(err)) {
        handleAuthExpiry();
        setAccessToken(null);
        setSheetsError('Sesi akun Google Anda telah berakhir. Silakan Masuk Kembali dengan Google.');
        showToast('Sesi Google kedaluwarsa. Membuka login Google...', 'error');
        handleSignIn();
        return;
      }
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
      if (isAuthError(err)) {
        handleAuthExpiry();
        setAccessToken(null);
        setSheetsError('Sesi akun Google Anda telah berakhir. Silakan Masuk Kembali dengan Google.');
        showToast('Sesi Google kedaluwarsa. Membuka login Google...', 'error');
        handleSignIn();
        return;
      }
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

  // Open modal popup for filtering download by Category, Brand, Custom checklist, or All
  const handleOpenDownloadModal = (prefix: string = 'balist') => {
    if (!parsedBalistXlsx && balistList.length === 0) {
      showToast('Unggah file Excel Balistshopee atau sinkronkan Google Sheets terlebih dahulu.', 'error');
      return;
    }
    setModalDownloadPrefix(prefix);
    setIsShopeeDownloadModalOpen(true);
  };

  // Execute download based on filter options selected by user in the modal
  const handleExecuteFilteredDownload = async (options: ExportFilterOptions) => {
    const prefix = options.storePrefix || modalDownloadPrefix || 'balist';

    try {
      const sourceItems = modalComparisonItems;
      if (!sourceItems || sourceItems.length === 0) {
        showToast('Tidak ada data produk yang dapat diunduh. Unggah file Excel atau sinkronkan data terlebih dahulu.', 'error');
        return;
      }

      // Filter items exactly according to the modal selection
      let filteredItems: BalistComparisonItem[] = [];
      if (options.mode === 'all') {
        filteredItems = [...sourceItems];
      } else if (options.mode === 'category') {
        const catSet = new Set(options.selectedCategories);
        filteredItems = sourceItems.filter((item) => {
          const cat = item.matchedStockItem?.category?.trim() || '(Tanpa Kategori)';
          return catSet.has(cat);
        });
      } else if (options.mode === 'brand') {
        const brandSet = new Set(options.selectedBrands);
        filteredItems = sourceItems.filter((item) => {
          const brand = item.matchedStockItem?.brand?.trim() || '(Tanpa Merk)';
          return brandSet.has(brand);
        });
      } else if (options.mode === 'custom') {
        const selectedIndicesSet = new Set(options.selectedRowIndices);
        filteredItems = sourceItems.filter((item) => selectedIndicesSet.has(item.rowIndex));
      }

      if (filteredItems.length === 0) {
        showToast('Tidak ada data produk yang cocok dengan pilihan filter yang dipilih.', 'error');
        return;
      }

      // Update target stock column with latest stock from Stock List if enabled
      const targetCol = balistStockColIndex !== undefined ? balistStockColIndex : 6;
      const dataRows = filteredItems.map((item) => {
        const row = [...item.rawRow];
        while (row.length <= targetCol) {
          row.push('');
        }
        if (updateStockFromStockList) {
          if (item.matchStatus === 'matched' && item.stockQty !== null) {
            row[targetCol] = item.stockQty;
          } else if (balistUnmatchedStockAction === 'zero') {
            row[targetCol] = 0;
          }
        }
        return row;
      });

      // Prepare 6 header rows (preserve from uploaded excel or use standard Shopee Mass Update template)
      let headerRows: any[][] = [];
      if (parsedBalistXlsx && parsedBalistXlsx.rows.length >= 6) {
        headerRows = parsedBalistXlsx.rows.slice(0, 6).map((r) => [...r]);
      } else {
        headerRows = [
          ['Pusat Edukasi Penjual > Pelajari Lebih Lanjut Tentang Update Massal Informasi Penjualan', '', '', '', '', '', '', '', '', ''],
          ['Kategori', 'Informasi Penjualan', '', '', '', '', '', '', '', ''],
          ['Catatan: 1. Jangan ubah format baris header (Baris 1-6) | 2. Jangan ubah data pada kolom bertanda bintang (*) | 3. Pastikan format file tetap .xlsx', '', '', '', '', '', '', '', '', ''],
          ['Kode Produk', 'Nama Produk', 'No. Integrasi Produk', 'Kode Variasi', 'Nama Variasi', 'Kode Integrasi', 'Stok', 'Harga', 'Status Produk', 'SKU Induk'],
          ['Wajib', 'Hanya baca', 'Hanya baca', 'Wajib', 'Hanya baca', 'Opsional', 'Wajib', 'Opsional', 'Hanya baca', 'Opsional'],
          ['Contoh: 12345678', 'Contoh: Produk A', 'Contoh: P001', 'Contoh: 87654321', 'Contoh: Standar', 'Contoh: SKU001', 'Contoh: 100', 'Contoh: 50000', 'Contoh: Aktif', 'Contoh: SKU000'],
        ];
      }

      const rowsToExport = [...headerRows, ...dataRows];
      const balistFileName = generateShopeeBalistFilename(prefix);
      const targetSheetName = parsedBalistXlsx?.selectedSheetName || config.balistSheetName || 'Template';

      downloadBalistRowsAsXlsx(rowsToExport, balistFileName, {
        sheetName: targetSheetName,
      });

      const filterDesc =
        options.mode === 'category'
          ? `Kategori (${options.selectedCategories.join(', ')})`
          : options.mode === 'brand'
          ? `Merk (${options.selectedBrands.join(', ')})`
          : options.mode === 'custom'
          ? `Pilihan Mandiri (${filteredItems.length} produk)`
          : 'Semua Produk';

      addLog(
        'download',
        `Unduh Excel Shopee ${prefix} (${filterDesc})`,
        `File "${balistFileName}" (${filteredItems.length} produk) berhasil diunduh dengan filter yang dipilih.`,
        'success',
        { rowCount: filteredItems.length, target: balistFileName }
      );

      showToast(`File "${balistFileName}" (${filteredItems.length} produk) berhasil diunduh!`);
    } catch (err: any) {
      console.error('Download error:', err);
      const msg = err?.message || 'Gagal membuat file unduhan Excel';
      showToast(msg, 'error');
      addLog('download', 'Gagal Unduh File', msg, 'error', { errorMessage: msg });
    }
  };

  // Direct handlers (legacy aliases to open modal)
  const handleDownloadUpdatedBalistXlsx = (prefix: string = 'balist') => {
    handleOpenDownloadModal(prefix);
  };

  const handleDownloadBalistFromSheetsXlsx = async (prefix: string = 'balist') => {
    handleOpenDownloadModal(prefix);
  };

  // UPLOAD 2: Handle Shopee XLSX Mass Update File
  const handleShopeeFileUpload = async (file: File) => {
    setUploadProgress({
      isOpen: true,
      fileName: file.name,
      fileSize: file.size,
      uploadType: 'shopee',
      step: 1,
      stepTitle: 'Membaca File Mass Update Shopee...',
      stepDescription: 'Mengekstrak worksheet & lembar data seller centre...',
      progressPercent: 20,
    });

    try {
      await sleep(200);
      setShopeeUploadedFile(file);
      const parsed = await parseShopeeXlsx(file);
      setParsedShopeeSheet(parsed);
      setSelectedSkuCol(parsed.skuColIndex);
      setSelectedStockCol(parsed.stockColIndex);

      setUploadProgress((prev) => ({
        ...prev,
        step: 2,
        stepTitle: 'Mendeteksi Kolom SKU & Kolom Stok...',
        stepDescription: `Menemukan ${parsed.rows.length.toLocaleString('id-ID')} baris data pada file Shopee...`,
        progressPercent: 55,
      }));

      await sleep(200);

      // Auto-detect store from filename (e.g. 31475604 -> balist, 56977507 -> Gomall)
      const detectedStore = detectStoreFromFilename(file.name);

      setUploadProgress((prev) => ({
        ...prev,
        step: 3,
        stepTitle: 'Mencocokkan SKU dengan STOCK LIST...',
        stepDescription: 'Menghitung perbedaan stok & menyusun perbandingan produk...',
        progressPercent: 85,
      }));

      await sleep(200);

      const storeDetectionMsg = detectedStore.storeName
        ? ` [Toko Terdeteksi: ${detectedStore.storeName} (${detectedStore.storeCode || detectedStore.storePrefix})] ➔ Format Unduh Otomatis: ${detectedStore.storePrefix}`
        : '';

      setUploadProgress((prev) => ({
        ...prev,
        step: 4,
        stepTitle: 'Pencocokan Data Selesai!',
        stepDescription: `${parsed.rows.length.toLocaleString('id-ID')} baris siap ditinjau & diunduh.`,
        progressPercent: 100,
      }));

      await sleep(350);

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
    } finally {
      setUploadProgress((prev) => ({ ...prev, isOpen: false }));
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

  // UPLOAD 3: Handle Tokopedia XLSX Mass Update File
  const handleTokopediaFileUpload = async (file: File) => {
    setUploadProgress({
      isOpen: true,
      fileName: file.name,
      fileSize: file.size,
      uploadType: 'tokopedia',
      step: 1,
      stepTitle: 'Membaca File Template Tokopedia...',
      stepDescription: 'Mengekstrak baris & worksheet Seller Center Tokopedia...',
      progressPercent: 25,
    });

    try {
      await sleep(200);
      setTokopediaUploadedFile(file);
      const parsed = await parseTokopediaXlsx(file);
      setParsedTokopediaSheet(parsed);
      setTokopediaSkuCol(parsed.skuColIndex); // default 3 (Kolom 4 / D)
      setTokopediaStockCol(parsed.stockColIndex); // default 8 (Kolom 9 / I)
      setTokopediaStartRow(parsed.dataStartRowIndex + 1); // default 4

      setUploadProgress((prev) => ({
        ...prev,
        step: 2,
        stepTitle: 'Mendeteksi Kolom SKU (Kolom 4) & Stok (Kolom 9)...',
        stepDescription: `Menemukan ${parsed.rows.length.toLocaleString('id-ID')} baris data Tokopedia...`,
        progressPercent: 60,
      }));

      await sleep(200);

      setUploadProgress((prev) => ({
        ...prev,
        step: 3,
        stepTitle: 'Mencocokkan SKU Tokopedia dengan STOCK LIST...',
        stepDescription: 'Menghitung perbedaan stok gudang & status kecocokan...',
        progressPercent: 90,
      }));

      await sleep(200);

      const dataRowCount = Math.max(0, parsed.rows.length - (parsed.dataStartRowIndex));

      setUploadProgress((prev) => ({
        ...prev,
        step: 4,
        stepTitle: 'Pencocokan Tokopedia Selesai!',
        stepDescription: `${dataRowCount.toLocaleString('id-ID')} produk siap ditinjau & diunduh.`,
        progressPercent: 100,
      }));

      await sleep(300);

      addLog(
        'upload',
        'Upload File Tokopedia Mass Update',
        `File "${file.name}" (${parsed.rows.length.toLocaleString('id-ID')} baris) berhasil dibaca. Kolom SKU: Kolom ${
          parsed.skuColIndex + 1
        } (D), Kolom Stok: Kolom ${parsed.stockColIndex + 1} (I), Mulai Baris: ${parsed.dataStartRowIndex + 1}.`,
        'success',
        { rowCount: dataRowCount, target: file.name }
      );
      showToast(`File Tokopedia "${file.name}" (${dataRowCount} produk) berhasil dicocokkan!`);
    } catch (err: any) {
      console.error('Tokopedia parsing error:', err);
      const errMsg = err?.message || 'Gagal membaca file Excel Tokopedia';
      addLog('upload', 'Gagal Baca File Tokopedia', errMsg, 'error', {
        errorMessage: errMsg,
        target: file.name,
      });
      showToast(errMsg, 'error');
    } finally {
      setUploadProgress((prev) => ({ ...prev, isOpen: false }));
    }
  };

  const handleClearTokopediaFile = () => {
    setTokopediaUploadedFile(null);
    setParsedTokopediaSheet(null);
    setTokopediaMatches([]);
    addLog('upload', 'File Tokopedia Dikosongkan', 'File unggahan Tokopedia telah dibersihkan.', 'info');
  };

  // Re-run matching whenever parsed Tokopedia sheet, maps, column choices, or action changes
  useEffect(() => {
    if (!parsedTokopediaSheet) {
      setTokopediaMatches([]);
      return;
    }

    const calculated = matchTokopediaFile(parsedTokopediaSheet, matchingMaps, {
      customSkuCol: tokopediaSkuCol,
      customStockCol: tokopediaStockCol,
      customStartRow: tokopediaStartRow,
      unmatchedAction: tokopediaUnmatchedAction,
    });

    setTokopediaMatches(calculated);
  }, [
    parsedTokopediaSheet,
    matchingMaps,
    tokopediaSkuCol,
    tokopediaStockCol,
    tokopediaStartRow,
    tokopediaUnmatchedAction,
  ]);

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
        {/* Marketplace Platform Selector Navigation Bar */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-2 border-b border-stone-200">
          <div className="flex items-center gap-2 bg-stone-100 p-1 rounded-xl border border-stone-200">
            {/* Tab 1: Shopee */}
            <button
              type="button"
              id="tab-platform-shopee"
              onClick={() => {
                setActivePlatform('shopee');
                localStorage.setItem('active_marketplace_platform', 'shopee');
              }}
              className={`inline-flex items-center gap-2 px-3.5 py-2 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                activePlatform === 'shopee'
                  ? 'bg-orange-600 text-white shadow-xs'
                  : 'text-stone-700 hover:text-stone-900 hover:bg-stone-200/70'
              }`}
            >
              <ShoppingBag className="w-4 h-4" />
              <span>Shopee &amp; Balistshopee</span>
              <span
                className={`px-1.5 py-0.2 rounded text-[10px] font-bold ${
                  activePlatform === 'shopee' ? 'bg-orange-700 text-white' : 'bg-orange-100 text-orange-800'
                }`}
              >
                Kolom 5 &amp; 6
              </span>
            </button>

            {/* Tab 2: Tokopedia */}
            <button
              type="button"
              id="tab-platform-tokopedia"
              onClick={() => {
                setActivePlatform('tokopedia');
                localStorage.setItem('active_marketplace_platform', 'tokopedia');
              }}
              className={`inline-flex items-center gap-2 px-3.5 py-2 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                activePlatform === 'tokopedia'
                  ? 'bg-emerald-600 text-white shadow-xs'
                  : 'text-stone-700 hover:text-stone-900 hover:bg-stone-200/70'
              }`}
            >
              <Package className="w-4 h-4" />
              <span>Tokopedia Mass Update</span>
              <span
                className={`px-1.5 py-0.2 rounded text-[10px] font-bold ${
                  activePlatform === 'tokopedia' ? 'bg-emerald-800 text-white' : 'bg-emerald-100 text-emerald-800'
                }`}
              >
                Kolom 4 &amp; 9
              </span>
            </button>

            {/* Tab 3: Database STOCK LIST */}
            <button
              type="button"
              id="tab-platform-stocklist"
              onClick={() => {
                setActivePlatform('stocklist');
                localStorage.setItem('active_marketplace_platform', 'stocklist');
              }}
              className={`inline-flex items-center gap-2 px-3.5 py-2 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                activePlatform === 'stocklist'
                  ? 'bg-stone-800 text-white shadow-xs'
                  : 'text-stone-700 hover:text-stone-900 hover:bg-stone-200/70'
              }`}
            >
              <Database className="w-4 h-4" />
              <span>STOCK LIST Gudang</span>
              <span
                className={`px-1.5 py-0.2 rounded text-[10px] font-bold ${
                  activePlatform === 'stocklist' ? 'bg-stone-700 text-stone-200' : 'bg-stone-200 text-stone-700'
                }`}
              >
                {stockList.length > 0 ? stockList.length.toLocaleString('id-ID') : '0'}
              </span>
            </button>
          </div>

          <div className="flex items-center gap-2 text-xs text-stone-500">
            <span className="hidden md:inline font-medium">Database Terhubung:</span>
            <span className="font-bold text-stone-700 bg-stone-100 px-2.5 py-1 rounded-md border border-stone-200">
              {config.stockSheetName} ({stockList.length.toLocaleString('id-ID')} items)
            </span>
          </div>
        </div>

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
            onRefreshStockList={handleRefreshStockList}
            isRefreshingStockList={isRefreshingStockList}
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
                <>
                  <button
                    type="button"
                    onClick={handleRefreshStockList}
                    disabled={isRefreshingStockList || isLoadingSheets}
                    className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-300 rounded-lg text-xs font-semibold transition-colors disabled:opacity-50"
                    title="Segarkan data terbaru khusus sheet STOCK LIST"
                  >
                    <RefreshCw className={`w-3 h-3 ${isRefreshingStockList ? 'animate-spin' : ''}`} />
                    <span>{isRefreshingStockList ? 'Memperbarui...' : 'Refresh STOCK LIST'}</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => loadSheets()}
                    disabled={isLoadingSheets}
                    className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-white hover:bg-stone-100 text-stone-700 border border-stone-200 rounded-lg text-xs font-medium transition-colors"
                  >
                    <RefreshCw className={`w-3 h-3 ${isLoadingSheets ? 'animate-spin' : ''}`} />
                    <span>Muat Ulang Semua</span>
                  </button>
                </>
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
            </div>
          </div>
        )}

        {/* ========================================================================= */}
        {/* PLATFORM 1: SHOPEE & BALISTSHOPEE WORKFLOW                                */}
        {/* ========================================================================= */}
        {activePlatform === 'shopee' && (
          <div className="space-y-6">
            {/* SECTION 1: DATABASE STOCK LIST & PEMBARUAN BALISTSHOPEE */}
            {(activeTab === 'workflow' || activeTab === 'balist_comparison') && (
              <div className="space-y-4">
                {/* Step 1: Database STOCK LIST (Upload XLSX atau Google Sheets) */}
                <StockListUploadCard
                  stockList={stockList}
                  uploadedStockFileName={uploadedStockFileName}
                  uploadedStockFileSize={uploadedStockFileSize}
                  dataSource={stockDataSource}
                  lastUpdated={lastLoaded}
                  onFileUpload={handleStockListFileUpload}
                  onClearFile={handleClearStockListFile}
                  onRefreshGoogleSheets={handleRefreshStockList}
                  isLoadingSheets={isRefreshingStockList || isLoadingSheets}
                  isAuthenticated={!!user}
                  stockSheetName={config.stockSheetName}
                />

                {/* Step 2: Pembaruan Data Balistshopee / Shopee */}
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
                  onRefreshStockList={handleRefreshStockList}
                  isRefreshingStockList={isRefreshingStockList}
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
                    isAuthenticated={!!user}
                    onPromptSignIn={handleSignIn}
                    onRefreshStockList={handleRefreshStockList}
                    isRefreshingStockList={isRefreshingStockList}
                    onDownloadUpdatedBalistXlsx={
                      parsedBalistXlsx
                        ? handleDownloadUpdatedBalistXlsx
                        : handleDownloadBalistFromSheetsXlsx
                    }
                    stockColIndex={balistStockColIndex}
                  />
                )}
              </div>
            )}

            {/* SECTION 2: TOMBOL UPLOAD 2 (UPLOAD SHOPEE MASS UPDATE & UPDATE STOK) */}
            {(activeTab === 'shopee_match' || (activeTab === 'workflow' && showShopeeSection)) ? (
              <div className="space-y-4">
                {activeTab === 'shopee_match' && (
                  <StockListUploadCard
                    stockList={stockList}
                    uploadedStockFileName={uploadedStockFileName}
                    uploadedStockFileSize={uploadedStockFileSize}
                    dataSource={stockDataSource}
                    lastUpdated={lastLoaded}
                    onFileUpload={handleStockListFileUpload}
                    onClearFile={handleClearStockListFile}
                    onRefreshGoogleSheets={handleRefreshStockList}
                    isLoadingSheets={isRefreshingStockList || isLoadingSheets}
                    isAuthenticated={!!user}
                    stockSheetName={config.stockSheetName}
                  />
                )}

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
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-white hover:bg-stone-100 text-stone-700 border border-stone-300 rounded-lg font-medium shadow-2xs transition-colors cursor-pointer"
                >
                  <Eye className="w-3.5 h-3.5 text-stone-500" />
                  <span>Tampilkan Menu Shopee</span>
                </button>
              </div>
            ) : null}
          </div>
        )}

        {/* ========================================================================= */}
        {/* PLATFORM 2: TOKOPEDIA MASS UPDATE (SKU KOLOM 4, STOK KOLOM 9)             */}
        {/* ========================================================================= */}
        {activePlatform === 'tokopedia' && (
          <div className="space-y-6">
            {/* Step 1: Database STOCK LIST Check */}
            <StockListUploadCard
              stockList={stockList}
              uploadedStockFileName={uploadedStockFileName}
              uploadedStockFileSize={uploadedStockFileSize}
              dataSource={stockDataSource}
              lastUpdated={lastLoaded}
              onFileUpload={handleStockListFileUpload}
              onClearFile={handleClearStockListFile}
              onRefreshGoogleSheets={handleRefreshStockList}
              isLoadingSheets={isRefreshingStockList || isLoadingSheets}
              isAuthenticated={!!user}
              stockSheetName={config.stockSheetName}
            />

            {/* Step 2: Tokopedia Stock Upload & Sync Table */}
            <TokopediaStockSection
              parsedFile={parsedTokopediaSheet}
              uploadedFileName={tokopediaUploadedFile?.name || parsedTokopediaSheet?.fileName}
              matches={tokopediaMatches}
              stockCount={stockList.length}
              stockSheetName={config.stockSheetName}
              selectedSkuCol={tokopediaSkuCol}
              selectedStockCol={tokopediaStockCol}
              dataStartRow={tokopediaStartRow}
              unmatchedAction={tokopediaUnmatchedAction}
              onFileUpload={handleTokopediaFileUpload}
              onClearFile={handleClearTokopediaFile}
              onChangeSkuCol={setTokopediaSkuCol}
              onChangeStockCol={setTokopediaStockCol}
              onChangeDataStartRow={setTokopediaStartRow}
              onChangeUnmatchedAction={setTokopediaUnmatchedAction}
              onRefreshStockList={handleRefreshStockList}
              isRefreshingStockList={isRefreshingStockList}
            />
          </div>
        )}

        {/* ========================================================================= */}
        {/* PLATFORM 3: DATABASE STOCK LIST & ALL-IN-ONE CONNECTION DETAILS           */}
        {/* ========================================================================= */}
        {activePlatform === 'stocklist' && (
          <div className="space-y-6">
            <StockListUploadCard
              stockList={stockList}
              uploadedStockFileName={uploadedStockFileName}
              uploadedStockFileSize={uploadedStockFileSize}
              dataSource={stockDataSource}
              lastUpdated={lastLoaded}
              onFileUpload={handleStockListFileUpload}
              onClearFile={handleClearStockListFile}
              onRefreshGoogleSheets={handleRefreshStockList}
              isLoadingSheets={isRefreshingStockList || isLoadingSheets}
              isAuthenticated={!!user}
              stockSheetName={config.stockSheetName}
            />

            {/* Raw Stock List Preview */}
            <div className="bg-white rounded-xl border border-stone-200 shadow-xs p-5 space-y-4">
              <div className="flex items-center justify-between pb-3 border-b border-stone-100">
                <div className="flex items-center gap-2">
                  <Database className="w-5 h-5 text-emerald-600" />
                  <h3 className="text-sm font-bold text-stone-900">
                    Daftar Inventaris Gudang ({config.stockSheetName})
                  </h3>
                </div>
                <span className="text-xs font-semibold text-stone-600 bg-stone-100 px-2.5 py-1 rounded-md">
                  Total {stockList.length.toLocaleString('id-ID')} SKU Terdaftar
                </span>
              </div>

              {stockList.length === 0 ? (
                <p className="text-xs text-stone-500 py-6 text-center">
                  Belum ada data inventaris. Silakan hubungkan Google Sheets atau upload file Excel STOCK LIST di atas.
                </p>
              ) : (
                <div className="max-h-96 overflow-y-auto overflow-x-auto border border-stone-100 rounded-lg">
                  <table className="w-full text-left text-xs text-stone-700">
                    <thead className="bg-stone-50 text-stone-600 uppercase font-semibold text-[10px] sticky top-0 border-b border-stone-200">
                      <tr>
                        <th className="py-2.5 px-3">No</th>
                        <th className="py-2.5 px-3">Kode / SKU (Kol 1)</th>
                        <th className="py-2.5 px-3">Barcode (Kol 2)</th>
                        <th className="py-2.5 px-3">Nama Barang (Kol 3)</th>
                        <th className="py-2.5 px-3">Merk / Brand</th>
                        <th className="py-2.5 px-3">Kategori</th>
                        <th className="py-2.5 px-3 text-center">Stok Qty (Kol 15)</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-stone-100">
                      {stockList.slice(0, 100).map((item, idx) => (
                        <tr key={idx} className="hover:bg-stone-50/60">
                          <td className="py-2 px-3 font-mono text-stone-400 text-[11px]">{idx + 1}</td>
                          <td className="py-2 px-3 font-mono font-bold text-stone-900">{item.code}</td>
                          <td className="py-2 px-3 font-mono text-stone-500">{item.barcode || '-'}</td>
                          <td className="py-2 px-3 font-medium text-stone-800">{item.description || '-'}</td>
                          <td className="py-2 px-3 text-stone-600">{item.brand || '-'}</td>
                          <td className="py-2 px-3 text-stone-600">{item.category || '-'}</td>
                          <td className="py-2 px-3 text-center">
                            <span
                              className={`font-mono font-bold px-2 py-0.5 rounded-full text-xs ${
                                item.qty > 0
                                  ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                                  : 'bg-rose-50 text-rose-800 border border-rose-200'
                              }`}
                            >
                              {item.qty.toLocaleString('id-ID')}
                            </span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        )}
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

      {/* Shopee Download Filter Modal: Download by Category, Brand, Custom Checklist or All */}
      <ShopeeDownloadModal
        isOpen={isShopeeDownloadModalOpen}
        onClose={() => setIsShopeeDownloadModalOpen(false)}
        onConfirmDownload={handleExecuteFilteredDownload}
        comparisonItems={modalComparisonItems}
        initialPrefix={modalDownloadPrefix}
        stockSheetName={config.stockSheetName}
      />

      {/* File Upload & Processing Loading Popup Modal */}
      <UploadLoadingModal progress={uploadProgress} />

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
