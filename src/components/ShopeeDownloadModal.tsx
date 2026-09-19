import React, { useState, useMemo, useEffect } from 'react';
import {
  X,
  Download,
  Filter,
  CheckSquare,
  Square,
  Layers,
  Tag,
  Briefcase,
  Search,
  CheckCircle2,
  FileSpreadsheet,
  Store,
  Sliders,
  AlertCircle,
  Sparkles,
  Info,
} from 'lucide-react';
import {
  BalistComparisonItem,
  ExportFilterMode,
  ExportFilterOptions,
  StockListItem,
} from '../types';
import { generateShopeeBalistFilename } from '../lib/excelProcessor';

interface ShopeeDownloadModalProps {
  isOpen: boolean;
  onClose: () => void;
  onConfirmDownload: (options: ExportFilterOptions) => void;
  comparisonItems: BalistComparisonItem[];
  stockList?: StockListItem[];
  initialPrefix?: string;
  defaultPrefix?: string;
  stockSheetName?: string;
  isProcessing?: boolean;
}

export const ShopeeDownloadModal: React.FC<ShopeeDownloadModalProps> = ({
  isOpen,
  onClose,
  onConfirmDownload,
  comparisonItems,
  stockList,
  initialPrefix,
  defaultPrefix = 'balist',
  stockSheetName = 'STOCK LIST',
  isProcessing = false,
}) => {
  const [mode, setMode] = useState<ExportFilterMode>('all');
  const [selectedCategories, setSelectedCategories] = useState<string[]>([]);
  const [selectedBrands, setSelectedBrands] = useState<string[]>([]);
  const [selectedRowIndices, setSelectedRowIndices] = useState<Set<number>>(new Set());
  const effectivePrefix = initialPrefix || defaultPrefix;
  const [customPrefix, setCustomPrefix] = useState<string>(effectivePrefix);

  // Search queries within modal tabs
  const [categorySearch, setCategorySearch] = useState('');
  const [brandSearch, setBrandSearch] = useState('');
  const [itemSearch, setItemSearch] = useState('');
  const [itemStockFilter, setItemStockFilter] = useState<'all' | 'in_stock' | 'out_of_stock'>('all');

  // Sync defaultPrefix when modal opens
  useEffect(() => {
    if (isOpen) {
      setCustomPrefix(effectivePrefix);
      // Initialize selectedRowIndices with all items
      setSelectedRowIndices(new Set(comparisonItems.map((it) => it.rowIndex)));
    }
  }, [isOpen, effectivePrefix, comparisonItems]);

  // Count items matched via 5-digit analysis
  const fiveDigitMatchedCount = useMemo(() => {
    return comparisonItems.filter((it) => it.matchedStockItem && (it.matchedBy === '5digits_sku' || it.matchedStockItem.code)).length;
  }, [comparisonItems]);

  // 1. Extract unique Categories from items with count
  const categoryStats = useMemo(() => {
    const map = new Map<string, number>();
    for (const item of comparisonItems) {
      const cat = item.matchedStockItem?.category?.trim() || '(Tanpa Kategori)';
      map.set(cat, (map.get(cat) || 0) + 1);
    }
    return Array.from(map.entries())
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));
  }, [comparisonItems]);

  // 2. Extract unique Brands/Merks from items with count
  const brandStats = useMemo(() => {
    const map = new Map<string, number>();
    for (const item of comparisonItems) {
      const brand = item.matchedStockItem?.brand?.trim() || '(Tanpa Merk)';
      map.set(brand, (map.get(brand) || 0) + 1);
    }
    return Array.from(map.entries())
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));
  }, [comparisonItems]);

  // Filtered categories in UI
  const filteredCategories = useMemo(() => {
    if (!categorySearch.trim()) return categoryStats;
    const q = categorySearch.toLowerCase().trim();
    return categoryStats.filter((c) => c.name.toLowerCase().includes(q));
  }, [categoryStats, categorySearch]);

  // Filtered brands in UI
  const filteredBrands = useMemo(() => {
    if (!brandSearch.trim()) return brandStats;
    const q = brandSearch.toLowerCase().trim();
    return brandStats.filter((b) => b.name.toLowerCase().includes(q));
  }, [brandStats, brandSearch]);

  // Filtered individual items for Custom Checklist
  const filteredItems = useMemo(() => {
    let result = comparisonItems;

    if (itemStockFilter === 'in_stock') {
      result = result.filter((it) => (it.stockQty ?? 0) > 0);
    } else if (itemStockFilter === 'out_of_stock') {
      result = result.filter((it) => (it.stockQty ?? 0) <= 0);
    }

    if (itemSearch.trim()) {
      const q = itemSearch.toLowerCase().trim();
      result = result.filter((it) => {
        const sku5 = (it.skuCol5 || '').toLowerCase();
        const sku6 = (it.skuCol6 || '').toLowerCase();
        const name = (it.matchedStockItem?.description || (it.rawRow && it.rawRow[1] ? String(it.rawRow[1]) : '')).toLowerCase();
        const cat = (it.matchedStockItem?.category || '').toLowerCase();
        const brand = (it.matchedStockItem?.brand || '').toLowerCase();
        return (
          sku5.includes(q) ||
          sku6.includes(q) ||
          name.includes(q) ||
          cat.includes(q) ||
          brand.includes(q)
        );
      });
    }

    return result;
  }, [comparisonItems, itemSearch, itemStockFilter]);

  // Calculate total matching items to export based on selected mode
  const itemsToExportCount = useMemo(() => {
    if (mode === 'all') {
      return comparisonItems.length;
    }
    if (mode === 'category') {
      if (selectedCategories.length === 0) return 0;
      const set = new Set(selectedCategories);
      return comparisonItems.filter((it) => {
        const cat = it.matchedStockItem?.category?.trim() || '(Tanpa Kategori)';
        return set.has(cat);
      }).length;
    }
    if (mode === 'brand') {
      if (selectedBrands.length === 0) return 0;
      const set = new Set(selectedBrands);
      return comparisonItems.filter((it) => {
        const brand = it.matchedStockItem?.brand?.trim() || '(Tanpa Merk)';
        return set.has(brand);
      }).length;
    }
    if (mode === 'custom') {
      return selectedRowIndices.size;
    }
    return comparisonItems.length;
  }, [mode, selectedCategories, selectedBrands, selectedRowIndices, comparisonItems]);

  if (!isOpen) return null;

  const handleToggleCategory = (catName: string) => {
    setSelectedCategories((prev) =>
      prev.includes(catName) ? prev.filter((c) => c !== catName) : [...prev, catName]
    );
  };

  const handleSelectAllCategories = () => {
    setSelectedCategories(categoryStats.map((c) => c.name));
  };

  const handleDeselectAllCategories = () => {
    setSelectedCategories([]);
  };

  const handleToggleBrand = (brandName: string) => {
    setSelectedBrands((prev) =>
      prev.includes(brandName) ? prev.filter((b) => b !== brandName) : [...prev, brandName]
    );
  };

  const handleSelectAllBrands = () => {
    setSelectedBrands(brandStats.map((b) => b.name));
  };

  const handleDeselectAllBrands = () => {
    setSelectedBrands([]);
  };

  const handleToggleItem = (rowIndex: number) => {
    setSelectedRowIndices((prev) => {
      const next = new Set(prev);
      if (next.has(rowIndex)) {
        next.delete(rowIndex);
      } else {
        next.add(rowIndex);
      }
      return next;
    });
  };

  const handleSelectAllFilteredItems = () => {
    setSelectedRowIndices((prev) => {
      const next = new Set(prev);
      for (const it of filteredItems) {
        next.add(it.rowIndex);
      }
      return next;
    });
  };

  const handleDeselectAllFilteredItems = () => {
    setSelectedRowIndices((prev) => {
      const next = new Set(prev);
      for (const it of filteredItems) {
        next.delete(it.rowIndex);
      }
      return next;
    });
  };

  const handleConfirm = () => {
    onConfirmDownload({
      mode,
      selectedCategories,
      selectedBrands,
      selectedRowIndices: Array.from(selectedRowIndices),
      storePrefix: customPrefix.trim() || effectivePrefix,
    });
    onClose();
  };

  return (
    <div
      id="modal-shopee-download"
      className="fixed inset-0 z-50 overflow-y-auto bg-stone-900/60 backdrop-blur-xs flex items-center justify-center p-3 sm:p-4"
    >
      <div className="bg-white rounded-2xl shadow-2xl border border-stone-200 w-full max-w-3xl overflow-hidden flex flex-col max-h-[92vh] animate-in fade-in zoom-in-95 duration-150">
        {/* Header */}
        <div className="p-4 sm:p-5 border-b border-stone-200 flex items-center justify-between bg-stone-50/80 shrink-0">
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-xl bg-emerald-700 text-white flex items-center justify-center shadow-xs">
              <Download className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-base sm:text-lg font-bold text-stone-900 leading-tight">
                Pilih Format &amp; Filter Unduhan Shopee
              </h2>
              <p className="text-xs text-stone-500">
                Kategori dan Merk dianalisa otomatis berdasarkan persamaan kode SKU 5 digit dengan sheet {stockSheetName}.
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="text-stone-400 hover:text-stone-600 p-1.5 rounded-lg hover:bg-stone-200/60 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Body */}
        <div className="p-4 sm:p-5 overflow-y-auto space-y-4 flex-1">
          {/* Mode Selection Cards */}
          <div>
            <div className="flex items-center justify-between mb-2">
              <label className="block text-xs font-bold text-stone-700 uppercase tracking-wider">
                Pilih Ruang Lingkup Data (Mode Unduh):
              </label>
              <span className="text-[11px] text-emerald-800 font-medium bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-full flex items-center gap-1">
                <Sparkles className="w-3 h-3 text-emerald-600" />
                {fiveDigitMatchedCount} SKU teranalisa
              </span>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
              {/* Card 1: All */}
              <button
                type="button"
                onClick={() => setMode('all')}
                className={`p-3 rounded-xl border text-left flex flex-col justify-between transition-all cursor-pointer ${
                  mode === 'all'
                    ? 'bg-emerald-50/90 border-emerald-600 ring-2 ring-emerald-500/20 shadow-xs'
                    : 'bg-white border-stone-200 hover:bg-stone-50'
                }`}
              >
                <div className="flex items-center justify-between mb-1.5">
                  <Layers className={`w-4 h-4 ${mode === 'all' ? 'text-emerald-700' : 'text-stone-500'}`} />
                  <span className={`text-[10px] font-bold px-1.5 py-0.2 rounded-full ${mode === 'all' ? 'bg-emerald-200 text-emerald-900' : 'bg-stone-100 text-stone-600'}`}>
                    {comparisonItems.length}
                  </span>
                </div>
                <div>
                  <span className="text-xs font-bold text-stone-900 block">Semua Produk</span>
                  <span className="text-[11px] text-stone-500 block leading-tight">Seluruh baris file</span>
                </div>
              </button>

              {/* Card 2: By Category */}
              <button
                type="button"
                onClick={() => setMode('category')}
                className={`p-3 rounded-xl border text-left flex flex-col justify-between transition-all cursor-pointer ${
                  mode === 'category'
                    ? 'bg-emerald-50/90 border-emerald-600 ring-2 ring-emerald-500/20 shadow-xs'
                    : 'bg-white border-stone-200 hover:bg-stone-50'
                }`}
              >
                <div className="flex items-center justify-between mb-1.5">
                  <Tag className={`w-4 h-4 ${mode === 'category' ? 'text-emerald-700' : 'text-stone-500'}`} />
                  <span className={`text-[10px] font-bold px-1.5 py-0.2 rounded-full ${mode === 'category' ? 'bg-emerald-200 text-emerald-900' : 'bg-stone-100 text-stone-600'}`}>
                    {categoryStats.length} Kat
                  </span>
                </div>
                <div>
                  <span className="text-xs font-bold text-stone-900 block">Kategori</span>
                  <span className="text-[11px] text-stone-500 block leading-tight">Analisa 5 Digit SKU</span>
                </div>
              </button>

              {/* Card 3: By Brand */}
              <button
                type="button"
                onClick={() => setMode('brand')}
                className={`p-3 rounded-xl border text-left flex flex-col justify-between transition-all cursor-pointer ${
                  mode === 'brand'
                    ? 'bg-emerald-50/90 border-emerald-600 ring-2 ring-emerald-500/20 shadow-xs'
                    : 'bg-white border-stone-200 hover:bg-stone-50'
                }`}
              >
                <div className="flex items-center justify-between mb-1.5">
                  <Briefcase className={`w-4 h-4 ${mode === 'brand' ? 'text-emerald-700' : 'text-stone-500'}`} />
                  <span className={`text-[10px] font-bold px-1.5 py-0.2 rounded-full ${mode === 'brand' ? 'bg-emerald-200 text-emerald-900' : 'bg-stone-100 text-stone-600'}`}>
                    {brandStats.length} Merk
                  </span>
                </div>
                <div>
                  <span className="text-xs font-bold text-stone-900 block">Merk / Brand</span>
                  <span className="text-[11px] text-stone-500 block leading-tight">Analisa 5 Digit SKU</span>
                </div>
              </button>

              {/* Card 4: Custom Checklist */}
              <button
                type="button"
                onClick={() => setMode('custom')}
                className={`p-3 rounded-xl border text-left flex flex-col justify-between transition-all cursor-pointer ${
                  mode === 'custom'
                    ? 'bg-emerald-50/90 border-emerald-600 ring-2 ring-emerald-500/20 shadow-xs'
                    : 'bg-white border-stone-200 hover:bg-stone-50'
                }`}
              >
                <div className="flex items-center justify-between mb-1.5">
                  <CheckSquare className={`w-4 h-4 ${mode === 'custom' ? 'text-emerald-700' : 'text-stone-500'}`} />
                  <span className={`text-[10px] font-bold px-1.5 py-0.2 rounded-full ${mode === 'custom' ? 'bg-emerald-200 text-emerald-900' : 'bg-stone-100 text-stone-600'}`}>
                    {selectedRowIndices.size} SKU
                  </span>
                </div>
                <div>
                  <span className="text-xs font-bold text-stone-900 block">Checklist Sendiri</span>
                  <span className="text-[11px] text-stone-500 block leading-tight">Pilih manual per SKU</span>
                </div>
              </button>
            </div>
          </div>

          {/* DYNAMIC CONTENT PER MODE */}

          {/* MODE 1: ALL */}
          {mode === 'all' && (
            <div className="p-4 bg-stone-50 rounded-xl border border-stone-200 text-xs text-stone-700 space-y-2">
              <div className="flex items-center gap-2 text-stone-900 font-bold">
                <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                <span>Seluruh {comparisonItems.length.toLocaleString('id-ID')} produk akan dimasukkan ke dalam file Excel.</span>
              </div>
              <p className="text-stone-500 leading-relaxed">
                Format file mempertahankan 6 baris header resmi Shopee Mass Update dengan kolom Stok yang telah disesuaikan berdasarkan analisa 5 digit kode SKU dengan sheet {stockSheetName}.
              </p>
            </div>
          )}

          {/* MODE 2: CATEGORY FILTER */}
          {mode === 'category' && (
            <div className="space-y-3 p-4 bg-stone-50 rounded-xl border border-stone-200">
              <div className="p-2.5 bg-emerald-50/90 border border-emerald-200 rounded-lg flex items-start gap-2 text-xs text-emerald-900">
                <Sparkles className="w-4 h-4 text-emerald-700 shrink-0 mt-0.5" />
                <div>
                  <span className="font-bold">Analisa Kode SKU 5 Digit: </span>
                  Kategori produk di bawah didapatkan otomatis dengan mencocokkan kode 5 digit pada SKU Shopee dengan sheet <strong>{stockSheetName}</strong>.
                </div>
              </div>

              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                <div className="flex items-center gap-2">
                  <Tag className="w-4 h-4 text-emerald-700" />
                  <span className="text-xs font-bold text-stone-900">
                    Pilih Kategori Produk ({selectedCategories.length} dari {categoryStats.length} terpilih)
                  </span>
                </div>
                <div className="flex items-center gap-1.5">
                  <button
                    type="button"
                    onClick={handleSelectAllCategories}
                    className="text-[11px] font-semibold text-emerald-700 hover:text-emerald-900 bg-white px-2.5 py-1 rounded-md border border-stone-200 hover:bg-stone-100 transition-colors cursor-pointer"
                  >
                    Pilih Semua
                  </button>
                  <button
                    type="button"
                    onClick={handleDeselectAllCategories}
                    className="text-[11px] font-semibold text-stone-600 hover:text-stone-900 bg-white px-2.5 py-1 rounded-md border border-stone-200 hover:bg-stone-100 transition-colors cursor-pointer"
                  >
                    Hapus Pilihan
                  </button>
                </div>
              </div>

              {/* Search Bar for Category */}
              <div className="relative">
                <Search className="w-3.5 h-3.5 text-stone-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  placeholder="Cari nama kategori..."
                  value={categorySearch}
                  onChange={(e) => setCategorySearch(e.target.value)}
                  className="w-full bg-white border border-stone-300 rounded-lg pl-8 pr-3 py-1.5 text-xs text-stone-800 placeholder-stone-400 focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500"
                />
              </div>

              {/* Category Checkbox Grid */}
              <div className="max-h-56 overflow-y-auto grid grid-cols-1 sm:grid-cols-2 gap-1.5 pr-1">
                {filteredCategories.length === 0 ? (
                  <div className="col-span-2 py-4 text-center text-xs text-stone-400">
                    Tidak ada kategori yang cocok dengan pencarian.
                  </div>
                ) : (
                  filteredCategories.map((cat) => {
                    const isChecked = selectedCategories.includes(cat.name);
                    return (
                      <label
                        key={cat.name}
                        onClick={() => handleToggleCategory(cat.name)}
                        className={`flex items-center justify-between p-2 rounded-lg border text-xs cursor-pointer transition-colors ${
                          isChecked
                            ? 'bg-emerald-50/90 border-emerald-400 text-emerald-950 font-semibold'
                            : 'bg-white border-stone-200 hover:bg-stone-100/70 text-stone-700'
                        }`}
                      >
                        <div className="flex items-center gap-2 truncate pr-2">
                          <input
                            type="checkbox"
                            checked={isChecked}
                            onChange={() => {}}
                            className="rounded text-emerald-600 focus:ring-emerald-500 shrink-0"
                          />
                          <span className="truncate">{cat.name}</span>
                        </div>
                        <span className="text-[10px] font-mono font-medium px-1.5 py-0.2 rounded bg-stone-100 text-stone-600 shrink-0">
                          {cat.count} item
                        </span>
                      </label>
                    );
                  })
                )}
              </div>
            </div>
          )}

          {/* MODE 3: BRAND FILTER */}
          {mode === 'brand' && (
            <div className="space-y-3 p-4 bg-stone-50 rounded-xl border border-stone-200">
              <div className="p-2.5 bg-emerald-50/90 border border-emerald-200 rounded-lg flex items-start gap-2 text-xs text-emerald-900">
                <Sparkles className="w-4 h-4 text-emerald-700 shrink-0 mt-0.5" />
                <div>
                  <span className="font-bold">Analisa Kode SKU 5 Digit: </span>
                  Merk/Brand produk di bawah didapatkan otomatis dengan mencocokkan kode 5 digit pada SKU Shopee dengan sheet <strong>{stockSheetName}</strong>.
                </div>
              </div>

              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                <div className="flex items-center gap-2">
                  <Briefcase className="w-4 h-4 text-emerald-700" />
                  <span className="text-xs font-bold text-stone-900">
                    Pilih Merk / Brand Produk ({selectedBrands.length} dari {brandStats.length} terpilih)
                  </span>
                </div>
                <div className="flex items-center gap-1.5">
                  <button
                    type="button"
                    onClick={handleSelectAllBrands}
                    className="text-[11px] font-semibold text-emerald-700 hover:text-emerald-900 bg-white px-2.5 py-1 rounded-md border border-stone-200 hover:bg-stone-100 transition-colors cursor-pointer"
                  >
                    Pilih Semua
                  </button>
                  <button
                    type="button"
                    onClick={handleDeselectAllBrands}
                    className="text-[11px] font-semibold text-stone-600 hover:text-stone-900 bg-white px-2.5 py-1 rounded-md border border-stone-200 hover:bg-stone-100 transition-colors cursor-pointer"
                  >
                    Hapus Pilihan
                  </button>
                </div>
              </div>

              {/* Search Bar for Brand */}
              <div className="relative">
                <Search className="w-3.5 h-3.5 text-stone-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  placeholder="Cari nama merk / brand..."
                  value={brandSearch}
                  onChange={(e) => setBrandSearch(e.target.value)}
                  className="w-full bg-white border border-stone-300 rounded-lg pl-8 pr-3 py-1.5 text-xs text-stone-800 placeholder-stone-400 focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500"
                />
              </div>

              {/* Brand Checkbox Grid */}
              <div className="max-h-56 overflow-y-auto grid grid-cols-1 sm:grid-cols-2 gap-1.5 pr-1">
                {filteredBrands.length === 0 ? (
                  <div className="col-span-2 py-4 text-center text-xs text-stone-400">
                    Tidak ada merk yang cocok dengan pencarian.
                  </div>
                ) : (
                  filteredBrands.map((b) => {
                    const isChecked = selectedBrands.includes(b.name);
                    return (
                      <label
                        key={b.name}
                        onClick={() => handleToggleBrand(b.name)}
                        className={`flex items-center justify-between p-2 rounded-lg border text-xs cursor-pointer transition-colors ${
                          isChecked
                            ? 'bg-emerald-50/90 border-emerald-400 text-emerald-950 font-semibold'
                            : 'bg-white border-stone-200 hover:bg-stone-100/70 text-stone-700'
                        }`}
                      >
                        <div className="flex items-center gap-2 truncate pr-2">
                          <input
                            type="checkbox"
                            checked={isChecked}
                            onChange={() => {}}
                            className="rounded text-emerald-600 focus:ring-emerald-500 shrink-0"
                          />
                          <span className="truncate">{b.name}</span>
                        </div>
                        <span className="text-[10px] font-mono font-medium px-1.5 py-0.2 rounded bg-stone-100 text-stone-600 shrink-0">
                          {b.count} item
                        </span>
                      </label>
                    );
                  })
                )}
              </div>
            </div>
          )}

          {/* MODE 4: CUSTOM CHECKLIST */}
          {mode === 'custom' && (
            <div className="space-y-3 p-4 bg-stone-50 rounded-xl border border-stone-200">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                <div className="flex items-center gap-2">
                  <CheckSquare className="w-4 h-4 text-emerald-700" />
                  <span className="text-xs font-bold text-stone-900">
                    Checklist Mandiri ({selectedRowIndices.size} dari {comparisonItems.length} SKU terpilih)
                  </span>
                </div>
                <div className="flex items-center gap-1.5">
                  <button
                    type="button"
                    onClick={handleSelectAllFilteredItems}
                    className="text-[11px] font-semibold text-emerald-700 hover:text-emerald-900 bg-white px-2.5 py-1 rounded-md border border-stone-200 hover:bg-stone-100 transition-colors cursor-pointer"
                  >
                    Pilih Yang Tampil ({filteredItems.length})
                  </button>
                  <button
                    type="button"
                    onClick={handleDeselectAllFilteredItems}
                    className="text-[11px] font-semibold text-stone-600 hover:text-stone-900 bg-white px-2.5 py-1 rounded-md border border-stone-200 hover:bg-stone-100 transition-colors cursor-pointer"
                  >
                    Batal Pilih
                  </button>
                </div>
              </div>

              {/* Search & Stock Filter */}
              <div className="flex flex-col sm:flex-row gap-2">
                <div className="relative flex-1">
                  <Search className="w-3.5 h-3.5 text-stone-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                  <input
                    type="text"
                    placeholder="Cari SKU, Nama Produk, Kategori, Merk..."
                    value={itemSearch}
                    onChange={(e) => setItemSearch(e.target.value)}
                    className="w-full bg-white border border-stone-300 rounded-lg pl-8 pr-3 py-1.5 text-xs text-stone-800 placeholder-stone-400 focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500"
                  />
                </div>
                <div className="flex items-center gap-1 shrink-0 bg-white p-0.5 rounded-lg border border-stone-300 text-xs">
                  <button
                    type="button"
                    onClick={() => setItemStockFilter('all')}
                    className={`px-2 py-1 rounded text-[11px] font-medium transition-colors ${
                      itemStockFilter === 'all' ? 'bg-stone-800 text-white' : 'text-stone-600 hover:text-stone-900'
                    }`}
                  >
                    Semua
                  </button>
                  <button
                    type="button"
                    onClick={() => setItemStockFilter('in_stock')}
                    className={`px-2 py-1 rounded text-[11px] font-medium transition-colors ${
                      itemStockFilter === 'in_stock' ? 'bg-emerald-600 text-white' : 'text-stone-600 hover:text-stone-900'
                    }`}
                  >
                    Stok &gt; 0
                  </button>
                  <button
                    type="button"
                    onClick={() => setItemStockFilter('out_of_stock')}
                    className={`px-2 py-1 rounded text-[11px] font-medium transition-colors ${
                      itemStockFilter === 'out_of_stock' ? 'bg-rose-600 text-white' : 'text-stone-600 hover:text-stone-900'
                    }`}
                  >
                    Stok 0 / Kosong
                  </button>
                </div>
              </div>

              {/* Item List Table */}
              <div className="max-h-60 overflow-y-auto border border-stone-200 rounded-lg bg-white divide-y divide-stone-100">
                {filteredItems.length === 0 ? (
                  <div className="py-6 text-center text-xs text-stone-400">
                    Tidak ada produk yang cocok dengan kriteria pencarian.
                  </div>
                ) : (
                  filteredItems.map((item) => {
                    const isSelected = selectedRowIndices.has(item.rowIndex);
                    const skuDisplay = item.skuCol5 || item.skuCol6 || `Baris #${item.rowIndex}`;
                    const name = item.matchedStockItem?.description || (item.rawRow && item.rawRow[1] ? String(item.rawRow[1]) : '-');
                    const cat = item.matchedStockItem?.category || '';
                    const brand = item.matchedStockItem?.brand || '';

                    return (
                      <div
                        key={item.rowIndex}
                        onClick={() => handleToggleItem(item.rowIndex)}
                        className={`p-2.5 flex items-center justify-between text-xs cursor-pointer transition-colors ${
                          isSelected ? 'bg-emerald-50/50 hover:bg-emerald-50' : 'hover:bg-stone-50'
                        }`}
                      >
                        <div className="flex items-center gap-2.5 min-w-0 pr-2">
                          <input
                            type="checkbox"
                            checked={isSelected}
                            onChange={() => {}}
                            className="rounded text-emerald-600 focus:ring-emerald-500 shrink-0"
                          />
                          <div className="min-w-0">
                            <div className="flex items-center gap-1.5 flex-wrap">
                              <span className="font-mono font-bold text-stone-900 text-xs bg-stone-100 px-1.5 py-0.5 rounded">
                                {skuDisplay}
                              </span>
                              {cat && (
                                <span className="text-[10px] bg-sky-50 text-sky-800 border border-sky-200 px-1.5 py-0.2 rounded">
                                  {cat}
                                </span>
                              )}
                              {brand && (
                                <span className="text-[10px] bg-purple-50 text-purple-800 border border-purple-200 px-1.5 py-0.2 rounded">
                                  {brand}
                                </span>
                              )}
                              {item.matchedBy === '5digits_sku' && (
                                <span className="text-[10px] bg-emerald-100 text-emerald-800 font-semibold px-1 py-0.2 rounded">
                                  5 Digit
                                </span>
                              )}
                            </div>
                            <p className="text-stone-600 truncate text-[11px] mt-0.5">{name}</p>
                          </div>
                        </div>

                        <div className="text-right shrink-0">
                          {item.stockQty !== null ? (
                            <span
                              className={`text-xs font-mono font-bold px-2 py-0.5 rounded-full ${
                                item.stockQty > 0
                                  ? 'bg-emerald-100 text-emerald-800'
                                  : 'bg-rose-100 text-rose-800'
                              }`}
                            >
                              Stok: {item.stockQty}
                            </span>
                          ) : (
                            <span className="text-[10px] bg-stone-100 text-stone-500 px-1.5 py-0.5 rounded">
                              Tak Cocok
                            </span>
                          )}
                        </div>
                      </div>
                    );
                  })
                )}
              </div>
            </div>
          )}

          {/* Export Settings: Prefix & Summary */}
          <div className="p-3 bg-stone-50 rounded-xl border border-stone-200 space-y-2.5">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
              <div>
                <label className="block text-xs font-bold text-stone-700">
                  Nama Prefix File Toko:
                </label>
                <p className="text-[11px] text-stone-500">
                  Prefix di depan nama file (contoh: balist, gomall, dll.)
                </p>
              </div>
              <div className="flex items-center gap-1.5">
                <input
                  type="text"
                  value={customPrefix}
                  onChange={(e) => setCustomPrefix(e.target.value.toLowerCase().replace(/[^a-z0-9_-]/g, ''))}
                  placeholder="balist"
                  className="w-32 bg-white border border-stone-300 rounded-lg px-2.5 py-1 text-xs font-mono text-stone-800 focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 text-center"
                />
                <span className="text-xs font-mono text-stone-400">_YYYYMMDD_HHMMSS.xlsx</span>
              </div>
            </div>
          </div>
        </div>

        {/* Modal Footer */}
        <div className="p-4 sm:p-5 border-t border-stone-200 bg-stone-50/80 flex flex-col sm:flex-row items-center justify-between gap-3 shrink-0">
          <div className="text-xs text-stone-600 flex items-center gap-1.5">
            <span className="font-medium">Total Produk Siap Unduh:</span>
            <span className="font-bold text-emerald-700 text-sm font-mono bg-emerald-100/80 px-2 py-0.5 rounded-md">
              {itemsToExportCount} Produk
            </span>
          </div>

          <div className="flex items-center gap-2 w-full sm:w-auto">
            <button
              type="button"
              onClick={onClose}
              className="flex-1 sm:flex-none px-4 py-2 text-xs font-semibold text-stone-700 bg-white border border-stone-300 rounded-xl hover:bg-stone-100 transition-colors cursor-pointer"
            >
              Batal
            </button>
            <button
              type="button"
              onClick={handleConfirm}
              disabled={itemsToExportCount === 0 || isProcessing}
              className="flex-1 sm:flex-none px-5 py-2 text-xs font-bold text-white bg-emerald-700 hover:bg-emerald-800 disabled:bg-stone-300 disabled:cursor-not-allowed rounded-xl shadow-xs transition-all flex items-center justify-center gap-2 cursor-pointer"
            >
              <Download className="w-4 h-4" />
              <span>Unduh File Excel ({itemsToExportCount})</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
