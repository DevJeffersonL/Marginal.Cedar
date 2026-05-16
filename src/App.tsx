/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useEffect, useMemo } from 'react';
import { 
  Plus,
  LayoutDashboard, 
  Table,
  CalendarDays, 
  Trash2, 
  RefreshCw, 
  TrendingUp, 
  TrendingDown,
  BrainCircuit,
  Settings2,
  Pencil,
  XCircle
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { analyzeTrades } from './services/geminiService';

// --- Types ---
interface Trade {
  id: string;
  type: 'pair' | 'buy' | 'sell';
  buyDate?: string;
  buyAmount?: number;
  sellDate?: string;
  sellAmount?: number;
}

interface DailyReport {
  date: string;
  netCashflow: number;
}

type View = 'dashboard' | 'reports';

export default function App() {
  const [trades, setTrades] = useState<Trade[]>([]);
  const [activeView, setActiveView] = useState<View>('dashboard');
  const [isSyncing, setIsSyncing] = useState(false);
  const [gsUrl, setGsUrl] = useState<string>('');
  const [lastSynced, setLastSynced] = useState<string | null>(null);
  const [isAiAnalyzing, setIsAiAnalyzing] = useState(false);
  const [aiAnalysis, setAiAnalysis] = useState<string | null>(null);
  const [formMode, setFormMode] = useState<'pair' | 'buy' | 'sell'>('pair');
  const [editingId, setEditingId] = useState<string | null>(null);

  // Helper: Human-friendly date formatting
  const formatDate = (dateStr?: string) => {
    if (!dateStr || dateStr === 'undefined' || dateStr === 'null' || dateStr === '-') return '-';
    try {
      // If the string contains a full date like "Fri May 15...", try to parse just the date part
      const date = new Date(dateStr);
      if (isNaN(date.getTime())) return dateStr;
      
      // Use UTC to avoid timezone shifts when displaying simple dates
      const year = date.getUTCFullYear();
      const month = String(date.getUTCMonth() + 1).padStart(2, '0');
      const day = String(date.getUTCDate()).padStart(2, '0');
      
      // Check if it was an ISO string that came through
      if (dateStr.includes('T')) {
        return dateStr.split('T')[0];
      }
      
      return `${year}-${month}-${day}`;
    } catch {
      return dateStr;
    }
  };
  const [formData, setFormData] = useState({
    buyDate: '',
    buyAmount: '',
    sellDate: '',
    sellAmount: '',
  });

  // --- Persistence & Sync ---
  useEffect(() => {
    // Load local data
    const localTrades = localStorage.getItem('marginal_trades');
    if (localTrades) setTrades(JSON.parse(localTrades));

    const savedUrl = localStorage.getItem('gsUrl') || 'https://script.google.com/macros/s/AKfycbxf9pRWYkJwuyeZIN6umzQ2gSyjGoFDANTC64jrngK6aHmzuKAhp_ZWpVXZ44641NcBVw/exec';
    if (savedUrl) {
      setGsUrl(savedUrl);
      fetchRemoteTrades(savedUrl);
    }
    
    const savedSyncTime = localStorage.getItem('last_synced_time');
    if (savedSyncTime) setLastSynced(savedSyncTime);
  }, []);

  const saveLocally = (updatedTrades: Trade[]) => {
    localStorage.setItem('marginal_trades', JSON.stringify(updatedTrades));
    if (gsUrl) syncToSheets(updatedTrades);
  };

  const fetchRemoteTrades = async (url: string) => {
    if (!url) return;
    setIsSyncing(true);
    try {
      const response = await fetch(url);
      const data = await response.json();
      if (Array.isArray(data)) {
        // Defensive Mapping: Handle case-insensitive headers and common variations
        const sanitizedData = data.map((item: any) => {
          // Normalize keys to lowercase for matching
          const normalizedItem: any = {};
          Object.keys(item).forEach(k => normalizedItem[k.toLowerCase().replace(/\s+/g, '')] = item[k]);

          const buyAmt = Number(normalizedItem.buyamount || normalizedItem.buyprice || 0);
          const sellAmt = Number(normalizedItem.sellamount || normalizedItem.sellprice || 0);
          const buyD = normalizedItem.buydate || undefined;
          const sellD = normalizedItem.selldate || undefined;

          return {
            id: normalizedItem.id || crypto.randomUUID(),
            type: normalizedItem.type || (buyAmt > 0 && sellAmt > 0 ? 'pair' : buyAmt > 0 ? 'buy' : 'sell'),
            buyAmount: buyAmt,
            sellAmount: sellAmt,
            buyDate: buyD,
            sellDate: sellD
          };
        });
        
        setTrades(sanitizedData);
        localStorage.setItem('marginal_trades', JSON.stringify(sanitizedData));
        const time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        setLastSynced(time);
        localStorage.setItem('last_synced_time', time);
      }
    } catch (error) {
      console.error("Remote Sync Failed:", error);
    } finally {
      setIsSyncing(false);
    }
  };

  // Standardize values for sync
  const syncToSheets = async (data: Trade[]) => {
    if (!gsUrl) return;
    setIsSyncing(true);
    try {
      const mappedData = data.map(t => ({
        id: t.id,
        type: t.type,
        buyDate: t.buyDate || '',
        buyAmount: t.buyAmount || 0,
        sellDate: t.sellDate || '',
        sellAmount: t.sellAmount || 0
      }));

      await fetch(gsUrl, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'sync', data: mappedData }),
      });
      const time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      setLastSynced(time);
      localStorage.setItem('last_synced_time', time);
    } catch (error) {
      console.error("Sheets Sync Error:", error);
    } finally {
      setIsSyncing(false);
    }
  };

  const [validationErrors, setValidationErrors] = useState<string[]>([]);
  const [showMacroCode, setShowMacroCode] = useState(false);

  const macroCode = `function doGet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  var json = data.map(function(row) {
    var obj = {};
    headers.forEach(function(header, i) {
      obj[header.toLowerCase().replace(/\\s+/g, '')] = row[i];
    });
    return obj;
  });
  return ContentService.createTextOutput(JSON.stringify(json)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var params = JSON.parse(e.postData.contents);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (params.action === 'sync') {
    sheet.clear();
    sheet.appendRow(['ID', 'Type', 'Buy Date', 'Buy Amount', 'Sell Date', 'Sell Amount']);
    params.data.forEach(function(t) {
      sheet.appendRow([t.id, t.type, t.buyDate, t.buyAmount, t.sellDate, t.sellAmount]);
    });
    return ContentService.createTextOutput(JSON.stringify({status: 'success'})).setMimeType(ContentService.MimeType.JSON);
  }
}`;

  // --- Logic ---
  const handleAddTrade = (e: React.FormEvent) => {
    e.preventDefault();
    
    const errors: string[] = [];
    const buyPrice = Number(formData.buyAmount);
    const sellPrice = Number(formData.sellAmount);
    
    // Check if the user has touched/started filling either side
    const hasBuyInput = formData.buyDate.trim() !== '' || formData.buyAmount.trim() !== '';
    const hasSellInput = formData.sellDate.trim() !== '' || formData.sellAmount.trim() !== '';

    // Validation logic per mode
    if (formMode === 'pair') {
      if (!hasBuyInput && !hasSellInput) {
        errors.push('buyDate', 'buyAmount', 'sellDate', 'sellAmount');
      } else {
        if (hasBuyInput) {
          if (!formData.buyDate) errors.push('buyDate');
          if (buyPrice <= 0) errors.push('buyAmount');
        }
        if (hasSellInput) {
          if (!formData.sellDate) errors.push('sellDate');
          if (sellPrice <= 0) errors.push('sellAmount');
        }
        // Sequence check: Sell must be after or on Buy date
        if (formData.buyDate && formData.sellDate && new Date(formData.sellDate) < new Date(formData.buyDate)) {
          errors.push('sellDate');
        }
      }
    } else if (formMode === 'buy') {
      if (!formData.buyDate) errors.push('buyDate');
      if (buyPrice <= 0) errors.push('buyAmount');
    } else if (formMode === 'sell') {
      if (!formData.sellDate) errors.push('sellDate');
      if (sellPrice <= 0) errors.push('sellAmount');
    }

    if (errors.length > 0) {
      setValidationErrors(errors);
      return;
    }

    setValidationErrors([]);
    
    // Determine the actual type based on what was actually completed
    let effectiveType = formMode;
    if (formMode === 'pair') {
      const isBuyComplete = formData.buyDate && buyPrice > 0;
      const isSellComplete = formData.sellDate && sellPrice > 0;
      
      if (isBuyComplete && isSellComplete) effectiveType = 'pair';
      else if (isBuyComplete) effectiveType = 'buy';
      else if (isSellComplete) effectiveType = 'sell';
    }

    const tradeData: Partial<Trade> = {
      type: effectiveType as 'pair' | 'buy' | 'sell',
      buyDate: (effectiveType !== 'sell') ? formData.buyDate : undefined,
      buyAmount: (effectiveType !== 'sell') ? buyPrice : 0,
      sellDate: (effectiveType !== 'buy') ? formData.sellDate : undefined,
      sellAmount: (effectiveType !== 'buy') ? sellPrice : 0,
    };

    let updated: Trade[];
    if (editingId) {
      updated = trades.map(t => t.id === editingId ? { ...t, ...tradeData } as Trade : t);
      setEditingId(null);
    } else {
      const newTrade: Trade = {
        id: crypto.randomUUID(),
        ...tradeData as Trade
      };
      updated = [newTrade, ...trades];
    }

    setTrades(updated);
    saveLocally(updated);
    setFormData({ buyDate: '', buyAmount: '', sellDate: '', sellAmount: '' });
  };

  const handleEditTrade = (trade: Trade) => {
    setEditingId(trade.id);
    setFormMode(trade.type);
    setFormData({
      buyDate: trade.buyDate || '',
      buyAmount: trade.buyAmount?.toString() || '',
      sellDate: trade.sellDate || '',
      sellAmount: trade.sellAmount?.toString() || '',
    });
    // Scroll to form
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const cancelEdit = () => {
    setEditingId(null);
    setFormData({ buyDate: '', buyAmount: '', sellDate: '', sellAmount: '' });
  };

  const handleDeleteTrade = (id: string) => {
    const updated = trades.filter(t => t.id !== id);
    setTrades(updated);
    saveLocally(updated);
  };

  const totalNetProfit = useMemo(() => {
    return trades.reduce((acc, t) => acc + ((t.sellAmount || 0) - (t.buyAmount || 0)), 0);
  }, [trades]);

  const dailyReports = useMemo(() => {
    const reports: Record<string, number> = {};
    trades.forEach(t => {
      if (t.buyDate && t.buyAmount) {
        reports[t.buyDate] = (reports[t.buyDate] || 0) - t.buyAmount;
      }
      if (t.sellDate && t.sellAmount) {
        reports[t.sellDate] = (reports[t.sellDate] || 0) + t.sellAmount;
      }
    });
    return Object.entries(reports)
      .map(([date, netCashflow]) => ({ date, netCashflow }))
      .sort((a, b) => b.date.localeCompare(a.date));
  }, [trades]);

  const handleAiAnalysis = async () => {
    setIsAiAnalyzing(true);
    const analysis = await analyzeTrades(trades);
    setAiAnalysis(analysis);
    setIsAiAnalyzing(false);
  };

  // --- UI Components ---
  const NavTab = ({ id, label, icon: Icon }: { id: View, label: string, icon: any }) => (
    <button
      onClick={() => setActiveView(id)}
      className={`flex flex-col items-center gap-1.5 px-8 pt-3 pb-2 transition-all duration-300 relative haptic-interaction ${
        activeView === id ? 'text-white' : 'text-slate-500'
      }`}
    >
      <div className={`p-2 rounded-xl transition-all duration-300 ${activeView === id ? 'bg-white/10 shadow-[0_0_15px_rgba(255,255,255,0.05)]' : 'bg-transparent'}`}>
        <Icon size={22} strokeWidth={activeView === id ? 2.5 : 2} />
      </div>
      <span className={`text-[9px] font-black uppercase tracking-[0.2em] transition-all ${activeView === id ? 'opacity-100' : 'opacity-40'}`}>{label}</span>
      {activeView === id && (
        <motion.div
           layoutId="nav-dot"
           className="absolute -bottom-1 w-1 h-1 bg-accent rounded-full shadow-[0_0_8px_#6366f1]"
        />
      )}
    </button>
  );

  return (
    <div className="min-h-screen pb-24 selection:bg-accent/20">
      {/* Ambient Orbs */}
      <div className="bg-orb orb-1" />
      <div className="bg-orb orb-2" />

      {/* Header */}
      <header className="sticky top-0 z-50 glass border-b border-white/5 py-4 px-4 sm:px-12 flex items-center justify-between transition-all duration-500">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-2xl bg-accent flex items-center justify-center shadow-[0_0_20px_rgba(99,102,241,0.3)] shrink-0 haptic-interaction">
            <TrendingUp className="text-white" size={20} />
          </div>
          <div className="hidden xs:block">
            <h1 className="font-heading font-black text-xl tracking-tighter leading-none text-white uppercase">Marginal</h1>
            <p className="text-[9px] font-mono text-slate-500 mt-1 font-bold uppercase tracking-[0.2em]">Alpha Engine v2</p>
          </div>
        </div>

        <div className="flex items-center gap-3">
          <div className="flex flex-col items-end">
             <div className="flex items-center gap-2 bg-white/[0.03] px-3 py-1.5 rounded-full border border-white/5">
                <div className={`w-2 h-2 rounded-full shrink-0 ${isSyncing ? 'bg-amber-400 animate-pulse' : gsUrl ? 'bg-profit shadow-[0_0_10px_#10b981]' : 'bg-slate-600'}`} />
                <span className="text-[10px] uppercase font-black text-white/40 tracking-widest hidden sm:inline">
                  {isSyncing ? 'Syncing...' : gsUrl ? 'Cloud' : 'Local'}
                </span>
             </div>
          </div>
          
          <div className="flex items-center gap-1">
            <button onClick={() => setShowMacroCode(true)} className="p-2.5 hover:bg-white/5 rounded-2xl transition-colors text-slate-500 active:text-white haptic-interaction">
              <Table size={18} />
            </button>
            <button 
              onClick={() => {
                const url = prompt("Enter Google Apps Script Web App URL:", gsUrl);
                if (url !== null) {
                  setGsUrl(url);
                  localStorage.setItem('gsUrl', url);
                  if (url) fetchRemoteTrades(url);
                }
              }}
              className="p-2.5 bg-accent/10 border border-accent/20 rounded-2xl transition-colors text-accent haptic-interaction"
            >
              <Settings2 size={18} />
            </button>
          </div>
        </div>
      </header>

      {/* Macro Modal */}
      <AnimatePresence>
        {showMacroCode && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-md"
          >
            <motion.div 
              initial={{ scale: 0.9, y: 20 }}
              animate={{ scale: 1, y: 0 }}
              className="glass-card w-full max-w-2xl max-h-[80vh] overflow-hidden flex flex-col"
            >
              <div className="p-6 border-b border-white/5 flex justify-between items-center bg-white/5">
                <div>
                  <h3 className="font-heading font-bold text-lg">Sync Macro Code</h3>
                  <p className="text-xs text-slate-500">Copy this into Google Apps Script Extensions</p>
                  <p className="text-[10px] text-accent mt-1 uppercase font-bold">Tip: Format Columns D & F as 'Number' in Sheets</p>
                </div>
                <button onClick={() => setShowMacroCode(false)} className="p-2 hover:bg-white/10 rounded-full">
                  <XCircle size={20} />
                </button>
              </div>
              <div className="p-6 overflow-auto bg-slate-950">
                <pre className="text-[10px] sm:text-xs font-mono text-profit leading-relaxed whitespace-pre">
                  {macroCode}
                </pre>
              </div>
              <div className="p-6 border-t border-white/5 flex gap-4">
                <button 
                  onClick={() => {
                    navigator.clipboard.writeText(macroCode);
                    alert("Copied to clipboard!");
                  }}
                  className="flex-1 bg-accent py-3 rounded-xl font-bold text-xs uppercase tracking-widest hover:bg-accent/80 transition-all"
                >
                  Copy Code
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      <main className="max-w-7xl mx-auto p-4 sm:p-8 md:p-12">
        <AnimatePresence mode="wait">
          {activeView === 'dashboard' ? (
            <motion.div 
              key="dashboard"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="grid grid-cols-1 lg:grid-cols-[1fr_1.5fr] gap-8"
            >
              {/* Left Column */}
              <div className="space-y-6 sm:space-y-8">
                {/* Hero Card */}
                <div className="glass-card surface-hero p-8 sm:p-12 flex flex-col items-center justify-center text-center relative group overflow-hidden">
                  <div className="absolute top-2 w-16 h-1.5 bg-white/5 rounded-full blur-xs" />
                  
                  <div className="absolute top-6 right-6 opacity-0 group-hover:opacity-100 transition-opacity z-10">
                    <button 
                      onClick={handleAiAnalysis}
                      className="p-3 bg-white/5 hover:bg-accent rounded-2xl text-slate-400 hover:text-white transition-all flex items-center gap-2 group/ai border border-white/5 haptic-interaction"
                    >
                      <BrainCircuit size={18} className={isAiAnalyzing ? 'animate-spin' : ''} />
                    </button>
                  </div>

                  <h2 className="text-[10px] font-black uppercase tracking-[0.4em] text-slate-500 mb-8 flex items-center gap-3">
                    <span className="w-6 h-[1px] bg-white/10" />
                    Portfolio Valuation
                    <span className="w-6 h-[1px] bg-white/10" />
                  </h2>
                  
                  <div className="relative">
                    <div className={`text-6xl sm:text-9xl font-heading font-black tabular-nums tracking-[-0.06em] leading-none ${
                      totalNetProfit >= 0 ? 'gradient-text-profit' : 'text-loss drop-shadow-[0_0_40px_rgba(244,63,94,0.3)]'
                    }`}>
                      <span className="opacity-20 mr-1">$</span>
                      {totalNetProfit.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}
                    </div>
                  </div>
                  
                  <div className="mt-8 flex flex-wrap justify-center items-center gap-4">
                    <div className="flex items-center gap-2 bg-white/5 px-4 py-2 rounded-2xl border border-white/5">
                      <div className={`w-2 h-2 rounded-full ${totalNetProfit >= 0 ? 'bg-profit shadow-[0_0_12px_#10b981]' : 'bg-loss animate-pulse'}`} />
                      <span className="text-[10px] font-black uppercase tracking-widest text-slate-400">
                        {totalNetProfit >= 0 ? 'Bullish Sentiment' : 'Bearish Weighted'}
                      </span>
                    </div>
                    <div className="text-[10px] font-black uppercase tracking-widest text-slate-600">
                      {trades.length} Verified Nodes
                    </div>
                  </div>
                </div>
                
                {aiAnalysis && (
                    <motion.div 
                      initial={{ opacity: 0, scale: 0.95 }}
                      animate={{ opacity: 1, scale: 1 }}
                      className="mt-8 p-5 bg-accent/5 backdrop-blur-xl rounded-2xl text-left border border-accent/20 relative group/insight"
                    >
                       <div className="absolute -top-3 left-6 px-3 py-1 bg-accent text-[9px] font-bold text-white uppercase tracking-widest rounded-full shadow-lg">
                         Proprietary Insight
                       </div>
                       <p className="text-xs text-slate-300 leading-relaxed font-medium italic">
                         "{aiAnalysis}"
                       </p>
                    </motion.div>
                  )}
                </div>

                {/* Form Card */}
                <div className={`glass-card p-5 sm:p-8 transition-all duration-500 ${editingId ? 'ring-2 ring-accent ring-inset bg-accent/5' : ''}`}>
                  <div className="flex justify-between items-center mb-6 sm:mb-8">
                    <h3 className="font-heading font-bold text-base sm:text-lg flex items-center gap-2">
                       {editingId ? <Pencil className="text-accent" size={18} /> : <Plus className="text-accent" size={18} />}
                       {editingId ? 'Edit Transaction' : 'New Transaction'}
                    </h3>
                    
                    {editingId && (
                       <button 
                        onClick={cancelEdit}
                        className="flex items-center gap-2 text-[10px] uppercase font-bold text-slate-500 hover:text-loss transition-colors"
                       >
                         <XCircle size={14} /> Cancel
                       </button>
                    )}
                  </div>

                  <form onSubmit={handleAddTrade} className="space-y-6" noValidate>
                    <div className="space-y-6 sm:space-y-8">
                      {/* Buy Section */}
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-6">
                        <div className="relative group/field">
                          <input 
                            type="date" 
                            value={formData.buyDate}
                            onChange={e => {
                              setFormData({...formData, buyDate: e.target.value});
                              setValidationErrors(prev => prev.filter(err => err !== 'buyDate'));
                            }}
                            className={`peer w-full bg-slate-800/20 border rounded-2xl px-5 py-4 pt-8 text-sm focus:outline-hidden focus:ring-2 transition-all font-mono tracking-tight ${
                              validationErrors.includes('buyDate') ? 'border-loss/50 focus:border-loss ring-loss/10' : 'border-white/5 focus:border-accent focus:ring-accent/10'
                            }`}
                          />
                          <label className={`absolute left-5 top-2 text-[10px] uppercase font-black tracking-widest transition-all ${
                            validationErrors.includes('buyDate') ? 'text-loss' : 'text-slate-500 peer-focus:text-accent'
                          }`}>
                            Entry Timestamp
                          </label>
                        </div>
                        <div className="relative group/field">
                          <input 
                            type="number" 
                            placeholder=" "
                            step="0.01"
                            value={formData.buyAmount}
                            onChange={e => {
                              setFormData({...formData, buyAmount: e.target.value});
                              setValidationErrors(prev => prev.filter(err => err !== 'buyAmount'));
                            }}
                            className={`peer w-full bg-slate-800/20 border rounded-2xl px-5 py-4 pt-8 text-sm focus:outline-hidden focus:ring-2 transition-all font-mono placeholder:opacity-0 ${
                              validationErrors.includes('buyAmount') ? 'border-loss/50 focus:border-loss ring-loss/10' : 'border-white/5 focus:border-accent focus:ring-accent/10'
                            }`}
                          />
                          <label className={`absolute left-5 top-2 text-[10px] uppercase font-black tracking-widest transition-all peer-placeholder-shown:top-5 peer-placeholder-shown:text-sm peer-focus:top-2 peer-focus:text-[10px] ${
                            validationErrors.includes('buyAmount') ? 'text-loss' : 'text-slate-500 peer-focus:text-accent'
                          }`}>
                            Capital Allocation
                          </label>
                        </div>
                      </div>

                      {/* Sell Section */}
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-6">
                        <div className="relative group/field">
                          <input 
                            type="date" 
                            value={formData.sellDate}
                            onChange={e => {
                              setFormData({...formData, sellDate: e.target.value});
                              setValidationErrors(prev => prev.filter(err => err !== 'sellDate'));
                            }}
                            className={`peer w-full bg-slate-800/20 border rounded-2xl px-5 py-4 pt-8 text-sm focus:outline-hidden focus:ring-2 transition-all font-mono tracking-tight ${
                              validationErrors.includes('sellDate') ? 'border-loss/50 focus:border-loss ring-loss/10' : 'border-white/5 focus:border-accent focus:ring-accent/10'
                            }`}
                          />
                          <label className={`absolute left-5 top-2 text-[10px] uppercase font-black tracking-widest transition-all ${
                            validationErrors.includes('sellDate') ? 'text-loss' : 'text-slate-500 peer-focus:text-accent'
                          }`}>
                            Exit Timestamp
                          </label>
                        </div>
                        <div className="relative group/field">
                          <input 
                            type="number" 
                            placeholder=" "
                            step="0.01"
                            value={formData.sellAmount}
                            onChange={e => {
                              setFormData({...formData, sellAmount: e.target.value});
                              setValidationErrors(prev => prev.filter(err => err !== 'sellAmount'));
                            }}
                            className={`peer w-full bg-slate-800/20 border rounded-2xl px-5 py-4 pt-8 text-sm focus:outline-hidden focus:ring-2 transition-all font-mono placeholder:opacity-0 ${
                              validationErrors.includes('sellAmount') ? 'border-loss/50 focus:border-loss ring-loss/10' : 'border-white/5 focus:border-accent focus:ring-accent/10'
                            }`}
                          />
                          <label className={`absolute left-5 top-2 text-[10px] uppercase font-black tracking-widest transition-all peer-placeholder-shown:top-5 peer-placeholder-shown:text-sm peer-focus:top-2 peer-focus:text-[10px] ${
                            validationErrors.includes('sellAmount') ? 'text-loss' : 'text-slate-500 peer-focus:text-accent'
                          }`}>
                            Liquidation Value
                          </label>
                        </div>
                      </div>
                    </div>

                    <button 
                      type="submit"
                      className="w-full bg-accent hover:bg-slate-100 text-white hover:text-accent font-black py-5 rounded-[2rem] shadow-[0_20px_40px_rgba(99,102,241,0.2)] transition-all active:scale-[0.95] mt-6 flex items-center justify-center gap-3 uppercase tracking-[0.2em] text-[10px] group haptic-interaction"
                    >
                      {editingId ? <RefreshCw size={20} className="group-hover:rotate-180 transition-transform" /> : <Plus size={20} className="group-hover:rotate-90 transition-transform" strokeWidth={3} />}
                      {editingId ? 'Push Updates' : `${formMode === 'pair' ? 'Initialize Trade' : formMode === 'buy' ? 'Record Entry' : 'Record Exit'}`}
                    </button>
                  </form>
                </div>

              {/* Right Column - Ledger */}
              <div className="glass-card min-h-[400px] flex flex-col overflow-hidden">
                <div className="p-5 sm:p-8 border-b border-white/5 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2">
                   <h3 className="font-heading font-bold text-base sm:text-lg flex items-center gap-2">
                    <Table className="text-accent" size={20} />
                    Transaction Ledger
                  </h3>
                  <span className="text-[9px] sm:text-[10px] font-mono text-slate-500 uppercase tracking-widest">{trades.length} Records</span>
                </div>
                
                <div className="flex-1 p-4 sm:p-6">
                  {/* Desktop Table View */}
                  <div className="hidden sm:block overflow-x-auto">
                    <table className="w-full text-left border-separate border-spacing-y-3">
                      <thead>
                        <tr>
                          <th className="px-6 pb-3 text-[10px] uppercase tracking-widest font-bold text-slate-500">Buy Side</th>
                          <th className="px-6 pb-3 text-[10px] uppercase tracking-widest font-bold text-slate-500">Sell Side</th>
                          <th className="px-6 pb-3 text-[10px] uppercase tracking-widest font-bold text-slate-500">Profit</th>
                          <th className="px-6 pb-3 text-right text-[10px] uppercase tracking-widest font-bold text-slate-500 pr-8">Actions</th>
                        </tr>
                      </thead>
                      <tbody>
                        {trades.length === 0 && (
                          <tr>
                            <td colSpan={4} className="px-6 py-24 text-center text-slate-500 font-light italic">
                              No transactions recorded yet.
                            </td>
                          </tr>
                        )}
                        {trades.map((trade) => {
                          const profit = (trade.sellAmount || 0) - (trade.buyAmount || 0);
                          const isBuyAvailable = trade.type !== 'sell' && (trade.buyAmount !== 0 || trade.buyDate);
                          const isSellAvailable = trade.type !== 'buy' && (trade.sellAmount !== 0 || trade.sellDate);

                          return (
                            <motion.tr 
                              key={trade.id} 
                              initial={{ opacity: 0 }}
                              animate={{ opacity: 1 }}
                              className="hover:bg-white/10 transition-colors group bg-white/5"
                            >
                              <td className="px-6 py-4 rounded-l-xl">
                                <div className="flex items-center gap-3">
                                  <span className={`w-1 h-8 rounded-full shrink-0 ${
                                    trade.type === 'pair' ? 'bg-accent' : trade.type === 'buy' ? 'bg-profit' : 'bg-loss'
                                  }`} />
                                  <div className="min-w-0">
                                    {isBuyAvailable ? (
                                      <>
                                        <p className="text-sm font-bold truncate text-white">${Number(trade.buyAmount).toLocaleString()}</p>
                                        <p className="text-[10px] font-mono text-slate-500 uppercase tracking-tight">{formatDate(trade.buyDate)}</p>
                                      </>
                                    ) : (
                                      <div className="flex flex-col opacity-40">
                                        <span className="text-[10px] font-mono font-bold text-slate-400 tracking-tighter">NO_BUY</span>
                                        <span className="text-[8px] text-slate-500 uppercase font-bold tracking-widest">Standalone</span>
                                      </div>
                                    )}
                                  </div>
                                </div>
                              </td>
                              <td className="px-6 py-4">
                                <div className="min-w-0">
                                  {isSellAvailable ? (
                                    <>
                                      <p className="text-sm font-bold truncate text-white">${Number(trade.sellAmount).toLocaleString()}</p>
                                      <p className="text-[10px] font-mono text-slate-500 uppercase tracking-tight">{formatDate(trade.sellDate)}</p>
                                    </>
                                  ) : (
                                    <div className="flex flex-col opacity-40">
                                      <span className="text-[10px] font-mono font-bold text-slate-400 tracking-tighter">OPEN</span>
                                      <span className="text-[8px] text-slate-500 uppercase font-bold tracking-widest">Exit Req</span>
                                    </div>
                                  )}
                                </div>
                              </td>
                              <td className="px-6 py-4">
                                <div className="flex flex-col gap-1">
                                  <span className={`text-[8px] self-start px-1.5 py-0.5 rounded-full font-bold uppercase tracking-widest ${
                                    trade.type === 'pair' ? 'bg-accent/20 text-accent' : trade.type === 'buy' ? 'bg-profit/20 text-profit' : 'bg-loss/20 text-loss'
                                  }`}>
                                    {trade.type}
                                  </span>
                                  <p className={`font-mono text-sm font-bold ${profit >= 0 ? 'text-profit' : 'text-loss'}`}>
                                    {profit >= 0 ? '+' : ''}${Math.abs(profit).toLocaleString(undefined, { minimumFractionDigits: 2 })}
                                  </p>
                                </div>
                              </td>
                              <td className="px-6 py-4 text-right rounded-r-xl pr-6">
                                <div className="flex items-center justify-end gap-2">
                                  <button onClick={() => handleEditTrade(trade)} className="p-2 text-slate-600 hover:text-accent transition-colors hover:bg-accent/10 rounded-lg">
                                    <Pencil size={14} />
                                  </button>
                                  <button onClick={() => handleDeleteTrade(trade.id)} className="p-2 text-slate-600 hover:text-loss transition-colors hover:bg-loss/10 rounded-lg">
                                    <Trash2 size={14} />
                                  </button>
                                </div>
                              </td>
                            </motion.tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>

                  {/* Mobile Mobile/List View */}
                    <div className="sm:hidden space-y-4">
                      {trades.length === 0 && (
                        <p className="py-20 text-center text-slate-500 font-light italic">No records found.</p>
                      )}
                      {trades.map((trade) => {
                        const profit = (trade.sellAmount || 0) - (trade.buyAmount || 0);
                        const isBuyAvailable = trade.type !== 'sell' && (trade.buyAmount !== 0 || trade.buyDate);
                        const isSellAvailable = trade.type !== 'buy' && (trade.sellAmount !== 0 || trade.sellDate);

                        return (
                          <motion.div 
                            key={trade.id}
                            layout
                            className="bg-white/[0.04] p-5 rounded-[2rem] border border-white/5 shadow-xl relative overflow-hidden"
                          >
                            <div className={`absolute top-0 left-0 w-1.5 h-full ${trade.type === 'pair' ? 'bg-accent' : trade.type === 'buy' ? 'bg-profit' : 'bg-loss'}`} />
                            
                            <div className="flex justify-between items-center mb-5 pl-2">
                              <div className="flex items-center gap-2">
                                <span className={`text-[10px] px-3 py-1 rounded-full font-black uppercase tracking-widest ${
                                  trade.type === 'pair' ? 'bg-accent/20 text-accent' : trade.type === 'buy' ? 'bg-profit/20 text-profit' : 'bg-loss/20 text-loss'
                                }`}>
                                  {trade.type}
                                </span>
                              </div>
                              <div className="flex gap-2">
                                <button onClick={() => handleEditTrade(trade)} className="p-2.5 bg-white/5 rounded-xl text-slate-400 haptic-interaction"><Pencil size={14} /></button>
                                <button onClick={() => handleDeleteTrade(trade.id)} className="p-2.5 bg-white/10 rounded-xl text-slate-100 haptic-interaction"><Trash2 size={14} /></button>
                              </div>
                            </div>

                            <div className="grid grid-cols-2 gap-6 mb-6 pl-2">
                              <div>
                                <p className="text-[10px] uppercase font-black text-slate-500 tracking-widest mb-1.5">Capital</p>
                                {isBuyAvailable ? (
                                  <div>
                                    <p className="text-lg font-heading font-bold text-white">${Number(trade.buyAmount).toLocaleString(undefined, { maximumFractionDigits: 0 })}</p>
                                    <p className="text-[10px] font-mono text-slate-500 font-bold">{formatDate(trade.buyDate)}</p>
                                  </div>
                                ) : <p className="text-xs font-bold text-slate-600">None</p>}
                              </div>
                              <div>
                                <p className="text-[10px] uppercase font-black text-slate-500 tracking-widest mb-1.5">Return</p>
                                {isSellAvailable ? (
                                  <div>
                                    <p className="text-lg font-heading font-bold text-white">${Number(trade.sellAmount).toLocaleString(undefined, { maximumFractionDigits: 0 })}</p>
                                    <p className="text-[10px] font-mono text-slate-500 font-bold">{formatDate(trade.sellDate)}</p>
                                  </div>
                                ) : <p className="text-xs font-bold text-accent/40 italic">In Orbit</p>}
                              </div>
                            </div>

                            <div className={`p-4 rounded-2xl flex justify-between items-center ${profit >= 0 ? 'bg-profit/10' : 'bg-loss/10'} ml-2`}>
                              <span className="text-[10px] font-black uppercase text-white/40 tracking-widest">Net Differential</span>
                              <span className={`font-mono text-lg font-black ${profit >= 0 ? 'text-profit' : 'text-loss'}`}>
                                {profit >= 0 ? '+' : '-'}${Math.abs(profit).toLocaleString(undefined, { minimumFractionDigits: 2 })}
                              </span>
                            </div>
                          </motion.div>
                        );
                      })}
                    </div>
                </div>
              </div>
            </motion.div>
          ) : (
            <motion.div 
              key="reports"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="max-w-3xl mx-auto"
            >
              <div className="glass-card">
                 <div className="p-8 border-b border-white/5 flex justify-between items-center">
                   <h3 className="font-heading font-bold text-lg flex items-center gap-2">
                    <CalendarDays className="text-accent" size={20} />
                    Daily P&L Report
                  </h3>
                </div>
                <div className="overflow-hidden">
                  <table className="w-full text-left">
                    <thead className="bg-white/[0.02] border-b border-white/5">
                      <tr>
                        <th className="px-8 py-4 text-[10px] uppercase tracking-widest font-bold text-slate-500">Date</th>
                        <th className="px-8 py-4 text-[10px] uppercase tracking-widest font-bold text-slate-500 text-right">Net Cashflow</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-white/5 font-mono">
                       {dailyReports.length === 0 && (
                        <tr>
                          <td colSpan={2} className="px-6 py-24 text-center text-slate-500 font-light italic">
                            Insufficient data for daily reports.
                          </td>
                        </tr>
                      )}
                      {dailyReports.map((report, idx) => (
                        <tr key={idx} className="hover:bg-white/[0.02] transition-colors">
                          <td className="px-8 py-5 text-sm font-medium">{formatDate(report.date)}</td>
                          <td className={`px-8 py-5 text-sm font-bold text-right ${report.netCashflow >= 0 ? 'text-profit' : 'text-loss'}`}>
                             {report.netCashflow >= 0 ? '+' : ''}${report.netCashflow.toLocaleString(undefined, { minimumFractionDigits: 2 })}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Persistent Footer Nav */}
      <nav className="fixed bottom-6 left-0 right-0 z-50 flex justify-center pointer-events-none">
        <div className="glass shadow-[0_20px_50px_rgba(0,0,0,0.8)] rounded-full border border-white/10 flex gap-1 p-1 pointer-events-auto scale-90 sm:scale-100">
          <NavTab id="dashboard" label="Home" icon={LayoutDashboard} />
          <NavTab id="reports" label="Ledger" icon={Table} />
        </div>
      </nav>
    </div>
  );
}
