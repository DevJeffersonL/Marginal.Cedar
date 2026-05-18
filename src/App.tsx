/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo } from 'react';
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
  XCircle,
  Clock,
  LucideIcon
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import debounce from 'lodash.debounce';

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

// --- Constants ---
const MACRO_CODE = `/**
 * MARGINAL TRADE TRACKER - PRO MACRO
 * Version: 2.1 (Added SECRET_KEY Security)
 * 
 * Instructions:
 * 1. Create a Google Sheet.
 * 2. Extensions > Apps Script.
 * 3. Paste this code and Save.
 * 4. IMPORTANT: Change the SECRET_KEY below to a unique password.
 * 5. Deploy > New Deployment > Web App.
 * 6. Execute as: Me, Access: Anyone.
 * 7. Copy the URL and your Secret Key into the App Settings.
 */

var SECRET_KEY = "your_secret_password_here";

function setupSheet(sheet) {
  var headers = [["ID", "Type", "Buy Date", "Buy Amt ($)", "Sell Date", "Sell Amt ($)", "Net P&L ($)"]];
  var headerRange = sheet.getRange(1, 1, 1, 7);
  headerRange.setValues(headers);
  
  headerRange.setFontWeight("bold")
             .setBackground("#09090b")
             .setFontColor("#6366f1")
             .setHorizontalAlignment("center");
  
  sheet.setFrozenRows(1);
  sheet.getRange("D:D").setNumberFormat("$#,##0.00");
  sheet.getRange("F:G").setNumberFormat("$#,##0.00");
  sheet.autoResizeColumns(1, 7);
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  try {
    var payload = JSON.parse(e.postData.contents);
    if (payload.secret !== SECRET_KEY) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (payload.action === 'sync') {
      var data = payload.data || [];
      if (sheet.getLastRow() === 0) setupSheet(sheet);
      
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow, 7).clearContent();
        sheet.getRange(2, 1, lastRow, 7).setFontColor("#ffffff");
      }
      
      if (data.length > 0) {
        var rows = [];
        var colors = [];
        
        data.forEach(function(item) {
          var buy = parseFloat(item.buyAmount) || 0;
          var sell = parseFloat(item.sellAmount) || 0;
          var profit = sell - buy;
          
          rows.push([
            item.id,
            (item.type || "pair").toUpperCase(),
            item.buyDate || "-",
            buy || "",
            item.sellDate || "-",
            sell || "",
            profit
          ]);
          
          var rowColors = ["#ffffff", "#94a3b8", "#ffffff", "#ffffff", "#ffffff", "#ffffff", "#ffffff"];
          rowColors[6] = profit >= 0 ? "#10b981" : "#ef4444"; 
          colors.push(rowColors);
        });
        
        var targetRange = sheet.getRange(2, 1, rows.length, 7);
        targetRange.setValues(rows);
        targetRange.setFontColors(colors);
      }
      
      sheet.autoResizeColumns(1, 7);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', count: data.length }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  if (e.parameter.secret !== SECRET_KEY) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Unauthorized' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  try {
    if (sheet.getLastRow() <= 1) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
    var data = sheet.getDataRange().getValues().slice(1);
    var result = data.map(function(row) {
      return {
        id: row[0].toString(),
        type: row[1].toString().toLowerCase(),
        buyDate: row[2] === "-" ? undefined : row[2],
        buyAmount: parseFloat(row[3]) || 0,
        sellDate: row[4] === "-" ? undefined : row[4],
        sellAmount: parseFloat(row[5]) || 0
      };
    });
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
  }
}`;

// --- Helper Functions ---
const formatDate = (dateStr?: string) => {
  if (!dateStr || dateStr === 'undefined' || dateStr === 'null' || dateStr === '-') return '-';
  try {
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) return dateStr;
    const year = date.getUTCFullYear();
    const month = String(date.getUTCMonth() + 1).padStart(2, '0');
    const day = String(date.getUTCDate()).padStart(2, '0');
    if (dateStr.includes('T')) return dateStr.split('T')[0];
    return `${year}-${month}-${day}`;
  } catch {
    return dateStr;
  }
};

// --- Memoized UI Components ---
const NavTab = React.memo(({ 
  id, 
  label, 
  icon: Icon, 
  activeView, 
  onClick 
}: { 
  id: View, 
  label: string, 
  icon: LucideIcon, 
  activeView: View, 
  onClick: (id: View) => void 
}) => {
  const isActive = activeView === id;
  return (
    <button
      onClick={() => onClick(id)}
      className={`flex flex-col items-center justify-center gap-1.5 px-7 pt-3 pb-2.5 transition-all duration-300 relative haptic-interaction ${
        isActive ? 'text-white' : 'text-slate-500'
      }`}
    >
      <div className={`p-2 rounded-2xl transition-all duration-500 ${isActive ? 'bg-white/10 ring-1 ring-white/10 shadow-[0_0_20px_rgba(255,255,255,0.05)] scale-110' : 'bg-transparent hover:bg-white/5'}`}>
        <Icon size={22} strokeWidth={isActive ? 2.5 : 2} />
      </div>
      <span className={`text-[9px] font-bold uppercase tracking-[0.2em] transition-all ${isActive ? 'opacity-100' : 'opacity-30'}`}>{label}</span>
      {isActive && (
        <motion.div
           layoutId="nav-dot"
           className="absolute bottom-0 w-1 h-1 bg-accent rounded-full shadow-[0_0_10px_#6366f1]"
        />
      )}
    </button>
  );
});

const TradeRow = React.memo(({ 
  trade, 
  onEdit, 
  onDelete 
}: { 
  trade: Trade, 
  onEdit: (t: Trade) => void, 
  onDelete: (id: string) => void 
}) => {
  const profit = (trade.sellAmount || 0) - (trade.buyAmount || 0);
  const isBuyAvailable = trade.type !== 'sell' && (trade.buyAmount !== 0 || trade.buyDate);
  const isSellAvailable = trade.type !== 'buy' && (trade.sellAmount !== 0 || trade.sellDate);

  return (
    <motion.tr 
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      className="hover:bg-white/5 transition-colors group bg-[#0a0a0a]"
    >
      <td className="px-4 py-3 rounded-l-lg border-y border-l border-white/5 group-hover:border-white/10 transition-colors">
        <div className="flex items-center gap-2">
          <div className="min-w-0">
            {isBuyAvailable ? (
              <>
                <p className="text-sm font-medium text-white">₹{Number(trade.buyAmount).toLocaleString()}</p>
                <p className="text-[10px] font-mono text-slate-500">{formatDate(trade.buyDate)}</p>
              </>
            ) : <span className="text-xs text-slate-600">—</span>}
          </div>
        </div>
      </td>
      <td className="px-4 py-3 border-y border-white/5 group-hover:border-white/10 transition-colors">
        <div className="min-w-0">
          {isSellAvailable ? (
            <>
              <p className="text-sm font-medium text-white">₹{Number(trade.sellAmount).toLocaleString()}</p>
              <p className="text-[10px] font-mono text-slate-500">{formatDate(trade.sellDate)}</p>
            </>
          ) : <span className="text-xs text-accent/50 italic">Open</span>}
        </div>
      </td>
      <td className="px-4 py-3 border-y border-white/5 group-hover:border-white/10 transition-colors">
        <p className={`font-mono text-sm font-medium ${profit >= 0 ? 'text-profit' : 'text-loss'}`}>
          {profit >= 0 ? '+' : ''}₹{Math.abs(profit).toLocaleString(undefined, { minimumFractionDigits: 2 })}
        </p>
      </td>
      <td className="px-4 py-3 rounded-r-lg border-y border-r border-white/5 group-hover:border-white/10 text-right transition-colors pr-4">
        <div className="flex items-center justify-end gap-1">
          <button onClick={() => onEdit(trade)} className="p-1.5 hover:bg-white/10 rounded-lg text-slate-500 hover:text-white transition-all">
            <Pencil size={12} />
          </button>
          <button onClick={() => onDelete(trade.id)} className="p-1.5 hover:bg-loss/20 rounded-lg text-slate-500 hover:text-loss transition-all">
            <Trash2 size={12} />
          </button>
        </div>
      </td>
    </motion.tr>
  );
});

const MobileTradeCard = React.memo(({ 
  trade, 
  onEdit, 
  onDelete 
}: { 
  trade: Trade, 
  onEdit: (t: Trade) => void, 
  onDelete: (id: string) => void 
}) => {
  const profit = (trade.sellAmount || 0) - (trade.buyAmount || 0);
  const isBuyAvailable = trade.type !== 'sell' && (trade.buyAmount !== 0 || trade.buyDate);
  const isSellAvailable = trade.type !== 'buy' && (trade.sellAmount !== 0 || trade.sellDate);

  return (
    <motion.div 
      layout
      className="bg-[#0a0a0a] p-4 rounded-2xl border border-white/5 relative overflow-hidden active:scale-[0.98] transition-all duration-200"
    >
      <div className={`absolute top-0 left-0 w-1.5 h-full ${trade.type === 'pair' ? 'bg-accent/40' : trade.type === 'buy' ? 'bg-profit/40' : 'bg-loss/40'}`} />
      
      <div className="flex justify-between items-center mb-4 pl-2">
          <div className="flex items-center gap-2">
            <span className={`text-[8px] px-2.5 py-1 rounded-full font-bold uppercase tracking-widest ${
              trade.type === 'pair' ? 'bg-accent/10 text-accent ring-1 ring-accent/20' : trade.type === 'buy' ? 'bg-profit/10 text-profit ring-1 ring-profit/20' : 'bg-loss/10 text-loss ring-1 ring-loss/20'
            }`}>
              {trade.type}
            </span>
          </div>
          <div className="flex gap-2">
            <button onClick={() => onEdit(trade)} className="p-3 bg-white/5 hover:bg-white/10 rounded-xl text-slate-400 active:scale-90 transition-all haptic-interaction"><Pencil size={15} /></button>
            <button onClick={() => onDelete(trade.id)} className="p-3 bg-loss/10 hover:bg-loss/20 rounded-xl text-loss active:scale-90 transition-all haptic-interaction"><Trash2 size={15} /></button>
          </div>
      </div>

      <div className="flex justify-between items-end pl-2">
        <div className="grid grid-cols-2 gap-4 flex-1">
          <div>
            <p className="text-[9px] uppercase font-bold text-slate-600 mb-1 tracking-widest leading-none">Entry</p>
            <p className="text-base font-sans font-medium text-white tabular-nums leading-none">₹{Number(isBuyAvailable ? trade.buyAmount : 0).toLocaleString()}</p>
            {isBuyAvailable && <p className="text-[10px] font-mono text-slate-500 mt-1.5">{formatDate(trade.buyDate)}</p>}
          </div>
          <div>
            <p className="text-[9px] uppercase font-bold text-slate-600 mb-1 tracking-widest leading-none">Exit</p>
            {isSellAvailable ? (
              <>
                <p className="text-base font-sans font-medium text-white tabular-nums leading-none">₹{Number(trade.sellAmount).toLocaleString()}</p>
                <p className="text-[10px] font-mono text-slate-500 mt-1.5">{formatDate(trade.sellDate)}</p>
              </>
            ) : (
              <div className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded-lg bg-accent/10 border border-accent/20">
                <span className="w-1.5 h-1.5 bg-accent rounded-full animate-pulse" />
                <span className="text-[9px] font-bold text-accent uppercase tracking-tighter">Open</span>
              </div>
            )}
          </div>
        </div>
        <div className="text-right flex flex-col items-end">
          <p className="text-[9px] uppercase font-bold text-slate-600 mb-1 tracking-widest leading-none">Result</p>
          <p className={`font-sans text-xl font-bold tabular-nums leading-none ${profit >= 0 ? 'text-profit' : 'text-loss'}`}>
            {profit >= 0 ? '+' : ''}₹{Math.abs(profit).toLocaleString()}
          </p>
        </div>
      </div>
    </motion.div>
  );
});

export default function App() {
  const [trades, setTrades] = useState<Trade[]>([]);
  const [activeView, setActiveView] = useState<View>('dashboard');
  const [isSyncing, setIsSyncing] = useState(false);
  const [gsUrl, setGsUrl] = useState<string>(localStorage.getItem('gsUrl') || '');
  const [gsSecret, setGsSecret] = useState<string>(localStorage.getItem('gsSecret') || '');
  const [lastSynced, setLastSynced] = useState<string | null>(localStorage.getItem('last_synced_time'));
  const [isAiAnalyzing, setIsAiAnalyzing] = useState(false);
  const [aiAnalysis, setAiAnalysis] = useState<string | null>(null);
  const [formMode, setFormMode] = useState<'pair' | 'buy' | 'sell'>('pair');
  const [editingId, setEditingId] = useState<string | null>(null);
  const [isGroupedByDate, setIsGroupedByDate] = useState(true);
  const [validationErrors, setValidationErrors] = useState<string[]>([]);
  const [showMacroCode, setShowMacroCode] = useState(false);
  const [showSettings, setShowSettings] = useState(false);

  const [formData, setFormData] = useState({
    buyDate: '',
    buyAmount: '',
    sellDate: '',
    sellAmount: '',
  });

  // --- Persistence & Sync ---
  const debouncedSync = useMemo(
    () => debounce(async (data: Trade[]) => {
      const url = localStorage.getItem('gsUrl');
      const secret = localStorage.getItem('gsSecret');
      if (!url) return;
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

        await fetch(url, {
          method: 'POST',
          mode: 'no-cors',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ action: 'sync', data: mappedData, secret: secret }),
        });
        const time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        setLastSynced(time);
        localStorage.setItem('last_synced_time', time);
      } catch (error) {
        console.error("Sheets Sync Error:", error);
      } finally {
        setIsSyncing(false);
      }
    }, 2000),
    []
  );

  const fetchRemoteTrades = React.useCallback(async (url: string, secret?: string) => {
    if (!url) return;
    setIsSyncing(true);
    try {
      const targetUrl = new URL(url);
      const activeSecret = secret || localStorage.getItem('gsSecret');
      if (activeSecret) targetUrl.searchParams.set('secret', activeSecret);
      
      const response = await fetch(targetUrl.toString());
      const data = await response.json();
      
      if (Array.isArray(data)) {
        setTrades(data);
        localStorage.setItem('marginal_trades', JSON.stringify(data));
        const time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        setLastSynced(time);
        localStorage.setItem('last_synced_time', time);
      }
    } catch (error) {
      console.error("Remote Sync Failed:", error);
    } finally {
      setIsSyncing(false);
    }
  }, []);

  useEffect(() => {
    const localTrades = localStorage.getItem('marginal_trades');
    if (localTrades) {
      try {
        setTrades(JSON.parse(localTrades));
      } catch (e) {
        console.error("Failed to parse local trades:", e);
      }
    }

    const savedUrl = localStorage.getItem('gsUrl');
    if (savedUrl) fetchRemoteTrades(savedUrl);

    return () => {
      debouncedSync.cancel();
    };
  }, [fetchRemoteTrades, debouncedSync]);

  const saveLocally = React.useCallback((updatedTrades: Trade[]) => {
    localStorage.setItem('marginal_trades', JSON.stringify(updatedTrades));
    debouncedSync(updatedTrades);
  }, [debouncedSync]);

  const handleAddTrade = React.useCallback((e: React.FormEvent) => {
    e.preventDefault();
    
    const errors: string[] = [];
    const buyPrice = Number(formData.buyAmount);
    const sellPrice = Number(formData.sellAmount);
    
    const hasBuyInput = formData.buyDate.trim() !== '' || formData.buyAmount.trim() !== '';
    const hasSellInput = formData.sellDate.trim() !== '' || formData.sellAmount.trim() !== '';

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
      const newTrade: Trade = { id: crypto.randomUUID(), ...tradeData as Trade };
      updated = [newTrade, ...trades];
    }

    setTrades(updated);
    saveLocally(updated);
    setFormData({ buyDate: '', buyAmount: '', sellDate: '', sellAmount: '' });
  }, [formData, formMode, trades, editingId, saveLocally]);

  const handleEditTrade = React.useCallback((trade: Trade) => {
    setEditingId(trade.id);
    setFormMode(trade.type);
    setFormData({
      buyDate: trade.buyDate || '',
      buyAmount: trade.buyAmount?.toString() || '',
      sellDate: trade.sellDate || '',
      sellAmount: trade.sellAmount?.toString() || '',
    });
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }, []);

  const cancelEdit = React.useCallback(() => {
    setEditingId(null);
    setFormData({ buyDate: '', buyAmount: '', sellDate: '', sellAmount: '' });
  }, []);

  const handleDeleteTrade = React.useCallback((id: string) => {
    const updated = trades.filter(t => t.id !== id);
    setTrades(updated);
    saveLocally(updated);
  }, [trades, saveLocally]);

  const totalNetProfit = useMemo(() => {
    return trades.reduce((acc, t) => acc + ((t.sellAmount || 0) - (t.buyAmount || 0)), 0);
  }, [trades]);

  const dailyReports = useMemo(() => {
    const reports: Record<string, { buy: number, sell: number }> = {};
    trades.forEach(t => {
      if (t.buyDate && t.buyAmount) {
        const d = t.buyDate;
        if (!reports[d]) reports[d] = { buy: 0, sell: 0 };
        reports[d].buy += t.buyAmount;
      }
      if (t.sellDate && t.sellAmount) {
        const d = t.sellDate;
        if (!reports[d]) reports[d] = { buy: 0, sell: 0 };
        reports[d].sell += t.sellAmount;
      }
    });

    return Object.entries(reports)
      .map(([date, values]) => ({ 
        date, 
        buyTotal: values.buy, 
        sellTotal: values.sell,
        netCashflow: values.sell - values.buy 
      }))
      .sort((a, b) => b.date.localeCompare(a.date));
  }, [trades]);

  const handleAiAnalysis = async () => {
    setIsAiAnalyzing(true);
    try {
      const response = await fetch("/api/analyze", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ trades }),
      });
      const data = await response.json();
      if (data.error) {
        setAiAnalysis(data.error);
      } else {
        setAiAnalysis(data.analysis);
      }
    } catch (error) {
      console.error("AI Analysis Failed:", error);
      setAiAnalysis("Unable to perform AI analysis at this moment.");
    } finally {
      setIsAiAnalyzing(false);
    }
  };

  const saveSettings = React.useCallback((url: string, secret: string) => {
    localStorage.setItem('gsUrl', url);
    localStorage.setItem('gsSecret', secret);
    setGsUrl(url);
    setGsSecret(secret);
    if (url) fetchRemoteTrades(url, secret);
    setShowSettings(false);
  }, [fetchRemoteTrades]);


  return (
    <div className="min-h-screen pb-24 selection:bg-[#0066FF]/20 bg-black">
      {/* Header */}
      <header className="sticky top-0 z-50 glass border-b border-white/5 py-3 sm:py-4 px-4 sm:px-8 lg:px-12 flex items-center justify-between transition-all duration-500">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 sm:w-10 sm:h-10 rounded-xl sm:rounded-2xl bg-accent flex items-center justify-center shrink-0 haptic-interaction">
            <TrendingUp className="text-white" size={16} />
          </div>
          <div>
            <h1 className="font-heading font-bold text-base sm:text-xl tracking-tight leading-none text-white uppercase">Marginal</h1>
            <p className="text-[8px] sm:text-[9px] font-mono text-slate-500 mt-0.5 sm:mt-1 uppercase tracking-[0.2em] hidden xs:block">Alpha Engine v2</p>
          </div>
        </div>

          <div className="flex items-center gap-3">
          <div className="flex flex-col items-end">
          </div>
        </div>
      </header>

      {/* Settings Modal */}
      <AnimatePresence>
        {showSettings && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-md"
          >
            <motion.div 
              initial={{ scale: 0.9, y: 20 }}
              animate={{ scale: 1, y: 0 }}
              className="glass-card w-full max-w-md overflow-hidden flex flex-col bg-[#0a0a0a]"
            >
              <div className="p-6 border-b border-white/5 flex justify-between items-center bg-white/5">
                <div>
                  <h3 className="font-heading font-bold text-lg">Sync Settings</h3>
                  <p className="text-xs text-slate-500">Configure your Google Sheets integration</p>
                </div>
                <button onClick={() => setShowSettings(false)} className="p-2 hover:bg-white/10 rounded-full">
                  <XCircle size={20} />
                </button>
              </div>
              
              <div className="p-6 space-y-6">
                <div className="space-y-2">
                  <label className="text-[10px] uppercase font-bold tracking-widest text-slate-500">Script Web App URL</label>
                  <input 
                    type="text" 
                    value={gsUrl}
                    onChange={(e) => setGsUrl(e.target.value)}
                    placeholder="https://script.google.com/macros/s/.../exec"
                    className="w-full bg-black border border-white/10 rounded-xl px-4 py-3 text-xs font-mono focus:outline-hidden focus:border-accent transition-all"
                  />
                </div>
                
                <div className="space-y-2">
                  <label className="text-[10px] uppercase font-bold tracking-widest text-slate-500">Security Secret Key</label>
                  <input 
                    type="password" 
                    value={gsSecret}
                    onChange={(e) => setGsSecret(e.target.value)}
                    placeholder="Same as SECRET_KEY in your macro"
                    className="w-full bg-black border border-white/10 rounded-xl px-4 py-3 text-xs font-mono focus:outline-hidden focus:border-accent transition-all"
                  />
                  <p className="text-[9px] text-slate-500 leading-relaxed">
                    Used to authenticate your app with your private Google Sheet script.
                  </p>
                </div>

                <div className="pt-2">
                  <button 
                    onClick={() => setShowMacroCode(true)}
                    className="text-[10px] uppercase font-bold text-accent hover:underline flex items-center gap-2"
                  >
                    View / Copy Macro Code
                  </button>
                </div>
              </div>

              <div className="p-6 border-t border-white/5 bg-white/5 flex gap-4">
                <button 
                  onClick={() => saveSettings(gsUrl, gsSecret)}
                  className="flex-1 bg-accent py-3 rounded-xl font-bold text-xs uppercase tracking-widest hover:bg-accent/80 transition-all text-white"
                >
                  Save Settings
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

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
                  {MACRO_CODE}
                </pre>
              </div>
              <div className="p-6 border-t border-white/5 flex gap-4">
                <button 
                  onClick={() => {
                    navigator.clipboard.writeText(MACRO_CODE);
                    alert("Copied to clipboard!");
                  }}
                  className="flex-1 bg-accent py-3 rounded-xl font-bold text-xs uppercase tracking-widest hover:bg-accent/80 transition-all text-white"
                >
                  Copy Code
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 sm:py-8 md:py-12">
        <AnimatePresence mode="wait">
          {activeView === 'dashboard' ? (
            <motion.div 
              key="dashboard"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="grid grid-cols-1 lg:grid-cols-[380px_1fr] xl:grid-cols-[420px_1fr] gap-6 lg:gap-8 items-start"
            >
              {/* Left Column */}
              <div className="space-y-4">
                {/* Hero Card - Compact */}
                <div className="glass-card p-6 sm:p-5 flex flex-col sm:flex-row items-start sm:items-center justify-between relative group overflow-hidden bg-[#0a0a0a] gap-4">
                  <div className="flex flex-col">
                    <h2 className="text-[10px] font-bold uppercase tracking-[0.2em] text-slate-500 mb-2 sm:mb-1">
                      Portfolio Valuation
                    </h2>
                    <div className={`text-5xl sm:text-4xl font-sans font-light tabular-nums tracking-tight leading-none ${
                        totalNetProfit >= 0 ? 'text-white' : 'text-loss'
                    }`}>
                      <span className="opacity-80 mr-1 text-2xl sm:text-2xl font-light text-slate-400">₹</span>
                      {totalNetProfit.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}
                    </div>
                  </div>

                  <div className="flex flex-row sm:flex-col items-center sm:items-end justify-between w-full sm:w-auto gap-2">
                    <div className="flex items-center gap-2 bg-[#111111] px-4 py-2 sm:px-3 sm:py-1.5 rounded-full border border-white/10 shadow-inner">
                      <div className={`w-1.5 h-1.5 rounded-full ${
                        totalNetProfit > 0 ? 'bg-profit shadow-[0_0_8px_#10b981]' : totalNetProfit < 0 ? 'bg-loss animate-pulse shadow-[0_0_8px_#ef4444]' : 'bg-slate-500'
                      }`} />
                      <span className="text-[10px] sm:text-[9px] uppercase font-bold tracking-widest text-[#e5e5e5]">
                        {totalNetProfit > 0 ? 'Bullish' : totalNetProfit < 0 ? 'Bearish' : 'Neutral'}
                      </span>
                    </div>
                    {lastSynced && (
                      <div className="flex items-center gap-1.5 text-[9px] sm:text-[8px] uppercase font-bold tracking-[0.1em] text-slate-600">
                        <Clock size={10} /> {lastSynced}
                      </div>
                    )}
                  </div>
                  
                  <div className="absolute top-2 right-2 opacity-0 group-hover:opacity-100 transition-opacity">
                    <button 
                      onClick={handleAiAnalysis}
                      disabled={isAiAnalyzing}
                      className="p-1.5 bg-[#111111] hover:bg-white/10 rounded-lg text-slate-400 hover:text-white transition-all border border-white/10 haptic-interaction"
                    >
                      <BrainCircuit size={14} className={isAiAnalyzing ? 'animate-spin' : ''} />
                    </button>
                  </div>
                </div>
                
                {aiAnalysis && (
                    <motion.div 
                      initial={{ opacity: 0, y: -10 }}
                      animate={{ opacity: 1, y: 0 }}
                      className="p-4 bg-[#0a0a0a] rounded-xl border border-white/10 relative"
                    >
                       <div className="absolute -top-2 left-4 px-2 py-0.5 bg-accent text-[8px] font-semibold text-white uppercase tracking-widest rounded-sm">
                         Insight
                       </div>
                       <p className="text-xs text-slate-300 leading-relaxed italic">
                         "{aiAnalysis}"
                       </p>
                    </motion.div>
                )}

                <div className={`glass-card p-6 sm:p-5 transition-all duration-500 ${editingId ? 'ring-2 ring-accent bg-[#111111]' : 'bg-[#0a0a0a]'}`}>
                  <div className="flex justify-between items-center mb-6 sm:mb-4">
                    <h3 className="font-heading font-bold text-base sm:text-sm flex items-center gap-2 uppercase tracking-widest leading-none">
                       {editingId ? <Pencil className="text-accent" size={16} /> : <Plus className="text-accent" size={16} />}
                       {editingId ? 'Update Entry' : 'New Transaction'}
                    </h3>
                    
                    {editingId && (
                       <button 
                        onClick={cancelEdit}
                        className="flex items-center gap-2 text-[10px] uppercase font-bold text-slate-500 hover:text-loss transition-colors px-3 py-1.5 bg-white/5 rounded-lg haptic-interaction"
                       >
                         <XCircle size={14} /> Cancel
                       </button>
                    )}
                  </div>

                  <form onSubmit={handleAddTrade} className="space-y-6 sm:space-y-4" noValidate>
                    <div className="space-y-6 sm:space-y-4">
                      {/* Buy Section */}
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 flex-wrap">
                        <div className="relative group/field flex-1">
                          <input 
                            type="date" 
                            value={formData.buyDate}
                            onChange={e => {
                              setFormData({...formData, buyDate: e.target.value});
                              setValidationErrors(prev => prev.filter(err => err !== 'buyDate'));
                            }}
                            className={`peer w-full bg-black border rounded-2xl sm:rounded-xl px-4 py-3 sm:py-1.5 pt-7 sm:pt-5 text-sm focus:outline-hidden focus:ring-2 transition-all font-mono tracking-tight appearance-none ${
                              validationErrors.includes('buyDate') ? 'border-loss focus:border-loss ring-loss/10' : 'border-white/10 focus:border-accent focus:ring-accent/20'
                            }`}
                          />
                          <label className={`absolute left-4 top-3 sm:top-2 text-[10px] uppercase font-bold tracking-[0.2em] transition-all leading-none ${
                            validationErrors.includes('buyDate') ? 'text-loss' : 'text-slate-500 peer-focus:text-accent'
                          }`}>
                            Entry Date
                          </label>
                        </div>
                        <div className="relative group/field flex-1">
                          <input 
                            type="number" 
                            placeholder=" "
                            step="0.01"
                            inputMode="decimal"
                            value={formData.buyAmount}
                            onChange={e => {
                              setFormData({...formData, buyAmount: e.target.value});
                              setValidationErrors(prev => prev.filter(err => err !== 'buyAmount'));
                            }}
                            className={`peer w-full bg-black border rounded-2xl sm:rounded-xl px-4 py-3 sm:py-1.5 pt-7 sm:pt-5 text-base sm:text-sm focus:outline-hidden focus:ring-2 transition-all font-mono placeholder:opacity-0 ${
                              validationErrors.includes('buyAmount') ? 'border-loss focus:border-loss ring-loss/10' : 'border-white/10 focus:border-accent focus:ring-accent/20'
                            }`}
                          />
                          <label className={`absolute left-4 top-3 sm:top-2 text-[10px] uppercase font-bold tracking-[0.2em] transition-all peer-placeholder-shown:top-5 sm:peer-placeholder-shown:top-4 peer-placeholder-shown:text-base sm:peer-placeholder-shown:text-sm peer-focus:top-3 sm:peer-focus:top-2 peer-focus:text-[10px] leading-none ${
                            validationErrors.includes('buyAmount') ? 'text-loss' : 'text-slate-500 peer-focus:text-accent'
                          }`}>
                            Entry Price (₹)
                          </label>
                        </div>
                      </div>

                      {/* Sell Section */}
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <div className="relative group/field flex-1">
                          <input 
                            type="date" 
                            value={formData.sellDate}
                            onChange={e => {
                              setFormData({...formData, sellDate: e.target.value});
                              setValidationErrors(prev => prev.filter(err => err !== 'sellDate'));
                            }}
                            className={`peer w-full bg-black border rounded-2xl sm:rounded-xl px-4 py-3 sm:py-1.5 pt-7 sm:pt-5 text-sm focus:outline-hidden focus:ring-2 transition-all font-mono tracking-tight appearance-none ${
                              validationErrors.includes('sellDate') ? 'border-loss focus:border-loss ring-loss/10' : 'border-white/10 focus:border-accent focus:ring-accent/20'
                            }`}
                          />
                          <label className={`absolute left-4 top-3 sm:top-2 text-[10px] uppercase font-bold tracking-[0.2em] transition-all leading-none ${
                            validationErrors.includes('sellDate') ? 'text-loss' : 'text-slate-500 peer-focus:text-accent'
                          }`}>
                            Exit Date
                          </label>
                        </div>
                        <div className="relative group/field flex-1">
                          <input 
                            type="number" 
                            placeholder=" "
                            step="0.01"
                            inputMode="decimal"
                            value={formData.sellAmount}
                            onChange={e => {
                              setFormData({...formData, sellAmount: e.target.value});
                              setValidationErrors(prev => prev.filter(err => err !== 'sellAmount'));
                            }}
                            className={`peer w-full bg-black border rounded-2xl sm:rounded-xl px-4 py-3 sm:py-1.5 pt-7 sm:pt-5 text-base sm:text-sm focus:outline-hidden focus:ring-2 transition-all font-mono placeholder:opacity-0 ${
                              validationErrors.includes('sellAmount') ? 'border-loss focus:border-loss ring-loss/10' : 'border-white/10 focus:border-accent focus:ring-accent/20'
                            }`}
                          />
                          <label className={`absolute left-4 top-3 sm:top-2 text-[10px] uppercase font-bold tracking-[0.2em] transition-all peer-placeholder-shown:top-5 sm:peer-placeholder-shown:top-4 peer-placeholder-shown:text-base sm:peer-placeholder-shown:text-sm peer-focus:top-3 sm:peer-focus:top-2 peer-focus:text-[10px] leading-none ${
                            validationErrors.includes('sellAmount') ? 'text-loss' : 'text-slate-500 peer-focus:text-accent'
                          }`}>
                            Exit Price (₹)
                          </label>
                        </div>
                      </div>
                    </div>

                    <button 
                      type="submit"
                      className={`w-full py-4 sm:py-3.5 rounded-2xl sm:rounded-xl font-bold uppercase tracking-[0.3em] text-xs transition-all shadow-lg active:scale-[0.98] haptic-interaction flex items-center justify-center gap-3 ${
                        editingId ? 'bg-white text-black hover:bg-slate-200' : 'bg-[#0066FF] text-white hover:bg-[#0055DD]'
                      }`}
                    >
                      {editingId ? <RefreshCw size={18} strokeWidth={2.5} /> : <Plus size={18} strokeWidth={2.5} />}
                      {editingId ? 'Update Deal' : `${formMode === 'pair' ? 'Add Deal' : formMode === 'buy' ? 'Add Buy Entry' : 'Add Sell Entry'}`}
                    </button>
                  </form>
                </div>
              </div>

              {/* Right Column - Ledger */}
              <div className="glass-card min-h-[400px] flex flex-col overflow-hidden bg-[#000000]">
                <div className="p-4 sm:p-5 border-b border-white/10 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 bg-[#0a0a0a]">
                   <div className="flex items-center gap-3">
                    <h3 className="font-heading font-semibold text-base flex items-center gap-2 uppercase tracking-widest">
                      <Table className="text-accent" size={16} />
                      Transactions
                    </h3>
                    <div className="flex bg-black p-0.5 rounded-lg border border-white/5">
                      <button 
                        onClick={() => setIsGroupedByDate(true)}
                        className={`px-3 py-1 rounded-md text-[8px] uppercase font-bold tracking-widest transition-all ${isGroupedByDate ? 'bg-white/10 text-white shadow-lg' : 'text-slate-500 hover:text-slate-300'}`}
                      >
                        Daily
                      </button>
                      <button 
                        onClick={() => setIsGroupedByDate(false)}
                        className={`px-3 py-1 rounded-md text-[8px] uppercase font-bold tracking-widest transition-all ${!isGroupedByDate ? 'bg-white/10 text-white shadow-lg' : 'text-slate-500 hover:text-slate-300'}`}
                      >
                        Raw
                      </button>
                    </div>
                  </div>
                  <span className="text-[9px] font-mono text-slate-500 uppercase tracking-widest">
                    {isGroupedByDate ? `${dailyReports.length} Dates` : `${trades.length} Records`}
                  </span>
                </div>
                
                <div className="flex-1 p-3 sm:p-4">
                  {/* Desktop Table View */}
                  <div className="hidden sm:block overflow-x-auto scroll-hide">
                    <table className="w-full text-left border-separate border-spacing-y-2">
                      <thead>
                        <tr>
                          <th className="px-4 pb-2 text-[9px] uppercase tracking-wider font-semibold text-slate-600">
                            {isGroupedByDate ? 'Aggregate Date' : 'Entry'}
                          </th>
                          <th className="px-4 pb-2 text-[9px] uppercase tracking-wider font-semibold text-slate-600">
                            {isGroupedByDate ? 'Daily Buy Total' : 'Exit'}
                          </th>
                          <th className="px-4 pb-2 text-[9px] uppercase tracking-wider font-semibold text-slate-600">
                            {isGroupedByDate ? 'Daily Sell Total' : 'Net Return'}
                          </th>
                          <th className={`px-4 pb-2 text-right text-[9px] uppercase tracking-wider font-semibold text-slate-600 ${isGroupedByDate ? '' : 'pr-4'}`}>
                            {isGroupedByDate ? 'Day P&L' : 'Actions'}
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {isGroupedByDate ? (
                          // Grouped Desktop View
                          <>
                            {dailyReports.length === 0 && (
                              <tr><td colSpan={4} className="px-4 py-20 text-center text-slate-500 font-light italic">No summaries available.</td></tr>
                            )}
                            {dailyReports.map((report) => (
                              <tr key={report.date} className="bg-[#0a0a0a] hover:bg-white/5 transition-all">
                                <td className="px-4 py-3 rounded-l-lg border-y border-l border-white/5">
                                  <span className="text-xs font-mono font-medium text-slate-300">{formatDate(report.date)}</span>
                                </td>
                                <td className="px-4 py-3 border-y border-white/5 text-sm font-medium text-white">
                                  ₹{report.buyTotal.toLocaleString()}
                                </td>
                                <td className="px-4 py-3 border-y border-white/5 text-sm font-medium text-white">
                                  ₹{report.sellTotal.toLocaleString()}
                                </td>
                                <td className={`px-4 py-3 border-y border-r border-white/5 text-right rounded-r-lg font-mono text-sm font-bold ${report.netCashflow >= 0 ? 'text-profit' : 'text-loss'}`}>
                                  {report.netCashflow >= 0 ? '+' : ''}₹{Math.abs(report.netCashflow).toLocaleString()}
                                </td>
                              </tr>
                            ))}
                          </>
                        ) : (
                          // Raw Desktop View (Existing)
                          <>
                            {trades.length === 0 && (
                              <tr><td colSpan={4} className="px-4 py-20 text-center text-slate-500 font-light italic">No records found.</td></tr>
                            )}
                            {trades.map((trade) => (
                              <TradeRow 
                                key={trade.id} 
                                trade={trade} 
                                onEdit={handleEditTrade} 
                                onDelete={handleDeleteTrade} 
                              />
                            ))}
                          </>
                        )}
                      </tbody>
                    </table>
                  <div className="sm:hidden space-y-3">
                      {isGroupedByDate ? (
                        dailyReports.map((report) => (
                          <motion.div 
                           key={report.date}
                           layout
                           className="bg-[#0a0a0a] p-4 rounded-2xl border border-white/5 relative overflow-hidden active:scale-[0.98] transition-transform"
                          >
                             <div className={`absolute top-0 left-0 w-1 h-full ${report.netCashflow >= 0 ? 'bg-profit/40' : 'bg-loss/40'}`} />
                             <div className="flex justify-between items-center mb-3 pl-2">
                               <div className="flex flex-col">
                                 <span className="text-[9px] uppercase font-bold text-slate-500 tracking-widest leading-none mb-1">Session Date</span>
                                 <span className="text-sm font-mono font-medium text-slate-200 leading-none">{formatDate(report.date)}</span>
                               </div>
                               <div className="text-right">
                                 <span className="text-[9px] uppercase font-bold text-slate-500 tracking-widest leading-none mb-1 block">Net Profit</span>
                                 <span className={`text-base font-bold font-mono leading-none ${report.netCashflow >= 0 ? 'text-profit' : 'text-loss'}`}>
                                   {report.netCashflow >= 0 ? '+' : ''}₹{Math.abs(report.netCashflow).toLocaleString()}
                                 </span>
                               </div>
                             </div>
                             <div className="flex justify-between pl-2 pt-2 border-t border-white/5">
                               <div className="flex flex-col">
                                 <span className="text-[8px] uppercase font-bold text-slate-600 tracking-tighter">Buy Total</span>
                                 <span className="text-xs font-mono text-white">₹{report.buyTotal.toLocaleString()}</span>
                               </div>
                               <div className="flex flex-col items-end">
                                 <span className="text-[8px] uppercase font-bold text-slate-600 tracking-tighter">Sell Total</span>
                                 <span className="text-xs font-mono text-white">₹{report.sellTotal.toLocaleString()}</span>
                               </div>
                             </div>
                          </motion.div>
                        ))
                      ) : (
                        trades.map((trade) => (
                          <MobileTradeCard 
                            key={trade.id} 
                            trade={trade} 
                            onEdit={handleEditTrade} 
                            onDelete={handleDeleteTrade} 
                          />
                        ))
                      )}
                  </div>
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
              className="max-w-4xl mx-auto"
            >
              <div className="glass-card overflow-hidden">
                 <div className="p-6 sm:p-8 border-b border-white/5 flex justify-between items-center bg-[#0a0a0a]">
                   <h3 className="font-heading font-bold text-base sm:text-lg flex items-center gap-2 uppercase tracking-widest">
                    <CalendarDays className="text-accent" size={20} />
                    Daily Performance
                  </h3>
                </div>
                
                {/* Desktop View */}
                <div className="hidden sm:block overflow-hidden">
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
                          <td colSpan={2} className="px-8 py-24 text-center text-slate-500 font-light italic">
                            Insufficient data for daily reports.
                          </td>
                        </tr>
                      )}
                      {dailyReports.map((report, idx) => (
                        <tr key={idx} className="hover:bg-white/[0.02] transition-colors">
                          <td className="px-8 py-5 text-sm font-medium text-slate-300">{formatDate(report.date)}</td>
                          <td className="px-8 py-5 text-sm font-mono text-white text-right">
                             <div className="flex flex-col items-end gap-1">
                                <span className={`font-bold ${report.netCashflow >= 0 ? 'text-profit' : 'text-loss'}`}>
                                  {report.netCashflow >= 0 ? '+' : ''}₹{Math.abs(report.netCashflow).toLocaleString(undefined, { minimumFractionDigits: 2 })}
                                </span>
                                <div className="flex gap-4 text-[10px] opacity-40 uppercase tracking-tighter">
                                   <span>In: ₹{report.buyTotal.toLocaleString()}</span>
                                   <span>Out: ₹{report.sellTotal.toLocaleString()}</span>
                                </div>
                             </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* Mobile Card View */}
                <div className="sm:hidden p-4 space-y-4 bg-black">
                  {dailyReports.length === 0 && (
                    <div className="py-24 text-center text-slate-500 font-light italic text-sm">
                      Insufficient data for daily reports.
                    </div>
                  )}
                  {dailyReports.map((report, idx) => (
                    <motion.div 
                      key={idx} 
                      initial={{ opacity: 0, scale: 0.95 }}
                      animate={{ opacity: 1, scale: 1 }}
                      className="bg-[#0a0a0a] p-5 rounded-2xl border border-white/5 flex flex-col gap-4 relative overflow-hidden active:scale-[0.98] transition-transform"
                    >
                      <div className={`absolute top-0 left-0 w-1.5 h-full ${report.netCashflow >= 0 ? 'bg-profit/40' : 'bg-loss/40'}`} />
                      <div className="flex justify-between items-start pl-2">
                        <div className="flex flex-col">
                          <span className="text-[10px] uppercase font-bold text-slate-500 tracking-[0.2em] mb-1">Trade Date</span>
                          <span className="text-base font-mono font-medium text-slate-200">{formatDate(report.date)}</span>
                        </div>
                        <div className="text-right">
                          <span className="text-[10px] uppercase font-bold text-slate-500 tracking-[0.2em] mb-1 block">Day P&L</span>
                          <span className={`text-xl font-bold font-mono ${report.netCashflow >= 0 ? 'text-profit' : 'text-loss'}`}>
                            {report.netCashflow >= 0 ? '+' : ''}₹{Math.abs(report.netCashflow).toLocaleString()}
                          </span>
                        </div>
                      </div>
                      <div className="grid grid-cols-2 gap-4 pl-2 pt-4 border-t border-white/5">
                        <div className="flex flex-col">
                          <span className="text-[9px] uppercase font-bold text-slate-600 tracking-widest mb-1">Buy Volume</span>
                          <span className="text-sm font-mono text-white">₹{report.buyTotal.toLocaleString()}</span>
                        </div>
                        <div className="flex flex-col items-end">
                          <span className="text-[9px] uppercase font-bold text-slate-600 tracking-widest mb-1">Sell Volume</span>
                          <span className="text-sm font-mono text-white">₹{report.sellTotal.toLocaleString()}</span>
                        </div>
                      </div>
                    </motion.div>
                  ))}
                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Persistent Footer Nav */}
      <nav className="fixed bottom-6 left-0 right-0 z-50 flex justify-center pointer-events-none">
        <div className="glass shadow-[0_20px_50px_rgba(0,0,0,0.8)] rounded-full border border-white/10 flex gap-1 p-1 pointer-events-auto scale-90 sm:scale-100">
          <NavTab id="dashboard" label="Home" icon={LayoutDashboard} activeView={activeView} onClick={setActiveView} />
          <NavTab id="reports" label="Ledger" icon={Table} activeView={activeView} onClick={setActiveView} />
          <button
            onClick={() => setShowSettings(true)}
            className="flex flex-col items-center gap-1 px-6 pt-2.5 pb-2 transition-all duration-300 relative haptic-interaction text-slate-500 hover:text-white"
          >
            <div className="p-1.5 rounded-xl bg-transparent">
              <Settings2 size={20} strokeWidth={2} />
            </div>
            <span className="text-[8px] font-black uppercase tracking-[0.2em] opacity-40">Setup</span>
          </button>
        </div>
      </nav>
    </div>
  );
}
