/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useEffect, useMemo } from 'react';
import { 
  PlusCircle, 
  LayoutDashboard, 
  CalendarDays, 
  Trash2, 
  RefreshCw, 
  TrendingUp, 
  TrendingDown,
  BrainCircuit,
  Settings2,
  Table as TableIcon,
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

  // Form State
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
        setTrades(data);
        localStorage.setItem('marginal_trades', JSON.stringify(data));
        setLastSynced(new Date().toLocaleTimeString());
      }
    } catch (error) {
      console.error("Remote Sync Failed:", error);
    } finally {
      setIsSyncing(false);
    }
  };

  const syncToSheets = async (data: Trade[]) => {
    if (!gsUrl) return;
    setIsSyncing(true);
    try {
      await fetch(gsUrl, {
        method: 'POST',
        mode: 'no-cors', // Apps Script often requires no-cors for simple posts
        body: JSON.stringify({ action: 'sync', data }),
      });
      setLastSynced(new Date().toLocaleTimeString());
    } catch (error) {
      console.error("Sheets Sync Error:", error);
    } finally {
      setIsSyncing(false);
    }
  };

  // --- Logic ---
  const handleAddTrade = (e: React.FormEvent) => {
    e.preventDefault();
    
    const buyPrice = Number(formData.buyAmount);
    const sellPrice = Number(formData.sellAmount);

    if (formMode === 'pair') {
      if (!formData.buyDate || !buyPrice || !formData.sellDate || !sellPrice) return;
    } else if (formMode === 'buy') {
      if (!formData.buyDate || !buyPrice) return;
    } else if (formMode === 'sell') {
      if (!formData.sellDate || !sellPrice) return;
    }

    const tradeData: Partial<Trade> = {
      type: formMode,
      buyDate: formMode !== 'sell' ? formData.buyDate : undefined,
      buyAmount: formMode !== 'sell' ? buyPrice : 0,
      sellDate: formMode !== 'buy' ? formData.sellDate : undefined,
      sellAmount: formMode !== 'buy' ? sellPrice : 0,
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
      className={`flex flex-col items-center gap-1 px-6 py-2 transition-all duration-300 relative ${
        activeView === id ? 'text-accent' : 'text-slate-500 hover:text-slate-300'
      }`}
    >
      <Icon size={20} />
      <span className="text-[10px] font-medium uppercase tracking-widest">{label}</span>
      {activeView === id && (
        <motion.div
           layoutId="nav-pill"
           className="absolute -top-1 w-1 h-1 bg-accent rounded-full shadow-[0_0_8px_#6366f1]"
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
      <header className="sticky top-0 z-50 glass border-b border-white/5 py-4 px-6 md:px-12 flex justify-between items-center transition-all duration-500">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-accent flex items-center justify-center shadow-lg shadow-accent/20">
            <TrendingUp className="text-white" size={20} />
          </div>
          <div>
            <h1 className="font-heading font-bold text-xl tracking-tight leading-none">MARGINAL</h1>
            <p className="text-[10px] font-mono text-slate-500 mt-1 uppercase tracking-widest">Trade Intelligence</p>
          </div>
        </div>

        <div className="flex items-center gap-4">
          <div className="flex flex-col items-end">
            <div className="flex items-center gap-2 glass-card px-3 py-1.5 border-white/5">
              <div className={`w-2 h-2 rounded-full ${isSyncing ? 'bg-amber-400 animate-pulse shadow-[0_0_8px_#fbbf24]' : gsUrl ? 'bg-profit shadow-[0_0_8px_#10b981]' : 'bg-slate-600'}`} />
              <span className="text-[10px] uppercase font-semibold text-slate-400 tracking-wider">
                {isSyncing ? 'Synchronizing' : gsUrl ? 'Cloud Connected' : 'Local Storage'}
              </span>
              {gsUrl && (
                <button 
                  onClick={() => fetchRemoteTrades(gsUrl)}
                  disabled={isSyncing}
                  className="ml-2 p-1 hover:bg-white/10 rounded transition-colors text-slate-500 hover:text-white disabled:opacity-50"
                  title="Manual Refresh"
                >
                  <RefreshCw size={12} className={isSyncing ? 'animate-spin' : ''} />
                </button>
              )}
            </div>
            {lastSynced && (
              <span className="text-[8px] text-slate-500 uppercase tracking-tighter mt-1">Last sync: {lastSynced}</span>
            )}
          </div>
          
          <button 
            onClick={() => {
              const url = prompt("Enter Google Apps Script Web App URL:", gsUrl);
              if (url !== null) {
                setGsUrl(url);
                localStorage.setItem('gsUrl', url);
                if (url) fetchRemoteTrades(url);
              }
            }}
            className="p-2 hover:bg-white/5 rounded-full transition-colors text-slate-400 hover:text-white"
            title="Settings"
          >
            <Settings2 size={18} />
          </button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto p-6 md:p-12">
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
              <div className="space-y-8">
                {/* Hero Card */}
                <div className="glass-card p-10 flex flex-col items-center justify-center text-center relative group min-h-48">
                  <div className="absolute top-4 right-4 opacity-0 group-hover:opacity-100 transition-opacity">
                    <button 
                      onClick={handleAiAnalysis}
                      disabled={isAiAnalyzing}
                      className="p-2 bg-white/5 hover:bg-accent rounded-lg text-slate-400 hover:text-white transition-all flex items-center gap-2 group/ai btn-glow"
                    >
                      <BrainCircuit size={16} className={isAiAnalyzing ? 'animate-spin' : ''} />
                      <span className="text-[10px] uppercase font-bold tracking-tighter">Analyze</span>
                    </button>
                  </div>

                  <p className="text-[10px] uppercase tracking-[0.3em] text-slate-400 mb-2">Total Net Profit</p>
                  <h2 className={`text-6xl font-heading font-black tabular-nums transition-all duration-500 ${
                    totalNetProfit >= 0 ? 'gradient-text-profit' : 'text-loss drop-shadow-[0_0_10px_rgba(239,68,68,0.3)]'
                  }`}>
                    ${totalNetProfit.toLocaleString(undefined, { minimumFractionDigits: 2 })}
                  </h2>
                  
                  {totalNetProfit > 0 && (
                    <div className="mt-4 flex items-center gap-2 text-xs text-profit font-medium">
                      <span className="bg-profit/20 px-2 py-0.5 rounded">Bullish</span>
                      <span>Trending Up</span>
                    </div>
                  )}
                  
                  {aiAnalysis && (
                    <motion.div 
                      initial={{ opacity: 0, height: 0 }}
                      animate={{ opacity: 1, height: 'auto' }}
                      className="mt-8 p-4 bg-white/5 rounded-xl text-left border border-white/5"
                    >
                       <p className="text-[10px] uppercase font-bold text-accent mb-2 flex items-center gap-2">
                         <BrainCircuit size={12} /> AI Insights
                       </p>
                       <p className="text-xs text-slate-400 leading-relaxed font-light italic">
                         "{aiAnalysis}"
                       </p>
                    </motion.div>
                  )}
                </div>

                {/* Form Card */}
                <div className={`glass-card p-8 transition-all duration-500 ${editingId ? 'ring-2 ring-accent ring-inset bg-accent/5' : ''}`}>
                  <div className="flex justify-between items-center mb-8">
                    <h3 className="font-heading font-bold text-lg flex items-center gap-2">
                       {editingId ? <Pencil className="text-accent" size={20} /> : <PlusCircle className="text-accent" size={20} />}
                       {editingId ? 'Edit Transaction' : 'New Transaction'}
                    </h3>
                    
                    {!editingId && (
                      <div className="flex bg-white/5 p-1 rounded-lg border border-white/5">
                        {(['pair', 'buy', 'sell'] as const).map((mode) => (
                          <button
                            key={mode}
                            onClick={() => setFormMode(mode)}
                            className={`px-3 py-1 text-[10px] uppercase font-bold rounded-md transition-all ${
                              formMode === mode ? 'bg-accent text-white shadow-lg' : 'text-slate-500 hover:text-slate-300'
                            }`}
                          >
                            {mode === 'pair' ? 'Trade' : mode}
                          </button>
                        ))}
                      </div>
                    )}

                    {editingId && (
                       <button 
                        onClick={cancelEdit}
                        className="flex items-center gap-2 text-[10px] uppercase font-bold text-slate-500 hover:text-loss transition-colors"
                       >
                         <XCircle size={14} /> Cancel
                       </button>
                    )}
                  </div>

                  <form onSubmit={handleAddTrade} className="space-y-6">
                    <AnimatePresence mode="wait">
                      <motion.div 
                        key={formMode}
                        initial={{ opacity: 0, x: 10 }}
                        animate={{ opacity: 1, x: 0 }}
                        exit={{ opacity: 0, x: -10 }}
                        transition={{ duration: 0.2 }}
                        className="space-y-6"
                      >
                        {(formMode === 'pair' || formMode === 'buy') && (
                          <div className="grid grid-cols-2 gap-4">
                            {/* Buy Date */}
                            <div className="relative group/field">
                              <input 
                                type="date" 
                                value={formData.buyDate}
                                onChange={e => setFormData({...formData, buyDate: e.target.value})}
                                className="peer w-full bg-white/5 border border-white/10 rounded-xl px-4 py-3 pt-6 text-sm focus:outline-hidden focus:border-accent/50 focus:ring-1 focus:ring-accent/50 transition-all font-mono"
                              />
                              <label className="absolute left-4 top-1 text-[9px] uppercase font-bold tracking-widest text-slate-500 transition-all peer-focus:text-accent">
                                Buy Date
                              </label>
                            </div>
                            {/* Buy Amount */}
                            <div className="relative group/field">
                              <input 
                                type="number" 
                                placeholder=" "
                                step="0.01"
                                value={formData.buyAmount}
                                onChange={e => setFormData({...formData, buyAmount: e.target.value})}
                                className="peer w-full bg-white/5 border border-white/10 rounded-xl px-4 py-3 pt-6 text-sm focus:outline-hidden focus:border-accent/50 focus:ring-1 focus:ring-accent/50 transition-all font-mono placeholder:opacity-0"
                              />
                              <label className="absolute left-4 top-1 text-[9px] uppercase font-bold tracking-widest text-slate-500 transition-all peer-focus:text-accent peer-placeholder-shown:top-4 peer-placeholder-shown:text-xs peer-focus:top-1 peer-focus:text-[9px]">
                                Buy Price ($)
                              </label>
                            </div>
                          </div>
                        )}

                        {(formMode === 'pair' || formMode === 'sell') && (
                          <div className="grid grid-cols-2 gap-4">
                            {/* Sell Date */}
                            <div className="relative group/field">
                              <input 
                                type="date" 
                                value={formData.sellDate}
                                onChange={e => setFormData({...formData, sellDate: e.target.value})}
                                className="peer w-full bg-white/5 border border-white/10 rounded-xl px-4 py-3 pt-6 text-sm focus:outline-hidden focus:border-accent/50 focus:ring-1 focus:ring-accent/50 transition-all font-mono"
                              />
                              <label className="absolute left-4 top-1 text-[9px] uppercase font-bold tracking-widest text-slate-500 transition-all peer-focus:text-accent">
                                Sell Date
                              </label>
                            </div>
                            {/* Sell Amount */}
                            <div className="relative group/field">
                              <input 
                                type="number" 
                                placeholder=" "
                                step="0.01"
                                value={formData.sellAmount}
                                onChange={e => setFormData({...formData, sellAmount: e.target.value})}
                                className="peer w-full bg-white/5 border border-white/10 rounded-xl px-4 py-3 pt-6 text-sm focus:outline-hidden focus:border-accent/50 focus:ring-1 focus:ring-accent/50 transition-all font-mono placeholder:opacity-0"
                              />
                              <label className="absolute left-4 top-1 text-[9px] uppercase font-bold tracking-widest text-slate-500 transition-all peer-focus:text-accent peer-placeholder-shown:top-4 peer-placeholder-shown:text-xs peer-focus:top-1 peer-focus:text-[9px]">
                                Sell Price ($)
                              </label>
                            </div>
                          </div>
                        )}
                      </motion.div>
                    </AnimatePresence>
                    
                    <button 
                      type="submit"
                      className="w-full bg-accent hover:bg-accent/80 text-white font-bold py-4 rounded-xl shadow-lg shadow-accent/20 transition-all active:scale-[0.98] mt-4 flex items-center justify-center gap-2 uppercase tracking-widest text-xs group btn-glow"
                    >
                      {editingId ? <RefreshCw size={18} className="group-hover:rotate-180 transition-transform" /> : <PlusCircle size={18} className="group-hover:rotate-90 transition-transform" />}
                      {editingId ? 'Update Transaction' : `Add ${formMode === 'pair' ? 'Deal' : formMode === 'buy' ? 'Buy Entry' : 'Sell Entry'}`}
                    </button>
                  </form>
                </div>
              </div>

              {/* Right Column - Ledger */}
              <div className="glass-card min-h-[500px] flex flex-col">
                <div className="p-8 border-b border-white/5 flex justify-between items-center">
                   <h3 className="font-heading font-bold text-lg flex items-center gap-2">
                    <TableIcon className="text-accent" size={20} />
                    Transaction Ledger
                  </h3>
                  <span className="text-[10px] font-mono text-slate-500 uppercase tracking-widest">{trades.length} Records</span>
                </div>
                
                <div className="overflow-x-auto flex-1 p-6 scroll-hide">
                  <table className="w-full text-left border-separate border-spacing-y-3">
                    <thead className="sticky top-0 bg-background/80 backdrop-blur-sm z-20">
                      <tr>
                        <th className="px-6 pb-3 text-[10px] uppercase tracking-widest font-bold text-slate-500">Buy Info</th>
                        <th className="px-6 pb-3 text-[10px] uppercase tracking-widest font-bold text-slate-500">Sell Info</th>
                        <th className="px-6 pb-3 text-[10px] uppercase tracking-widest font-bold text-slate-500">Net Profit</th>
                        <th className="px-6 pb-3 text-right pr-12 text-[10px] uppercase tracking-widest font-bold text-slate-500">Action</th>
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
                        return (
                          <motion.tr 
                            key={trade.id} 
                            initial={{ opacity: 0 }}
                            animate={{ opacity: 1 }}
                            className="hover:bg-white/10 transition-colors group bg-white/5"
                          >
                            <td className="px-6 py-4 rounded-l-xl">
                              {trade.buyAmount ? (
                                <>
                                  <p className="text-sm font-semibold">${trade.buyAmount.toLocaleString()}</p>
                                  <p className="text-[10px] font-mono text-slate-500">{trade.buyDate}</p>
                                </>
                              ) : (
                                <p className="text-xs text-slate-600 font-mono italic">Independent Sell</p>
                              )}
                            </td>
                            <td className="px-6 py-4">
                              {trade.sellAmount ? (
                                <>
                                  <p className="text-sm font-semibold">${trade.sellAmount.toLocaleString()}</p>
                                  <p className="text-[10px] font-mono text-slate-500">{trade.sellDate}</p>
                                </>
                              ) : (
                                <p className="text-xs text-slate-600 font-mono italic">Independent Buy</p>
                              )}
                            </td>
                            <td className={`px-6 py-4 font-mono text-sm font-bold ${profit >= 0 ? 'text-profit' : 'text-loss'}`}>
                              {profit >= 0 ? '+' : ''}${profit.toLocaleString(undefined, { minimumFractionDigits: 2 })}
                            </td>
                            <td className="px-6 py-4 text-right rounded-r-xl pr-6">
                              <div className="flex items-center justify-end gap-2">
                                <button 
                                  onClick={() => handleEditTrade(trade)}
                                  className="p-2 text-slate-600 hover:text-accent transition-colors group-hover:bg-accent/10 rounded-lg"
                                  title="Edit"
                                >
                                  <Pencil size={16} />
                                </button>
                                <button 
                                  onClick={() => handleDeleteTrade(trade.id)}
                                  className="p-2 text-slate-600 hover:text-loss transition-colors group-hover:bg-loss/10 rounded-lg"
                                  title="Delete"
                                >
                                  <Trash2 size={16} />
                                </button>
                              </div>
                            </td>
                          </motion.tr>
                        );
                      })}
                    </tbody>
                  </table>
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
                          <td className="px-8 py-5 text-sm font-medium">{report.date}</td>
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
      <nav className="fixed bottom-0 left-0 right-0 z-50 flex justify-center pb-8 pt-4 pointer-events-none">
        <div className="glass shadow-2xl rounded-2xl border border-white/10 flex gap-4 p-2 pointer-events-auto">
          <NavTab id="dashboard" label="Dashboard" icon={LayoutDashboard} />
          <NavTab id="reports" label="Daily P&L" icon={CalendarDays} />
        </div>
      </nav>
    </div>
  );
}
