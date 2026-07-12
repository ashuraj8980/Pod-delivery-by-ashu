
"use client";

import React, { useState, useEffect, useMemo, useCallback } from "react";
import { 
  Truck, 
  Calendar,
  ChevronDown,
  ChevronRight,
  Package,
  RefreshCcw,
  ArrowRight,
  Save,
  Search,
  X,
  Download,
  CheckCircle2,
  Clock,
  AlertCircle,
  TrendingUp
} from "lucide-react";
import Link from "next/link";
import { cn } from "@/lib/utils";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";

/**
 * @fileOverview Professional Historical Dashboard
 * Reads from pod_monthly_records and allows viewing details in a Modal.
 * Includes Global Search for AWB, DSP, or FE Name.
 */

interface PODRow {
  id: string;
  awb: string;
  client: string;
  orderId: string;
  status: string;
  remark: string;
  feName: string;
  dspId: string;
  date: string;
  returnAddress?: string;
}

interface MonthlyRecord {
  id: string;
  feName: string;
  dspId: string;
  date: string;
  timestamp: number;
  data: PODRow[];
  stats: {
    total: number;
    pending: number;
    dispatched: number;
    rto: number;
    dto: number;
  };
}

type MonthlyStorage = Record<string, Record<string, MonthlyRecord[]>>;

export default function Dashboard() {
  const [monthlyRecords, setMonthlyRecords] = useState<MonthlyStorage>({});
  const [selectedMonth, setSelectedMonth] = useState<string>("");
  const [expandedDates, setExpandedDates] = useState<Set<string>>(new Set());
  const [hasMounted, setHasMounted] = useState(false);
  const [currentTime, setCurrentTime] = useState<Date | null>(null);
  const [pendingSessionsCount, setPendingSessionsCount] = useState(0);
  const [globalSearch, setGlobalSearch] = useState("");
  
  // Modal State
  const [viewingSession, setViewingSession] = useState<MonthlyRecord | null>(null);
  const [modalSearch, setModalSearch] = useState("");

  useEffect(() => {
    setHasMounted(true);
    setCurrentTime(new Date());
    loadData();
    const interval = setInterval(() => {
      setCurrentTime(new Date());
    }, 1000);

    return () => clearInterval(interval);
  }, []);

  const loadData = () => {
    try {
      const saved = localStorage.getItem('pod_monthly_records');
      if (saved && saved.trim() !== "") {
        const parsed = JSON.parse(saved);
        if (parsed && typeof parsed === 'object') {
          setMonthlyRecords(parsed);
          const months = Object.keys(parsed).sort().reverse();
          if (months.length > 0 && !selectedMonth) {
            setSelectedMonth(months[0]);
          }
        }
      }

      const pending = localStorage.getItem('pod_sessions');
      if (pending && pending.trim() !== "") {
        const parsedPending = JSON.parse(pending);
        setPendingSessionsCount(Array.isArray(parsedPending) ? parsedPending.length : 0);
      } else {
        setPendingSessionsCount(0);
      }
    } catch (e) {
      console.error("Failed to load records:", e);
    }
  };

  const showToast = useCallback((msg: string) => {
    if (typeof document === 'undefined') return;
    const toast = document.createElement('div');
    toast.className = `fixed bottom-8 left-1/2 -translate-x-1/2 px-6 py-3 rounded-xl text-[13px] font-bold z-[2000] shadow-2xl transition-all duration-300 border bg-slate-900 text-white border-white/10`;
    toast.innerHTML = msg;
    document.body.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      setTimeout(() => toast.remove(), 300);
    }, 3000);
  }, []);

  const handleSaveAllSessions = () => {
    try {
      const saved = localStorage.getItem('pod_sessions');
      if (!saved || saved.trim() === "") {
        showToast("No pending sessions to save");
        return;
      }
      
      const sessions = JSON.parse(saved);
      if (!Array.isArray(sessions) || sessions.length === 0) {
        showToast("No pending sessions to save");
        return;
      }

      const nextMonthly = { ...monthlyRecords };
      
      sessions.forEach((session: any) => {
        if (!session.date) return;
        
        let dateObj = session.date;
        let parts = dateObj.split('-');
        if (parts.length !== 3) return;

        const yearMonth = `${parts[2]}-${parts[1]}`;
        const fullDateKey = `${parts[2]}-${parts[1]}-${parts[0]}`;
        
        if (!nextMonthly[yearMonth]) nextMonthly[yearMonth] = {};
        if (!nextMonthly[yearMonth][fullDateKey]) nextMonthly[yearMonth][fullDateKey] = [];

        const newRecord = { ...session };
        const filtered = nextMonthly[yearMonth][fullDateKey].filter((s: any) => s.dspId !== session.dspId || s.feName !== session.feName);
        nextMonthly[yearMonth][fullDateKey] = [newRecord, ...filtered];
      });

      localStorage.setItem('pod_monthly_records', JSON.stringify(nextMonthly));
      localStorage.removeItem('pod_sessions');
      
      setMonthlyRecords(nextMonthly);
      setPendingSessionsCount(0);
      showToast("All sessions saved to archive & cleared from tool");
      
      const months = Object.keys(nextMonthly).sort().reverse();
      if (months.length > 0) setSelectedMonth(months[0]);
    } catch (error) {
      console.error("Save error:", error);
      showToast("Error saving sessions");
    }
  };

  const toggleDate = (date: string) => {
    setExpandedDates(prev => {
      const next = new Set(prev);
      if (next.has(date)) next.delete(date); else next.add(date);
      return next;
    });
  };

  const getMonthName = (yearMonth: string) => {
    if (!yearMonth) return "";
    const [y, m] = yearMonth.split('-');
    const date = new Date(parseInt(y), parseInt(m) - 1, 1);
    return date.toLocaleString('default', { month: 'long', year: 'numeric' });
  };

  const formatTimeStr = (d: Date | null) => d ? d.toLocaleTimeString() : "--:--:--";
  const formatDateStr = (d: Date | null) => d ? d.toLocaleDateString() : "--/--/----";

  const renderSessionBadges = (stats: any) => (
    <div className="flex flex-wrap gap-1.5 mt-2">
      <span className="px-2 py-0.5 rounded-md bg-slate-100 text-[9px] font-black text-slate-600 border border-slate-200 uppercase">{stats?.total || 0} PKT</span>
      {stats?.pending > 0 && <span className="px-2 py-0.5 rounded-md bg-amber-50 text-[9px] font-black text-amber-600 border border-amber-200 uppercase">{stats.pending} PENDING</span>}
      {stats?.rto > 0 && <span className="px-2 py-0.5 rounded-md bg-rose-50 text-[9px] font-black text-rose-600 border border-rose-200 uppercase">{stats.rto} RTO</span>}
      {stats?.dto > 0 && <span className="px-2 py-0.5 rounded-md bg-emerald-50 text-[9px] font-black text-emerald-600 border border-emerald-200 uppercase">{stats.dto} DTO</span>}
      {stats?.dispatched > 0 && <span className="px-2 py-0.5 rounded-md bg-blue-50 text-[9px] font-black text-blue-600 border border-blue-200 uppercase">{stats.dispatched} DISPATCHED</span>}
    </div>
  );

  const searchResults = useMemo(() => {
    if (!globalSearch.trim()) return [];
    const query = globalSearch.toLowerCase().trim();
    const found: MonthlyRecord[] = [];
    
    Object.values(monthlyRecords).forEach(dates => {
      Object.values(dates).forEach(sessions => {
        sessions.forEach(session => {
          const hasMatch = 
            session.feName.toLowerCase().includes(query) || 
            session.dspId.toLowerCase().includes(query) || 
            session.data.some(r => r.awb.toLowerCase().includes(query));
          
          if (hasMatch) {
            found.push(session);
          }
        });
      });
    });
    return found.sort((a, b) => b.timestamp - a.timestamp);
  }, [monthlyRecords, globalSearch]);

  const filteredModalData = useMemo(() => {
    if (!viewingSession) return [];
    if (!modalSearch) return viewingSession.data;
    const s = modalSearch.toLowerCase();
    return viewingSession.data.filter(r => 
      r.awb.toLowerCase().includes(s) || 
      r.client.toLowerCase().includes(s) || 
      r.remark.toLowerCase().includes(s)
    );
  }, [viewingSession, modalSearch]);

  if (!hasMounted) return null;

  return (
    <div className="min-h-screen bg-[#F0F4FA] font-body">
      <div className="h-1 bg-gradient-to-r from-blue-600 via-amber-400 to-emerald-500" />
      <header className="bg-[#1C2333] text-white py-4 px-6 sticky top-0 z-50 shadow-xl">
        <div className="max-w-[1400px] mx-auto flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="w-10 h-10 bg-blue-600 rounded-xl flex items-center justify-center shadow-lg">
              <Truck className="w-6 h-6" />
            </div>
            <div>
              <h1 className="text-lg font-black tracking-tight text-white uppercase">POD Management Tool — Archive</h1>
              <p className="text-[10px] font-bold text-blue-400 uppercase tracking-widest">Delhivery · Palam Vihar RPC · By Ashu</p>
            </div>
          </div>
          <div className="flex items-center gap-6">
            <div className="text-right hidden md:block">
              <p className="text-[13px] font-black text-white">{formatDateStr(currentTime)}</p>
              <p className="text-[11px] font-bold text-slate-400 uppercase tracking-widest">{formatTimeStr(currentTime)}</p>
            </div>
            
            {pendingSessionsCount > 0 && (
              <button 
                onClick={handleSaveAllSessions}
                className="bg-amber-500 text-[#1C2333] px-4 py-2 rounded-xl text-[12px] font-black flex items-center gap-2 hover:bg-amber-400 transition-all shadow-lg active:scale-95 animate-pulse"
              >
                <Save className="w-4 h-4" />
                SAVE {pendingSessionsCount} SESSIONS
              </button>
            )}

            <button onClick={loadData} className="p-2 hover:bg-white/10 rounded-full transition-all active:rotate-180">
              <RefreshCcw className="w-5 h-5" />
            </button>
            <Link href="/eod" className="bg-[#FFD700] text-[#1C2333] px-5 py-2.5 rounded-xl text-[13px] font-black flex items-center gap-2 hover:scale-105 transition-all shadow-lg active:scale-95">
              <Truck className="w-4 h-4" />
              EOD REJECTION TOOL
            </Link>
          </div>
        </div>
      </header>

      <main className="max-w-[1400px] mx-auto p-6 space-y-8">
        <div className="bg-white rounded-[2rem] p-8 shadow-sm border border-slate-100 min-h-[600px]">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-6 mb-8">
            <div className="flex items-center gap-3">
              <Calendar className="w-6 h-6 text-blue-600" />
              <h2 className="text-xl font-black text-slate-800 tracking-tight uppercase">Monthly Records Archive</h2>
            </div>
            
            <div className="flex items-center gap-4">
              <div className="relative w-[300px]">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                <input 
                  type="text" 
                  placeholder="Search AWB, DSP, FE Name..." 
                  value={globalSearch}
                  onChange={(e) => setGlobalSearch(e.target.value)}
                  className="w-full bg-slate-50 border-2 border-slate-100 rounded-2xl pl-10 pr-4 h-12 text-[14px] font-bold outline-none focus:border-blue-500 transition-all"
                />
              </div>
              <select 
                value={selectedMonth}
                onChange={(e) => setSelectedMonth(e.target.value)}
                className="bg-slate-50 border-2 border-slate-100 rounded-2xl px-6 h-12 text-[14px] font-black outline-none focus:border-blue-500 transition-all min-w-[200px] cursor-pointer"
              >
                {Object.keys(monthlyRecords).length === 0 && <option value="">No history found</option>}
                {Object.keys(monthlyRecords).sort().reverse().map(m => (
                  <option key={m} value={m}>{getMonthName(m)}</option>
                ))}
              </select>
            </div>
          </div>

          {Object.keys(monthlyRecords).length === 0 ? (
            <div className="text-center py-24 bg-slate-50 rounded-[2.5rem] border-2 border-dashed border-slate-200">
              <Package className="w-16 h-16 text-slate-300 mx-auto mb-4" />
              <p className="text-lg font-black text-slate-500">No sessions saved in history yet.</p>
              <p className="text-sm text-slate-400 mb-8 font-bold">Use the EOD tool to upload and then click "SAVE DAILY REPORT" here.</p>
              <Link href="/eod" className="bg-[#1C2333] text-white px-8 py-3 rounded-2xl font-black text-[14px] inline-flex items-center gap-2 hover:bg-slate-800 transition-all">
                Go to EOD Tool <ArrowRight className="w-4 h-4" />
              </Link>
            </div>
          ) : globalSearch.trim() ? (
            <div className="space-y-6">
              <div className="flex items-center justify-between">
                <p className="text-sm font-black text-slate-500 uppercase tracking-widest">Search Results — {searchResults.length} Found</p>
                <button onClick={() => setGlobalSearch("")} className="text-blue-600 font-bold text-sm">Clear Search</button>
              </div>
              {searchResults.length === 0 ? (
                <div className="py-12 text-center text-slate-400 font-bold">No results found for "{globalSearch}"</div>
              ) : (
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                  {searchResults.map((s: MonthlyRecord) => (
                    <div 
                      key={s.id} 
                      onClick={() => setViewingSession(s)}
                      className="p-5 border-[1.5px] border-slate-100 rounded-3xl hover:border-blue-500 hover:shadow-lg transition-all group relative cursor-pointer bg-white"
                    >
                      <div className="absolute left-0 top-0 bottom-0 w-1.5 bg-blue-600" />
                      <div className="flex items-center justify-between mb-2">
                        <p className="text-[17px] font-black text-slate-900 leading-tight group-hover:text-blue-600 transition-colors">{s.feName}</p>
                        <ArrowRight className="w-4 h-4 text-slate-300 group-hover:text-blue-600 group-hover:translate-x-1 transition-all" />
                      </div>
                      <p className="text-[11px] text-slate-400 font-black mb-1 uppercase tracking-wider">{s.dspId} — {s.date}</p>
                      {renderSessionBadges(s.stats)}
                    </div>
                  ))}
                </div>
              )}
            </div>
          ) : (
            <div className="space-y-4">
              {selectedMonth && monthlyRecords[selectedMonth] && Object.keys(monthlyRecords[selectedMonth]).sort().reverse().map(dateKey => {
                const isExpanded = expandedDates.has(dateKey);
                const sessionsOnDate = monthlyRecords[selectedMonth][dateKey];
                return (
                  <div key={dateKey} className="border border-slate-100 rounded-3xl overflow-hidden bg-white shadow-sm hover:shadow-md transition-shadow">
                    <button 
                      onClick={() => toggleDate(dateKey)}
                      className="w-full px-8 py-6 flex items-center justify-between bg-slate-50/50 hover:bg-slate-50 transition-colors"
                    >
                      <div className="flex items-center gap-4">
                        <span className="text-lg font-black text-slate-800">{dateKey}</span>
                        <span className="px-3 py-1 bg-blue-600 text-white rounded-full text-[10px] font-black uppercase tracking-tight">{sessionsOnDate.length} Sessions</span>
                      </div>
                      {isExpanded ? <ChevronDown className="w-6 h-6 text-slate-400" /> : <ChevronRight className="w-6 h-6 text-slate-400" />}
                    </button>
                    {isExpanded && (
                      <div className="p-6 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 bg-white border-t border-slate-50">
                        {sessionsOnDate.map((s: MonthlyRecord) => (
                          <div 
                            key={s.id} 
                            onClick={() => setViewingSession(s)}
                            className="p-5 border-[1.5px] border-slate-100 rounded-3xl hover:border-blue-500 hover:shadow-lg transition-all group relative cursor-pointer bg-white"
                          >
                            <div className="absolute left-0 top-0 bottom-0 w-1.5 bg-slate-200 group-hover:bg-blue-600 transition-colors" />
                            <div className="flex items-center justify-between mb-2">
                              <p className="text-[17px] font-black text-slate-900 leading-tight group-hover:text-blue-600 transition-colors">{s.feName}</p>
                              <ArrowRight className="w-4 h-4 text-slate-300 group-hover:text-blue-600 group-hover:translate-x-1 transition-all" />
                            </div>
                            <p className="text-[11px] text-slate-400 font-black mb-1 uppercase tracking-wider">{s.dspId} — {s.date}</p>
                            {renderSessionBadges(s.stats)}
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </main>

      {/* Session Details Modal */}
      <Dialog open={!!viewingSession} onOpenChange={(open) => !open && setViewingSession(null)}>
        <DialogContent className="max-w-[95vw] w-[1400px] max-h-[90vh] overflow-hidden flex flex-col p-0 gap-0 border-none rounded-[2rem]">
          {viewingSession && (
            <>
              <DialogHeader className="p-8 bg-[#0f172a] text-white space-y-4">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-4">
                    <div className="w-12 h-12 bg-blue-600 rounded-2xl flex items-center justify-center text-white shadow-xl">
                      <Truck className="w-7 h-7" />
                    </div>
                    <div>
                      <DialogTitle className="text-2xl font-black tracking-tight">{viewingSession.feName}</DialogTitle>
                      <p className="text-[11px] font-bold text-blue-400 uppercase tracking-widest">{viewingSession.dspId} • {viewingSession.date}</p>
                    </div>
                  </div>
                  <button onClick={() => setViewingSession(null)} className="p-2 hover:bg-white/10 rounded-full text-slate-400 hover:text-white transition-all">
                    <X className="w-6 h-6" />
                  </button>
                </div>

                <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
                  {[
                    { label: 'Total', val: viewingSession.stats.total, icon: Package, color: 'text-slate-900', bg: 'bg-white' },
                    { label: 'Dispatched', val: viewingSession.stats.dispatched, icon: Truck, color: 'text-blue-600', bg: 'bg-blue-50' },
                    { label: 'Pending', val: viewingSession.stats.pending, icon: Clock, color: 'text-amber-600', bg: 'bg-amber-50' },
                    { label: 'RTO', val: viewingSession.stats.rto, icon: AlertCircle, color: 'text-rose-600', bg: 'bg-rose-50' },
                    { label: 'DTO', val: viewingSession.stats.dto, icon: CheckCircle2, color: 'text-emerald-600', bg: 'bg-emerald-50' }
                  ].map((m, i) => (
                    <div key={i} className={cn("p-3 rounded-2xl flex flex-col items-center justify-center border border-white/10", m.bg)}>
                      <p className={cn("text-[18px] font-black leading-none mb-1", m.color)}>{m.val}</p>
                      <p className="text-[9px] font-black text-slate-500 uppercase tracking-widest">{m.label}</p>
                    </div>
                  ))}
                </div>
              </DialogHeader>

              <div className="bg-slate-50 p-4 border-b flex items-center gap-4">
                <div className="relative flex-1 max-w-md">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                  <input 
                    type="text" 
                    placeholder="Search AWB, Client or Remark..." 
                    value={modalSearch}
                    onChange={(e) => setModalSearch(e.target.value)}
                    className="w-full bg-white border border-slate-200 rounded-xl pl-10 pr-4 h-10 text-sm font-bold focus:border-blue-500 outline-none transition-all shadow-sm"
                  />
                </div>
                <div className="flex-1" />
                <button className="h-10 px-4 bg-emerald-600 text-white rounded-xl text-[11px] font-black uppercase flex items-center gap-2 hover:bg-emerald-700 transition-all">
                  <Download className="w-4 h-4" /> Download Excel
                </button>
              </div>

              <div className="flex-1 overflow-auto bg-white custom-scrollbar">
                <table className="w-full text-center border-collapse">
                  <thead className="sticky top-0 bg-[#f8fafc] text-slate-500 border-b z-10">
                    <tr className="h-12">
                      <th className="px-4 text-[10px] font-black uppercase tracking-widest">Waybill</th>
                      <th className="px-4 text-[10px] font-black uppercase tracking-widest">Client</th>
                      <th className="px-4 text-[10px] font-black uppercase tracking-widest">Order ID</th>
                      <th className="px-4 text-[10px] font-black uppercase tracking-widest text-left">Remark</th>
                      <th className="px-4 text-[10px] font-black uppercase tracking-widest text-left">Return Address</th>
                      <th className="px-4 text-[10px] font-black uppercase tracking-widest">Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredModalData.map((row) => (
                      <tr key={row.id} className="border-b hover:bg-slate-50 transition-colors group">
                        <td className="px-4 py-3 text-[13px] font-black font-mono text-blue-700 cursor-pointer" onClick={() => { navigator.clipboard.writeText(row.awb); showToast("Waybill Copied"); }}>{row.awb}</td>
                        <td className="px-4 py-3 text-[13px] font-bold text-slate-700">{row.client}</td>
                        <td className="px-4 py-3 text-[12px] font-medium text-slate-400">{row.orderId}</td>
                        <td className="px-4 py-3 text-[11px] font-black text-left uppercase text-slate-600">{row.remark}</td>
                        <td className="px-4 py-3 text-[11px] text-left text-slate-500 leading-relaxed max-w-[300px]">{row.returnAddress}</td>
                        <td className="px-4 py-3">
                          <span className={cn(
                            "px-2 py-0.5 rounded text-[9px] font-black uppercase border",
                            row.status === 'Pending' ? "bg-amber-50 text-amber-700 border-amber-200" :
                            row.status === 'RTO' ? "bg-rose-50 text-rose-700 border-rose-200" :
                            row.status === 'DTO' ? "bg-emerald-50 text-emerald-700 border-emerald-200" :
                            "bg-slate-50 text-slate-700 border-slate-200"
                          )}>
                            {row.status}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          )}
        </DialogContent>
      </Dialog>

      <footer className="py-12 text-center text-slate-400 text-[11px] font-black uppercase tracking-widest">
        Enterprise Return Center Operations Tool • Built by Ashu v3.0
      </footer>
    </div>
  );
}
