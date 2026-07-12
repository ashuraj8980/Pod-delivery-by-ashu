
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
  Save
} from "lucide-react";
import Link from "next/link";
import { cn } from "@/lib/utils";

/**
 * @fileOverview Professional Historical Dashboard - Simplified Archive View
 * Reads from pod_monthly_records to provide a searchable archive.
 */

interface PODRow {
  id: string;
  awb: string;
  client: string;
  status: string;
  remark: string;
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

  useEffect(() => {
    setHasMounted(true);
    setCurrentTime(new Date());
    loadData();
    const interval = setInterval(() => {
      setCurrentTime(new Date());
    }, 1000);

    const refreshInterval = setInterval(loadData, 30000);
    return () => {
      clearInterval(interval);
      clearInterval(refreshInterval);
    };
  }, []);

  const loadData = () => {
    try {
      // Load Archive
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

      // Load Pending Count
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
    toast.className = `fixed bottom-8 left-1/2 -translate-x-1/2 px-6 py-3 rounded-xl text-[13px] font-bold z-[1000] shadow-2xl transition-all duration-300 border bg-slate-900 text-white border-white/10`;
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
        if (!session.date || !/^\d{2}-\d{2}-\d{4}$/.test(session.date)) return;
        
        const parts = session.date.split('-');
        const yearMonth = `${parts[2]}-${parts[1]}`;
        const fullDateKey = `${parts[2]}-${parts[1]}-${parts[0]}`;
        
        if (!nextMonthly[yearMonth]) nextMonthly[yearMonth] = {};
        if (!nextMonthly[yearMonth][fullDateKey]) nextMonthly[yearMonth][fullDateKey] = [];

        const sessionStats = session.stats || {
          total: session.data?.length || 0,
          pending: (session.data || []).filter((r: any) => r.status === 'Pending').length,
          dispatched: (session.data || []).filter((r: any) => r.status === 'Dispatched').length,
          rto: (session.data || []).filter((r: any) => r.status === 'RTO').length,
          dto: (session.data || []).filter((r: any) => r.status === 'DTO').length,
        };

        const newRecord = { ...session, stats: sessionStats };
        const filtered = nextMonthly[yearMonth][fullDateKey].filter((s: any) => s.dspId !== session.dspId);
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

  const monthOptions = useMemo(() => {
    return Object.keys(monthlyRecords).sort().reverse();
  }, [monthlyRecords]);

  const getMonthName = (yearMonth: string) => {
    const [y, m] = yearMonth.split('-');
    const date = new Date(parseInt(y), parseInt(m) - 1, 1);
    return date.toLocaleString('default', { month: 'long', year: 'numeric' });
  };

  const toggleDate = (date: string) => {
    setExpandedDates(prev => {
      const next = new Set(prev);
      if (next.has(date)) next.delete(date); else next.add(date);
      return next;
    });
  };

  const formatTimeStr = (d: Date | null) => {
    if (!d) return "--:--:--";
    return d.toLocaleTimeString();
  };

  const formatDateStr = (d: Date | null) => {
    if (!d) return "--/--/----";
    return d.toLocaleDateString();
  };

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
            <select 
              value={selectedMonth}
              onChange={(e) => setSelectedMonth(e.target.value)}
              className="bg-slate-50 border-2 border-slate-100 rounded-2xl px-6 h-14 text-[15px] font-black outline-none focus:border-blue-500 transition-all min-w-[250px] cursor-pointer"
            >
              {monthOptions.length === 0 && <option value="">No history found</option>}
              {monthOptions.map(m => (
                <option key={m} value={m}>{getMonthName(m)}</option>
              ))}
            </select>
          </div>

          {monthOptions.length === 0 ? (
            <div className="text-center py-24 bg-slate-50 rounded-[2.5rem] border-2 border-dashed border-slate-200">
              <Package className="w-16 h-16 text-slate-300 mx-auto mb-4" />
              <p className="text-lg font-black text-slate-500">No sessions saved in history yet.</p>
              <p className="text-sm text-slate-400 mb-8 font-bold">Use the EOD tool to upload and then click "SAVE DAILY REPORT" here.</p>
              <Link href="/eod" className="bg-[#1C2333] text-white px-8 py-3 rounded-2xl font-black text-[14px] inline-flex items-center gap-2 hover:bg-slate-800 transition-all">
                Go to EOD Tool <ArrowRight className="w-4 h-4" />
              </Link>
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
                        {sessionsOnDate.map((s: any) => (
                          <Link 
                            key={s.id} 
                            href={`/eod?sessionId=${s.id}`}
                            className="p-5 border-[1.5px] border-slate-100 rounded-3xl hover:border-blue-500 hover:shadow-lg transition-all group relative overflow-hidden bg-white"
                          >
                            <div className="absolute left-0 top-0 bottom-0 w-1.5 bg-slate-200 group-hover:bg-blue-600 transition-colors" />
                            <div className="flex items-center justify-between mb-2">
                              <p className="text-[17px] font-black text-slate-900 leading-tight group-hover:text-blue-600 transition-colors">{s.feName}</p>
                              <ArrowRight className="w-4 h-4 text-slate-300 group-hover:text-blue-600 group-hover:translate-x-1 transition-all" />
                            </div>
                            <p className="text-[11px] text-slate-400 font-black mb-4 uppercase tracking-wider">{s.dspId} — {s.date}</p>
                            <div className="flex flex-wrap gap-2">
                              <span className="px-2 py-0.5 rounded-lg bg-slate-100 text-[9px] font-black text-slate-600 border border-slate-200 uppercase">{s.stats?.total || 0} PKT</span>
                              {s.stats?.pending > 0 && <span className="px-2 py-0.5 rounded-lg bg-amber-50 text-[9px] font-black text-amber-600 border border-amber-200 uppercase">{s.stats.pending} PENDING</span>}
                              {s.stats?.rto > 0 && <span className="px-2 py-0.5 rounded-lg bg-rose-50 text-[9px] font-black text-rose-600 border border-rose-200 uppercase">{s.stats.rto} RTO</span>}
                              {s.stats?.dto > 0 && <span className="px-2 py-0.5 rounded-lg bg-emerald-50 text-[9px] font-black text-emerald-600 border border-emerald-200 uppercase">{s.stats.dto} DTO</span>}
                            </div>
                          </Link>
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

      <footer className="py-12 text-center text-slate-400 text-[11px] font-black uppercase tracking-widest">
        Enterprise Return Center Operations Tool • Built by Ashu v3.0
      </footer>
    </div>
  );
}
