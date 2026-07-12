
"use client";

import React, { useState, useEffect, useMemo } from "react";
import { 
  Truck, 
  Calendar,
  ChevronDown,
  ChevronRight,
  Search,
  Users,
  Package,
  Clock,
  CheckCircle2,
  RefreshCcw,
  ArrowRight
} from "lucide-react";
import Link from "next/link";
import { cn } from "@/lib/utils";

/**
 * @fileOverview Professional Historical Dashboard
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

  useEffect(() => {
    setHasMounted(true);
    setCurrentTime(new Date());
    loadData();
    const interval = setInterval(() => {
      setCurrentTime(new Date());
      loadData();
    }, 30000);
    return () => clearInterval(interval);
  }, []);

  const loadData = () => {
    const saved = localStorage.getItem('pod_monthly_records');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        setMonthlyRecords(parsed);
        const months = Object.keys(parsed).sort().reverse();
        if (months.length > 0 && !selectedMonth) {
          setSelectedMonth(months[0]);
        }
      } catch (e) {
        setMonthlyRecords({});
      }
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

  const metrics = useMemo(() => {
    let totalFe = 0;
    let totalPkt = 0;
    let totalPending = 0;
    let totalReturned = 0;
    let totalDispatched = 0;

    const dataSet = selectedMonth ? { [selectedMonth]: monthlyRecords[selectedMonth] } : monthlyRecords;

    Object.values(dataSet).forEach(days => {
      Object.values(days).forEach(sessions => {
        sessions.forEach(s => {
          totalFe++;
          totalPkt += s.stats.total;
          totalPending += s.stats.pending;
          totalReturned += (s.stats.rto + s.stats.dto);
          totalDispatched += s.stats.dispatched;
        });
      });
    });

    return { totalFe, totalPkt, totalPending, totalReturned, totalDispatched };
  }, [monthlyRecords, selectedMonth]);

  const toggleDate = (date: string) => {
    setExpandedDates(prev => {
      const next = new Set(prev);
      if (next.has(date)) next.delete(date); else next.add(date);
      return next;
    });
  };

  if (!hasMounted) return null;

  return (
    <div className="min-h-screen bg-[#F0F4FA] font-body">
      {/* Header */}
      <div className="h-1 bg-gradient-to-r from-blue-600 via-amber-400 to-emerald-500" />
      <header className="bg-[#1C2333] text-white py-4 px-6 sticky top-0 z-50 shadow-xl">
        <div className="max-w-[1400px] mx-auto flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="w-10 h-10 bg-blue-600 rounded-xl flex items-center justify-center shadow-lg">
              <Truck className="w-6 h-6" />
            </div>
            <div>
              <h1 className="text-lg font-black tracking-tight">POD Management Tool — Archive</h1>
              <p className="text-[10px] font-bold text-blue-400 uppercase tracking-widest">Delhivery · Palam Vihar RPC · By Ashu</p>
            </div>
          </div>
          <div className="flex items-center gap-6">
            <div className="text-right hidden md:block">
              <p className="text-[13px] font-black">{currentTime?.toLocaleDateString()}</p>
              <p className="text-[11px] font-bold text-slate-400 uppercase tracking-widest">{currentTime?.toLocaleTimeString()}</p>
            </div>
            <button onClick={loadData} className="p-2 hover:bg-white/10 rounded-full transition-all active:rotate-180">
              <RefreshCcw className="w-5 h-5" />
            </button>
            <Link href="/eod" className="bg-[#FFD700] text-[#1C2333] px-5 py-2.5 rounded-xl text-[13px] font-black flex items-center gap-2 hover:scale-105 transition-all shadow-lg active:scale-95">
              <Truck className="w-4 h-4" />
              EOD REJECTION REPORT
            </Link>
          </div>
        </div>
      </header>

      <main className="max-w-[1400px] mx-auto p-6 space-y-8">
        {/* Metrics Grid */}
        <div className="grid grid-cols-1 md:grid-cols-5 gap-4">
          {[
            { label: 'Total FE Saved', val: metrics.totalFe, color: 'text-blue-600', icon: Users, bg: 'bg-white' },
            { label: 'Total Shipments', val: metrics.totalPkt, color: 'text-slate-900', icon: Package, bg: 'bg-white' },
            { label: 'Total Pending', val: metrics.totalPending, color: 'text-amber-600', icon: Clock, bg: 'bg-white' },
            { label: 'Total RTO/DTO', val: metrics.totalReturned, color: 'text-emerald-600', icon: CheckCircle2, bg: 'bg-white' },
            { label: 'Total Dispatched', val: metrics.totalDispatched, color: 'text-rose-600', icon: Truck, bg: 'bg-white' }
          ].map((m, i) => (
            <div key={i} className={cn("p-6 rounded-3xl shadow-sm border border-slate-200 flex flex-col items-center text-center", m.bg)}>
              <div className={cn("w-10 h-10 rounded-full flex items-center justify-center mb-3 bg-slate-50", m.color)}>
                <m.icon className="w-5 h-5" />
              </div>
              <p className="text-[32px] font-black leading-none mb-1 tracking-tighter">{m.val}</p>
              <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{m.label}</p>
            </div>
          ))}
        </div>

        {/* Month Selector */}
        <div className="bg-white rounded-[2rem] p-8 shadow-sm border border-slate-100">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-6 mb-8">
            <div className="flex items-center gap-3">
              <Calendar className="w-6 h-6 text-blue-600" />
              <h2 className="text-xl font-black text-slate-800 tracking-tight">Saved Monthly Archive</h2>
            </div>
            <select 
              value={selectedMonth}
              onChange={(e) => setSelectedMonth(e.target.value)}
              className="bg-slate-50 border-2 border-slate-100 rounded-2xl px-6 h-14 text-[15px] font-black outline-none focus:border-blue-500 transition-all min-w-[250px] cursor-pointer"
            >
              {monthOptions.length === 0 && <option value="">No data found</option>}
              {monthOptions.map(m => (
                <option key={m} value={m}>{getMonthName(m)}</option>
              ))}
            </select>
          </div>

          {monthOptions.length === 0 ? (
            <div className="text-center py-20 bg-slate-50 rounded-[2.5rem] border-2 border-dashed border-slate-200">
              <Package className="w-16 h-16 text-slate-300 mx-auto mb-4" />
              <p className="text-lg font-black text-slate-500">No sessions saved in history yet.</p>
              <p className="text-sm text-slate-400 mb-8 font-bold">Upload and "Save Daily Report" in the EOD tool to see data here.</p>
              <Link href="/eod" className="bg-[#1C2333] text-white px-8 py-3 rounded-2xl font-black text-[14px] inline-flex items-center gap-2 hover:bg-slate-800 transition-all">
                Go to EOD Tool <ArrowRight className="w-4 h-4" />
              </Link>
            </div>
          ) : (
            <div className="space-y-4">
              {selectedMonth && monthlyRecords[selectedMonth] && Object.keys(monthlyRecords[selectedMonth]).sort().reverse().map(dateKey => {
                const isExpanded = expandedDates.has(dateKey);
                const sessions = monthlyRecords[selectedMonth][dateKey];
                return (
                  <div key={dateKey} className="border border-slate-100 rounded-3xl overflow-hidden bg-white hover:shadow-md transition-shadow">
                    <button 
                      onClick={() => toggleDate(dateKey)}
                      className="w-full px-8 py-6 flex items-center justify-between bg-slate-50/50 hover:bg-slate-50 transition-colors"
                    >
                      <div className="flex items-center gap-4">
                        <span className="text-lg font-black text-slate-800">{dateKey}</span>
                        <span className="px-3 py-1 bg-blue-600 text-white rounded-full text-[10px] font-black uppercase">{sessions.length} Sessions</span>
                      </div>
                      {isExpanded ? <ChevronDown className="w-6 h-6 text-slate-400" /> : <ChevronRight className="w-6 h-6 text-slate-400" />}
                    </button>
                    {isExpanded && (
                      <div className="p-6 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 bg-white border-t border-slate-50">
                        {sessions.map(s => (
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
                              <span className="px-2.5 py-1 rounded-lg bg-slate-100 text-[9px] font-black text-slate-600 border border-slate-200">{s.stats.total} PKT</span>
                              {s.stats.pending > 0 && <span className="px-2.5 py-1 rounded-lg bg-amber-50 text-[9px] font-black text-amber-600 border border-amber-200">{s.stats.pending} PENDING</span>}
                              {s.stats.rto > 0 && <span className="px-2.5 py-1 rounded-lg bg-rose-50 text-[9px] font-black text-rose-600 border border-rose-200">{s.stats.rto} RTO</span>}
                              {s.stats.dto > 0 && <span className="px-2.5 py-1 rounded-lg bg-emerald-50 text-[9px] font-black text-emerald-600 border border-emerald-200">{s.stats.dto} DTO</span>}
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
