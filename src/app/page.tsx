
"use client";

import React, { useState, useEffect, useMemo } from "react";
import { 
  Truck, 
  Trash2, 
  Download, 
  Copy, 
  X, 
  AlertCircle, 
  Filter,
  User,
  Hash,
  CheckCircle2,
  FileSpreadsheet,
  MousePointerClick,
  Layers,
  Search,
  Settings
} from "lucide-react";
import * as XLSX from "xlsx";
import { cn } from "@/lib/utils";

/**
 * @fileOverview Delhivery POD Management Tool v3.2
 * Features: Session Persistence, Duplicate Highlighting, Scientific AWB Fix, Filtered Export.
 * Removed "Current Session" card as per user request.
 * Author: Ashu (ashuraj9771@gmail.com)
 */

const STORAGE_KEY = "pod_master_v1";

const STATUS_MAP: Record<string, string> = {
  "pending": "pending",
  "dispatched": "dispatched",
  "dispatch": "dispatched",
  "rto": "rto",
  "dto": "dto",
  "delivered": "dto"
};

const REMARK_MAPPING: Record<string, string> = {
  "Incomplete address & contact details": "Return Address Not Found (Need To New Contact Number)",
  "On Hold. Recipient unable to Accept Delivery": "Client Out Of Station (Receive The Shipment After 3 Days)",
  "On Hold. Recipient unable to Accept Delivery for 5 days": "Client Out Of Station (Receive The Shipment After 3 Days)",
  "Not Attempted": "Not Attempted To Client",
  "Seller/CWH permanently closed": "Seller/CWH permanently closed",
  "Recipient unavailable.Establishment closed": "Client Office Found Close",
  "Reject but package intact": "Client Not Share OTP",
  "Reject - RID not found": "Not Traced In Client System",
  "Barcode/QR mismatch": "Client Rejected Due To Barcode/QR Mismatch",
  "Content mismatch/missing - package tampered": "Client Rejected Due To Content Mismatch",
  "Short shipment": "Short Shipment Received By Fe",
  "Recipient wants delivery at a different address": "Return Address Shifted (Need To New Return Address)"
};

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
}

interface Session {
  id: string;
  feName: string;
  dspId: string;
  date: string;
  data: PODRow[];
  timestamp: number;
}

export default function PODTool() {
  const [activeTab, setActiveTab] = useState<"eod" | "remark">("eod");
  const [sessions, setSessions] = useState<Session[]>([]);
  const [selectedSessionId, setSelectedSessionId] = useState<string | null>(null);
  const [statusFilter, setStatusFilter] = useState<string>("all");
  const [remarkFilter, setRemarkFilter] = useState<string | null>(null);
  const [setupData, setSetupData] = useState({ feName: "", dspId: "", date: "" });
  const [isMounted, setIsMounted] = useState(false);
  
  // Module 2 State
  const [replacerData, setReplacerData] = useState<any[]>([]);
  const [replacerStats, setReplacerStats] = useState({ total: 0, replaced: 0, missing: 0 });

  useEffect(() => {
    setIsMounted(true);
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      const parsed = JSON.parse(saved);
      setSessions(parsed);
      if (parsed.length > 0) setSelectedSessionId(parsed[0].id);
    }
    setSetupData(prev => ({ ...prev, date: new Date().toISOString().split('T')[0] }));
  }, []);

  useEffect(() => {
    if (isMounted) {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(sessions));
    }
  }, [sessions, isMounted]);

  // Keyboard Navigation
  useEffect(() => {
    if (!isMounted) return;
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.target instanceof HTMLInputElement || e.target instanceof HTMLTextAreaElement) return;
      const tabs = ["all", "pending", "dispatched", "rto", "dto"];
      const idx = tabs.indexOf(statusFilter);
      if (e.key === "ArrowLeft" && idx > 0) { setStatusFilter(tabs[idx - 1]); setRemarkFilter(null); }
      if (e.key === "ArrowRight" && idx < tabs.length - 1) { setStatusFilter(tabs[idx + 1]); setRemarkFilter(null); }
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [statusFilter, isMounted]);

  const fixAWB = (val: any) => {
    let str = String(val).trim().replace(/['",]/g, ""); 
    if (/^[\d.]+[eE][+\-]?\d+$/.test(str)) {
      str = BigInt(Math.round(Number(val))).toString();
    }
    return str.replace(/\.0$/, "");
  };

  const showToast = (msg: string, type: 'ok' | 'err' | 'info') => {
    if (!isMounted) return;
    const toast = document.createElement('div');
    toast.className = `fixed bottom-10 left-1/2 -translate-x-1/2 px-6 py-3 rounded-2xl text-[12px] font-black z-[200] shadow-2xl animate-in slide-in-from-bottom-5 duration-300 ${
      type === 'ok' ? 'bg-[#052E0F] text-[#6EE7A6]' : 
      type === 'err' ? 'bg-[#2D0808] text-[#FCA5A5]' : 
      'bg-[#1C2333] text-[#93C5FD]'
    }`;
    toast.innerHTML = `<span class="flex items-center gap-2">${type === 'ok' ? '✓' : type === 'err' ? '✕' : 'ℹ'} ${msg}</span>`;
    document.body.appendChild(toast);
    setTimeout(() => {
      toast.classList.add('animate-out', 'fade-out', 'duration-500');
      setTimeout(() => toast.remove(), 500);
    }, 2800);
  };

  const copyToClipboard = (text: string, label: string = "Copied") => {
    navigator.clipboard.writeText(text).then(() => {
      showToast(`${label}: ${text}`, "info");
    });
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!setupData.feName || !setupData.dspId) {
      showToast("Palam Vihar Biker, DSP ID aur FE Name bharein!", "err");
      e.target.value = "";
      return;
    }
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(ws);
        const parsedRows: PODRow[] = rawData.map((row: any) => {
          const keys = Object.keys(row);
          const findVal = (regex: RegExp) => {
            const key = keys.find(k => regex.test(k.toLowerCase().replace(/[\s_-]/g, "")));
            return key ? row[key] : "";
          };
          const awb = fixAWB(findVal(/waybill|awb|awbnumber/));
          const statusRaw = String(findVal(/status|currentstatus/)).toLowerCase().trim();
          const status = STATUS_MAP[statusRaw] || "unknown";
          return {
            id: crypto.randomUUID(),
            awb,
            client: String(findVal(/client|clientname/)),
            orderId: String(findVal(/order|orderid/)),
            status,
            remark: String(findVal(/remark|remarks/)),
            feName: setupData.feName,
            dspId: setupData.dspId,
            date: setupData.date,
          };
        }).filter(row => row.awb.length >= 3 && row.status !== "unknown");

        if (parsedRows.length > 0) {
          const newSession: Session = {
            id: crypto.randomUUID(),
            feName: setupData.feName,
            dspId: setupData.dspId,
            date: setupData.date,
            data: parsedRows,
            timestamp: Date.now()
          };
          setSessions(prev => [newSession, ...prev]);
          setSelectedSessionId(newSession.id);
          showToast(`Imported ${parsedRows.length} rows!`, "ok");
        } else {
          showToast("No valid rows found!", "err");
        }
      } catch (err) {
        showToast("Error parsing file!", "err");
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  const handleReplacerUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const rawData = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        let replaced = 0, missing = 0;
        const processed = rawData.map((row: any, idx) => {
          const findVal = (regex: RegExp) => {
            const key = Object.keys(row).find(k => regex.test(k.toLowerCase().replace(/[\s_-]/g, "")));
            return key ? row[key] : "";
          };
          const oldRemark = String(findVal(/remark/)).trim();
          let official = REMARK_MAPPING[oldRemark];
          if (!official) {
            const key = Object.keys(REMARK_MAPPING).find(k => oldRemark.toLowerCase().includes(k.toLowerCase()));
            if (key) official = REMARK_MAPPING[key];
          }
          if (official) replaced++; else missing++;
          return {
            id: idx + 1,
            date: String(findVal(/date/)),
            dsp: String(findVal(/dsp|no/)),
            awb: fixAWB(findVal(/awb|waybill/)),
            client: String(findVal(/client/)),
            oldRemark,
            officialRemark: official || oldRemark,
            isReplaced: !!official,
            feName: String(findVal(/fe|biker/))
          };
        });
        setReplacerData(processed);
        setReplacerStats({ total: processed.length, replaced, missing });
        showToast(`Processed ${processed.length} rows`, "ok");
      } catch (err) {
        showToast("Replacer error!", "err");
      }
    };
    reader.readAsBinaryString(file);
  };

  const currentSession = useMemo(() => sessions.find(s => s.id === selectedSessionId) || null, [sessions, selectedSessionId]);

  const filteredRows = useMemo(() => {
    if (!currentSession) return [];
    let rows = currentSession.data;
    if (statusFilter !== "all") rows = rows.filter(r => r.status === statusFilter);
    if (remarkFilter !== null) rows = rows.filter(r => (r.remark || "") === remarkFilter);
    return rows;
  }, [currentSession, statusFilter, remarkFilter]);

  const duplicateAWBs = useMemo(() => {
    if (!currentSession) return new Set<string>();
    const seen = new Set<string>();
    const dupes = new Set<string>();
    currentSession.data.forEach(r => {
      if (seen.has(r.awb)) dupes.add(r.awb);
      seen.add(r.awb);
    });
    return dupes;
  }, [currentSession]);

  const duplicateFEs = useMemo(() => {
    const counts: Record<string, number> = {};
    sessions.forEach(s => {
      const key = `${s.feName}-${s.dspId}`;
      counts[key] = (counts[key] || 0) + 1;
    });
    return counts;
  }, [sessions]);

  const downloadExcel = (dataRows: PODRow[], session: Session, type: string = "EOD") => {
    if (dataRows.length === 0) return;
    const header = ['Date', 'DSP ID', 'AWB Number', 'Client', 'Order ID', 'Remark', 'FE Name'];
    const rows = dataRows.map((r, i) => [
      r.date,
      i === 0 ? session.dspId : "",
      { v: r.awb, t: 's' },
      r.client,
      r.orderId,
      r.remark,
      r.feName
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    ws['!cols'] = [{ wch: 13 }, { wch: 12 }, { wch: 26 }, { wch: 20 }, { wch: 20 }, { wch: 36 }, { wch: 18 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, `Delhivery_${type}_${session.feName}.xlsx`);
  };

  const copyTable = (dataRows: PODRow[]) => {
    if (dataRows.length === 0) return;
    const text = dataRows.map((r, i) => {
      const dsp = i === 0 ? r.dspId : "";
      return `${r.date}\t${dsp}\t${r.awb}\t${r.client}\t${r.orderId}\t${r.remark}\t${r.feName}`;
    }).join("\n");
    copyToClipboard(text, "Table Data Copied (Paste in Excel)");
  };

  const deleteRow = (rowId: string) => {
    if (!selectedSessionId) return;
    setSessions(prev => prev.map(s => {
      if (s.id === selectedSessionId) return { ...s, data: s.data.filter(r => r.id !== rowId) };
      return s;
    }));
  };

  const stats = useMemo(() => {
    if (!currentSession) return { total: 0, pending: 0, dispatched: 0, rto: 0, dto: 0, intact: 0 };
    return {
      total: currentSession.data.length,
      pending: currentSession.data.filter(r => r.status === 'pending').length,
      dispatched: currentSession.data.filter(r => r.status === 'dispatched').length,
      rto: currentSession.data.filter(r => r.status === 'rto').length,
      dto: currentSession.data.filter(r => r.status === 'dto').length,
      intact: currentSession.data.filter(r => r.status === 'pending' && (r.remark.toLowerCase().includes('intact') || r.remark.toLowerCase().includes('reject but package'))).length
    };
  }, [currentSession]);

  if (!isMounted) return null;

  return (
    <div className="min-h-screen bg-[#F0F4FA] font-body text-[#1C2333] select-auto transition-all duration-500 overflow-x-hidden">
      <div className="h-[4px] w-full bg-gradient-to-r from-[#1565C0] via-[#F9A825] via-[#2E7D32] to-[#D32F2F] sticky top-0 z-[100]" />
      
      <header className="h-[65px] bg-[#1C2333] px-8 flex items-center justify-between text-white shadow-2xl relative z-[90]">
        <div className="flex items-center gap-4">
          <div className="bg-gradient-to-br from-blue-400 to-blue-700 p-2 rounded-xl shadow-lg">
            <Truck className="w-7 h-7 text-white" />
          </div>
          <div className="flex flex-col">
            <h1 className="text-[17px] font-black tracking-tight leading-none">POD Management Tool</h1>
            <p className="text-[11px] text-slate-400 font-bold mt-1 uppercase tracking-wider">
              Delhivery · Palam Vihar RPC · <span className="text-[#F9A825] font-black italic">By Ashu</span>
            </p>
          </div>
        </div>
        <div className="flex items-center gap-6">
          <div className="flex items-center gap-2 bg-green-500/20 px-4 py-1.5 rounded-full border border-green-500/30">
            <div className="w-2 h-2 bg-green-500 rounded-full animate-pulse shadow-[0_0_10px_rgba(34,197,94,0.8)]" />
            <span className="text-[10px] font-black text-green-400 tracking-[0.2em] uppercase">System Live</span>
          </div>
          <div className="bg-[#F9A825] px-4 py-1.5 rounded-xl text-[#1C2333] font-code text-[12px] font-black shadow-inner">
            {currentSession?.data.length || 0} ROWS
          </div>
        </div>
      </header>

      <nav className="bg-[#1C2333] px-8 flex gap-12 border-t border-white/10 shadow-lg">
        {[
          { id: "eod", label: "Daily EOD Rejection" },
          { id: "remark", label: "EOD Rejection Remark" }
        ].map(tab => (
          <button 
            key={tab.id}
            onClick={() => setActiveTab(tab.id as any)}
            className={cn(
              "py-4 text-[12px] font-black uppercase tracking-[0.2em] transition-all relative outline-none",
              activeTab === tab.id ? "text-white" : "text-slate-500 hover:text-slate-300"
            )}
          >
            {tab.label}
            {activeTab === tab.id && <div className="absolute bottom-0 left-0 w-full h-[4px] bg-[#F9A825] rounded-t-full shadow-[0_-4px_15px_rgba(249,168,37,0.6)]" />}
          </button>
        ))}
      </nav>

      <main className="p-8 max-w-[1400px] mx-auto animate-in fade-in slide-in-from-top-4 duration-700">
        {activeTab === "eod" ? (
          <div className="space-y-8">
            {/* Session Setup Card - Full Width Clean Layout */}
            <div className="bg-white rounded-[2rem] p-10 shadow-xl border border-[#E2E8F0] space-y-10">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-10">
                <div className="space-y-3">
                  <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest flex items-center gap-2"><Hash className="w-4 h-4" /> DSP ID (Numbers Only)</label>
                  <input 
                    type="number" 
                    value={setupData.dspId} 
                    onChange={(e) => setSetupData({...setupData, dspId: e.target.value})}
                    className="w-full bg-slate-50 border-2 border-slate-200 rounded-2xl px-6 py-5 text-[14px] font-code font-bold outline-none focus:border-[#1565C0] focus:ring-4 ring-blue-500/10 transition-all shadow-sm" 
                    placeholder="Enter DSP No..." 
                  />
                </div>
                <div className="space-y-3">
                  <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest flex items-center gap-2"><User className="w-4 h-4" /> FE / Biker Name</label>
                  <input 
                    type="text" 
                    value={setupData.feName} 
                    onChange={(e) => setSetupData({...setupData, feName: e.target.value})}
                    className="w-full bg-slate-50 border-2 border-slate-200 rounded-2xl px-6 py-5 text-[14px] font-bold outline-none focus:border-[#1565C0] focus:ring-4 ring-blue-500/10 transition-all shadow-sm" 
                    placeholder="Enter FE Name..." 
                  />
                </div>
              </div>
              
              <div 
                className={cn(
                  "border-4 border-dashed rounded-[3rem] p-16 text-center transition-all relative group overflow-hidden shadow-inner",
                  (!setupData.feName || !setupData.dspId) ? "bg-slate-50 border-slate-200 opacity-40 cursor-not-allowed" : "border-blue-200 hover:border-blue-500 hover:bg-blue-50/50 cursor-pointer"
                )}
              >
                <input 
                  type="file" 
                  disabled={!setupData.feName || !setupData.dspId} 
                  onChange={handleFileUpload} 
                  className="absolute inset-0 opacity-0 cursor-pointer disabled:cursor-not-allowed z-10" 
                  accept=".xlsx,.xls,.csv" 
                />
                <div className="space-y-6">
                  <Download className={cn("w-16 h-16 mx-auto transition-all duration-300", (!setupData.feName || !setupData.dspId) ? "text-slate-300" : "text-blue-500 group-hover:scale-110")} />
                  <div>
                    <p className="text-[14px] font-black text-slate-700 uppercase tracking-widest">
                      {(!setupData.feName || !setupData.dspId) ? "Pehele DSP ID aur FE Name bharein!" : "Drop Delhivery export file here or click to upload"}
                    </p>
                    <p className="text-[10px] text-slate-400 font-bold uppercase mt-2 tracking-[0.2em]">Supports Excel & CSV</p>
                  </div>
                </div>
              </div>
            </div>

            {sessions.length > 0 && (
              <div className="space-y-6">
                <div className="flex items-center justify-between px-2">
                  <h2 className="text-[11px] font-black text-slate-500 uppercase tracking-[0.4em] flex items-center gap-3"><Layers className="w-5 h-5" /> Saved Records Database</h2>
                  <button onClick={() => { if(confirm("Kya aap saare sessions delete karna chahte hain?")) setSessions([]); }} className="text-[11px] font-black text-red-600 hover:text-red-700 uppercase tracking-widest flex items-center gap-2 border-b-2 border-red-600/20 hover:border-red-600 pb-0.5"><Trash2 className="w-4 h-4" /> Clear All Data</button>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
                  {sessions.map(s => {
                    const feKey = `${s.feName}-${s.dspId}`;
                    const isFEDupe = duplicateFEs[feKey] > 1;
                    return (
                      <div 
                        key={s.id} 
                        onClick={() => setSelectedSessionId(s.id)}
                        className={cn(
                          "bg-white p-6 rounded-[2.5rem] border-2 transition-all cursor-pointer relative group shadow-sm",
                          selectedSessionId === s.id ? "border-blue-600 shadow-2xl scale-[1.02] bg-blue-50/20" : "border-slate-100 hover:border-blue-400"
                        )}
                      >
                        <button 
                          onClick={(e) => { e.stopPropagation(); setSessions(prev => prev.filter(x => x.id !== s.id)); }}
                          className="absolute top-5 right-5 opacity-0 group-hover:opacity-100 p-2 hover:bg-red-50 rounded-xl text-red-500 transition-all shadow-sm"
                        >
                          <Trash2 className="w-5 h-5" />
                        </button>
                        <div className="flex items-center gap-4">
                          <div className="w-2 h-14 rounded-full bg-gradient-to-b from-blue-500 to-indigo-600" />
                          <div>
                            <p className={cn("text-lg font-black italic", isFEDupe ? "text-red-600" : "text-slate-800")}>{s.feName}</p>
                            <p className="text-[11px] font-code text-slate-400 font-black tracking-tight mt-1 uppercase">DSP: {s.dspId}</p>
                          </div>
                        </div>
                        <div className="mt-8 flex gap-6 border-t pt-5 border-slate-50">
                          <div className="text-center">
                            <p className="text-[20px] font-black text-slate-900">{s.data.length}</p>
                            <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest">Total</p>
                          </div>
                          <div className="text-center ml-auto">
                            <p className="text-[20px] font-black text-amber-600">{s.data.filter(r => r.status === 'pending').length}</p>
                            <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest">Pend</p>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {currentSession && (
              <div className="bg-white rounded-[2.5rem] shadow-2xl border border-slate-200 overflow-hidden animate-in slide-in-from-bottom-6 duration-700">
                <div className="p-8 bg-slate-50 border-b-2 border-slate-100 flex flex-wrap gap-6 items-center justify-between">
                  <div className="flex flex-wrap gap-4">
                    <button onClick={() => downloadExcel(filteredRows, currentSession)} className="bg-[#2E7D32] hover:bg-[#1B5E20] text-white px-8 py-4 rounded-[1.5rem] font-black text-[11px] uppercase tracking-widest flex items-center gap-3 shadow-xl active:scale-95 transition-all"><Download className="w-5 h-5" /> Export {statusFilter.toUpperCase()}</button>
                    <button onClick={() => copyTable(filteredRows)} className="bg-[#1565C0] hover:bg-[#0D47A1] text-white px-8 py-4 rounded-[1.5rem] font-black text-[11px] uppercase tracking-widest flex items-center gap-3 shadow-xl active:scale-95 transition-all"><Copy className="w-5 h-5" /> Copy List</button>
                  </div>
                  <div className="flex items-center gap-5">
                    {[
                      { id: 'all', label: 'All', val: stats.total, col: 'blue' },
                      { id: 'pending', label: 'Pending', val: stats.pending, col: 'amber' },
                      { id: 'rto', label: 'RTO', val: stats.rto, col: 'red' },
                      { id: 'dto', label: 'DTO', val: stats.dto, col: 'green' }
                    ].map(t => (
                      <button 
                        key={t.id} 
                        onClick={() => { setStatusFilter(t.id); setRemarkFilter(null); }}
                        className={cn(
                          "px-7 py-3 rounded-2xl border-2 transition-all text-center",
                          statusFilter === t.id ? `bg-${t.col}-600 border-${t.col}-600 text-white shadow-xl scale-105` : `border-slate-200 text-slate-400 hover:border-slate-300`
                        )}
                      >
                        <p className="text-[16px] font-black leading-none">{t.val}</p>
                        <p className="text-[8px] font-black uppercase mt-1 tracking-widest">{t.label}</p>
                      </button>
                    ))}
                  </div>
                </div>
                
                <div className="overflow-x-auto max-h-[750px] custom-scrollbar">
                  <table className="w-full text-left border-collapse">
                    <thead className="sticky top-0 z-30">
                      <tr className="bg-[#1C2333] text-white text-[11px] font-black uppercase tracking-[0.2em]">
                        <th className="p-6 w-[60px] text-center"><Hash className="w-4 h-4 mx-auto" /></th>
                        <th className="p-6 w-[120px]">DSP ID</th>
                        <th className="p-6 w-[200px]">AWB Number</th>
                        <th className="p-6 w-[160px]">Client</th>
                        <th className="p-6 w-[160px]">Order ID</th>
                        <th className="p-6">Official Remark</th>
                        <th className="p-6 w-[150px]">FE Name</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredRows.map((row, idx) => {
                        const isAWBDupe = duplicateAWBs.has(row.awb);
                        const isIntact = row.remark.toLowerCase().includes("intact") || row.remark.toLowerCase().includes("reject but package");
                        return (
                          <tr key={row.id} className={cn(
                            "border-b-2 border-slate-50 hover:bg-slate-50 transition-colors group relative",
                            isIntact ? "bg-red-50/40" : ""
                          )}>
                            <td className="p-6 text-center">
                              <button onClick={() => deleteRow(row.id)} className="text-slate-300 hover:text-red-600 transition-colors">
                                <X className="w-5 h-5" />
                              </button>
                            </td>
                            <td className="p-6 font-code text-[12px] font-bold text-slate-400">{idx === 0 ? row.dspId : ""}</td>
                            <td 
                              className={cn(
                                "p-6 font-code text-[14px] font-black tracking-wider cursor-pointer hover:underline flex items-center gap-3",
                                isAWBDupe ? "text-red-600" : "text-[#1565C0]"
                              )}
                              onClick={() => copyToClipboard(row.awb, "AWB Copied")}
                            >
                              {row.awb}
                              <MousePointerClick className="w-4 h-4 text-slate-300 opacity-0 group-hover:opacity-100 transition-opacity" />
                            </td>
                            <td className="p-6 text-[12px] font-bold text-slate-700">{row.client}</td>
                            <td className="p-6 text-[12px] text-slate-500 font-medium">{row.orderId}</td>
                            <td className="p-6">
                              <span className={cn(
                                "inline-flex px-5 py-2 rounded-full text-[10px] font-black uppercase tracking-widest border shadow-sm",
                                isIntact ? "bg-red-100 text-red-700 border-red-200" : "bg-[#FFFDE7] text-amber-700 border-amber-200"
                              )}>
                                {row.remark || "No Remark"}
                              </span>
                            </td>
                            <td className="p-6 text-[12px] font-black text-slate-500">{row.feName}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        ) : (
          <div className="space-y-10">
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-10">
              <div className="bg-white rounded-[3rem] p-12 border-2 border-slate-100 shadow-2xl space-y-12">
                <h2 className="text-3xl font-black tracking-tighter flex items-center gap-4"><FileSpreadsheet className="w-10 h-10 text-green-600" /> Remark Replacer AI</h2>
                <div className="border-8 border-dashed border-green-50 rounded-[4rem] p-20 text-center hover:border-green-300 hover:bg-green-50/50 transition-all cursor-pointer relative group overflow-hidden shadow-inner">
                  <input type="file" onChange={handleReplacerUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" accept=".xlsx,.xls,.csv" />
                  <Download className="w-20 h-20 text-green-200 mx-auto transition-transform group-hover:scale-125 group-hover:text-green-500" />
                  <div className="mt-8">
                    <p className="text-xl font-black text-slate-700 uppercase italic">Upload Delhivery EOD Export</p>
                    <p className="text-[11px] font-bold text-slate-400 mt-3 tracking-[0.4em] uppercase">Excel & CSV Supported</p>
                  </div>
                </div>

                {replacerData.length > 0 && (
                  <div className="grid grid-cols-3 gap-8 pt-8 animate-in zoom-in-95 duration-500">
                    <div className="bg-slate-900 p-8 rounded-[2.5rem] text-center shadow-xl">
                       <p className="text-4xl font-black text-white leading-none">{replacerStats.total}</p>
                       <p className="text-[10px] font-black text-slate-400 uppercase mt-3 tracking-widest">Total</p>
                    </div>
                    <div className="bg-green-600 p-8 rounded-[2.5rem] text-center shadow-xl">
                       <p className="text-4xl font-black text-white leading-none">{replacerStats.replaced}</p>
                       <p className="text-[10px] font-black text-white/60 uppercase mt-3 tracking-widest">Replaced</p>
                    </div>
                    <div className="bg-amber-500 p-8 rounded-[2.5rem] text-center shadow-xl">
                       <p className="text-4xl font-black text-white leading-none">{replacerStats.missing}</p>
                       <p className="text-[10px] font-black text-white/60 uppercase mt-3 tracking-widest">No Match</p>
                    </div>
                  </div>
                )}
              </div>

              <div className="bg-[#1C2333] rounded-[3rem] p-12 text-white shadow-2xl relative overflow-hidden">
                <div className="absolute top-0 left-0 w-full h-[8px] bg-gradient-to-r from-blue-500 via-green-500 to-amber-500" />
                <h3 className="text-[12px] font-black text-[#F9A825] uppercase tracking-[0.5em] mb-12 flex items-center gap-4">
                  <AlertCircle className="w-6 h-6" /> Official Remark Guide
                </h3>
                <div className="overflow-y-auto max-h-[550px] space-y-4 custom-scrollbar pr-6">
                  {Object.entries(REMARK_MAPPING).map(([old, official]) => (
                    <div key={old} className="p-6 bg-white/5 rounded-3xl text-[11px] border border-white/10 hover:bg-white/10 transition-colors shadow-sm">
                      <p className="text-slate-500 italic font-bold mb-3 uppercase tracking-tight">{old}</p>
                      <p className="text-green-400 font-black flex items-center gap-3 text-[14px]">
                        <CheckCircle2 className="w-5 h-5 text-green-500" /> {official}
                      </p>
                    </div>
                  ))}
                </div>
              </div>
            </div>
            
            {replacerData.length > 0 && (
              <div className="bg-white rounded-[3rem] shadow-2xl border-2 border-slate-100 overflow-hidden animate-in slide-in-from-bottom-8 duration-700">
                <div className="p-10 bg-slate-50 border-b-2 border-slate-100 flex items-center justify-between">
                  <div className="flex gap-8">
                    <button onClick={() => {
                        const header = ['Date', 'DSP No', 'AWB Number', 'Client', 'Original Remark', 'Official Remark', 'FE Name'];
                        const rows = replacerData.map(r => [r.date, r.dsp, { v: r.awb, t: 's' }, r.client, r.oldRemark, r.officialRemark, r.feName]);
                        const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
                        const wb = XLSX.utils.book_new();
                        XLSX.utils.book_append_sheet(wb, ws, "Replaced");
                        XLSX.writeFile(wb, `Delhivery_Replaced_Remarks.xlsx`);
                    }} className="bg-[#2E7D32] text-white px-12 py-5 rounded-[1.75rem] font-black text-[11px] uppercase tracking-widest flex items-center gap-4 shadow-2xl active:scale-95 transition-all"><Download className="w-6 h-6" /> Download Replaced</button>
                    <button onClick={() => {
                      const text = replacerData.map(r => `${r.date}\t${r.dsp}\t${r.awb}\t${r.client}\t${r.officialRemark}\t${r.feName}`).join("\n");
                      copyToClipboard(text, "Replaced List Copied");
                    }} className="bg-[#1565C0] text-white px-12 py-5 rounded-[1.75rem] font-black text-[11px] uppercase tracking-widest flex items-center gap-4 shadow-2xl active:scale-95 transition-all"><Copy className="w-6 h-6" /> Copy For Excel</button>
                  </div>
                </div>
                
                <div className="overflow-x-auto max-h-[750px] custom-scrollbar">
                  <table className="w-full text-left">
                    <thead className="sticky top-0 z-30">
                      <tr className="bg-[#1C2333] text-white text-[11px] font-black uppercase tracking-[0.2em]">
                        <th className="p-7 w-[80px]">#</th>
                        <th className="p-7 w-[140px]">Date</th>
                        <th className="p-7 w-[200px]">AWB Number</th>
                        <th className="p-7">Official Remark</th>
                        <th className="p-7 w-[180px]">Original Remark</th>
                        <th className="p-7 w-[150px]">FE Name</th>
                      </tr>
                    </thead>
                    <tbody>
                      {replacerData.map((row, idx) => (
                        <tr key={idx} className={cn(
                          "border-b-2 border-slate-50 hover:bg-slate-50 transition-colors group",
                          row.isReplaced ? "bg-green-50/20" : "bg-amber-50/20"
                        )}>
                          <td className="p-7 text-[12px] font-black text-slate-400">{row.id}</td>
                          <td className="p-7 text-[12px] text-slate-600 font-black">{row.date}</td>
                          <td 
                            className="p-7 font-code text-[14px] font-black text-[#1565C0] cursor-pointer hover:underline flex items-center gap-3"
                            onClick={() => copyToClipboard(row.awb, "AWB Copied")}
                          >
                            {row.awb}
                            <MousePointerClick className="w-4 h-4 opacity-0 group-hover:opacity-100 transition-opacity" />
                          </td>
                          <td className="p-7">
                            <span className={cn(
                              "inline-flex px-6 py-2.5 rounded-full text-[10px] font-black uppercase tracking-widest border shadow-sm",
                              row.isReplaced ? "bg-green-600 text-white border-green-700" : "bg-amber-500 text-white border-amber-600"
                            )}>
                              {row.officialRemark}
                            </span>
                          </td>
                          <td className="p-7 text-[10px] text-slate-400 italic font-bold truncate max-w-[180px]">{row.oldRemark}</td>
                          <td className="p-7 text-[12px] font-black text-slate-700">{row.feName}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        )}
      </main>

      <style jsx global>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&family=IBM+Plex+Mono:wght@500;600&display=swap');
        
        body { font-family: 'Inter', sans-serif; overflow-x: hidden; }
        .font-code { font-family: 'IBM Plex Mono', monospace; }
        
        .custom-scrollbar::-webkit-scrollbar { width: 10px; height: 10px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #F8FAFC; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 20px; border: 3px solid #F8FAFC; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #94A3B8; }

        input[type="number"]::-webkit-inner-spin-button,
        input[type="number"]::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }
        
        ::selection { background: #BFDBFE; color: #1E40AF; }
      `}</style>
    </div>
  );
}
