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
  Settings,
  ChevronRight,
  ChevronLeft,
  Calendar,
  User,
  Hash,
  ArrowUpDown,
  CheckCircle2,
  FileSpreadsheet,
  MousePointerClick
} from "lucide-react";
import * as XLSX from "xlsx";
import { cn } from "@/lib/utils";

/**
 * @fileOverview Delhivery POD Management Tool v1.2
 * Features: Click-to-Copy, Module 2 Data Fix, Smooth Transitions, Filtered Export.
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
  awb: string;
  client: string;
  orderId: string;
  status: string;
  remark: string;
  feName: string;
  dspId: string;
  date: string;
  selected?: boolean;
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
  const [currentSession, setCurrentSession] = useState<Session | null>(null);
  const [statusFilter, setStatusFilter] = useState<string>("all");
  const [remarkFilter, setRemarkFilter] = useState<string | null>(null);
  const [setupData, setSetupData] = useState({ feName: "", dspId: "", date: "" });
  const [showIntactModal, setShowIntactModal] = useState(false);
  const [isMounted, setIsMounted] = useState(false);
  
  // Module 2 State
  const [replacerData, setReplacerData] = useState<any[]>([]);
  const [replacerStats, setReplacerStats] = useState({ total: 0, replaced: 0, missing: 0 });

  useEffect(() => {
    setIsMounted(true);
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) setSessions(JSON.parse(saved));
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
      
      if (e.key === "ArrowLeft") {
        const tabs = ["all", "pending", "dispatched", "rto", "dto"];
        const idx = tabs.indexOf(statusFilter);
        if (idx > 0) {
          setStatusFilter(tabs[idx - 1]);
          setRemarkFilter(null);
        }
      }
      if (e.key === "ArrowRight") {
        const tabs = ["all", "pending", "dispatched", "rto", "dto"];
        const idx = tabs.indexOf(statusFilter);
        if (idx < tabs.length - 1) {
          setStatusFilter(tabs[idx + 1]);
          setRemarkFilter(null);
        }
      }
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [statusFilter, isMounted]);

  const fixAWB = (val: any) => {
    let str = String(val).trim();
    if (/^[\d.]+[eE][+\-]?\d+$/.test(str)) {
      str = BigInt(Math.round(Number(val))).toString();
    }
    return str.replace(/\.0$/, "");
  };

  const showToast = (msg: string, type: 'ok' | 'err' | 'info') => {
    if (!isMounted) return;
    const toast = document.createElement('div');
    toast.className = `fixed bottom-10 left-1/2 -translate-x-1/2 px-6 py-3 rounded-2xl text-[11px] font-black z-[100] shadow-2xl animate-in slide-in-from-bottom-5 duration-300 ${
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
      showToast("Bhai, DSP ID aur FE Name fill karein!", "err");
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
            awb,
            client: String(findVal(/client|clientname/)),
            orderId: String(findVal(/order|orderid/)),
            status,
            remark: String(findVal(/remark|remarks/)),
            feName: setupData.feName,
            dspId: setupData.dspId,
            date: setupData.date,
            selected: false
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
          setCurrentSession(newSession);
          setSessions(prev => [newSession, ...prev]);
          showToast(`Imported ${parsedRows.length} rows!`, "ok");
        } else {
          showToast("No valid rows found in file!", "err");
        }
      } catch (err) {
        showToast("Error parsing file!", "err");
      }
    };
    reader.readAsBinaryString(file);
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

        let replaced = 0;
        let missing = 0;

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
        showToast(`Processed ${processed.length} rows for replacement`, "ok");
      } catch (err) {
        showToast("Replacer error!", "err");
      }
    };
    reader.readAsBinaryString(file);
  };

  const filteredRows = useMemo(() => {
    if (!currentSession) return [];
    let rows = currentSession.data;
    if (statusFilter !== "all") {
      rows = rows.filter(r => r.status === statusFilter);
    }
    if (remarkFilter !== null) {
      rows = rows.filter(r => (r.remark || "") === remarkFilter);
    }
    return rows;
  }, [currentSession, statusFilter, remarkFilter]);

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
    XLSX.writeFile(wb, `Delhivery_${type}_${session.feName}_${statusFilter}.xlsx`);
    showToast("Exporting Excel...", "ok");
  };

  const downloadReplacerExcel = () => {
    if (replacerData.length === 0) return;
    const header = ['Date', 'DSP No', 'AWB Number', 'Client', 'Original Remark', 'Official Remark', 'FE Name'];
    const rows = replacerData.map(r => [
      r.date,
      r.dsp,
      { v: r.awb, t: 's' },
      r.client,
      r.oldRemark,
      r.officialRemark,
      r.feName
    ]);

    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Replaced_Remarks");
    XLSX.writeFile(wb, `Delhivery_Official_Remarks.xlsx`);
  };

  const copyTable = (dataRows: PODRow[]) => {
    if (dataRows.length === 0) return;
    const text = dataRows.map((r, i) => {
      const dsp = i === 0 ? r.dspId : "";
      return `${r.date}\t${dsp}\t${r.awb}\t${r.client}\t${r.orderId}\t${r.remark}\t${r.feName}`;
    }).join("\n");
    copyToClipboard(text, "Table Data Copied");
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

  const uniqueRemarks = useMemo(() => {
    if (!currentSession || statusFilter !== "pending") return [];
    const counts: Record<string, number> = {};
    currentSession.data.filter(r => r.status === 'pending').forEach(r => {
      const rem = r.remark || "";
      counts[rem] = (counts[rem] || 0) + 1;
    });
    return Object.entries(counts).sort((a, b) => b[1] - a[1]);
  }, [currentSession, statusFilter]);

  if (!isMounted) return null;

  return (
    <div className="min-h-screen bg-[#F0F4FA] font-body text-[#1C2333] select-none transition-colors duration-500">
      <div className="h-[3px] w-full bg-gradient-to-r from-[#1565C0] via-[#F9A825] via-[#2E7D32] to-[#D32F2F] sticky top-0 z-[100]" />
      
      <header className="h-[58px] bg-[#1C2333] px-6 flex items-center justify-between text-white shadow-xl relative z-[90]">
        <div className="flex items-center gap-3">
          <div className="bg-gradient-to-br from-blue-400 to-blue-600 p-1.5 rounded-lg shadow-lg">
            <Truck className="w-6 h-6 text-white" />
          </div>
          <div className="flex flex-col">
            <h1 className="text-[15px] font-bold tracking-tight">POD Management Tool</h1>
            <p className="text-[10px] text-slate-400 font-medium">
              Delhivery · Palam Vihar RPC · <span className="text-[#F9A825] font-black italic">By Ashu</span>
            </p>
          </div>
        </div>
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2 bg-green-500/10 px-3 py-1 rounded-full border border-green-500/20">
            <div className="w-1.5 h-1.5 bg-green-500 rounded-full animate-pulse shadow-[0_0_8px_rgba(34,197,94,0.6)]" />
            <span className="text-[9px] font-black text-green-400 tracking-widest uppercase">Live</span>
          </div>
          <div className="bg-[#F9A825] px-3 py-1 rounded-lg text-[#1C2333] font-code text-[11px] font-black shadow-inner">
            {currentSession?.data.length || 0} ROWS
          </div>
        </div>
      </header>

      <nav className="bg-[#1C2333] px-6 flex gap-10 border-t border-white/5">
        <button 
          onClick={() => setActiveTab("eod")}
          className={cn(
            "py-3 text-[11px] font-black uppercase tracking-widest transition-all relative outline-none",
            activeTab === "eod" ? "text-white" : "text-slate-500 hover:text-slate-300"
          )}
        >
          Daily EOD Rejection
          {activeTab === "eod" && <div className="absolute bottom-0 left-0 w-full h-[3px] bg-[#F9A825] rounded-t-full shadow-[0_-2px_10px_rgba(249,168,37,0.5)]" />}
        </button>
        <button 
          onClick={() => setActiveTab("remark")}
          className={cn(
            "py-3 text-[11px] font-black uppercase tracking-widest transition-all relative outline-none",
            activeTab === "remark" ? "text-white" : "text-slate-500 hover:text-slate-300"
          )}
        >
          EOD Rejection Remark
          {activeTab === "remark" && <div className="absolute bottom-0 left-0 w-full h-[3px] bg-[#F9A825] rounded-t-full shadow-[0_-2px_10px_rgba(249,168,37,0.5)]" />}
        </button>
      </nav>

      <main className="p-6 max-w-[1440px] mx-auto animate-in fade-in slide-in-from-top-4 duration-700">
        {activeTab === "eod" ? (
          <div className="space-y-6">
            <div className="bg-white rounded-[1.5rem] p-6 shadow-sm border border-[#E2E8F0] grid grid-cols-1 md:grid-cols-3 gap-6">
              <div className="md:col-span-2 grid grid-cols-2 gap-4">
                <div className="space-y-1.5">
                  <label className="text-[9px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-1"><Hash className="w-3 h-3" /> DSP ID (Required)</label>
                  <input 
                    type="number" 
                    value={setupData.dspId} 
                    onChange={(e) => setSetupData({...setupData, dspId: e.target.value})}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 text-xs font-code font-bold outline-none focus:ring-2 ring-blue-500/20 transition-all" 
                    placeholder="Enter DSP No..." 
                  />
                </div>
                <div className="space-y-1.5">
                  <label className="text-[9px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-1"><User className="w-3 h-3" /> FE / Biker Name</label>
                  <input 
                    type="text" 
                    value={setupData.feName} 
                    onChange={(e) => setSetupData({...setupData, feName: e.target.value})}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 text-xs font-bold outline-none focus:ring-2 ring-blue-500/20 transition-all" 
                    placeholder="Enter FE Name..." 
                  />
                </div>
                <div 
                  className={cn(
                    "col-span-2 border-2 border-dashed rounded-2xl p-8 text-center transition-all relative group overflow-hidden",
                    (!setupData.feName || !setupData.dspId) ? "bg-slate-50 border-slate-100 opacity-50 cursor-not-allowed" : "border-blue-200 hover:border-blue-500 hover:bg-blue-50/40 cursor-pointer"
                  )}
                >
                  <input 
                    type="file" 
                    disabled={!setupData.feName || !setupData.dspId} 
                    onChange={handleFileUpload} 
                    className="absolute inset-0 opacity-0 cursor-pointer disabled:cursor-not-allowed z-10" 
                    accept=".xlsx,.xls,.csv" 
                  />
                  <div className="space-y-3">
                    <Download className={cn("w-10 h-10 mx-auto transition-all duration-300", (!setupData.feName || !setupData.dspId) ? "text-slate-300" : "text-blue-400 group-hover:scale-110 group-hover:text-blue-600")} />
                    <p className="text-[11px] font-bold text-slate-600">
                      {(!setupData.feName || !setupData.dspId) ? "Bhai, pehle DSP ID aur FE Name fill karein!" : "Drop Delhivery export file here or click to upload"}
                    </p>
                  </div>
                </div>
              </div>
              
              <div className="bg-slate-900 rounded-[1.25rem] p-6 text-white flex flex-col justify-between shadow-2xl relative overflow-hidden group">
                <div className="absolute top-0 right-0 w-32 h-32 bg-blue-500/10 rounded-full -mr-10 -mt-10 blur-3xl group-hover:bg-blue-500/20 transition-all" />
                <div className="space-y-4 relative z-10">
                  <p className="text-[10px] font-black text-blue-400 uppercase tracking-widest">Active Session</p>
                  <div>
                    <h3 className="text-xl font-black">{setupData.feName || "FE Name"}</h3>
                    <p className="text-xs font-code text-slate-400 mt-1">DSP: {setupData.dspId || "----"}</p>
                  </div>
                </div>
                <div className="relative z-10 pt-6 border-t border-white/10 flex justify-between items-end">
                   <div className="text-center">
                      <p className="text-[24px] font-black text-amber-500 leading-none">{stats.total}</p>
                      <p className="text-[8px] font-black uppercase text-slate-500 mt-1">Total Pkts</p>
                   </div>
                   <div className="text-center">
                      <p className="text-[24px] font-black text-red-500 leading-none">{stats.intact}</p>
                      <p className="text-[8px] font-black uppercase text-slate-500 mt-1">Intact</p>
                   </div>
                </div>
              </div>
            </div>

            {currentSession && (
              <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
                {[
                  { id: 'all', label: 'All', val: stats.total, col: 'blue' },
                  { id: 'pending', label: 'Pending', val: stats.pending, col: 'amber' },
                  { id: 'dispatched', label: 'Dispatch', val: stats.dispatched, col: 'blue' },
                  { id: 'rto', label: 'RTO', val: stats.rto, col: 'red' },
                  { id: 'dto', label: 'DTO', val: stats.dto, col: 'green' }
                ].map(t => (
                  <button 
                    key={t.id} 
                    onClick={() => { setStatusFilter(t.id); setRemarkFilter(null); }}
                    className={cn(
                      "bg-white p-5 rounded-[1.5rem] shadow-sm border-b-4 transition-all text-center group active:scale-95",
                      statusFilter === t.id ? `border-${t.col}-500 bg-${t.col}-50/30 ring-4 ring-${t.col}-500/5` : "border-transparent border-slate-100 hover:bg-slate-50"
                    )}
                  >
                    <p className={cn("text-3xl font-[900] leading-none transition-transform group-hover:scale-110", `text-${t.col}-600`)}>{t.val}</p>
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-[0.2em] mt-2">{t.label}</p>
                  </button>
                ))}
              </div>
            )}

            {statusFilter === "pending" && uniqueRemarks.length > 0 && (
              <div className="bg-white rounded-[1.5rem] p-6 border border-slate-200 shadow-sm animate-in zoom-in-95 duration-300">
                <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-4 flex items-center gap-2">
                  <Filter className="w-3 h-3" /> REMARK BREAKDOWN — PENDING
                </p>
                <div className="flex flex-wrap gap-3">
                  {uniqueRemarks.map(([rem, count]) => (
                    <button 
                      key={rem} 
                      onClick={() => setRemarkFilter(rem)}
                      className={cn(
                        "px-5 py-2.5 rounded-full text-[10px] font-black uppercase tracking-widest border-2 transition-all flex items-center gap-2 active:scale-95",
                        remarkFilter === rem ? "bg-blue-600 border-blue-600 text-white shadow-lg" : 
                        (rem.toLowerCase().includes('intact') || rem.toLowerCase().includes('reject but package')) ? "bg-red-50 border-red-100 text-red-600 hover:bg-red-100" :
                        "bg-slate-50 border-slate-200 text-slate-500 hover:border-slate-300"
                      )}
                    >
                      {rem || "No Remark"}
                      <span className={cn("px-2 py-0.5 rounded-md text-[9px]", remarkFilter === rem ? "bg-white/20" : "bg-black/5")}>{count}</span>
                    </button>
                  ))}
                  {remarkFilter !== null && (
                    <button onClick={() => setRemarkFilter(null)} className="text-[10px] font-black text-red-500 hover:underline uppercase transition-all px-4">← Show All Pending</button>
                  )}
                </div>
              </div>
            )}

            {currentSession && (
              <div className="bg-white rounded-[1.5rem] shadow-2xl border border-slate-200 overflow-hidden animate-in slide-in-from-bottom-4 duration-500">
                <div className="p-6 bg-slate-50 border-b border-slate-200 flex flex-wrap gap-4 items-center justify-between">
                  <div className="flex flex-wrap gap-3">
                    <button onClick={() => downloadExcel(filteredRows, currentSession)} className="bg-[#2E7D32] hover:bg-[#1B5E20] text-white px-6 py-3 rounded-2xl font-black text-[10px] uppercase tracking-widest flex items-center gap-2 shadow-lg active:scale-95 transition-all"><Download className="w-4 h-4" /> Download Excel</button>
                    <button onClick={() => copyTable(filteredRows)} className="bg-[#1565C0] hover:bg-[#0D47A1] text-white px-6 py-3 rounded-2xl font-black text-[10px] uppercase tracking-widest flex items-center gap-2 shadow-lg active:scale-95 transition-all"><Copy className="w-4 h-4" /> Copy Table</button>
                  </div>
                  <div className="text-[10px] font-black text-slate-400 uppercase tracking-[0.2em]">
                    Showing {filteredRows.length} shipments
                  </div>
                </div>
                
                <div className="overflow-x-auto max-h-[600px] custom-scrollbar">
                  <table className="w-full text-left">
                    <thead className="sticky top-0 z-30">
                      <tr className="bg-[#1C2333] text-white text-[10px] font-black uppercase tracking-[0.2em]">
                        <th className="p-4 w-[105px]">DSP ID</th>
                        <th className="p-4 w-[165px]">AWB Number</th>
                        <th className="p-4 w-[120px]">Client</th>
                        <th className="p-4 w-[120px]">Order ID</th>
                        <th className="p-4">Remark</th>
                        <th className="p-4 w-[115px]">FE Name</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredRows.map((row, idx) => (
                        <tr key={row.awb} className={cn(
                          "border-b border-slate-50 hover:bg-slate-50 transition-colors group",
                          (row.remark.toLowerCase().includes("intact") || row.remark.toLowerCase().includes("reject but package")) && "bg-red-50/40"
                        )}>
                          <td className="p-4 font-code text-xs font-bold text-slate-400">{idx === 0 ? row.dspId : ""}</td>
                          <td 
                            className="p-4 font-code text-[11.5px] font-black text-[#1565C0] tracking-wider cursor-pointer hover:underline flex items-center gap-2"
                            onClick={() => copyToClipboard(row.awb, "AWB Copied")}
                          >
                            {row.awb}
                            <MousePointerClick className="w-3 h-3 text-slate-300 opacity-0 group-hover:opacity-100 transition-opacity" />
                          </td>
                          <td className="p-4 text-[11px] font-bold text-slate-700">{row.client}</td>
                          <td className="p-4 text-[11px] text-slate-500">{row.orderId}</td>
                          <td className="p-4">
                            <span className={cn(
                              "inline-flex px-3 py-1 rounded-full text-[9px] font-black uppercase tracking-widest border shadow-sm",
                              (row.remark.toLowerCase().includes("intact") || row.remark.toLowerCase().includes("reject but package")) ? "bg-red-100 text-red-600 border-red-200" : "bg-[#FFFDE7] text-amber-600 border-amber-200"
                            )}>
                              {row.remark || "No Remark"}
                            </span>
                          </td>
                          <td className="p-4 text-[11px] font-bold text-slate-500">{row.feName}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        ) : (
          <div className="space-y-6">
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              <div className="bg-white rounded-[1.5rem] p-8 border border-slate-200 shadow-sm space-y-8 h-full">
                <h2 className="text-lg font-black tracking-tight flex items-center gap-2"><FileSpreadsheet className="w-5 h-5 text-green-600" /> EOD Remark Replacer</h2>
                <div className="border-4 border-dashed border-green-50 rounded-[2.5rem] p-12 text-center hover:border-green-200 hover:bg-green-50/20 transition-all cursor-pointer relative group overflow-hidden">
                  <input type="file" onChange={handleReplacerUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" accept=".xlsx,.xls,.csv" />
                  <Download className="w-12 h-12 text-green-200 mx-auto transition-transform group-hover:scale-110 group-hover:text-green-400" />
                  <div className="mt-4">
                    <p className="text-sm font-black text-slate-700">Upload EOD Export File</p>
                    <p className="text-[10px] font-bold text-slate-400 mt-1 uppercase tracking-widest">Supports .xlsx, .csv</p>
                  </div>
                </div>

                {replacerData.length > 0 && (
                  <div className="grid grid-cols-3 gap-4 pt-4 animate-in fade-in duration-500">
                    <div className="bg-slate-900 p-4 rounded-2xl text-center">
                       <p className="text-xl font-black text-white leading-none">{replacerStats.total}</p>
                       <p className="text-[8px] font-black text-slate-400 uppercase mt-1">Total</p>
                    </div>
                    <div className="bg-green-50 p-4 rounded-2xl text-center border border-green-100">
                       <p className="text-xl font-black text-green-600 leading-none">{replacerStats.replaced}</p>
                       <p className="text-[8px] font-black text-green-400 uppercase mt-1">Replaced</p>
                    </div>
                    <div className="bg-amber-50 p-4 rounded-2xl text-center border border-amber-100">
                       <p className="text-xl font-black text-amber-600 leading-none">{replacerStats.missing}</p>
                       <p className="text-[8px] font-black text-amber-400 uppercase mt-1">Manual</p>
                    </div>
                  </div>
                )}
              </div>

              <div className="bg-[#1C2333] rounded-[1.5rem] p-8 text-white shadow-2xl overflow-hidden relative">
                <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-blue-500 to-green-500" />
                <h3 className="text-[11px] font-black text-yellow-400 uppercase tracking-[0.3em] mb-8 flex items-center gap-2">
                  <AlertCircle className="w-4 h-4" /> Official Mapping Reference
                </h3>
                <div className="overflow-y-auto max-h-[450px] space-y-2 custom-scrollbar pr-2">
                  {Object.entries(REMARK_MAPPING).map(([old, official]) => (
                    <div key={old} className="p-3 bg-white/5 rounded-xl text-[10px] border border-white/5 hover:bg-white/10 transition-colors">
                      <p className="text-slate-500 italic font-medium">{old}</p>
                      <p className="text-green-400 font-bold mt-1.5 flex items-center gap-2">
                        <CheckCircle2 className="w-3 h-3" /> {official}
                      </p>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {replacerData.length > 0 && (
              <div className="bg-white rounded-[1.5rem] shadow-2xl border border-slate-200 overflow-hidden animate-in slide-in-from-bottom-5 duration-700">
                <div className="p-6 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
                  <div className="flex gap-3">
                    <button onClick={downloadReplacerExcel} className="bg-[#2E7D32] text-white px-6 py-3 rounded-2xl font-black text-[10px] uppercase tracking-widest flex items-center gap-2 active:scale-95 transition-all"><Download className="w-4 h-4" /> Download Replaced Excel</button>
                    <button onClick={() => {
                      const text = replacerData.map(r => `${r.date}\t${r.dsp}\t${r.awb}\t${r.client}\t${r.officialRemark}\t${r.feName}`).join("\n");
                      copyToClipboard(text, "Replaced Data Copied");
                    }} className="bg-[#1565C0] text-white px-6 py-3 rounded-2xl font-black text-[10px] uppercase tracking-widest flex items-center gap-2 active:scale-95 transition-all"><Copy className="w-4 h-4" /> Copy Replaced Table</button>
                  </div>
                </div>
                
                <div className="overflow-x-auto max-h-[600px] custom-scrollbar">
                  <table className="w-full text-left">
                    <thead className="sticky top-0 z-30">
                      <tr className="bg-[#1C2333] text-white text-[10px] font-black uppercase tracking-[0.2em]">
                        <th className="p-4 w-[60px]">#</th>
                        <th className="p-4 w-[110px]">Date</th>
                        <th className="p-4 w-[145px]">AWB Number</th>
                        <th className="p-4">Official Remark</th>
                        <th className="p-4 w-[120px]">Original Remark</th>
                        <th className="p-4 w-[100px]">FE Name</th>
                      </tr>
                    </thead>
                    <tbody>
                      {replacerData.map((row, idx) => (
                        <tr key={idx} className={cn(
                          "border-b border-slate-50 hover:bg-slate-50 transition-colors group",
                          row.isReplaced ? "bg-green-50/20" : "bg-amber-50/20"
                        )}>
                          <td className="p-4 text-[10px] font-bold text-slate-400">{row.id}</td>
                          <td className="p-4 text-[10px] text-slate-500 font-medium">{row.date}</td>
                          <td 
                            className="p-4 font-code text-[11px] font-black text-[#1565C0] cursor-pointer hover:underline flex items-center gap-1.5"
                            onClick={() => copyToClipboard(row.awb)}
                          >
                            {row.awb}
                            <MousePointerClick className="w-2.5 h-2.5 opacity-0 group-hover:opacity-100 transition-opacity" />
                          </td>
                          <td className="p-4">
                            <span className={cn(
                              "inline-flex px-3 py-1 rounded-full text-[9px] font-black uppercase tracking-widest border",
                              row.isReplaced ? "bg-green-100 text-green-700 border-green-200" : "bg-amber-100 text-amber-700 border-amber-200"
                            )}>
                              {row.officialRemark}
                            </span>
                          </td>
                          <td className="p-4 text-[9px] text-slate-400 italic truncate max-w-[120px]">{row.oldRemark}</td>
                          <td className="p-4 text-[10px] font-bold text-slate-500">{row.feName}</td>
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

      {stats.intact > 0 && activeTab === 'eod' && (
        <button 
          onClick={() => setShowIntactModal(true)}
          className="fixed bottom-10 right-10 bg-red-600 hover:bg-red-700 text-white px-8 py-5 rounded-full shadow-2xl flex items-center gap-3 z-40 border-b-4 border-red-800 active:scale-95 active:translate-y-1 transition-all animate-cta-pulse"
        >
          <AlertCircle className="w-6 h-6 fill-current" />
          <div className="flex flex-col items-start">
            <span className="text-[10px] font-black uppercase tracking-widest leading-none">Intact Alert</span>
            <span className="text-2xl font-black leading-none mt-1">{stats.intact} PKTS</span>
          </div>
        </button>
      )}

      {showIntactModal && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-6 bg-[#1C2333]/90 backdrop-blur-md animate-in fade-in duration-300" onClick={() => setShowIntactModal(false)}>
          <div className="bg-white w-full max-w-4xl rounded-[2.5rem] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-300" onClick={e => e.stopPropagation()}>
            <div className="bg-gradient-to-r from-red-600 to-red-800 p-8 text-white flex justify-between items-center">
              <h2 className="text-2xl font-black uppercase tracking-tighter">Intact Packet Summary</h2>
              <button onClick={() => setShowIntactModal(false)} className="hover:bg-white/10 p-2 rounded-full transition-colors"><X className="w-6 h-6" /></button>
            </div>
            <div className="p-8 space-y-6">
              <div className="grid grid-cols-2 gap-4">
                <button 
                  onClick={() => {
                    const rows = currentSession?.data.filter(r => r.status === 'pending' && (r.remark.toLowerCase().includes('intact') || r.remark.toLowerCase().includes('reject but package'))) || [];
                    copyTable(rows);
                  }} 
                  className="bg-slate-900 hover:bg-black text-white p-5 rounded-2xl font-black text-[11px] uppercase tracking-widest transition-all active:scale-95"
                >
                  Copy All Intact
                </button>
                <button 
                  onClick={() => {
                    const rows = currentSession?.data.filter(r => r.status === 'pending' && (r.remark.toLowerCase().includes('intact') || r.remark.toLowerCase().includes('reject but package'))) || [];
                    downloadExcel(rows, currentSession!, "INTACT");
                  }} 
                  className="bg-red-600 hover:bg-red-700 text-white p-5 rounded-2xl font-black text-[11px] uppercase tracking-widest transition-all active:scale-95"
                >
                  Download Intact Excel
                </button>
              </div>
              <div className="max-h-[350px] overflow-y-auto custom-scrollbar pr-2">
                <table className="w-full text-left">
                  <thead className="sticky top-0 bg-white z-10">
                    <tr className="text-[10px] font-black text-slate-400 uppercase border-b">
                      <th className="py-2">AWB</th>
                      <th className="py-2">Client</th>
                      <th className="py-2">Remark</th>
                    </tr>
                  </thead>
                  <tbody>
                    {currentSession?.data.filter(r => r.status === 'pending' && (r.remark.toLowerCase().includes('intact') || r.remark.toLowerCase().includes('reject but package'))).map(r => (
                      <tr key={r.awb} className="border-b border-slate-50 hover:bg-red-50/30 transition-colors">
                        <td className="py-3 font-code text-[11px] font-black text-red-600 cursor-pointer" onClick={() => copyToClipboard(r.awb)}>{r.awb}</td>
                        <td className="py-3 text-[10px] font-bold text-slate-600">{r.client}</td>
                        <td className="py-3 text-[10px] text-red-400">{r.remark}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      )}

      <style jsx global>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=IBM+Plex+Mono:wght@500;600&display=swap');
        
        body { font-family: 'Inter', sans-serif; overflow-x: hidden; }
        .font-code { font-family: 'IBM Plex Mono', monospace; }
        
        .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: transparent; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #E2E8F0; border-radius: 10px; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #CBD5E1; }

        @keyframes pulse-custom {
          0%, 100% { transform: scale(1); box-shadow: 0 0 0 0 rgba(220, 38, 38, 0.4); }
          50% { transform: scale(1.05); box-shadow: 0 0 20px 5px rgba(220, 38, 38, 0.2); }
        }
        .animate-cta-pulse { animation: pulse-custom 3s infinite; }
      `}</style>
    </div>
  );
}

