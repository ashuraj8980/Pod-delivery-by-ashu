
"use client";

import React, { useState, useEffect, useMemo, useRef } from "react";
import { 
  Truck, 
  Trash2, 
  Download, 
  Copy, 
  Plus, 
  X, 
  CheckCircle2, 
  AlertCircle, 
  Search,
  ChevronLeft,
  ChevronRight,
  Filter,
  Lock,
  Settings
} from "lucide-react";
import * as XLSX from "xlsx";
import { cn } from "@/lib/utils";

/**
 * @fileOverview POD Management Tool for Delhivery.
 * Exact specifications: Module 1 (EOD) & Module 2 (Replacer).
 */

// --- Constants & Config ---
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
  const [setupData, setSetupData] = useState({ feName: "", dspId: "", date: new Date().toISOString().split('T')[0] });
  const [showIntactModal, setShowIntactModal] = useState(false);
  const [focusedIndex, setFocusedIndex] = useState(-1);
  
  // Module 2 State
  const [replacerData, setReplacerData] = useState<any[]>([]);
  const [replacerStats, setReplacerStats] = useState({ total: 0, replaced: 0, missing: 0 });

  const setupRef = useRef<any>(null);
  const debounceTimer = useRef<NodeJS.Timeout | null>(null);

  // Persistence
  useEffect(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) setSessions(JSON.parse(saved));
  }, []);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(sessions));
  }, [sessions]);

  // Keyboard Shortcuts
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Tab Navigation
      if (e.key === "ArrowLeft") {
        const tabs = ["all", "pending", "dispatched", "rto", "dto"];
        const idx = tabs.indexOf(statusFilter);
        if (idx > 0) setStatusFilter(tabs[idx - 1]);
      }
      if (e.key === "ArrowRight") {
        const tabs = ["all", "pending", "dispatched", "rto", "dto"];
        const idx = tabs.indexOf(statusFilter);
        if (idx < tabs.length - 1) setStatusFilter(tabs[idx + 1]);
      }

      // Table Navigation
      if (currentSession && focusedIndex !== -1) {
        if (e.key === "ArrowDown") setFocusedIndex(Math.min(filteredRows.length - 1, focusedIndex + 1));
        if (e.key === "ArrowUp") setFocusedIndex(Math.max(0, focusedIndex - 1));
        if (e.key === "Enter") toggleRowSelection(filteredRows[focusedIndex].awb);
        if (e.key === "Delete") removeRow(filteredRows[focusedIndex].awb);
      }
      if (e.key === "Escape") setFocusedIndex(-1);
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [statusFilter, focusedIndex, currentSession]);

  // Auto-save session debounce
  useEffect(() => {
    if (setupData.feName && setupData.dspId) {
      if (debounceTimer.current) clearTimeout(debounceTimer.current);
      debounceTimer.current = setTimeout(() => {
        // Automatically handled when file is uploaded or session created
      }, 800);
    }
  }, [setupData.feName, setupData.dspId]);

  // --- Logic ---

  const fixAWB = (val: any) => {
    let str = String(val);
    if (/^[\d.]+[eE][+\-]?\d+$/.test(str)) {
      str = BigInt(Math.round(Number(val))).toString();
    }
    return str.replace(/\.0$/, "");
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!setupData.feName || !setupData.dspId) {
      showToast("Bhai, FE aur DSP Number enter karein!", "err");
      e.target.value = "";
      return;
    }

    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
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
        showToast("No valid rows found!", "err");
      }
    };
    reader.readAsBinaryString(file);
  };

  const toggleRowSelection = (awb: string) => {
    if (!currentSession) return;
    const newData = currentSession.data.map(r => r.awb === awb ? { ...r, selected: !r.selected } : r);
    setCurrentSession({ ...currentSession, data: newData });
  };

  const removeRow = (awb: string) => {
    if (!currentSession) return;
    const newData = currentSession.data.filter(r => r.awb !== awb);
    setCurrentSession({ ...currentSession, data: newData });
  };

  const deleteSelected = () => {
    if (!currentSession) return;
    const newData = currentSession.data.filter(r => !r.selected);
    setCurrentSession({ ...currentSession, data: newData });
  };

  const copyTable = (dataRows: PODRow[]) => {
    const text = dataRows.map((r, i) => {
      const dsp = i === 0 ? r.dspId : "";
      return `${r.date}\t${dsp}\t${r.awb}\t${r.client}\t${r.orderId}\t${r.remark}\t${r.feName}`;
    }).join("\n");
    
    navigator.clipboard.writeText(text).then(() => {
      showToast("Copied — Paste in Excel", "info");
    }).catch(() => {
      const el = document.createElement('textarea');
      el.value = text;
      document.body.appendChild(el);
      el.select();
      document.execCommand('copy');
      document.body.removeChild(el);
      showToast("Copied — Paste in Excel", "info");
    });
  };

  const downloadExcel = (session: Session) => {
    const header = ['Date', 'DSP ID', 'AWB Number', 'Client', 'Order ID', 'Remark', 'FE Name'];
    const rows = session.data.map((r, i) => [
      r.date,
      i === 0 ? r.dspId : "",
      { v: r.awb, t: 's' },
      r.client,
      r.orderId,
      r.remark,
      r.feName
    ]);

    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    
    // Formatting
    const wscols = [{ wch: 13 }, { wch: 12 }, { wch: 26 }, { wch: 20 }, { wch: 20 }, { wch: 36 }, { wch: 18 }];
    ws['!cols'] = wscols;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "POD_Report");
    XLSX.writeFile(wb, `Delhivery_POD_${session.feName}_${session.date}.xlsx`);
    showToast("Downloading Excel...", "ok");
  };

  // --- Module 2 Logic ---
  const handleReplacerUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
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

        const oldRemark = String(findVal(/remark/));
        // Exact case-insensitive match or partial include
        let official = REMARK_MAPPING[oldRemark];
        if (!official) {
          const key = Object.keys(REMARK_MAPPING).find(k => oldRemark.toLowerCase().includes(k.toLowerCase()));
          if (key) official = REMARK_MAPPING[key];
        }

        if (official) replaced++; else missing++;

        return {
          id: idx + 1,
          date: String(findVal(/date/)),
          dsp: String(findVal(/dsp/)),
          awb: fixAWB(findVal(/awb/)),
          client: String(findVal(/client/)),
          oldRemark,
          officialRemark: official || oldRemark,
          isReplaced: !!official,
          feName: String(findVal(/fe/))
        };
      });

      setReplacerData(processed);
      setReplacerStats({ total: processed.length, replaced, missing });
      showToast(`Processed ${processed.length} rows`, "ok");
    };
    reader.readAsBinaryString(file);
  };

  const showToast = (msg: string, type: 'ok' | 'err' | 'info') => {
    const toast = document.createElement('div');
    toast.className = `fixed bottom-10 left-1/2 -translate-x-1/2 px-6 py-3 rounded-2xl text-xs font-bold z-[100] shadow-2xl animate-in slide-in-from-bottom-2 duration-300 ${
      type === 'ok' ? 'bg-[#052E0F] text-[#6EE7A6]' : 
      type === 'err' ? 'bg-[#2D0808] text-[#FCA5A5]' : 
      'bg-[#1C2333] text-[#93C5FD]'
    }`;
    toast.innerText = msg;
    document.body.appendChild(toast);
    setTimeout(() => {
      toast.classList.add('animate-out', 'fade-out', 'duration-500');
      setTimeout(() => toast.remove(), 500);
    }, 2800);
  };

  // --- Computed ---

  const filteredRows = useMemo(() => {
    if (!currentSession) return [];
    let rows = currentSession.data;
    if (statusFilter !== "all") {
      rows = rows.filter(r => r.status === statusFilter);
    }
    if (remarkFilter) {
      rows = rows.filter(r => (r.remark || "") === (remarkFilter === "" ? "" : remarkFilter));
    }
    return rows;
  }, [currentSession, statusFilter, remarkFilter]);

  const stats = useMemo(() => {
    if (!currentSession) return { total: 0, pending: 0, dispatched: 0, rto: 0, dto: 0, intact: 0 };
    return {
      total: currentSession.data.length,
      pending: currentSession.data.filter(r => r.status === 'pending').length,
      dispatched: currentSession.data.filter(r => r.status === 'dispatched').length,
      rto: currentSession.data.filter(r => r.status === 'rto').length,
      dto: currentSession.data.filter(r => r.status === 'dto').length,
      intact: currentSession.data.filter(r => r.status === 'pending' && r.remark.toLowerCase().includes('intact')).length
    };
  }, [currentSession]);

  const uniqueRemarks = useMemo(() => {
    if (!currentSession || statusFilter !== "pending") return [];
    const counts: Record<string, number> = {};
    currentSession.data.filter(r => r.status === 'pending').forEach(r => {
      counts[r.remark] = (counts[r.remark] || 0) + 1;
    });
    return Object.entries(counts).sort((a, b) => b[1] - a[1]);
  }, [currentSession, statusFilter]);

  return (
    <div className="min-h-screen bg-[#F0F4FA] font-body text-[#1C2333] select-none">
      {/* Rainbow Stripe */}
      <div className="h-[3px] w-full bg-gradient-to-r from-[#1565C0] via-[#F9A825] via-[#2E7D32] to-[#D32F2F]" />
      
      {/* Header */}
      <header className="h-[58px] bg-[#1C2333] px-6 flex items-center justify-between text-white shadow-xl relative z-20">
        <div className="flex items-center gap-3">
          <div className="bg-gradient-to-br from-blue-400 to-blue-600 p-1.5 rounded-lg shadow-inner">
            <Truck className="w-6 h-6 text-white" />
          </div>
          <div className="flex flex-col">
            <h1 className="text-[15px] font-bold tracking-tight">POD Management Tool</h1>
            <p className="text-[10px] text-slate-400 font-medium">
              Delhivery · Palam Vihar RPC · <span className="text-[#F9A825] font-bold italic">By Ashu</span>
            </p>
          </div>
        </div>
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2 bg-green-500/10 px-3 py-1 rounded-full border border-green-500/20">
            <div className="w-1.5 h-1.5 bg-green-500 rounded-full animate-pulse shadow-[0_0_8px_rgba(34,197,94,0.6)]" />
            <span className="text-[9px] font-black text-green-400 tracking-widest uppercase">Live</span>
          </div>
          <div className="bg-[#F9A825] px-3 py-1 rounded-lg text-[#1C2333] font-code text-[11px] font-black shadow-lg">
            {currentSession?.data.length || 0} ROWS
          </div>
          <button className="text-slate-400 hover:text-white transition-colors">
            <Settings className="w-4 h-4" />
          </button>
        </div>
      </header>

      {/* Nav Tabs */}
      <nav className="bg-[#1C2333] px-6 flex gap-10">
        <button 
          onClick={() => setActiveTab("eod")}
          className={cn(
            "py-3 text-[11px] font-bold uppercase tracking-widest transition-all relative",
            activeTab === "eod" ? "text-white" : "text-slate-500 hover:text-slate-300"
          )}
        >
          Daily EOD Rejection
          {activeTab === "eod" && <div className="absolute bottom-0 left-0 w-full h-[3px] bg-[#F9A825] rounded-t-full shadow-[0_-2px_8px_rgba(249,168,37,0.4)]" />}
        </button>
        <button 
          onClick={() => setActiveTab("remark")}
          className={cn(
            "py-3 text-[11px] font-bold uppercase tracking-widest transition-all relative",
            activeTab === "remark" ? "text-white" : "text-slate-500 hover:text-slate-300"
          )}
        >
          EOD Rejection Remark
          {activeTab === "remark" && <div className="absolute bottom-0 left-0 w-full h-[3px] bg-[#F9A825] rounded-t-full" />}
        </button>
      </nav>

      <main className="p-6 max-w-[1400px] mx-auto space-y-6">
        {activeTab === "eod" ? (
          <>
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              {/* Setup Card */}
              <div className="bg-white rounded-[1.25rem] p-6 shadow-sm border border-[#E2E8F0] space-y-6">
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <label className="text-[9px] font-black text-slate-400 uppercase tracking-widest">DSP ID</label>
                    <input 
                      type="number" 
                      value={setupData.dspId} 
                      onChange={(e) => setSetupData({...setupData, dspId: e.target.value})}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-xs font-code font-bold outline-none focus:ring-2 ring-blue-500/10" 
                      placeholder="Enter DSP No..." 
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[9px] font-black text-slate-400 uppercase tracking-widest">FE / Biker Name</label>
                    <input 
                      type="text" 
                      value={setupData.feName} 
                      onChange={(e) => setSetupData({...setupData, feName: e.target.value})}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-xs font-bold outline-none" 
                      placeholder="Name..." 
                    />
                  </div>
                </div>
                <div className="space-y-1">
                  <label className="text-[9px] font-black text-slate-400 uppercase tracking-widest">Date</label>
                  <input 
                    type="date" 
                    value={setupData.date} 
                    onChange={(e) => setSetupData({...setupData, date: e.target.value})}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-xs font-bold" 
                  />
                </div>

                <div 
                  className={cn(
                    "border-2 border-dashed rounded-2xl p-8 text-center transition-all relative group",
                    (!setupData.feName || !setupData.dspId) ? "bg-slate-50 border-slate-200 opacity-60" : "border-blue-200 hover:border-blue-500 hover:bg-blue-50/30 cursor-pointer"
                  )}
                >
                  <input 
                    type="file" 
                    disabled={!setupData.feName || !setupData.dspId} 
                    onChange={handleFileUpload} 
                    className="absolute inset-0 opacity-0 cursor-pointer disabled:cursor-not-allowed" 
                    accept=".xlsx,.xls,.csv,.tsv,.ods" 
                  />
                  <div className="space-y-3">
                    <Download className={cn("w-10 h-10 mx-auto transition-colors", (!setupData.feName || !setupData.dspId) ? "text-slate-300" : "text-blue-400 group-hover:text-blue-600")} />
                    <div>
                      <p className="text-[11px] font-bold text-slate-600">Drop Delhivery export file here or click to upload</p>
                      <div className="mt-2 flex items-center justify-center gap-1.5">
                        {['.XLSX', '.XLS', '.CSV', '.ODS'].map(tag => (
                          <span key={tag} className="text-[8px] font-code bg-slate-100 px-1.5 py-0.5 rounded text-slate-400">{tag}</span>
                        ))}
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              {/* Saved Sessions Grid */}
              <div className="lg:col-span-2 space-y-4">
                <div className="flex items-center justify-between">
                  <h3 className="text-xs font-black text-slate-400 uppercase tracking-[0.2em]">All FE Sessions — Saved Data</h3>
                  <button 
                    onClick={() => { if(confirm("Clear All Saved Data?")) {setSessions([]); localStorage.removeItem(STORAGE_KEY); showToast("All sessions cleared", "err");}}}
                    className="text-[9px] font-black text-red-500 hover:bg-red-50 px-3 py-1.5 rounded-full uppercase transition-all"
                  >
                    Clear All Sessions
                  </button>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
                  {sessions.length === 0 && (
                    <div className="col-span-full py-16 text-center text-slate-400 font-bold italic border-2 border-dashed rounded-[2rem]">No Saved Data Found</div>
                  )}
                  {sessions.map(s => (
                    <div 
                      key={s.id} 
                      onClick={() => { setCurrentSession(s); setStatusFilter("all"); showToast(`Loaded session: ${s.feName}`, "info"); }}
                      className={cn(
                        "bg-white border p-5 rounded-[1.5rem] shadow-sm hover:shadow-md transition-all cursor-pointer relative group",
                        currentSession?.id === s.id ? "border-blue-500 ring-4 ring-blue-500/5 bg-blue-50/20" : "border-slate-200"
                      )}
                    >
                      <div className="absolute left-0 top-1/2 -translate-y-1/2 w-1 h-12 bg-gradient-to-b from-blue-500 to-green-500 rounded-r-full" />
                      <button 
                        onClick={(e) => { e.stopPropagation(); setSessions(sessions.filter(ses => ses.id !== s.id)); if(currentSession?.id === s.id) setCurrentSession(null); }}
                        className="absolute top-4 right-4 text-slate-300 hover:text-red-500 p-1 transition-colors"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                      <div className="space-y-4">
                        <div>
                          <p className="text-sm font-black text-[#1C2333] tracking-tight">{s.feName}</p>
                          <p className="text-[10px] font-code font-bold text-slate-400">DSP: {s.dspId} • {s.date}</p>
                        </div>
                        <div className="flex gap-2">
                          {[{ l: 'PKT', v: s.data.length, c: 'bg-slate-50 text-slate-600' }, { l: 'PEND', v: s.data.filter(r => r.status === 'pending').length, c: 'bg-amber-50 text-amber-600' }, { l: 'RTO', v: s.data.filter(r => r.status === 'rto').length, c: 'bg-red-50 text-red-600' }].map(item => (
                            <div key={item.l} className={cn("flex-1 p-2 rounded-xl text-center", item.c)}>
                              <p className="text-[8px] font-black uppercase opacity-60">{item.l}</p>
                              <p className="text-[11px] font-black">{item.v}</p>
                            </div>
                          ))}
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* Status Tabs */}
            {currentSession && (
              <div className="grid grid-cols-5 gap-4">
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
                      "bg-white p-5 rounded-[1.5rem] shadow-sm border-b-4 transition-all text-center",
                      statusFilter === t.id ? `border-${t.col}-500 bg-${t.col}-50/30 ring-4 ring-${t.col}-500/5` : "border-transparent border-slate-100"
                    )}
                  >
                    <p className={cn("text-3xl font-[800] leading-none", `text-${t.col}-600`)}>{t.val}</p>
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-[0.2em] mt-2">{t.label}</p>
                  </button>
                ))}
              </div>
            )}

            {/* Remark Breakdown (Pending only) */}
            {statusFilter === "pending" && uniqueRemarks.length > 0 && (
              <div className="bg-white rounded-[2rem] p-6 border border-slate-200 shadow-sm animate-in slide-in-from-top-2 duration-300">
                <div className="flex items-center justify-between mb-6">
                  <h4 className="text-[11px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-2">
                    <Filter className="w-3.5 h-3.5" /> Remark Breakdown — Pending
                  </h4>
                  {remarkFilter !== null && (
                    <button onClick={() => setRemarkFilter(null)} className="text-[10px] font-black text-blue-600 hover:underline uppercase tracking-widest">← All Pending</button>
                  )}
                </div>
                <div className="flex flex-wrap gap-3">
                  {uniqueRemarks.map(([rem, count]) => {
                    const isIntact = rem.toLowerCase().includes("intact") || rem.toLowerCase().includes("reject but package");
                    return (
                      <button 
                        key={rem} 
                        onClick={() => setRemarkFilter(rem)}
                        className={cn(
                          "px-5 py-2.5 rounded-full text-[10px] font-black uppercase tracking-widest border-2 transition-all flex items-center gap-2",
                          remarkFilter === rem 
                            ? (isIntact ? "bg-red-600 border-red-600 text-white shadow-lg" : "bg-blue-600 border-blue-600 text-white shadow-lg")
                            : (isIntact ? "bg-red-50 border-red-100 text-red-500 hover:bg-red-100" : "bg-slate-100 border-slate-200 text-slate-500 hover:bg-slate-200")
                        )}
                      >
                        {rem || "No Remark"}
                        <span className={cn("px-2 py-0.5 rounded-md text-[9px]", remarkFilter === rem ? "bg-white/20" : "bg-slate-200 text-slate-600")}>{count}</span>
                      </button>
                    );
                  })}
                </div>
              </div>
            )}

            {/* Action Bar & Table */}
            {currentSession && (
              <div className="bg-white rounded-[2rem] shadow-2xl border border-slate-200 overflow-hidden">
                <div className="p-6 bg-slate-50 border-b border-slate-200 flex items-center justify-between flex-wrap gap-4">
                  <div className="flex gap-3">
                    <button onClick={() => downloadExcel(currentSession)} className="bg-[#2E7D32] text-white px-6 py-3 rounded-2xl font-black text-[10px] uppercase tracking-widest hover:bg-[#1B5E20] transition-all flex items-center gap-2 shadow-lg"><Download className="w-4 h-4" /> Download Excel</button>
                    <button onClick={() => copyTable(filteredRows)} className="bg-[#1565C0] text-white px-6 py-3 rounded-2xl font-black text-[10px] uppercase tracking-widest hover:bg-[#0D47A1] transition-all flex items-center gap-2 shadow-lg"><Copy className="w-4 h-4" /> Copy Table</button>
                    {currentSession.data.some(r => r.selected) && (
                      <button onClick={deleteSelected} className="bg-red-600 text-white px-6 py-3 rounded-2xl font-black text-[10px] uppercase tracking-widest hover:bg-red-700 transition-all flex items-center gap-2 shadow-lg animate-in zoom-in">
                        <Trash2 className="w-4 h-4" /> Delete ({currentSession.data.filter(r => r.selected).length})
                      </button>
                    )}
                  </div>
                  <button onClick={() => { if(confirm("Clear current session?")) setCurrentSession(null); }} className="text-[10px] font-black text-slate-400 hover:text-red-500 uppercase tracking-widest">Clear Session</button>
                </div>
                
                <div className="overflow-x-auto max-h-[500px] scrollbar-thin scrollbar-thumb-slate-200">
                  <table className="w-full text-left border-collapse">
                    <thead className="sticky top-0 z-30">
                      <tr className="bg-[#1C2333] text-white text-[10px] font-black uppercase tracking-[0.2em]">
                        <th className="p-4 w-[34px] text-center"><input type="checkbox" onChange={(e) => setCurrentSession({...currentSession, data: currentSession.data.map(r => ({...r, selected: e.target.checked}))})} /></th>
                        <th className="p-4 w-[30px] text-center">✕</th>
                        <th className="p-4 w-[105px]">DSP ID</th>
                        <th className="p-4 w-[155px]">AWB Number</th>
                        <th className="p-4 w-[120px]">Client</th>
                        <th className="p-4 w-[120px]">Order ID</th>
                        <th className="p-4">Remark</th>
                        <th className="p-4 w-[115px]">FE Name</th>
                      </tr>
                    </thead>
                    <tbody>
                      {/* DSP Group Banner */}
                      <tr className="bg-gradient-to-r from-[#0D1B2E] to-[#1A2F4A] text-white">
                        <td colSpan={8} className="p-3">
                          <div className="flex items-center justify-between px-4">
                            <div className="flex items-center gap-4">
                              <span className="text-[#F9A825] font-code font-black text-sm tracking-widest uppercase">DSP: {currentSession.dspId}</span>
                              <div className="px-3 py-1 rounded-full bg-white/5 border border-white/10 text-[10px] font-black tracking-widest text-amber-400">
                                {currentSession.data.length} PACKETS
                              </div>
                            </div>
                            <div className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
                              {currentSession.feName} · {currentSession.date}
                            </div>
                          </div>
                        </td>
                      </tr>
                      {filteredRows.length === 0 && (
                        <tr><td colSpan={8} className="py-20 text-center text-slate-400 font-bold italic">No records in this status</td></tr>
                      )}
                      {filteredRows.map((row, idx) => {
                        const isIntact = row.remark.toLowerCase().includes("intact") || row.remark.toLowerCase().includes("reject but package");
                        return (
                          <tr 
                            key={row.awb} 
                            onClick={() => setFocusedIndex(idx)}
                            className={cn(
                              "border-b border-slate-50 transition-colors group",
                              row.selected ? "bg-[#DBEAFE]" : isIntact ? "bg-[#FFF5F5] hover:bg-[#FFEBEE]" : "hover:bg-[#F0F7FF]",
                              focusedIndex === idx && "ring-2 ring-blue-500 ring-inset outline-none"
                            )}
                          >
                            <td className="p-4 text-center" onClick={(e) => e.stopPropagation()}><input type="checkbox" checked={row.selected} onChange={() => toggleRowSelection(row.awb)} className="w-4 h-4 rounded" /></td>
                            <td className="p-4 text-center"><button onClick={(e) => { e.stopPropagation(); removeRow(row.awb); }} className="text-slate-200 hover:text-red-500 transition-colors"><X className="w-4 h-4" /></button></td>
                            <td className="p-4 font-code text-xs font-bold text-slate-400">{idx === 0 ? row.dspId : ""}</td>
                            <td className="p-4 font-code text-[11.5px] font-black text-[#1565C0] tracking-wider">{row.awb}</td>
                            <td className="p-4 text-[11px] font-bold text-slate-700">{row.client}</td>
                            <td className="p-4 text-[11px] text-slate-500 font-medium">{row.orderId}</td>
                            <td className="p-4">
                              <span className={cn(
                                "inline-flex px-3 py-1 rounded-full text-[9px] font-black uppercase tracking-widest border",
                                isIntact ? "bg-red-100 text-red-600 border-red-200" : "bg-[#FFFDE7] text-amber-600 border-amber-200"
                              )}>
                                {row.remark || "No Remark"}
                              </span>
                            </td>
                            <td className="p-4 text-[11px] font-bold text-slate-500">{row.feName}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </>
        ) : (
          <div className="space-y-6">
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              {/* Replacer Upload */}
              <div className="bg-white rounded-[2rem] p-8 border border-slate-200 shadow-sm space-y-8">
                <div className="flex items-center gap-3">
                  <div className="p-2 bg-green-50 rounded-xl text-green-600"><AlertCircle className="w-5 h-5" /></div>
                  <h2 className="text-lg font-black tracking-tight">EOD Remark Replacer</h2>
                </div>
                
                <div className="border-4 border-dashed border-green-50 rounded-[2.5rem] p-12 text-center hover:border-green-200 transition-all cursor-pointer group relative">
                  <input type="file" onChange={handleReplacerUpload} className="absolute inset-0 opacity-0 cursor-pointer" accept=".xlsx,.xls,.csv" />
                  <div className="space-y-4">
                    <Download className="w-12 h-12 text-green-200 group-hover:text-green-500 mx-auto transition-colors" />
                    <div>
                      <p className="text-sm font-black text-slate-700">Upload EOD Export File</p>
                      <p className="text-xs text-slate-400 mt-1 uppercase tracking-widest font-bold">Auto-detect NSL Remarks</p>
                    </div>
                  </div>
                </div>

                <div className="grid grid-cols-3 gap-4">
                  {[
                    { l: 'Total Rows', v: replacerStats.total, c: 'text-slate-900 bg-slate-50' },
                    { l: 'Replaced', v: replacerStats.replaced, c: 'text-green-600 bg-green-50' },
                    { l: 'Missing', v: replacerStats.missing, c: 'text-amber-600 bg-amber-50' }
                  ].map(stat => (
                    <div key={stat.l} className={cn("p-5 rounded-[1.5rem] text-center", stat.c)}>
                      <p className="text-[10px] font-black uppercase tracking-widest opacity-60">{stat.l}</p>
                      <p className="text-3xl font-black">{stat.v}</p>
                    </div>
                  ))}
                </div>

                {replacerData.length > 0 && (
                  <div className="flex gap-4">
                    <button className="flex-1 bg-green-600 text-white font-black py-4 rounded-2xl shadow-xl hover:bg-green-700 transition-all flex items-center justify-center gap-2 uppercase tracking-widest text-[11px]">
                      <Download className="w-4 h-4" /> Download Official File
                    </button>
                    <button className="bg-blue-600 text-white p-4 rounded-2xl shadow-xl hover:bg-blue-700 transition-all"><Copy className="w-5 h-5" /></button>
                  </div>
                )}
              </div>

              {/* Mapping Reference */}
              <div className="bg-[#1C2333] rounded-[2.5rem] p-8 text-white shadow-2xl relative overflow-hidden">
                <div className="absolute top-0 right-0 w-32 h-32 bg-green-500/10 blur-[60px] rounded-full" />
                <h3 className="text-sm font-black text-yellow-400 uppercase tracking-[0.3em] mb-8 flex items-center gap-2">
                  <CheckCircle2 className="w-5 h-5" /> Official Mapping Reference
                </h3>
                <div className="overflow-y-auto max-h-[450px] pr-4 scrollbar-thin scrollbar-thumb-white/10">
                  <table className="w-full text-[11px]">
                    <thead className="text-slate-500 font-black uppercase tracking-widest border-b border-white/5">
                      <tr><th className="py-4 text-left w-1/2">NSL / Export Remark</th><th className="py-4 text-left">Official Replacement</th></tr>
                    </thead>
                    <tbody className="divide-y divide-white/5">
                      {Object.entries(REMARK_MAPPING).map(([old, official]) => (
                        <tr key={old} className="hover:bg-white/5 transition-colors group">
                          <td className="py-4 pr-6 text-slate-400 italic group-hover:text-white transition-colors">{old}</td>
                          <td className="py-4 font-bold text-green-400">{official}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

            {/* Preview Table */}
            {replacerData.length > 0 && (
              <div className="bg-white rounded-[2.5rem] shadow-2xl border border-slate-200 overflow-hidden animate-in slide-in-from-bottom-4 duration-500">
                <div className="p-8 border-b border-slate-200 flex items-center justify-between">
                  <h3 className="font-black text-xs uppercase tracking-widest text-slate-400">Replacement Preview — Quick Audit</h3>
                  <div className="flex items-center gap-3">
                    <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Rows: {replacerData.length}</span>
                  </div>
                </div>
                <div className="overflow-x-auto max-h-[600px] scrollbar-thin">
                  <table className="w-full text-left">
                    <thead className="bg-slate-50 sticky top-0 z-10 border-b border-slate-200">
                      <tr className="text-[9px] font-black text-slate-400 uppercase tracking-widest">
                        <th className="p-5">#</th>
                        <th className="p-5">AWB</th>
                        <th className="p-5">Export Remark</th>
                        <th className="p-5">Official Remark</th>
                        <th className="p-5">FE Name</th>
                      </tr>
                    </thead>
                    <tbody>
                      {replacerData.map(r => (
                        <tr key={r.id} className={cn("border-b border-slate-50 text-[11px] transition-colors", r.isReplaced ? "bg-green-50/40 hover:bg-green-100/50" : "bg-amber-50/40 hover:bg-amber-100/50")}>
                          <td className="p-5 text-slate-400 font-bold">{r.id}</td>
                          <td className="p-5 font-code font-black text-[#1C2333] tracking-tighter">{r.awb}</td>
                          <td className="p-5 text-slate-500 italic">{r.oldRemark}</td>
                          <td className="p-5">
                            <span className={cn("font-black uppercase tracking-tighter", r.isReplaced ? "text-green-600" : "text-amber-600")}>
                              {r.officialRemark}
                            </span>
                          </td>
                          <td className="p-5 font-bold text-slate-500">{r.feName}</td>
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

      {/* Intact Modal */}
      {showIntactModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-6 bg-[#1C2333]/90 backdrop-blur-lg animate-in fade-in duration-300" onClick={() => setShowIntactModal(false)}>
          <div className="bg-white w-full max-w-4xl rounded-[3rem] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-300" onClick={e => e.stopPropagation()}>
            <div className="bg-gradient-to-r from-red-600 to-red-800 p-10 text-white relative">
              <button onClick={() => setShowIntactModal(false)} className="absolute top-10 right-10 text-white/40 hover:text-white transition-colors"><X className="w-8 h-8" /></button>
              <h2 className="text-4xl font-black tracking-tighter italic uppercase">Intact Packet Summary</h2>
              <p className="text-[10px] font-black text-white/60 uppercase tracking-widest mt-2">Found {stats.intact} pending intact shipments for verification</p>
            </div>
            <div className="p-10 space-y-8">
              <div className="grid grid-cols-2 gap-4">
                <button className="h-20 bg-slate-900 text-white rounded-3xl font-black text-[11px] uppercase tracking-widest flex items-center justify-center gap-3 shadow-2xl active:scale-95 transition-all"><Copy className="w-6 h-6" /> Copy All AWBs</button>
                <button className="h-20 bg-red-600 text-white rounded-3xl font-black text-[11px] uppercase tracking-widest flex items-center justify-center gap-3 shadow-2xl active:scale-95 transition-all"><Download className="w-6 h-6" /> Download List</button>
              </div>
              <div className="max-h-[400px] overflow-y-auto border border-slate-100 rounded-[2rem] scrollbar-thin">
                <table className="w-full text-left">
                  <thead className="bg-slate-50 sticky top-0"><tr className="text-[10px] font-black text-slate-400 uppercase tracking-widest border-b"><th className="p-5">AWB Number</th><th className="p-5">Client Name</th><th className="p-5">Remark</th><th className="p-5 text-center">Action</th></tr></thead>
                  <tbody>
                    {currentSession?.data.filter(r => r.status === 'pending' && r.remark.toLowerCase().includes('intact')).map(row => (
                      <tr key={row.awb} className="border-b border-slate-50 hover:bg-slate-50 transition-colors">
                        <td className="p-5 font-code text-sm font-black text-red-600 tracking-tighter">{row.awb}</td>
                        <td className="p-5 text-[11px] font-bold text-slate-700">{row.client}</td>
                        <td className="p-5"><span className="bg-red-50 text-red-600 text-[9px] font-black px-2.5 py-1 rounded-full uppercase border border-red-100">{row.remark}</span></td>
                        <td className="p-5 text-center"><button onClick={() => removeRow(row.awb)} className="text-slate-200 hover:text-red-500 transition-colors"><Trash2 className="w-4 h-4" /></button></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Intact Float Badge */}
      {currentSession && stats.intact > 0 && activeTab === 'eod' && (
        <button 
          onClick={() => setShowIntactModal(true)}
          className="fixed bottom-10 right-10 bg-red-600 text-white px-8 py-5 rounded-full shadow-2xl animate-cta-pulse flex items-center gap-3 z-40 border-b-4 border-red-800"
        >
          <AlertCircle className="w-6 h-6 fill-current" />
          <div className="flex flex-col items-start">
            <span className="text-[10px] font-black uppercase tracking-widest leading-none">Intact Alert</span>
            <span className="text-2xl font-black leading-none mt-1">{stats.intact} PKTS</span>
          </div>
        </button>
      )}

      <style jsx global>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500;600&display=swap');
        body { font-family: 'Inter', sans-serif; }
        .font-code { font-family: 'IBM Plex Mono', monospace; }
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 10px; }
        ::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
        
        @keyframes pulse-custom {
          0%, 100% { transform: scale(1); box-shadow: 0 0 0 0 rgba(220, 38, 38, 0.4); }
          50% { transform: scale(1.05); box-shadow: 0 0 30px 10px rgba(220, 38, 38, 0.2); }
        }
        .animate-cta-pulse { animation: pulse-custom 3s infinite; }
      `}</style>
    </div>
  );
}

