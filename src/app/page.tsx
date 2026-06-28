
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
  Filter
} from "lucide-react";
import * as XLSX from "xlsx";

/**
 * @fileOverview POD Management Tool for Delhivery.
 * Consolidates Module 1 (EOD Rejection) and Module 2 (Remark Replacer).
 * Uses localStorage for persistence.
 */

// --- Constants & Types ---

const STORAGE_KEY = "pod_master_v1";

const STATUS_MAPPING: Record<string, string> = {
  "pending": "pending",
  "dispatched": "dispatch",
  "dispatch": "dispatch",
  "rto": "rto",
  "dto": "dto",
  "delivered": "delivered"
};

const REMARK_DICTIONARY: Record<string, string> = {
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
  
  // Remark Replacer State
  const [remarkReplacerData, setRemarkReplacerData] = useState<any[]>([]);
  const [remarkStats, setRemarkStats] = useState({ total: 0, replaced: 0, missing: 0 });

  // Load persistence
  useEffect(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        setSessions(JSON.parse(saved));
      } catch (e) {
        console.error("Failed to load sessions", e);
      }
    }
  }, []);

  // Save persistence
  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(sessions));
  }, [sessions]);

  // Auto-save session setup
  useEffect(() => {
    const timer = setTimeout(() => {
      if (setupData.feName && setupData.dspId) {
        // Just a hint that data is ready to be saved with file
      }
    }, 800);
    return () => clearTimeout(timer);
  }, [setupData.feName, setupData.dspId]);

  // Keyboard Navigation
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === "ArrowLeft") {
        const tabs = ["all", "pending", "dispatch", "rto", "dto"];
        const idx = tabs.indexOf(statusFilter);
        if (idx > 0) setStatusFilter(tabs[idx - 1]);
      }
      if (e.key === "ArrowRight") {
        const tabs = ["all", "pending", "dispatch", "rto", "dto"];
        const idx = tabs.indexOf(statusFilter);
        if (idx < tabs.length - 1) setStatusFilter(tabs[idx + 1]);
      }
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [statusFilter]);

  // --- Handlers ---

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const rawData = XLSX.utils.sheet_to_json(ws);

      const parsedRows: PODRow[] = rawData.map((row: any) => {
        // Fuzzy column detection
        const keys = Object.keys(row);
        const findVal = (regex: RegExp) => {
          const key = keys.find(k => regex.test(k.toLowerCase()));
          return key ? row[key] : "";
        };

        let awb = String(findVal(/awb|waybill|number/));
        // AWB Scientific Notation Fix
        if (awb.includes('E+') || awb.includes('e+')) {
          awb = Math.round(Number(awb)).toString();
        }

        const statusRaw = String(findVal(/status|currentstatus/)).toLowerCase();
        const status = STATUS_MAPPING[statusRaw] || "unknown";

        return {
          awb,
          client: String(findVal(/client|name/)),
          orderId: String(findVal(/order|id/)),
          status,
          remark: String(findVal(/remark/)),
          feName: setupData.feName || String(findVal(/fe|biker/)),
          dspId: setupData.dspId || String(findVal(/dsp/)),
          date: setupData.date || String(findVal(/date/)),
          selected: false
        };
      }).filter(row => row.awb.length >= 3 && row.status !== "unknown");

      if (parsedRows.length > 0) {
        const newSession: Session = {
          id: crypto.randomUUID(),
          feName: setupData.feName || "Unknown FE",
          dspId: setupData.dspId || "Unknown DSP",
          date: setupData.date,
          data: parsedRows,
          timestamp: Date.now()
        };
        setCurrentSession(newSession);
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleSaveSession = () => {
    if (currentSession) {
      setSessions([currentSession, ...sessions]);
      setCurrentSession(null);
      setSetupData({ feName: "", dspId: "", date: new Date().toISOString().split('T')[0] });
    }
  };

  const deleteSession = (id: string) => {
    setSessions(sessions.filter(s => s.id !== id));
    if (currentSession?.id === id) setCurrentSession(null);
  };

  const toggleRowSelection = (awb: string) => {
    if (!currentSession) return;
    const newData = currentSession.data.map(r => r.awb === awb ? { ...r, selected: !r.selected } : r);
    setCurrentSession({ ...currentSession, data: newData });
  };

  const deleteSelectedRows = () => {
    if (!currentSession) return;
    const newData = currentSession.data.filter(r => !r.selected);
    setCurrentSession({ ...currentSession, data: newData });
  };

  const removeRow = (awb: string) => {
    if (!currentSession) return;
    const newData = currentSession.data.filter(r => r.awb !== awb);
    setCurrentSession({ ...currentSession, data: newData });
  };

  // --- Module 2 Replacer Logic ---
  const handleRemarkReplacerUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const rawData = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);

      let replacedCount = 0;
      let missingCount = 0;

      const processed = rawData.map((row: any, idx) => {
        const keys = Object.keys(row);
        const findVal = (regex: RegExp) => {
          const key = keys.find(k => regex.test(k.toLowerCase()));
          return key ? row[key] : "";
        };

        const oldRemark = String(findVal(/remark/));
        const officialRemark = REMARK_DICTIONARY[oldRemark];
        
        if (officialRemark) replacedCount++;
        else missingCount++;

        return {
          id: idx + 1,
          date: String(findVal(/date/)),
          dsp: String(findVal(/dsp/)),
          awb: String(findVal(/awb/)),
          client: String(findVal(/client/)),
          oldRemark,
          officialRemark: officialRemark || "No Mapping Found",
          isReplaced: !!officialRemark,
          feName: String(findVal(/fe/))
        };
      });

      setRemarkReplacerData(processed);
      setRemarkStats({ total: processed.length, replaced: replacedCount, missing: missingCount });
    };
    reader.readAsBinaryString(file);
  };

  // --- Exports ---

  const copyTableToClipboard = (dataRows: PODRow[]) => {
    const text = dataRows.map((r, i) => {
      const dsp = i === 0 ? r.dspId : "";
      return `${r.date}\t${dsp}\t${r.awb}\t${r.client}\t${r.orderId}\t${r.remark}\t${r.feName}`;
    }).join("\n");
    
    navigator.clipboard.writeText(text).then(() => {
      alert("Table copied successfully (No Headers)!");
    }).catch(() => {
      // fallback
      const el = document.createElement('textarea');
      el.value = text;
      document.body.appendChild(el);
      el.select();
      document.execCommand('copy');
      document.body.removeChild(el);
    });
  };

  const downloadExcel = (session: Session) => {
    const wsData = session.data.map(r => ({
      "Date": r.date,
      "DSP No": r.dspId,
      "AWB Number": r.awb,
      "Client": r.client,
      "Order ID": r.orderId,
      "Remark": r.remark,
      "FE Name": r.feName
    }));

    const ws = XLSX.utils.json_to_sheet(wsData);
    
    // Formatting
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cell_ref = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[cell_ref]) continue;
        
        // Force AWB as text
        if (C === 2 && R > 0) {
          ws[cell_ref].t = 's';
          ws[cell_ref].z = '@';
        }
      }
    }

    ws['!cols'] = [
      { wch: 13 }, { wch: 12 }, { wch: 26 }, { wch: 20 }, { wch: 20 }, { wch: 36 }, { wch: 18 }
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "POD_Report");
    XLSX.writeFile(wb, `Delhivery_POD_${session.feName}_${session.date}.xlsx`);
  };

  // --- Computed ---

  const filteredRows = useMemo(() => {
    if (!currentSession) return [];
    let rows = currentSession.data;
    if (statusFilter !== "all") {
      rows = rows.filter(r => r.status === statusFilter);
    }
    if (remarkFilter) {
      rows = rows.filter(r => (r.remark || "") === (remarkFilter === "No Remark" ? "" : remarkFilter));
    }
    return rows;
  }, [currentSession, statusFilter, remarkFilter]);

  const stats = useMemo(() => {
    if (!currentSession) return { total: 0, pending: 0, dispatch: 0, rto: 0, dto: 0, intact: 0 };
    return {
      total: currentSession.data.length,
      pending: currentSession.data.filter(r => r.status === "pending").length,
      dispatch: currentSession.data.filter(r => r.status === "dispatch").length,
      rto: currentSession.data.filter(r => r.status === "rto").length,
      dto: currentSession.data.filter(r => r.status === "dto").length,
      intact: currentSession.data.filter(r => r.status === "pending" && (r.remark || "").toLowerCase().includes("intact")).length
    };
  }, [currentSession]);

  const uniqueRemarks = useMemo(() => {
    if (!currentSession || statusFilter !== "pending") return [];
    const pending = currentSession.data.filter(r => r.status === "pending");
    const counts: Record<string, number> = {};
    pending.forEach(r => {
      const key = r.remark || "No Remark";
      counts[key] = (counts[key] || 0) + 1;
    });
    return Object.entries(counts).map(([label, count]) => ({ label, count }));
  }, [currentSession, statusFilter]);

  // --- Render ---

  return (
    <div className="min-h-screen bg-[#F0F4FA] font-body text-[#1C2333]">
      {/* Rainbow Header */}
      <div className="h-[3px] w-full bg-gradient-to-r from-blue-500 via-yellow-400 to-green-500 red-500" />
      <header className="bg-[#1C2333] px-6 py-4 flex flex-col md:flex-row items-center justify-between text-white shadow-lg">
        <div className="flex items-center gap-4">
          <div className="bg-gradient-to-br from-blue-400 to-blue-600 p-2 rounded-xl shadow-md">
            <Truck className="w-8 h-8 text-white" />
          </div>
          <div className="flex flex-col">
            <div className="flex items-center gap-2">
              <h1 className="text-xl font-bold tracking-tight">POD Management Tool</h1>
              <span className="flex items-center gap-1 bg-green-900/40 text-green-400 text-[10px] font-bold px-2 py-0.5 rounded-full border border-green-500/30">
                <span className="w-1.5 h-1.5 bg-green-400 rounded-full animate-pulse" /> LIVE
              </span>
            </div>
            <p className="text-xs text-slate-400 font-medium">Delhivery · Palam Vihar RPC · By Ashu</p>
          </div>
        </div>

        <div className="mt-4 md:mt-0 flex items-center gap-2 bg-slate-800/50 px-4 py-2 rounded-2xl border border-white/10">
          <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Active Records</span>
          <span className="text-sm font-black text-yellow-400">{currentSession?.data.length || 0}</span>
        </div>
      </header>

      {/* Tabs Nav */}
      <nav className="bg-[#1C2333] px-6 border-t border-white/5 flex gap-8">
        <button 
          onClick={() => setActiveTab("eod")}
          className={`py-4 text-xs font-black tracking-widest uppercase transition-all relative ${activeTab === 'eod' ? 'text-yellow-400' : 'text-slate-400 hover:text-white'}`}
        >
          Daily EOD Rejection
          {activeTab === 'eod' && <div className="absolute bottom-0 left-0 w-full h-1 bg-yellow-400 rounded-t-full" />}
        </button>
        <button 
          onClick={() => setActiveTab("remark")}
          className={`py-4 text-xs font-black tracking-widest uppercase transition-all relative ${activeTab === 'remark' ? 'text-yellow-400' : 'text-slate-400 hover:text-white'}`}
        >
          EOD Rejection Remark
          {activeTab === 'remark' && <div className="absolute bottom-0 left-0 w-full h-1 bg-yellow-400 rounded-t-full" />}
        </button>
      </nav>

      <main className="p-6 max-w-[1400px] mx-auto space-y-6">
        
        {activeTab === "eod" ? (
          <>
            {/* Session Setup */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              <div className="bg-white border border-slate-200 rounded-3xl p-6 shadow-sm space-y-6">
                <div className="flex items-center gap-3">
                  <div className="p-2 bg-blue-50 rounded-lg text-blue-600">
                    <Plus className="w-5 h-5" />
                  </div>
                  <h2 className="font-bold">Session Setup</h2>
                </div>
                
                <div className="space-y-4">
                  <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1">
                      <label className="text-[10px] font-bold text-slate-400 uppercase">FE / Biker Name</label>
                      <input 
                        type="text" 
                        value={setupData.feName}
                        onChange={(e) => setSetupData({...setupData, feName: e.target.value})}
                        className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-sm focus:ring-2 ring-blue-500/20 outline-none transition-all"
                        placeholder="Ashu..."
                      />
                    </div>
                    <div className="space-y-1">
                      <label className="text-[10px] font-bold text-slate-400 uppercase">DSP ID</label>
                      <input 
                        type="text" 
                        value={setupData.dspId}
                        onChange={(e) => setSetupData({...setupData, dspId: e.target.value})}
                        className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-sm focus:ring-2 ring-blue-500/20 outline-none transition-all"
                        placeholder="DSP001"
                      />
                    </div>
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] font-bold text-slate-400 uppercase">Date</label>
                    <input 
                      type="date" 
                      value={setupData.date}
                      onChange={(e) => setSetupData({...setupData, date: e.target.value})}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-sm outline-none"
                    />
                  </div>
                  
                  <div className="border-2 border-dashed border-slate-200 rounded-2xl p-6 text-center hover:border-blue-500 transition-colors cursor-pointer group relative">
                    <input type="file" onChange={handleFileUpload} className="absolute inset-0 opacity-0 cursor-pointer" accept=".xlsx,.xls,.csv,.tsv,.ods" />
                    <div className="space-y-2">
                      <Download className="w-8 h-8 text-slate-300 group-hover:text-blue-500 mx-auto transition-colors" />
                      <p className="text-xs font-bold text-slate-500">Drag & Drop or Click to Upload POD</p>
                      <p className="text-[10px] text-slate-400">Excel, CSV, TSV Supported</p>
                    </div>
                  </div>

                  <button 
                    onClick={handleSaveSession}
                    disabled={!currentSession}
                    className="w-full bg-[#1C2333] text-white font-bold py-3 rounded-2xl shadow-lg hover:bg-slate-800 disabled:opacity-50 transition-all flex items-center justify-center gap-2"
                  >
                    Save Current Session
                  </button>
                </div>
              </div>

              {/* Sessions Grid */}
              <div className="lg:col-span-2 space-y-4">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-3">
                    <div className="p-2 bg-yellow-50 rounded-lg text-yellow-600">
                      <Filter className="w-5 h-5" />
                    </div>
                    <h2 className="font-bold">Recent Sessions</h2>
                  </div>
                  <button 
                    onClick={() => { if(confirm("Clear All?")) {setSessions([]); localStorage.removeItem(STORAGE_KEY);}}}
                    className="text-[10px] font-bold text-red-500 uppercase tracking-widest hover:bg-red-50 px-3 py-1.5 rounded-full transition-colors"
                  >
                    Clear All Sessions
                  </button>
                </div>
                
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {sessions.length === 0 && <div className="col-span-2 py-12 text-center text-slate-400 font-bold border-2 border-dashed rounded-3xl">No Saved Sessions</div>}
                  {sessions.map(s => (
                    <div 
                      key={s.id} 
                      onClick={() => setCurrentSession(s)}
                      className={`bg-white border p-5 rounded-[2rem] shadow-sm hover:shadow-md transition-all cursor-pointer group relative ${currentSession?.id === s.id ? 'border-blue-500 ring-4 ring-blue-500/5' : 'border-slate-200'}`}
                    >
                      <button 
                        onClick={(e) => { e.stopPropagation(); deleteSession(s.id); }}
                        className="absolute top-4 right-4 p-2 text-slate-300 hover:text-red-500 transition-colors"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                      <div className="space-y-3">
                        <div className="flex items-center gap-2">
                          <div className="w-8 h-8 bg-blue-100 rounded-full flex items-center justify-center text-blue-600 text-xs font-black">
                            {s.feName[0]}
                          </div>
                          <div>
                            <p className="text-sm font-black text-slate-900">{s.feName}</p>
                            <p className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter">DSP: {s.dspId} • {s.date}</p>
                          </div>
                        </div>
                        <div className="flex gap-2">
                          <div className="flex-1 bg-slate-50 p-2 rounded-xl text-center">
                            <p className="text-[9px] font-bold text-slate-400 uppercase">Pkt</p>
                            <p className="text-xs font-black">{s.data.length}</p>
                          </div>
                          <div className="flex-1 bg-amber-50 p-2 rounded-xl text-center">
                            <p className="text-[9px] font-bold text-amber-500 uppercase">Pend</p>
                            <p className="text-xs font-black text-amber-600">{s.data.filter(r => r.status === 'pending').length}</p>
                          </div>
                          <div className="flex-1 bg-red-50 p-2 rounded-xl text-center">
                            <p className="text-[9px] font-bold text-red-500 uppercase">RTO</p>
                            <p className="text-xs font-black text-red-600">{s.data.filter(r => r.status === 'rto').length}</p>
                          </div>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* Main Table Area */}
            {currentSession && (
              <div className="bg-white border border-slate-200 rounded-[2.5rem] shadow-xl overflow-hidden transition-all animate-in fade-in slide-in-from-bottom-4">
                {/* Stats Bar */}
                <div className="p-1 flex flex-wrap bg-[#1C2333]">
                  {[
                    { id: 'all', label: 'All Records', count: stats.total, color: 'bg-blue-600', text: 'text-white' },
                    { id: 'pending', label: 'Pending', count: stats.pending, color: 'bg-amber-500', text: 'text-white' },
                    { id: 'dispatch', label: 'Dispatched', count: stats.dispatch, color: 'bg-blue-500', text: 'text-white' },
                    { id: 'rto', label: 'RTO / DTO', count: stats.rto + stats.dto, color: 'bg-red-500', text: 'text-white' },
                  ].map(tab => (
                    <button 
                      key={tab.id}
                      onClick={() => { setStatusFilter(tab.id); setRemarkFilter(null); }}
                      className={`flex-1 min-w-[120px] p-4 text-center transition-all ${statusFilter === tab.id ? 'bg-white text-[#1C2333] rounded-t-2xl' : 'text-slate-400 hover:text-white'}`}
                    >
                      <p className="text-[10px] font-bold uppercase tracking-widest">{tab.label}</p>
                      <p className="text-xl font-black">{tab.count}</p>
                    </button>
                  ))}
                  {stats.intact > 0 && (
                    <button 
                      onClick={() => setShowIntactModal(true)}
                      className="flex-1 min-w-[120px] p-4 text-center bg-red-900/40 text-red-400 animate-pulse hover:bg-red-800 transition-colors"
                    >
                      <p className="text-[10px] font-bold uppercase tracking-widest">Intact Pkts</p>
                      <p className="text-xl font-black">{stats.intact}</p>
                    </button>
                  )}
                </div>

                {/* Filters & Actions */}
                <div className="p-6 bg-slate-50 border-b border-slate-200 flex flex-col md:flex-row items-center justify-between gap-4">
                  <div className="flex items-center gap-2 flex-wrap">
                    {statusFilter === "pending" && (
                      <>
                        <button 
                          onClick={() => setRemarkFilter(null)}
                          className={`text-[10px] font-black px-4 py-2 rounded-full uppercase tracking-widest border-2 transition-all ${!remarkFilter ? 'bg-[#1C2333] text-white border-[#1C2333]' : 'bg-white border-slate-200 text-slate-500'}`}
                        >
                          All Pending
                        </button>
                        {uniqueRemarks.map(rem => (
                          <button 
                            key={rem.label}
                            onClick={() => setRemarkFilter(rem.label)}
                            className={`text-[10px] font-black px-4 py-2 rounded-full uppercase tracking-widest border-2 transition-all flex items-center gap-2 ${remarkFilter === rem.label ? 'bg-amber-500 text-white border-amber-500 shadow-md' : 'bg-white border-slate-200 text-slate-500 hover:border-amber-200'}`}
                          >
                            {rem.label}
                            <span className={`px-1.5 py-0.5 rounded-md text-[9px] ${remarkFilter === rem.label ? 'bg-white/20' : 'bg-slate-100'}`}>{rem.count}</span>
                          </button>
                        ))}
                      </>
                    )}
                  </div>

                  <div className="flex items-center gap-3">
                    {currentSession.data.some(r => r.selected) && (
                      <button 
                        onClick={deleteSelectedRows}
                        className="bg-red-500 text-white p-3 rounded-2xl hover:bg-red-600 transition-colors shadow-lg shadow-red-500/20"
                        title="Delete Selected"
                      >
                        <Trash2 className="w-5 h-5" />
                      </button>
                    )}
                    <button 
                      onClick={() => copyTableToClipboard(filteredRows)}
                      className="bg-blue-600 text-white px-6 py-3 rounded-2xl font-black text-xs uppercase tracking-widest hover:bg-blue-700 transition-all shadow-lg shadow-blue-500/20 flex items-center gap-2"
                    >
                      <Copy className="w-4 h-4" /> Copy Table
                    </button>
                    <button 
                      onClick={() => downloadExcel(currentSession)}
                      className="bg-green-600 text-white px-6 py-3 rounded-2xl font-black text-xs uppercase tracking-widest hover:bg-green-700 transition-all shadow-lg shadow-green-500/20 flex items-center gap-2"
                    >
                      <Download className="w-4 h-4" /> Download Excel
                    </button>
                  </div>
                </div>

                {/* Table */}
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="bg-[#1C2333] text-white text-[10px] font-black uppercase tracking-[0.2em] sticky top-0 z-10">
                        <th className="p-5 w-12 text-center">
                          <input 
                            type="checkbox" 
                            onChange={(e) => {
                              const checked = e.target.checked;
                              setCurrentSession({...currentSession, data: currentSession.data.map(r => ({...r, selected: checked}))});
                            }}
                          />
                        </th>
                        <th className="p-5 w-12">✕</th>
                        <th className="p-5 w-32">DSP ID</th>
                        <th className="p-5 w-48">AWB Number</th>
                        <th className="p-5">Client Name</th>
                        <th className="p-5">Order ID</th>
                        <th className="p-5">Remark</th>
                        <th className="p-5">FE Name</th>
                      </tr>
                    </thead>
                    <tbody>
                      {/* DSP Group Header */}
                      <tr className="bg-gradient-to-r from-slate-800 to-slate-900 text-white">
                        <td colSpan={8} className="p-3">
                          <div className="flex items-center justify-between px-4">
                            <div className="flex items-center gap-4">
                              <span className="bg-yellow-400 text-black text-[10px] font-black px-3 py-1 rounded-full uppercase tracking-widest">
                                DSP: {currentSession.dspId}
                              </span>
                              <span className="text-xs font-bold text-slate-300">
                                Total Pkts: <span className="text-white text-sm font-black">{currentSession.data.length}</span>
                              </span>
                            </div>
                            <p className="text-xs font-bold text-slate-400 uppercase tracking-[0.3em]">Managed By: {currentSession.feName}</p>
                          </div>
                        </td>
                      </tr>

                      {filteredRows.map((row, idx) => {
                        const isIntact = (row.remark || "").toLowerCase().includes("intact");
                        return (
                          <tr 
                            key={row.awb} 
                            className={`group border-b border-slate-100 transition-colors ${row.selected ? 'bg-blue-50' : isIntact ? 'bg-red-50/50' : 'hover:bg-slate-50'}`}
                          >
                            <td className="p-4 text-center">
                              <input 
                                type="checkbox" 
                                checked={row.selected}
                                onChange={() => toggleRowSelection(row.awb)}
                                className="w-4 h-4 rounded cursor-pointer"
                              />
                            </td>
                            <td className="p-4">
                              <button 
                                onClick={() => removeRow(row.awb)}
                                className="text-slate-300 hover:text-red-500 opacity-0 group-hover:opacity-100 transition-all"
                              >
                                <X className="w-4 h-4" />
                              </button>
                            </td>
                            <td className="p-4 text-xs font-bold text-slate-400">
                              {idx === 0 ? row.dspId : ""}
                            </td>
                            <td className="p-4 font-code text-sm font-black text-blue-600 tracking-wider">
                              {row.awb}
                            </td>
                            <td className="p-4 text-sm font-bold text-slate-700">{row.client}</td>
                            <td className="p-4 text-sm text-slate-500 font-medium">{row.orderId}</td>
                            <td className="p-4">
                              <span className={`inline-flex px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-widest ${isIntact ? 'bg-red-100 text-red-600 border border-red-200 shadow-sm' : 'bg-amber-100 text-amber-600 border border-amber-200'}`}>
                                {row.remark || "No Remark"}
                              </span>
                            </td>
                            <td className="p-4 text-xs font-bold text-slate-500">{row.feName}</td>
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
          /* Module 2: Remark Replacer */
          <div className="space-y-6">
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              {/* Controls */}
              <div className="bg-white border border-slate-200 rounded-3xl p-6 shadow-sm space-y-6">
                <div className="flex items-center gap-3">
                  <div className="p-2 bg-blue-50 rounded-lg text-blue-600">
                    <AlertCircle className="w-5 h-5" />
                  </div>
                  <h2 className="font-bold">EOD Remark Replacer</h2>
                </div>

                <div className="space-y-4">
                  <div className="border-2 border-dashed border-slate-200 rounded-2xl p-8 text-center hover:border-blue-500 transition-colors cursor-pointer group relative">
                    <input type="file" onChange={handleRemarkReplacerUpload} className="absolute inset-0 opacity-0 cursor-pointer" accept=".xlsx,.xls,.csv" />
                    <div className="space-y-3">
                      <Download className="w-10 h-10 text-slate-300 group-hover:text-blue-500 mx-auto transition-colors" />
                      <p className="text-sm font-bold text-slate-600">Upload EOD Export File</p>
                      <p className="text-xs text-slate-400">Auto-detect NSL Remarks</p>
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <div className="bg-slate-50 p-4 rounded-2xl text-center">
                      <p className="text-[10px] font-bold text-slate-400 uppercase">Total Rows</p>
                      <p className="text-2xl font-black">{remarkStats.total}</p>
                    </div>
                    <div className="bg-green-50 p-4 rounded-2xl text-center">
                      <p className="text-[10px] font-bold text-green-500 uppercase">Replaced</p>
                      <p className="text-2xl font-black text-green-600">{remarkStats.replaced}</p>
                    </div>
                  </div>

                  {remarkReplacerData.length > 0 && (
                    <button 
                      onClick={() => {
                        const wsData = remarkReplacerData.map(r => ({
                          "Date": r.date,
                          "DSP No": r.dsp,
                          "AWB Number": r.awb,
                          "Client Name": r.client,
                          "Official Remark": r.officialRemark,
                          "FE Name": r.feName
                        }));
                        const ws = XLSX.utils.json_to_sheet(wsData);
                        const wb = XLSX.utils.book_new();
                        XLSX.utils.book_append_sheet(wb, ws, "Replaced_Remarks");
                        XLSX.writeFile(wb, "Delhivery_Official_Remarks.xlsx");
                      }}
                      className="w-full bg-green-600 text-white font-bold py-4 rounded-2xl shadow-xl hover:bg-green-700 transition-all flex items-center justify-center gap-2"
                    >
                      <Download className="w-5 h-5" /> Download Official File
                    </button>
                  )}
                </div>
              </div>

              {/* Mapping Reference Table */}
              <div className="lg:col-span-2 bg-[#1C2333] rounded-3xl p-6 text-white shadow-2xl space-y-4">
                <div className="flex items-center justify-between">
                  <h3 className="font-bold flex items-center gap-2 text-yellow-400">
                    <CheckCircle2 className="w-5 h-5" /> Official Mapping Reference
                  </h3>
                  <span className="text-[10px] font-black text-slate-400 uppercase tracking-[0.2em]">12 Rules Loaded</span>
                </div>
                <div className="overflow-y-auto max-h-[400px] scrollbar-thin scrollbar-thumb-slate-700">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="text-slate-500 border-b border-white/5">
                        <th className="py-3 text-left w-1/2">NSL / Export Remark</th>
                        <th className="py-3 text-left">Official Replacement</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-white/5">
                      {Object.entries(REMARK_DICTIONARY).map(([old, official]) => (
                        <tr key={old} className="hover:bg-white/5 transition-colors">
                          <td className="py-3 pr-4 text-slate-400 italic">{old}</td>
                          <td className="py-3 font-bold text-green-400">{official}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

            {/* Preview Table */}
            {remarkReplacerData.length > 0 && (
              <div className="bg-white border border-slate-200 rounded-[2.5rem] shadow-xl overflow-hidden">
                <div className="p-6 border-b border-slate-200 flex items-center justify-between">
                  <h3 className="font-bold">Replacement Preview</h3>
                  <button 
                    onClick={() => {
                      const text = remarkReplacerData.map(r => `${r.date}\t${r.dsp}\t${r.awb}\t${r.client}\t${r.officialRemark}\t${r.feName}`).join("\n");
                      navigator.clipboard.writeText(text).then(() => alert("Preview copied!"));
                    }}
                    className="flex items-center gap-2 text-[10px] font-black uppercase tracking-widest text-blue-600 bg-blue-50 px-4 py-2 rounded-full hover:bg-blue-100 transition-all"
                  >
                    <Copy className="w-4 h-4" /> Copy Preview Table
                  </button>
                </div>
                <div className="overflow-x-auto max-h-[600px]">
                  <table className="w-full text-left">
                    <thead className="bg-slate-50 sticky top-0 z-10">
                      <tr className="text-[10px] font-black text-slate-400 uppercase tracking-widest border-b border-slate-200">
                        <th className="p-4">#</th>
                        <th className="p-4">AWB</th>
                        <th className="p-4">Export Remark</th>
                        <th className="p-4">Official Remark</th>
                        <th className="p-4">FE Name</th>
                      </tr>
                    </thead>
                    <tbody>
                      {remarkReplacerData.map(r => (
                        <tr key={r.id} className={`border-b border-slate-50 text-xs ${r.isReplaced ? 'bg-green-50/50' : 'bg-yellow-50/50'}`}>
                          <td className="p-4 text-slate-400 font-bold">{r.id}</td>
                          <td className="p-4 font-code font-black text-[#1C2333]">{r.awb}</td>
                          <td className="p-4 text-slate-500 italic">{r.oldRemark}</td>
                          <td className="p-4">
                            <span className={`font-black uppercase tracking-tighter ${r.isReplaced ? 'text-green-600' : 'text-amber-600'}`}>
                              {r.officialRemark}
                            </span>
                          </td>
                          <td className="p-4 font-bold text-slate-500">{r.feName}</td>
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
        <div className="fixed inset-0 z-50 flex items-center justify-center p-6 bg-[#1C2333]/80 backdrop-blur-md animate-in fade-in duration-300" onClick={() => setShowIntactModal(false)}>
          <div className="bg-white w-full max-w-4xl rounded-[3rem] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-300" onClick={e => e.stopPropagation()}>
            <div className="bg-red-600 p-8 text-white relative">
              <button onClick={() => setShowIntactModal(false)} className="absolute top-8 right-8 text-white/60 hover:text-white transition-colors">
                <X className="w-8 h-8" />
              </button>
              <h2 className="text-3xl font-black font-headline tracking-tighter">INTACT PKT SUMMARY</h2>
              <p className="text-xs font-bold text-white/60 uppercase tracking-widest mt-1">Found {stats.intact} pending intact shipments</p>
            </div>
            
            <div className="p-8 space-y-6">
              <div className="flex gap-4">
                <button 
                  onClick={() => {
                    const intacts = currentSession?.data.filter(r => r.status === "pending" && (r.remark || "").toLowerCase().includes("intact")) || [];
                    const awbs = intacts.map(r => r.awb).join("\n");
                    navigator.clipboard.writeText(awbs);
                    alert("AWBs Copied!");
                  }}
                  className="flex-1 bg-slate-900 text-white h-16 rounded-2xl font-black text-xs uppercase tracking-widest flex items-center justify-center gap-3 hover:bg-slate-800 transition-all shadow-xl shadow-slate-900/20"
                >
                  <Copy className="w-5 h-5" /> Copy AWBs
                </button>
                <button 
                  onClick={() => {
                    if (!currentSession) return;
                    const intacts = currentSession.data.filter(r => r.status === "pending" && (r.remark || "").toLowerCase().includes("intact"));
                    const ws = XLSX.utils.json_to_sheet(intacts);
                    const wb = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb, ws, "Intact_Pkts");
                    XLSX.writeFile(wb, "Delhivery_Intact_Report.xlsx");
                  }}
                  className="flex-1 bg-red-600 text-white h-16 rounded-2xl font-black text-xs uppercase tracking-widest flex items-center justify-center gap-3 hover:bg-red-700 transition-all shadow-xl shadow-red-500/20"
                >
                  <Download className="w-5 h-5" /> Download List
                </button>
              </div>

              <div className="max-h-[400px] overflow-y-auto border border-slate-100 rounded-2xl scrollbar-thin scrollbar-thumb-slate-200">
                <table className="w-full text-left">
                  <thead className="bg-slate-50 sticky top-0">
                    <tr className="text-[10px] font-black text-slate-400 uppercase tracking-widest border-b">
                      <th className="p-4">AWB Number</th>
                      <th className="p-4">Client</th>
                      <th className="p-4">Remark</th>
                      <th className="p-4 text-center">Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {currentSession?.data.filter(r => r.status === "pending" && (r.remark || "").toLowerCase().includes("intact")).map(row => (
                      <tr key={row.awb} className="border-b border-slate-50 hover:bg-slate-50 transition-colors">
                        <td className="p-4 font-code text-sm font-black text-red-600">{row.awb}</td>
                        <td className="p-4 text-xs font-bold text-slate-700">{row.client}</td>
                        <td className="p-4">
                          <span className="bg-red-100 text-red-600 text-[10px] font-black px-2 py-1 rounded-md uppercase border border-red-200">{row.remark}</span>
                        </td>
                        <td className="p-4 text-center">
                          <button onClick={() => removeRow(row.awb)} className="text-slate-300 hover:text-red-500">
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Font & Global CSS */}
      <style jsx global>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;900&family=IBM+Plex+Mono:wght@500;700&display=swap');
        
        body {
          font-family: 'Inter', sans-serif;
        }

        .font-code {
          font-family: 'IBM Plex Mono', monospace;
        }

        ::-webkit-scrollbar {
          width: 8px;
          height: 8px;
        }
        ::-webkit-scrollbar-track {
          background: #f1f5f9;
        }
        ::-webkit-scrollbar-thumb {
          background: #cbd5e1;
          border-radius: 10px;
        }
        ::-webkit-scrollbar-thumb:hover {
          background: #94a3b8;
        }
      `}</style>
    </div>
  );
}
