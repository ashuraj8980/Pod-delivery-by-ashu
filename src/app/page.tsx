
"use client";

import React, { useState, useEffect, useMemo, useCallback } from "react";
import { 
  Truck, 
  Trash2, 
  Download, 
  Copy, 
  X, 
  AlertCircle, 
  User,
  Hash,
  CheckCircle2,
  FileSpreadsheet,
  Layers,
  Search,
  Check
} from "lucide-react";
import * as XLSX from "xlsx";
import { cn } from "@/lib/utils";

/**
 * @fileOverview Delhivery POD Management Tool v5.0 (Elite Edition)
 * Finalized for Palam Vihar RPC.
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
  const [selectedSessionId, setSelectedSessionId] = useState<string | null>(null);
  const [statusFilter, setStatusFilter] = useState<string>("all");
  const [remarkFilter, setRemarkFilter] = useState<string | null>(null);
  const [setupData, setSetupData] = useState({ feName: "", dspId: "", date: "" });
  const [isMounted, setIsMounted] = useState(false);
  const [focusedIndex, setFocusedIndex] = useState(-1);
  
  // Module 2 State
  const [replacerData, setReplacerData] = useState<any[]>([]);
  const [replacerStats, setReplacerStats] = useState({ total: 0, replaced: 0, missing: 0 });

  useEffect(() => {
    setIsMounted(true);
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        setSessions(Array.isArray(parsed) ? parsed : []);
        if (parsed.length > 0) setSelectedSessionId(parsed[0].id);
      } catch (e) {
        console.error("Failed to load sessions");
      }
    }
    setSetupData(prev => ({ ...prev, date: new Date().toISOString().split('T')[0] }));
  }, []);

  useEffect(() => {
    if (isMounted) {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(sessions));
    }
  }, [sessions, isMounted]);

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
    toast.className = `fixed bottom-6 left-1/2 -translate-x-1/2 px-5 py-2.5 rounded-2xl text-[11px] font-black z-[200] shadow-2xl animate-in slide-in-from-bottom-5 duration-300 ${
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
      showToast(`${label}: ${text}`, "ok");
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
        
        if (!rawData || rawData.length === 0) throw new Error("File empty");

        const parsedRows: PODRow[] = rawData.map((row: any) => {
          const keys = Object.keys(row);
          const findVal = (regex: RegExp) => {
            const key = keys.find(k => regex.test(k.toLowerCase().replace(/[\s_-]/g, "")));
            return key ? row[key] : "";
          };
          const awb = fixAWB(findVal(/waybill|awb|awbnumber|waybillno/));
          const statusRaw = String(findVal(/status|currentstatus|current status/)).toLowerCase().trim();
          const status = STATUS_MAP[statusRaw] || "unknown";
          return {
            id: crypto.randomUUID(),
            awb,
            client: String(findVal(/client|clientname/)),
            orderId: String(findVal(/order|orderid|orderno/)),
            status,
            remark: String(findVal(/remark|remarks|remark1/)),
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
          setSessions(prev => [newSession, ...prev]);
          setSelectedSessionId(newSession.id);
          showToast(`Imported ${parsedRows.length} rows!`, "ok");
        } else {
          showToast("No valid rows found! Unknown status dropped.", "err");
        }
      } catch (err) {
        showToast("Error parsing file! Please upload valid Delhivery export.", "err");
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
        showToast("Replacer error! Check your file format.", "err");
      }
    };
    reader.readAsBinaryString(file);
  };

  const currentSession = useMemo(() => sessions.find(s => s.id === selectedSessionId) || null, [sessions, selectedSessionId]);

  const filteredRows = useMemo(() => {
    if (!currentSession) return [];
    let rows = currentSession.data;
    if (statusFilter !== "all") rows = rows.filter(r => r.status === statusFilter);
    if (remarkFilter) rows = rows.filter(r => r.remark === (remarkFilter === "No Remark" ? "" : remarkFilter));
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

  const remarkStats = useMemo(() => {
    if (!currentSession || statusFilter !== 'pending') return [];
    const counts: Record<string, number> = {};
    currentSession.data.filter(r => r.status === 'pending').forEach(r => {
      const label = r.remark || "No Remark";
      counts[label] = (counts[label] || 0) + 1;
    });
    return Object.entries(counts).sort((a,b) => b[1] - a[1]);
  }, [currentSession, statusFilter]);

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

  const toggleRowSelect = (rowId: string) => {
    setSessions(prev => prev.map(s => {
      if (s.id === selectedSessionId) {
        return { ...s, data: s.data.map(r => r.id === rowId ? { ...r, selected: !r.selected } : r) };
      }
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
      intact: currentSession.data.filter(r => r.remark.toLowerCase().includes("intact")).length
    };
  }, [currentSession]);

  // Keyboard Navigation
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (document.activeElement?.tagName === "INPUT") return;

      if (e.key === "ArrowRight") {
        const tabs = ["all", "pending", "dispatched", "rto", "dto"];
        const next = tabs[(tabs.indexOf(statusFilter) + 1) % tabs.length];
        setStatusFilter(next);
      } else if (e.key === "ArrowLeft") {
        const tabs = ["all", "pending", "dispatched", "rto", "dto"];
        const prev = tabs[(tabs.indexOf(statusFilter) - 1 + tabs.length) % tabs.length];
        setStatusFilter(prev);
      } else if (e.key === "Delete" && focusedIndex !== -1 && filteredRows[focusedIndex]) {
        deleteRow(filteredRows[focusedIndex].id);
      } else if (e.key === "Enter" && focusedIndex !== -1 && filteredRows[focusedIndex]) {
        toggleRowSelect(filteredRows[focusedIndex].id);
      } else if (e.key === "ArrowDown") {
        setFocusedIndex(prev => Math.min(prev + 1, filteredRows.length - 1));
      } else if (e.key === "ArrowUp") {
        setFocusedIndex(prev => Math.max(prev - 1, 0));
      }
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [statusFilter, filteredRows, focusedIndex]);

  if (!isMounted) return null;

  return (
    <div className="min-h-screen bg-[#F0F4FA] font-body text-[#1C2333] select-auto overflow-x-hidden">
      {/* Rainbow Stripe */}
      <div className="h-[3px] w-full bg-gradient-to-r from-[#1565C0] via-[#F9A825] via-[#2E7D32] to-[#D32F2F] sticky top-0 z-[100]" />
      
      {/* Header */}
      <header className="h-[55px] bg-[#1C2333] px-6 flex items-center justify-between text-white shadow-lg relative z-[90]">
        <div className="flex items-center gap-3">
          <div className="bg-gradient-to-br from-blue-400 to-blue-700 p-1.5 rounded-lg">
            <Truck className="w-5 h-5 text-white" />
          </div>
          <div className="flex flex-col">
            <h1 className="text-[14px] font-black tracking-tight leading-none uppercase">POD Management Tool</h1>
            <p className="text-[9px] text-slate-400 font-bold mt-1 uppercase tracking-wider">
              Delhivery · Palam Vihar RPC · <span className="text-[#F9A825] font-black italic">By Ashu</span>
            </p>
          </div>
        </div>
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2 bg-green-500/10 px-3 py-1 rounded-full border border-green-500/20">
            <div className="w-1.5 h-1.5 bg-green-500 rounded-full animate-pulse" />
            <span className="text-[9px] font-black text-green-400 uppercase tracking-widest">Live</span>
          </div>
          <div className="bg-[#F9A825] px-3 py-1 rounded-lg text-[#1C2333] font-code text-[11px] font-black">
            {currentSession?.data.length || 0} ROWS
          </div>
        </div>
      </header>

      {/* Tabs */}
      <nav className="bg-[#1C2333] px-6 flex gap-8 border-t border-white/5 shadow-md sticky top-[55px] z-[85]">
        {[
          { id: "eod", label: "Daily EOD Rejection" },
          { id: "remark", label: "EOD Rejection Remark" }
        ].map(tab => (
          <button 
            key={tab.id}
            onClick={() => setActiveTab(tab.id as any)}
            className={cn(
              "py-3 text-[11px] font-black uppercase tracking-[0.1em] transition-all relative outline-none",
              activeTab === tab.id ? "text-white" : "text-slate-500 hover:text-slate-300"
            )}
          >
            {tab.label}
            {activeTab === tab.id && <div className="absolute bottom-0 left-0 w-full h-[3px] bg-[#F9A825] rounded-t-full shadow-lg" />}
          </button>
        ))}
      </nav>

      <main className="p-3 md:p-5 max-w-[1300px] mx-auto space-y-5">
        {activeTab === "eod" ? (
          <div className="space-y-5">
            {/* Setup Card */}
            <div className="bg-white rounded-[1.25rem] p-5 shadow-sm border border-slate-200">
              <div className="grid grid-cols-1 md:grid-cols-3 gap-5 mb-5">
                <div className="space-y-2">
                  <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest flex items-center gap-2"><Hash className="w-3.5 h-3.5" /> DSP ID</label>
                  <input 
                    type="number" 
                    value={setupData.dspId} 
                    onChange={(e) => setSetupData({...setupData, dspId: e.target.value})}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-[12px] font-code font-bold outline-none focus:border-[#1565C0] transition-all" 
                    placeholder="DSP No..." 
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest flex items-center gap-2"><User className="w-3.5 h-3.5" /> FE / Biker Name</label>
                  <input 
                    type="text" 
                    value={setupData.feName} 
                    onChange={(e) => setSetupData({...setupData, feName: e.target.value})}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-[12px] font-bold outline-none focus:border-[#1565C0] transition-all" 
                    placeholder="FE Name..." 
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest flex items-center gap-2">Date</label>
                  <input 
                    type="date" 
                    value={setupData.date} 
                    onChange={(e) => setSetupData({...setupData, date: e.target.value})}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-[12px] font-bold outline-none focus:border-[#1565C0]" 
                  />
                </div>
              </div>
              
              <div 
                className={cn(
                  "border-2 border-dashed rounded-[1.25rem] p-6 text-center transition-all relative group overflow-hidden",
                  (!setupData.feName || !setupData.dspId) ? "bg-slate-50 border-slate-200 opacity-50 cursor-not-allowed" : "border-blue-200 hover:border-blue-500 hover:bg-blue-50/30 cursor-pointer"
                )}
              >
                <input 
                  type="file" 
                  disabled={!setupData.feName || !setupData.dspId} 
                  onChange={handleFileUpload} 
                  className="absolute inset-0 opacity-0 cursor-pointer disabled:cursor-not-allowed z-10" 
                  accept=".xlsx,.xls,.csv" 
                />
                <div className="space-y-2">
                  <Download className={cn("w-6 h-6 mx-auto transition-all", (!setupData.feName || !setupData.dspId) ? "text-slate-300" : "text-blue-500")} />
                  <div>
                    <p className="text-[11px] font-black text-slate-700 uppercase tracking-widest">
                      {(!setupData.feName || !setupData.dspId) ? "Required: Fill DSP & FE Name" : "Upload Delhivery Export File"}
                    </p>
                    <p className="text-[8px] text-slate-400 font-bold uppercase mt-1 tracking-widest">Excel & CSV Only</p>
                  </div>
                </div>
              </div>
            </div>

            {/* Sessions Grid */}
            {sessions.length > 0 && (
              <div className="space-y-3">
                <div className="flex items-center justify-between px-1">
                  <h2 className="text-[9px] font-black text-slate-500 uppercase tracking-[0.3em] flex items-center gap-2"><Layers className="w-4 h-4" /> Saved Sessions</h2>
                  <button onClick={() => { if(confirm("Clear all?")) setSessions([]); }} className="text-[9px] font-black text-red-600 hover:text-red-700 uppercase tracking-widest flex items-center gap-1"><Trash2 className="w-3 h-3" /> Clear All</button>
                </div>
                <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-3">
                  {sessions.map(s => (
                    <div 
                      key={s.id} 
                      onClick={() => setSelectedSessionId(s.id)}
                      className={cn(
                        "bg-white p-3 rounded-xl border transition-all cursor-pointer relative group",
                        selectedSessionId === s.id ? "border-blue-500 ring-2 ring-blue-500/10 shadow-sm" : "border-slate-200 hover:border-blue-300"
                      )}
                    >
                      <button 
                        onClick={(e) => { e.stopPropagation(); setSessions(prev => prev.filter(x => x.id !== s.id)); }}
                        className="absolute top-1 right-1 opacity-0 group-hover:opacity-100 p-1 text-red-400 hover:text-red-600"
                      >
                        <X className="w-3 h-3" />
                      </button>
                      <p className="text-[10px] font-black truncate text-slate-800">{s.feName}</p>
                      <p className="text-[8px] font-code text-slate-400 font-bold">DSP: {s.dspId}</p>
                      <div className="mt-2 text-[11px] font-black text-blue-600">{s.data.length} PKTS</div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Main Dashboard */}
            {currentSession && (
              <div className="bg-white rounded-[1.25rem] shadow-lg border border-slate-200 overflow-hidden">
                <div className="p-3 bg-slate-50 border-b border-slate-100 flex flex-wrap gap-3 items-center justify-between">
                  <div className="flex gap-2">
                    <button onClick={() => downloadExcel(filteredRows, currentSession)} className="bg-[#2E7D32] text-white px-4 py-2 rounded-xl font-black text-[9px] uppercase tracking-widest flex items-center gap-2"><Download className="w-3.5 h-3.5" /> Export</button>
                    <button onClick={() => copyTable(filteredRows)} className="bg-[#1565C0] text-white px-4 py-2 rounded-xl font-black text-[9px] uppercase tracking-widest flex items-center gap-2"><Copy className="w-3.5 h-3.5" /> Copy</button>
                    {filteredRows.some(r => r.selected) && (
                      <button onClick={() => {
                        setSessions(prev => prev.map(s => {
                          if (s.id === selectedSessionId) return { ...s, data: s.data.filter(r => !r.selected) };
                          return s;
                        }));
                      }} className="bg-red-600 text-white px-4 py-2 rounded-xl font-black text-[9px] uppercase tracking-widest">Delete Selected</button>
                    )}
                  </div>
                  <div className="flex items-center gap-2">
                    {[
                      { id: 'all', label: 'All', val: stats.total, col: 'blue' },
                      { id: 'pending', label: 'Pend', val: stats.pending, col: 'amber' },
                      { id: 'dispatched', label: 'Disp', val: stats.dispatched, col: 'blue' },
                      { id: 'rto', label: 'RTO', val: stats.rto, col: 'red' },
                      { id: 'dto', label: 'DTO', val: stats.dto, col: 'green' }
                    ].map(t => (
                      <button 
                        key={t.id} 
                        onClick={() => { setStatusFilter(t.id); setRemarkFilter(null); }}
                        className={cn(
                          "px-3 py-1.5 rounded-lg border transition-all text-center min-w-[45px]",
                          statusFilter === t.id ? `bg-${t.col}-600 border-${t.col}-600 text-white shadow-md` : `border-slate-200 text-slate-400`
                        )}
                      >
                        <p className="text-[11px] font-black leading-none">{t.val}</p>
                        <p className="text-[7px] font-black uppercase mt-0.5 tracking-tighter">{t.label}</p>
                      </button>
                    ))}
                  </div>
                </div>

                {/* Remark Chips */}
                {statusFilter === 'pending' && remarkStats.length > 0 && (
                  <div className="p-3 bg-white border-b border-slate-100 flex flex-wrap gap-2">
                    <p className="w-full text-[8px] font-black text-slate-400 uppercase tracking-[0.2em] mb-1">Filter by Remark:</p>
                    {remarkFilter && (
                      <button onClick={() => setRemarkFilter(null)} className="px-3 py-1 bg-blue-50 text-blue-600 rounded-full text-[8px] font-black uppercase tracking-widest border border-blue-200">← All Pending</button>
                    )}
                    {remarkStats.map(([remark, count]) => {
                      const isIntact = remark.toLowerCase().includes("intact");
                      return (
                        <button 
                          key={remark} 
                          onClick={() => setRemarkFilter(remark)}
                          className={cn(
                            "px-3 py-1 rounded-full text-[8px] font-black uppercase tracking-widest border transition-all",
                            remarkFilter === remark 
                              ? (isIntact ? "bg-red-600 border-red-700 text-white" : "bg-blue-600 border-blue-700 text-white")
                              : (isIntact ? "bg-red-50 border-red-100 text-red-600 hover:bg-red-100" : "bg-slate-50 border-slate-100 text-slate-600 hover:bg-slate-100")
                          )}
                        >
                          {remark} <span className="opacity-50 ml-1">({count})</span>
                        </button>
                      );
                    })}
                  </div>
                )}
                
                <div className="overflow-x-auto max-h-[400px] custom-scrollbar">
                  <table className="w-full text-left border-collapse">
                    <thead className="sticky top-0 z-30">
                      <tr className="bg-[#1C2333] text-white text-[9px] font-black uppercase tracking-widest">
                        <th className="p-3 w-[40px] text-center">Sel</th>
                        <th className="p-3 w-[40px] text-center">Del</th>
                        <th className="p-3 w-[90px]">DSP ID</th>
                        <th className="p-3 w-[150px]">AWB Number</th>
                        <th className="p-3 w-[130px]">Client</th>
                        <th className="p-3 w-[130px]">Order ID</th>
                        <th className="p-3">Remark</th>
                        <th className="p-3 w-[110px]">FE Name</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredRows.map((row, idx) => {
                        const isAWBDupe = duplicateAWBs.has(row.awb);
                        const isIntact = row.remark.toLowerCase().includes("intact");
                        const isFocused = focusedIndex === idx;
                        return (
                          <tr key={row.id} className={cn(
                            "border-b border-slate-50 hover:bg-blue-50/50 transition-colors group",
                            row.selected ? "bg-blue-100/50" : "",
                            isFocused ? "ring-2 ring-blue-500 ring-inset" : "",
                            isIntact ? "bg-red-50/30" : ""
                          )}>
                            <td className="p-2 text-center">
                              <input type="checkbox" checked={row.selected} onChange={() => toggleRowSelect(row.id)} className="w-3.5 h-3.5 accent-blue-600" />
                            </td>
                            <td className="p-2 text-center">
                              <button onClick={() => deleteRow(row.id)} className="text-slate-300 hover:text-red-500"><X className="w-3.5 h-3.5" /></button>
                            </td>
                            <td className="p-2 font-code text-[10px] font-bold text-slate-400">{idx === 0 ? row.dspId : ""}</td>
                            <td 
                              className={cn(
                                "p-2 font-code text-[11px] font-black tracking-wider cursor-pointer hover:underline",
                                isAWBDupe ? "text-red-600" : "text-[#1565C0]"
                              )}
                              onClick={() => copyToClipboard(row.awb, "AWB")}
                            >
                              {row.awb}
                            </td>
                            <td className="p-2 text-[10px] font-bold text-slate-700 truncate max-w-[130px]">{row.client}</td>
                            <td className="p-2 text-[10px] text-slate-500 font-medium truncate max-w-[130px]">{row.orderId}</td>
                            <td className="p-2">
                              <span className={cn(
                                "inline-flex px-2 py-0.5 rounded-full text-[7.5px] font-black uppercase tracking-widest border",
                                isIntact ? "bg-red-100 text-red-700 border-red-200" : "bg-amber-50 text-amber-700 border-amber-200"
                              )}>
                                {row.remark || "N/A"}
                              </span>
                            </td>
                            <td className="p-2 text-[10px] font-black text-slate-500 truncate">{row.feName}</td>
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
          <div className="space-y-5">
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
              <div className="bg-white rounded-[1.25rem] p-5 border border-slate-200 shadow-sm space-y-5">
                <h2 className="text-lg font-black tracking-tighter flex items-center gap-2"><FileSpreadsheet className="w-5 h-5 text-green-600" /> Remark Replacer</h2>
                <div className="border-2 border-dashed border-green-50 rounded-2xl p-8 text-center hover:border-green-300 hover:bg-green-50/20 transition-all cursor-pointer relative group overflow-hidden">
                  <input type="file" onChange={handleReplacerUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" accept=".xlsx,.xls,.csv" />
                  <Download className="w-8 h-8 text-green-200 mx-auto" />
                  <div className="mt-3">
                    <p className="text-[11px] font-black text-slate-700 uppercase tracking-widest">Upload EOD Export</p>
                    <p className="text-[7px] font-bold text-slate-400 mt-1 tracking-widest uppercase">Excel & CSV Only</p>
                  </div>
                </div>

                {replacerData.length > 0 && (
                  <div className="grid grid-cols-3 gap-2">
                    <div className="bg-slate-900 p-3 rounded-xl text-center">
                       <p className="text-lg font-black text-white">{replacerStats.total}</p>
                       <p className="text-[7px] font-black text-slate-400 uppercase tracking-widest">Total</p>
                    </div>
                    <div className="bg-green-600 p-3 rounded-xl text-center">
                       <p className="text-lg font-black text-white">{replacerStats.replaced}</p>
                       <p className="text-[7px] font-black text-white/70 uppercase tracking-widest">Fixed</p>
                    </div>
                    <div className="bg-amber-500 p-3 rounded-xl text-center">
                       <p className="text-lg font-black text-white">{replacerStats.missing}</p>
                       <p className="text-[7px] font-black text-white/70 uppercase tracking-widest">N/A</p>
                    </div>
                  </div>
                )}
              </div>

              <div className="bg-[#1C2333] rounded-[1.25rem] p-5 text-white shadow-xl relative overflow-hidden">
                <h3 className="text-[9px] font-black text-[#F9A825] uppercase tracking-[0.3em] mb-4 flex items-center gap-2">
                  <AlertCircle className="w-3.5 h-3.5" /> Official Mapping Guide
                </h3>
                <div className="overflow-y-auto max-h-[250px] space-y-2 custom-scrollbar pr-2">
                  {Object.entries(REMARK_MAPPING).map(([old, official]) => (
                    <div key={old} className="p-2.5 bg-white/5 rounded-xl text-[8.5px] border border-white/10">
                      <p className="text-slate-500 italic font-bold mb-1">{old}</p>
                      <p className="text-green-400 font-black flex items-center gap-2">
                        <CheckCircle2 className="w-3 h-3 text-green-500" /> {official}
                      </p>
                    </div>
                  ))}
                </div>
              </div>
            </div>
            
            {replacerData.length > 0 && (
              <div className="bg-white rounded-[1.25rem] shadow-lg border border-slate-200 overflow-hidden">
                <div className="p-3 bg-slate-50 border-b border-slate-100 flex items-center justify-between">
                  <button onClick={() => {
                      const header = ['Date', 'DSP No', 'AWB Number', 'Client', 'Original Remark', 'Official Remark', 'FE Name'];
                      const rows = replacerData.map(r => [r.date, r.dsp, { v: r.awb, t: 's' }, r.client, r.oldRemark, r.officialRemark, r.feName]);
                      const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
                      const wb = XLSX.utils.book_new();
                      XLSX.utils.book_append_sheet(wb, ws, "Replaced");
                      XLSX.writeFile(wb, `Delhivery_Replaced_Remarks.xlsx`);
                  }} className="bg-[#2E7D32] text-white px-5 py-2 rounded-xl font-black text-[9px] uppercase tracking-widest flex items-center gap-2 shadow-sm"><Download className="w-3.5 h-3.5" /> Download Report</button>
                </div>
                
                <div className="overflow-x-auto max-h-[400px] custom-scrollbar">
                  <table className="w-full text-left">
                    <thead className="sticky top-0 z-30">
                      <tr className="bg-[#1C2333] text-white text-[9px] font-black uppercase tracking-widest">
                        <th className="p-3 w-[50px]">#</th>
                        <th className="p-3 w-[100px]">Date</th>
                        <th className="p-3 w-[150px]">AWB Number</th>
                        <th className="p-3">Official Remark</th>
                        <th className="p-3 w-[140px]">Original</th>
                        <th className="p-3 w-[110px]">FE Name</th>
                      </tr>
                    </thead>
                    <tbody>
                      {replacerData.map((row, idx) => (
                        <tr key={idx} className={cn(
                          "border-b border-slate-50 hover:bg-slate-50 transition-colors group",
                          row.isReplaced ? "bg-green-50/20" : "bg-amber-50/20"
                        )}>
                          <td className="p-2 text-[9px] font-black text-slate-400">{row.id}</td>
                          <td className="p-2 text-[9px] text-slate-600 font-bold">{row.date}</td>
                          <td className="p-2 font-code text-[11px] font-black text-[#1565C0]">{row.awb}</td>
                          <td className="p-2">
                            <span className={cn(
                              "inline-flex px-2 py-0.5 rounded-full text-[7.5px] font-black uppercase tracking-widest border",
                              row.isReplaced ? "bg-green-600 text-white border-green-700" : "bg-amber-500 text-white border-amber-600"
                            )}>
                              {row.officialRemark}
                            </span>
                          </td>
                          <td className="p-2 text-[8px] text-slate-400 italic font-bold truncate max-w-[140px]">{row.oldRemark}</td>
                          <td className="p-2 text-[9px] font-black text-slate-700 truncate">{row.feName}</td>
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
        
        body { font-family: 'Inter', sans-serif; background: #F0F4FA; }
        .font-code { font-family: 'IBM Plex Mono', monospace; }
        
        .custom-scrollbar::-webkit-scrollbar { width: 5px; height: 5px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #F1F5F9; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 10px; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #94A3B8; }

        input[type="number"]::-webkit-inner-spin-button,
        input[type="number"]::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }
        
        ::selection { background: #BFDBFE; color: #1E40AF; }
      `}</style>
    </div>
  );
}
