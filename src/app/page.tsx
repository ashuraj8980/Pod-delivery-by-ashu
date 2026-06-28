
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
  Check,
  AlertTriangle,
  MousePointer2,
  FileWarning
} from "lucide-react";
import * as XLSX from "xlsx";
import { cn } from "@/lib/utils";

/**
 * @fileOverview Delhivery POD Management Tool v8.0 (Robust Performance Edition)
 * Finalized for Palam Vihar RPC. Optimized for large datasets and error resilience.
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
  const [isProcessing, setIsProcessing] = useState(false);
  
  // Module 2 State
  const [replacerData, setReplacerData] = useState<any[]>([]);
  const [replacerStats, setReplacerStats] = useState({ total: 0, replaced: 0, missing: 0 });

  useEffect(() => {
    setIsMounted(true);
    setSetupData(prev => ({ ...prev, date: new Date().toISOString().split('T')[0] }));
    // Fresh Start: We do NOT auto-load from localStorage on refresh anymore.
    // Data only stays during active session until cleared manually.
  }, []);

  const fixAWB = (val: any) => {
    if (val === null || val === undefined) return "";
    let str = String(val).trim().replace(/['",]/g, ""); 
    if (/^[\d.]+[eE][+\-]?\d+$/.test(str)) {
      try {
        str = BigInt(Math.round(Number(val))).toString();
      } catch (e) {
        str = String(val);
      }
    }
    return str.replace(/\.0$/, "");
  };

  const showToast = (msg: string, type: 'ok' | 'err' | 'info') => {
    if (!isMounted) return;
    const toast = document.createElement('div');
    toast.className = `fixed bottom-10 left-1/2 -translate-x-1/2 px-5 py-2.5 rounded-full text-[11px] font-black z-[200] shadow-2xl transition-all duration-300 transform scale-95 opacity-0 animate-in fade-in slide-in-from-bottom-5 ${
      type === 'ok' ? 'bg-[#052E0F] text-[#6EE7A6] border border-[#6EE7A6]/20' : 
      type === 'err' ? 'bg-[#2D0808] text-[#FCA5A5] border border-[#FCA5A5]/20' : 
      'bg-[#1C2333] text-[#93C5FD] border border-[#93C5FD]/20'
    }`;
    toast.style.opacity = '1';
    toast.style.transform = 'translateX(-50%) scale(1)';
    toast.innerHTML = `<span class="flex items-center gap-2">${type === 'ok' ? '✓' : type === 'err' ? '✕' : 'ℹ'} ${msg}</span>`;
    document.body.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateX(-50%) scale(0.9)';
      setTimeout(() => toast.remove(), 300);
    }, 2800);
  };

  const copyToClipboard = (text: string, label: string = "Copied") => {
    if (!text) return;
    navigator.clipboard.writeText(text).then(() => {
      showToast(`${label}: ${text}`, "ok");
    }).catch(() => {
      const textArea = document.createElement("textarea");
      textArea.value = text;
      document.body.appendChild(textArea);
      textArea.select();
      document.execCommand("copy");
      textArea.remove();
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

    setIsProcessing(true);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary', cellDates: true });
        if (!wb || !wb.SheetNames.length) throw new Error("Invalid workbook");
        
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(ws);
        
        if (!Array.isArray(rawData) || rawData.length === 0) {
          throw new Error("File empty or invalid format");
        }

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
          showToast("No valid rows found! Only Pending/RTO/DTO accepted.", "err");
        }
      } catch (err: any) {
        console.error("Upload Error:", err);
        showToast(err.message || "Error parsing file! Please upload valid Delhivery export.", "err");
      } finally {
        setIsProcessing(false);
      }
    };
    reader.onerror = () => {
      showToast("File reading error!", "err");
      setIsProcessing(false);
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  const handleReplacerUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsProcessing(true);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        if (!wb || !wb.SheetNames.length) throw new Error("Invalid workbook");
        
        const rawData = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        if (!Array.isArray(rawData) || rawData.length === 0) throw new Error("File empty");
        
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
      } catch (err: any) {
        showToast("Replacer error! Check your file format.", "err");
      } finally {
        setIsProcessing(false);
      }
    };
    reader.onerror = () => {
      showToast("File reading error!", "err");
      setIsProcessing(false);
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
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
    if (!dataRows || dataRows.length === 0) return;
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
    if (!dataRows || dataRows.length === 0) return;
    const text = dataRows.map((r, i) => {
      const dsp = i === 0 ? r.dspId : "";
      return `${r.date}\t${dsp}\t${r.awb}\t${r.client}\t${r.orderId}\t${r.remark}\t${r.feName}`;
    }).join("\n");
    
    copyToClipboard(text, "Table Data Copied (Paste in Excel)");
  };

  const clearAllSessions = () => {
    if (confirm("Kya aap saara data browser se hatana chahte hain? Ye hamesha ke liye clear ho jayega.")) {
      setSessions([]);
      setSelectedSessionId(null);
      localStorage.removeItem(STORAGE_KEY);
      setReplacerData([]);
      setReplacerStats({ total: 0, replaced: 0, missing: 0 });
      showToast("All sessions cleared freshly!", "ok");
    }
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

  if (!isMounted) return null;

  return (
    <div className="min-h-screen bg-[#F0F4FA] font-body text-[#1C2333] select-auto overflow-x-hidden">
      {/* Rainbow Stripe */}
      <div className="h-[3px] w-full bg-gradient-to-r from-[#1565C0] via-[#F9A825] via-[#2E7D32] to-[#D32F2F] sticky top-0 z-[100]" />
      
      {/* Header */}
      <header className="h-[52px] bg-[#1C2333] px-5 flex items-center justify-between text-white shadow-lg relative z-[90]">
        <div className="flex items-center gap-2.5">
          <div className="bg-gradient-to-br from-blue-400 to-blue-700 p-1 rounded-lg">
            <Truck className="w-4 h-4 text-white" />
          </div>
          <div className="flex flex-col">
            <h1 className="text-[13px] font-black tracking-tight leading-none uppercase">POD Management Tool</h1>
            <p className="text-[8px] text-slate-400 font-bold mt-0.5 uppercase tracking-wider">
              Delhivery · Palam Vihar RPC · <span className="text-[#F9A825] font-black italic">By Ashu</span>
            </p>
          </div>
        </div>
        <div className="flex items-center gap-3">
          {isProcessing && (
            <div className="flex items-center gap-2 bg-blue-500/10 px-2 py-1 rounded-full border border-blue-500/20">
              <div className="w-1.5 h-1.5 bg-blue-500 rounded-full animate-spin" />
              <span className="text-[8px] font-black text-blue-400 uppercase">Processing...</span>
            </div>
          )}
          <div className="flex items-center gap-2 bg-green-500/10 px-2.5 py-1 rounded-full border border-green-500/20">
            <div className="w-1.5 h-1.5 bg-green-500 rounded-full animate-pulse" />
            <span className="text-[8px] font-black text-green-400 uppercase tracking-widest">Live</span>
          </div>
        </div>
      </header>

      {/* Tabs */}
      <nav className="bg-[#1C2333] px-5 flex gap-6 border-t border-white/5 shadow-md sticky top-[52px] z-[85]">
        {[
          { id: "eod", label: "Daily EOD Rejection" },
          { id: "remark", label: "EOD Rejection Remark" }
        ].map(tab => (
          <button 
            key={tab.id}
            onClick={() => { setActiveTab(tab.id as any); setStatusFilter("all"); setRemarkFilter(null); }}
            className={cn(
              "py-2.5 text-[10px] font-black uppercase tracking-[0.05em] transition-all relative outline-none",
              activeTab === tab.id ? "text-white" : "text-slate-500 hover:text-slate-300"
            )}
          >
            {tab.label}
            {activeTab === tab.id && <div className="absolute bottom-0 left-0 w-full h-[2.5px] bg-[#F9A825] rounded-t-full shadow-lg" />}
          </button>
        ))}
      </nav>

      <main className="p-3 md:p-4 max-w-[1200px] mx-auto space-y-4">
        {activeTab === "eod" ? (
          <div className="space-y-4 animate-in fade-in duration-300">
            {/* Setup Card */}
            <div className="bg-white rounded-2xl p-4 shadow-sm border border-slate-200">
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
                <div className="space-y-1.5">
                  <label className="text-[8px] font-black text-slate-500 uppercase tracking-widest flex items-center gap-1.5"><Hash className="w-3 h-3" /> DSP Number</label>
                  <input 
                    type="number" 
                    value={setupData.dspId} 
                    onChange={(e) => setSetupData({...setupData, dspId: e.target.value})}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-[11px] font-code font-bold outline-none focus:border-[#1565C0] focus:bg-white transition-all" 
                    placeholder="DSP No..." 
                  />
                </div>
                <div className="space-y-1.5">
                  <label className="text-[8px] font-black text-slate-500 uppercase tracking-widest flex items-center gap-1.5"><User className="w-3 h-3" /> FE / Biker Name</label>
                  <input 
                    type="text" 
                    value={setupData.feName} 
                    onChange={(e) => setSetupData({...setupData, feName: e.target.value})}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-[11px] font-bold outline-none focus:border-[#1565C0] focus:bg-white transition-all" 
                    placeholder="FE Name..." 
                  />
                </div>
                <div className="space-y-1.5">
                  <label className="text-[8px] font-black text-slate-500 uppercase tracking-widest flex items-center gap-1.5">Operation Date</label>
                  <input 
                    type="date" 
                    value={setupData.date} 
                    onChange={(e) => setSetupData({...setupData, date: e.target.value})}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-[11px] font-bold outline-none focus:border-[#1565C0] focus:bg-white" 
                  />
                </div>
              </div>
              
              <div 
                className={cn(
                  "border-2 border-dashed rounded-2xl p-5 text-center transition-all relative group overflow-hidden",
                  (!setupData.feName || !setupData.dspId) ? "bg-slate-50 border-slate-100 opacity-50 cursor-not-allowed" : "border-blue-200 hover:border-blue-500 hover:bg-blue-50/30 cursor-pointer"
                )}
              >
                <input 
                  type="file" 
                  disabled={!setupData.feName || !setupData.dspId || isProcessing} 
                  onChange={handleFileUpload} 
                  className="absolute inset-0 opacity-0 cursor-pointer disabled:cursor-not-allowed z-10" 
                  accept=".xlsx,.xls,.csv" 
                />
                <div className="space-y-1">
                  {isProcessing ? (
                    <div className="flex flex-col items-center gap-2">
                       <div className="w-8 h-8 border-4 border-blue-100 border-t-blue-500 rounded-full animate-spin" />
                       <p className="text-[10px] font-black text-blue-500 uppercase tracking-widest">Processing Data...</p>
                    </div>
                  ) : (
                    <>
                      <Download className={cn("w-5 h-5 mx-auto transition-all", (!setupData.feName || !setupData.dspId) ? "text-slate-200" : "text-blue-500")} />
                      <div>
                        <p className="text-[10px] font-black text-slate-700 uppercase tracking-widest">
                          {(!setupData.feName || !setupData.dspId) ? "Enter DSP & FE to Upload" : "Upload Delhivery Export File"}
                        </p>
                        <p className="text-[7px] text-slate-400 font-bold uppercase mt-0.5 tracking-widest">Excel & CSV Only</p>
                      </div>
                    </>
                  )}
                </div>
              </div>
            </div>

            {/* Sessions Grid */}
            {sessions.length > 0 && (
              <div className="space-y-3 animate-in slide-in-from-top-2 duration-500">
                <div className="flex items-center justify-between px-1">
                  <h2 className="text-[8px] font-black text-slate-500 uppercase tracking-[0.2em] flex items-center gap-1.5"><Layers className="w-3 h-3" /> All FE Sessions — Saved Data</h2>
                  <button onClick={clearAllSessions} className="text-[9px] font-black text-red-600 hover:text-red-700 uppercase tracking-widest flex items-center gap-1.5 bg-red-50 px-3 py-1 rounded-full transition-all active:scale-95"><Trash2 className="w-3.5 h-3.5" /> Clear All Sessions</button>
                </div>
                <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-5 gap-3">
                  {sessions.map(s => (
                    <div 
                      key={s.id} 
                      onClick={() => { setSelectedSessionId(s.id); setStatusFilter("all"); setRemarkFilter(null); }}
                      className={cn(
                        "bg-white p-3 rounded-xl border transition-all cursor-pointer relative group",
                        selectedSessionId === s.id ? "border-blue-500 ring-2 ring-blue-500/10 shadow-md bg-blue-50/10" : "border-slate-200 hover:border-blue-300"
                      )}
                    >
                      <button 
                        onClick={(e) => { e.stopPropagation(); setSessions(prev => prev.filter(x => x.id !== s.id)); if(selectedSessionId === s.id) setSelectedSessionId(null); }}
                        className="absolute top-2 right-2 opacity-0 group-hover:opacity-100 p-1 text-red-400 hover:text-red-600 transition-opacity"
                      >
                        <X className="w-3.5 h-3.5" />
                      </button>
                      <p className="text-[10px] font-black truncate text-slate-800 uppercase tracking-tighter">{s.feName}</p>
                      <p className="text-[8px] font-code text-slate-400 font-bold">DSP: {s.dspId}</p>
                      <div className="mt-2 flex justify-between items-end">
                        <span className="text-[11px] font-black text-blue-600">{s.data.length} PKTS</span>
                        <span className="text-[7px] font-bold text-slate-300 uppercase">{new Date(s.timestamp).toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})}</span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Main Dashboard */}
            {currentSession && (
              <div className="bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden animate-in fade-in zoom-in-95 duration-500">
                <div className="p-3 bg-slate-50 border-b border-slate-100 flex flex-wrap gap-2.5 items-center justify-between">
                  <div className="flex gap-2">
                    <button onClick={() => downloadExcel(filteredRows, currentSession, statusFilter.toUpperCase())} className="bg-[#2E7D32] text-white px-4 py-2 rounded-lg font-black text-[9px] uppercase tracking-widest flex items-center gap-2 shadow-sm hover:bg-[#1B5E20] transition-all active:scale-95"><Download className="w-3.5 h-3.5" /> Export {statusFilter.toUpperCase()}</button>
                    <button onClick={() => copyTable(filteredRows)} className="bg-[#1565C0] text-white px-4 py-2 rounded-lg font-black text-[9px] uppercase tracking-widest flex items-center gap-2 shadow-sm hover:bg-[#0D47A1] transition-all active:scale-95"><Copy className="w-3.5 h-3.5" /> Copy For Excel</button>
                    {filteredRows.some(r => r.selected) && (
                      <button onClick={() => {
                        setSessions(prev => prev.map(s => {
                          if (s.id === selectedSessionId) return { ...s, data: s.data.filter(r => !r.selected) };
                          return s;
                        }));
                      }} className="bg-red-600 text-white px-4 py-2 rounded-lg font-black text-[9px] uppercase tracking-widest shadow-md hover:bg-red-700 transition-all active:scale-95">Delete ({filteredRows.filter(r => r.selected).length})</button>
                    )}
                  </div>
                  <div className="flex items-center gap-1.5">
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
                          "px-3 py-1.5 rounded-xl border transition-all text-center min-w-[50px]",
                          statusFilter === t.id ? `bg-${t.col}-600 border-${t.col}-600 text-white shadow-lg transform scale-105` : `border-slate-200 text-slate-400 hover:bg-white`
                        )}
                      >
                        <p className="text-[12px] font-black leading-none">{t.val}</p>
                        <p className="text-[7px] font-black uppercase mt-0.5 tracking-tighter">{t.label}</p>
                      </button>
                    ))}
                  </div>
                </div>

                {/* Remark Chips */}
                {statusFilter === 'pending' && remarkStats.length > 0 && (
                  <div className="p-3 bg-white border-b border-slate-100 flex flex-wrap gap-2">
                    <p className="w-full text-[8px] font-black text-slate-400 uppercase tracking-[0.2em] mb-1">Remark Breakdown — Pending</p>
                    {remarkFilter && (
                      <button onClick={() => setRemarkFilter(null)} className="px-3 py-1 bg-blue-50 text-blue-600 rounded-full text-[8px] font-black uppercase tracking-widest border border-blue-200 hover:bg-blue-100 transition-all">← Show All Pending</button>
                    )}
                    {remarkStats.map(([remark, count]) => {
                      const isIntact = remark.toLowerCase().includes("intact");
                      return (
                        <button 
                          key={remark} 
                          onClick={() => setRemarkFilter(remark)}
                          className={cn(
                            "px-3 py-1 rounded-full text-[8px] font-black uppercase tracking-widest border transition-all hover:scale-105 active:scale-95",
                            remarkFilter === remark 
                              ? (isIntact ? "bg-red-600 border-red-700 text-white shadow-md" : "bg-blue-600 border-blue-700 text-white shadow-md")
                              : (isIntact ? "bg-red-50 border-red-100 text-red-600 hover:bg-red-100" : "bg-slate-50 border-slate-100 text-slate-600 hover:bg-slate-100")
                          )}
                        >
                          {remark || "No Remark"} <span className="opacity-50 ml-1">({count})</span>
                        </button>
                      );
                    })}
                  </div>
                )}
                
                <div className="overflow-x-auto max-h-[450px] custom-scrollbar">
                  <table className="w-full text-left border-collapse">
                    <thead className="sticky top-0 z-30">
                      <tr className="bg-[#1C2333] text-white text-[9px] font-black uppercase tracking-widest">
                        <th className="p-3 w-[40px] text-center border-r border-white/5">Sel</th>
                        <th className="p-3 w-[40px] text-center border-r border-white/5">Del</th>
                        <th className="p-3 w-[90px] border-r border-white/5">DSP ID</th>
                        <th className="p-3 w-[150px] border-r border-white/5">AWB Number</th>
                        <th className="p-3 w-[130px] border-r border-white/5">Client</th>
                        <th className="p-3 w-[130px] border-r border-white/5">Order ID</th>
                        <th className="p-3 border-r border-white/5">Remark</th>
                        <th className="p-3 w-[110px]">FE Name</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredRows.map((row, idx) => {
                        const isAWBDupe = duplicateAWBs.has(row.awb);
                        const isIntact = row.remark.toLowerCase().includes("intact");
                        return (
                          <tr key={row.id} className={cn(
                            "border-b border-slate-100 hover:bg-blue-50/50 transition-colors group",
                            row.selected ? "bg-blue-100/50" : "",
                            isIntact ? "bg-red-50/30" : ""
                          )}>
                            <td className="p-2.5 text-center">
                              <input type="checkbox" checked={row.selected} onChange={() => {
                                setSessions(prev => prev.map(s => {
                                  if (s.id === selectedSessionId) {
                                    return { ...s, data: s.data.map(r => r.id === row.id ? { ...r, selected: !r.selected } : r) };
                                  }
                                  return s;
                                }));
                              }} className="w-4 h-4 accent-blue-600 cursor-pointer" />
                            </td>
                            <td className="p-2.5 text-center">
                              <button onClick={() => {
                                setSessions(prev => prev.map(s => {
                                  if (s.id === selectedSessionId) return { ...s, data: s.data.filter(r => r.id !== row.id) };
                                  return s;
                                }));
                              }} className="text-slate-200 hover:text-red-500 transition-colors"><X className="w-4 h-4" /></button>
                            </td>
                            <td className="p-2.5 font-code text-[10px] font-bold text-slate-400">{idx === 0 ? row.dspId : ""}</td>
                            <td 
                              className={cn(
                                "p-2.5 font-code text-[12px] font-black tracking-wider cursor-pointer hover:underline flex items-center gap-1.5",
                                isAWBDupe ? "text-red-600" : "text-[#1565C0]"
                              )}
                              onClick={() => copyToClipboard(row.awb, "AWB")}
                            >
                              <MousePointer2 className="w-3 h-3 opacity-0 group-hover:opacity-100 text-slate-400" />
                              {row.awb}
                            </td>
                            <td className="p-2.5 text-[10px] font-bold text-slate-700 truncate max-w-[130px]">{row.client}</td>
                            <td className="p-2.5 text-[10px] text-slate-500 font-medium truncate max-w-[130px]">{row.orderId}</td>
                            <td className="p-2.5">
                              <span className={cn(
                                "inline-flex px-2.5 py-1 rounded-full text-[8px] font-black uppercase tracking-widest border",
                                isIntact ? "bg-red-100 text-red-700 border-red-200" : "bg-amber-50 text-amber-700 border-amber-200"
                              )}>
                                {row.remark || "N/A"}
                              </span>
                            </td>
                            <td className="p-2.5 text-[10px] font-black text-slate-500 truncate">{row.feName}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                  {filteredRows.length === 0 && (
                    <div className="py-24 text-center space-y-4 animate-in fade-in duration-500">
                       <FileWarning className="w-12 h-12 text-slate-200 mx-auto" />
                       <div>
                         <p className="text-[12px] font-black text-slate-400 uppercase tracking-widest">No matching rows found</p>
                         <p className="text-[8px] text-slate-300 font-bold uppercase mt-1 tracking-widest">Upload a file or change filters</p>
                       </div>
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        ) : (
          <div className="space-y-4 animate-in fade-in duration-300">
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
              <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm space-y-4">
                <h2 className="text-md font-black tracking-tighter flex items-center gap-2"><FileSpreadsheet className="w-5 h-5 text-green-600" /> Remark Replacer</h2>
                <div className="border-2 border-dashed border-green-100 rounded-2xl p-8 text-center hover:border-green-400 hover:bg-green-50/20 transition-all cursor-pointer relative group overflow-hidden">
                  <input type="file" disabled={isProcessing} onChange={handleReplacerUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" accept=".xlsx,.xls,.csv" />
                  {isProcessing ? (
                    <div className="flex flex-col items-center gap-2">
                       <div className="w-8 h-8 border-4 border-green-100 border-t-green-500 rounded-full animate-spin" />
                       <p className="text-[10px] font-black text-green-500 uppercase tracking-widest">Processing...</p>
                    </div>
                  ) : (
                    <>
                      <Download className="w-8 h-8 text-green-200 mx-auto group-hover:scale-110 transition-transform" />
                      <div className="mt-3">
                        <p className="text-[11px] font-black text-slate-700 uppercase tracking-widest">Upload EOD Export File</p>
                        <p className="text-[8px] font-bold text-slate-400 mt-1 tracking-widest uppercase">Excel & CSV Accepted</p>
                      </div>
                    </>
                  )}
                </div>

                {replacerData.length > 0 && (
                  <div className="grid grid-cols-3 gap-3">
                    <div className="bg-[#1C2333] p-3 rounded-2xl text-center shadow-lg border-b-4 border-white/5">
                       <p className="text-xl font-black text-white">{replacerStats.total}</p>
                       <p className="text-[7px] font-black text-slate-400 uppercase tracking-widest">Total Rows</p>
                    </div>
                    <div className="bg-green-600 p-3 rounded-2xl text-center shadow-lg border-b-4 border-black/10">
                       <p className="text-xl font-black text-white">{replacerStats.replaced}</p>
                       <p className="text-[7px] font-black text-white/70 uppercase tracking-widest">Replaced OK</p>
                    </div>
                    <div className="bg-amber-500 p-3 rounded-2xl text-center shadow-lg border-b-4 border-black/10">
                       <p className="text-xl font-black text-white">{replacerStats.missing}</p>
                       <p className="text-[7px] font-black text-white/70 uppercase tracking-widest">No Mapping</p>
                    </div>
                  </div>
                )}
              </div>

              <div className="bg-[#1C2333] rounded-2xl p-5 text-white shadow-xl relative overflow-hidden">
                <h3 className="text-[9px] font-black text-[#F9A825] uppercase tracking-[0.2em] mb-4 flex items-center gap-2">
                  <AlertCircle className="w-4 h-4" /> Official Remark Mapping Reference
                </h3>
                <div className="overflow-y-auto max-h-[250px] space-y-2 custom-scrollbar pr-2">
                  {Object.entries(REMARK_MAPPING).map(([old, official]) => (
                    <div key={old} className="p-3 bg-white/5 rounded-xl text-[9px] border border-white/10 hover:bg-white/10 transition-colors">
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
              <div className="bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden animate-in slide-in-from-bottom-4 duration-500">
                <div className="p-3 bg-slate-50 border-b border-slate-100 flex items-center justify-between">
                  <div className="flex gap-2">
                    <button onClick={() => {
                        const header = ['Date', 'DSP No', 'AWB Number', 'Client', 'Original Remark', 'Official Remark', 'FE Name'];
                        const rows = replacerData.map(r => [r.date, r.dsp, { v: r.awb, t: 's' }, r.client, r.oldRemark, r.officialRemark, r.feName]);
                        const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
                        const wb = XLSX.utils.book_new();
                        XLSX.utils.book_append_sheet(wb, ws, "Replaced");
                        XLSX.writeFile(wb, `Delhivery_Replaced_Remarks.xlsx`);
                    }} className="bg-[#2E7D32] text-white px-5 py-2 rounded-lg font-black text-[9px] uppercase tracking-widest flex items-center gap-2 shadow-sm hover:bg-[#1B5E20] transition-all"><Download className="w-3.5 h-3.5" /> Download Report</button>
                    <button onClick={() => {
                      const text = replacerData.map(r => `${r.date}\t${r.dsp}\t${r.awb}\t${r.client}\t${r.oldRemark}\t${r.officialRemark}\t${r.feName}`).join("\n");
                      copyToClipboard(text, "Replaced Table Copied");
                    }} className="bg-[#1565C0] text-white px-5 py-2 rounded-lg font-black text-[9px] uppercase tracking-widest flex items-center gap-2 shadow-sm hover:bg-[#0D47A1] transition-all"><Copy className="w-3.5 h-3.5" /> Copy Table</button>
                  </div>
                  <button onClick={() => setReplacerData([])} className="text-[9px] font-black text-slate-400 uppercase tracking-widest hover:text-red-500 transition-colors">Clear Replacer</button>
                </div>
                
                <div className="overflow-x-auto max-h-[450px] custom-scrollbar">
                  <table className="w-full text-left">
                    <thead className="sticky top-0 z-30">
                      <tr className="bg-[#1C2333] text-white text-[9px] font-black uppercase tracking-widest">
                        <th className="p-3 w-[45px]">#</th>
                        <th className="p-3 w-[100px]">Date</th>
                        <th className="p-3 w-[150px]">AWB Number</th>
                        <th className="p-3">Official Remark</th>
                        <th className="p-3 w-[140px]">Original Remark</th>
                        <th className="p-3 w-[110px]">FE Name</th>
                      </tr>
                    </thead>
                    <tbody>
                      {replacerData.map((row, idx) => (
                        <tr key={idx} className={cn(
                          "border-b border-slate-100 hover:bg-slate-50 transition-colors group",
                          row.isReplaced ? "bg-green-50/20" : "bg-amber-50/20"
                        )}>
                          <td className="p-2.5 text-[9px] font-black text-slate-400">{row.id}</td>
                          <td className="p-2.5 text-[9px] text-slate-600 font-bold">{row.date}</td>
                          <td 
                            className="p-2.5 font-code text-[11px] font-black text-[#1565C0] cursor-pointer hover:underline"
                            onClick={() => copyToClipboard(row.awb, "AWB")}
                          >
                            {row.awb}
                          </td>
                          <td className="p-2.5">
                            <span className={cn(
                              "inline-flex px-2.5 py-1 rounded-full text-[8px] font-black uppercase tracking-widest border shadow-sm",
                              row.isReplaced ? "bg-green-600 text-white border-green-700" : "bg-amber-500 text-white border-amber-600"
                            )}>
                              {row.officialRemark}
                            </span>
                          </td>
                          <td className="p-2.5 text-[8px] text-slate-400 italic font-bold truncate max-w-[140px]">{row.oldRemark}</td>
                          <td className="p-2.5 text-[9px] font-black text-slate-700 truncate">{row.feName}</td>
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
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=IBM+Plex+Mono:wght@500;600;700&display=swap');
        
        body { font-family: 'Inter', sans-serif; background: #F0F4FA; }
        .font-code { font-family: 'IBM Plex Mono', monospace; }
        
        .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #F1F5F9; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 10px; border: 2px solid #F1F5F9; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #94A3B8; }

        input[type="number"]::-webkit-inner-spin-button,
        input[type="number"]::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }
        
        ::selection { background: #BFDBFE; color: #1E40AF; }
      `}</style>
    </div>
  );
}
