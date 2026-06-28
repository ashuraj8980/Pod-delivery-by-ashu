
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
  FileWarning,
  ArrowRight,
  Info
} from "lucide-react";
import * as XLSX from "xlsx";
import { cn } from "@/lib/utils";

/**
 * @fileOverview Delhivery POD Management Tool v11.0 (Professional Logistics Edition)
 * Optimized for Palam Vihar RPC. Features Robust Remark Replacer & Crash Protection.
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

interface ReplacerMeta {
  total: number;
  replaced: number;
  missing: number;
  remarkKey: string;
  awbKey: string;
  headers: string[];
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
  const [uploadError, setUploadError] = useState<string | null>(null);
  
  // Module 2 State
  const [replacerData, setReplacerData] = useState<any[]>([]);
  const [replacerMeta, setReplacerMeta] = useState<ReplacerMeta | null>(null);

  useEffect(() => {
    setIsMounted(true);
    setSetupData(prev => ({ ...prev, date: new Date().toISOString().split('T')[0] }));
  }, []);

  const showToast = useCallback((msg: string, type: 'ok' | 'err' | 'info') => {
    if (!isMounted) return;
    const toast = document.createElement('div');
    toast.className = `fixed bottom-10 left-1/2 -translate-x-1/2 px-6 py-3 rounded-xl text-[12px] font-bold z-[200] shadow-2xl transition-all duration-300 transform scale-95 opacity-0 animate-in fade-in slide-in-from-bottom-5 border ${
      type === 'ok' ? 'bg-[#052E0F] text-[#6EE7A6] border-[#6EE7A6]/20' : 
      type === 'err' ? 'bg-[#2D0808] text-[#FCA5A5] border-[#FCA5A5]/20' : 
      'bg-[#1C2333] text-[#93C5FD] border-[#93C5FD]/20'
    }`;
    toast.style.opacity = '1';
    toast.style.transform = 'translateX(-50%) scale(1)';
    toast.innerHTML = `<span class="flex items-center gap-2">${type === 'ok' ? '✓' : type === 'err' ? '✕' : 'ℹ'} ${msg}</span>`;
    document.body.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateX(-50%) scale(0.9)';
      setTimeout(() => toast.remove(), 300);
    }, 3000);
  }, [isMounted]);

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

  const copyToClipboard = (text: string, label: string = "Copied") => {
    if (!text) return;
    navigator.clipboard.writeText(text).then(() => {
      showToast(`${label}: ${text.substring(0, 15)}...`, "ok");
    }).catch(() => {
      const textArea = document.createElement("textarea");
      textArea.value = text;
      document.body.appendChild(textArea);
      textArea.select();
      document.execCommand("copy");
      textArea.remove();
      showToast(`${label}: ${text.substring(0, 15)}...`, "ok");
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
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(ws);
        
        if (!Array.isArray(rawData) || rawData.length === 0) throw new Error("File empty or invalid");

        const parsedRows: PODRow[] = rawData.slice(0, 10000).map((row: any) => {
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
          showToast("No valid rows found!", "err");
        }
      } catch (err: any) {
        showToast("Error parsing file! Please upload valid Delhivery export.", "err");
      } finally {
        setIsProcessing(false);
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  const handleReplacerUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsProcessing(true);
    setUploadError(null);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const extension = file.name.split('.').pop()?.toLowerCase();
        let wb;
        
        if (['csv', 'tsv'].includes(extension || '')) {
          wb = XLSX.read(bstr, { type: 'string' });
        } else {
          wb = XLSX.read(bstr, { type: 'binary' });
        }

        const ws = wb.Sheets[wb.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(ws, { defval: "" });
        
        if (!rawData.length) throw new Error("File empty");

        const keys = Object.keys(rawData[0]);
        const remarkKey = keys.find(k => /remarksof nsl|remark|remarks|nsl remark/i.test(k.replace(/[\s_-]/g, "")));
        const awbKey = keys.find(k => /waybill|awb|awbnumber|waybillno/i.test(k.replace(/[\s_-]/g, "")));

        if (!remarkKey) {
          throw new Error("Remark column not found in this file. Please upload the correct EOD rejection file.");
        }
        if (!awbKey) {
          throw new Error("AWB column not found. Please upload the correct EOD rejection file.");
        }

        let replacedCount = 0;
        let missingCount = 0;

        const processed = rawData.map((row: any) => {
          const oldRemark = String(row[remarkKey]).trim();
          let newRemark = oldRemark;
          let isReplaced = false;

          // Mapping logic: Exact then Partial
          const exactKey = Object.keys(REMARK_MAPPING).find(k => k.toLowerCase() === oldRemark.toLowerCase());
          if (exactKey) {
            newRemark = REMARK_MAPPING[exactKey];
            isReplaced = true;
          } else {
            const partialKey = Object.keys(REMARK_MAPPING).find(k => oldRemark.toLowerCase().includes(k.toLowerCase()));
            if (partialKey) {
              newRemark = REMARK_MAPPING[partialKey];
              isReplaced = true;
            }
          }

          if (isReplaced) replacedCount++; else missingCount++;

          return {
            ...row,
            [remarkKey]: newRemark,
            [awbKey]: fixAWB(row[awbKey]),
            __isReplaced: isReplaced,
            __oldRemark: oldRemark
          };
        });

        setReplacerData(processed);
        setReplacerMeta({
          total: processed.length,
          replaced: replacedCount,
          missing: missingCount,
          remarkKey,
          awbKey,
          headers: keys
        });
        showToast(`Processed ${processed.length} rows successfully`, "ok");
      } catch (err: any) {
        setUploadError(err.message || "Failed to read file");
        showToast(err.message || "Replacer error!", "err");
      } finally {
        setIsProcessing(false);
      }
    };
    
    if (file.name.endsWith('.csv') || file.name.endsWith('.tsv')) {
      reader.readAsText(file);
    } else {
      reader.readAsBinaryString(file);
    }
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

  const downloadReplacerExcel = () => {
    if (!replacerData.length || !replacerMeta) return;
    
    // Clean data for export (remove internal keys)
    const exportData = replacerData.map(row => {
      const { __isReplaced, __oldRemark, ...rest } = row;
      // Force AWB as string
      if (replacerMeta.awbKey in rest) {
        return { ...rest, [replacerMeta.awbKey]: { v: String(rest[replacerMeta.awbKey]), t: 's' } };
      }
      return rest;
    });

    const ws = XLSX.utils.json_to_sheet(exportData);
    
    // Header Styling (Simulated via AOAA)
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Replaced Remarks");
    
    const date = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `EOD_Remarks_Replaced_${date}.xlsx`);
  };

  const copyReplacerTable = () => {
    if (!replacerData.length || !replacerMeta) return;
    const text = replacerData.map(r => {
      return replacerMeta.headers.map(h => r[h]).join("\t");
    }).join("\n");
    copyToClipboard(text, "Table Data Copied");
  };

  const clearAllSessions = () => {
    if (confirm("KYA AAP SAARA DATA CLEAR KARNA CHAHTE HAIN? Ye browser memory se bhi hat jayega.")) {
      setSessions([]);
      setSelectedSessionId(null);
      setReplacerData([]);
      setReplacerMeta(null);
      localStorage.removeItem(STORAGE_KEY);
      localStorage.clear();
      showToast("Browser storage cleared successfully!", "ok");
    }
  };

  const stats = useMemo(() => {
    if (!currentSession) return { total: 0, pending: 0, dispatched: 0, rto: 0, dto: 0 };
    return {
      total: currentSession.data.length,
      pending: currentSession.data.filter(r => r.status === 'pending').length,
      dispatched: currentSession.data.filter(r => r.status === 'dispatched').length,
      rto: currentSession.data.filter(r => r.status === 'rto').length,
      dto: currentSession.data.filter(r => r.status === 'dto').length
    };
  }, [currentSession]);

  if (!isMounted) return null;

  return (
    <div className="min-h-screen bg-[#F0F4FA] font-body text-[#374151] select-auto overflow-x-hidden">
      {/* Rainbow Stripe */}
      <div className="h-[3px] w-full bg-gradient-to-r from-[#1565C0] via-[#F9A825] via-[#2E7D32] to-[#D32F2F] sticky top-0 z-[100]" />
      
      {/* Header */}
      <header className="h-[58px] bg-[#1C2333] px-6 flex items-center justify-between text-white shadow-lg relative z-[90]">
        <div className="flex items-center gap-3">
          <div className="bg-gradient-to-br from-blue-500 to-blue-700 p-1.5 rounded-xl shadow-lg">
            <Truck className="w-5 h-5 text-white" />
          </div>
          <div className="flex flex-col">
            <h1 className="text-[14px] font-[800] tracking-tight leading-none uppercase">POD Management Tool</h1>
            <p className="text-[9px] text-slate-400 font-bold mt-1 uppercase tracking-wider">
              Delhivery · Palam Vihar RPC · <span className="text-[#F9A825] font-black italic">By Ashu</span>
            </p>
          </div>
        </div>
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2 bg-[#2E7D32]/20 px-3 py-1 rounded-full border border-[#2E7D32]/30">
            <div className="w-2 h-2 bg-[#2E7D32] rounded-full animate-pulse" />
            <span className="text-[9px] font-black text-[#6EE7A6] uppercase tracking-[0.1em]">Live</span>
          </div>
          {currentSession && (
            <div className="bg-[#1565C0]/20 px-3 py-1 rounded-full border border-[#1565C0]/30 text-[9px] font-black text-blue-400 uppercase">
              {currentSession.data.length} Shipments
            </div>
          )}
        </div>
      </header>

      {/* Navigation */}
      <nav className="bg-[#1C2333] px-6 flex gap-8 border-t border-white/5 shadow-md sticky top-[58px] z-[85]">
        {[
          { id: "eod", label: "Daily EOD Rejection" },
          { id: "remark", label: "EOD Rejection Remark" }
        ].map(tab => (
          <button 
            key={tab.id}
            onClick={() => { setActiveTab(tab.id as any); setStatusFilter("all"); setRemarkFilter(null); }}
            className={cn(
              "py-3.5 text-[11px] font-black uppercase tracking-[0.08em] transition-all relative outline-none",
              activeTab === tab.id ? "text-white" : "text-slate-500 hover:text-slate-300"
            )}
          >
            {tab.label}
            {activeTab === tab.id && <div className="absolute bottom-0 left-0 w-full h-[3px] bg-[#F9A825] rounded-t-full shadow-lg" />}
          </button>
        ))}
      </nav>

      <main className="p-5 max-w-[1360px] mx-auto space-y-6">
        {activeTab === "eod" ? (
          <div className="space-y-6 animate-in fade-in duration-300">
            {/* Setup & Upload Section */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
              <div className="lg:col-span-8 bg-white rounded-[14px] p-6 shadow-sm border border-slate-200 space-y-5">
                <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-[#64748B] uppercase tracking-[0.8px] flex items-center gap-2"><Hash className="w-3.5 h-3.5" /> DSP Number</label>
                    <input 
                      type="number" 
                      value={setupData.dspId} 
                      onChange={(e) => setSetupData({...setupData, dspId: e.target.value})}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-[13px] font-bold text-[#111827] outline-none focus:border-[#1565C0] focus:ring-1 focus:ring-blue-500/10 transition-all" 
                      placeholder="Enter DSP No..." 
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-[#64748B] uppercase tracking-[0.8px] flex items-center gap-2"><User className="w-3.5 h-3.5" /> FE / Biker Name</label>
                    <input 
                      type="text" 
                      value={setupData.feName} 
                      onChange={(e) => setSetupData({...setupData, feName: e.target.value})}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-[13px] font-bold text-[#111827] outline-none focus:border-[#1565C0] focus:ring-1 focus:ring-blue-500/10 transition-all" 
                      placeholder="Enter FE Name..." 
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-[#64748B] uppercase tracking-[0.8px]">Operation Date</label>
                    <input 
                      type="date" 
                      value={setupData.date} 
                      onChange={(e) => setSetupData({...setupData, date: e.target.value})}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-[13px] font-bold text-[#111827] outline-none focus:border-[#1565C0]" 
                    />
                  </div>
                </div>
                
                <div 
                  className={cn(
                    "border-2 border-dashed rounded-2xl p-8 text-center transition-all relative group overflow-hidden",
                    (!setupData.feName || !setupData.dspId) ? "bg-slate-50 border-slate-200 opacity-50 cursor-not-allowed" : "border-[#CBD5E1] bg-[#F8FAFC] hover:border-[#1976D2] hover:bg-[#E3F2FD] cursor-pointer"
                  )}
                >
                  <input 
                    type="file" 
                    disabled={!setupData.feName || !setupData.dspId || isProcessing} 
                    onChange={handleFileUpload} 
                    className="absolute inset-0 opacity-0 cursor-pointer disabled:cursor-not-allowed z-10" 
                    accept=".xlsx,.xls,.csv" 
                  />
                  <div className="space-y-3">
                    {isProcessing ? (
                      <div className="flex flex-col items-center gap-3">
                         <div className="w-10 h-10 border-[3px] border-blue-200 border-t-blue-600 rounded-full animate-spin" />
                         <p className="text-[12px] font-[700] text-blue-600 uppercase tracking-widest">Processing Data...</p>
                      </div>
                    ) : (
                      <>
                        <Download className={cn("w-8 h-8 mx-auto transition-all", (!setupData.feName || !setupData.dspId) ? "text-slate-300" : "text-[#1565C0]")} />
                        <div>
                          <p className="text-[14px] font-[600] text-[#111827]">
                            {(!setupData.feName || !setupData.dspId) ? "Complete Setup to Upload" : "Upload Delhivery EOD Export"}
                          </p>
                          <p className="text-[11px] text-[#64748B] mt-1 tracking-wide">Drop Excel or CSV file here</p>
                        </div>
                      </>
                    )}
                  </div>
                </div>
              </div>

              <div className="lg:col-span-4 bg-[#1C2333] rounded-[14px] p-6 shadow-xl text-white relative overflow-hidden">
                <div className="absolute top-0 right-0 w-32 h-32 bg-blue-500/10 rounded-full -mr-16 -mt-16 blur-3xl" />
                <h2 className="text-[10px] font-black text-[#F9A825] uppercase tracking-[0.2em] mb-6 flex items-center gap-2">
                  <Info className="w-4 h-4" /> System Quick Stats
                </h2>
                <div className="grid grid-cols-2 gap-4">
                  <div className="bg-white/5 p-4 rounded-xl border border-white/10">
                    <p className="text-[8px] font-black text-slate-400 uppercase tracking-widest">Active Sessions</p>
                    <p className="text-2xl font-black mt-1">{sessions.length}</p>
                  </div>
                  <div className="bg-white/5 p-4 rounded-xl border border-white/10">
                    <p className="text-[8px] font-black text-slate-400 uppercase tracking-widest">Total Packets</p>
                    <p className="text-2xl font-black mt-1">{sessions.reduce((acc, s) => acc + s.data.length, 0)}</p>
                  </div>
                </div>
                <div className="mt-6">
                  <button onClick={clearAllSessions} className="w-full py-3 bg-red-600/10 hover:bg-red-600 text-red-500 hover:text-white border border-red-500/20 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all active:scale-95 flex items-center justify-center gap-2">
                    <Trash2 className="w-4 h-4" /> Reset All System Data
                  </button>
                </div>
              </div>
            </div>

            {/* Session Cards */}
            {sessions.length > 0 && (
              <div className="space-y-4">
                <h2 className="text-[10px] font-black text-[#64748B] uppercase tracking-[0.2em] flex items-center gap-2 px-1"><Layers className="w-4 h-4" /> Recent Operation Sessions</h2>
                <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-4">
                  {sessions.map(s => (
                    <div 
                      key={s.id} 
                      onClick={() => { setSelectedSessionId(s.id); setStatusFilter("all"); setRemarkFilter(null); }}
                      className={cn(
                        "bg-white p-4 rounded-2xl border transition-all cursor-pointer relative group",
                        selectedSessionId === s.id ? "border-[#1565C0] ring-4 ring-blue-500/5 shadow-md" : "border-slate-200 hover:border-blue-300"
                      )}
                    >
                      <button 
                        onClick={(e) => { e.stopPropagation(); setSessions(prev => prev.filter(x => x.id !== s.id)); if(selectedSessionId === s.id) setSelectedSessionId(null); }}
                        className="absolute top-3 right-3 opacity-0 group-hover:opacity-100 p-1 text-red-400 hover:text-red-600 transition-opacity"
                      >
                        <X className="w-4 h-4" />
                      </button>
                      <p className="text-[11px] font-[800] truncate text-[#1C2333] uppercase">{s.feName}</p>
                      <p className="text-[9px] font-mono text-slate-400 mt-0.5">DSP: {s.dspId}</p>
                      <div className="mt-4 flex justify-between items-end">
                        <span className="text-[14px] font-[900] text-[#1565C0]">{s.data.length} PKT</span>
                        <span className="text-[8px] font-bold text-slate-300 uppercase">{new Date(s.timestamp).toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})}</span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Dashboard */}
            {currentSession && (
              <div className="bg-white rounded-[14px] shadow-sm border border-slate-200 overflow-hidden animate-in fade-in zoom-in-95 duration-500">
                <div className="p-4 bg-slate-50 border-b border-slate-200 flex flex-wrap gap-4 items-center justify-between">
                  <div className="flex gap-3">
                    <button onClick={() => {
                        const header = ['Date', 'DSP ID', 'AWB Number', 'Client', 'Order ID', 'Remark', 'FE Name'];
                        const rows = filteredRows.map((r, i) => [r.date, i === 0 ? currentSession.dspId : "", { v: r.awb, t: 's' }, r.client, r.orderId, r.remark, r.feName]);
                        const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
                        const wb = XLSX.utils.book_new();
                        XLSX.utils.book_append_sheet(wb, ws, "Report");
                        XLSX.writeFile(wb, `Delhivery_EOD_${currentSession.feName}.xlsx`);
                    }} className="bg-[#2E7D32] hover:bg-[#1B5E20] text-white px-5 py-2.5 rounded-lg font-[700] text-[12px] uppercase tracking-widest flex items-center gap-2 shadow-sm transition-all active:scale-95 hover:-translate-y-0.5"><Download className="w-4 h-4" /> Export {statusFilter.toUpperCase()}</button>
                    <button onClick={() => {
                        const text = filteredRows.map((r, i) => `${r.date}\t${i === 0 ? currentSession.dspId : ""}\t${r.awb}\t${r.client}\t${r.orderId}\t${r.remark}\t${r.feName}`).join("\n");
                        copyToClipboard(text, "Table Data Copied");
                    }} className="bg-[#1565C0] hover:bg-[#0D47A1] text-white px-5 py-2.5 rounded-lg font-[700] text-[12px] uppercase tracking-widest flex items-center gap-2 shadow-sm transition-all active:scale-95 hover:-translate-y-0.5"><Copy className="w-4 h-4" /> Copy For Excel</button>
                  </div>
                  <div className="flex items-center gap-2">
                    {[
                      { id: 'all', label: 'All', val: stats.total, color: '#1C2333' },
                      { id: 'pending', label: 'Pending', val: stats.pending, color: '#D97706' },
                      { id: 'dispatched', label: 'Disp', val: stats.dispatched, color: '#1565C0' },
                      { id: 'rto', label: 'RTO', val: stats.rto, color: '#DC2626' },
                      { id: 'dto', label: 'DTO', val: stats.dto, color: '#2E7D32' }
                    ].map(t => (
                      <button 
                        key={t.id} 
                        onClick={() => { setStatusFilter(t.id); setRemarkFilter(null); }}
                        className={cn(
                          "px-4 py-2 rounded-xl border transition-all text-center min-w-[70px]",
                          statusFilter === t.id ? `bg-[#1C2333] border-[#1C2333] text-white shadow-lg scale-105` : `border-slate-200 text-[#64748B] hover:bg-white`
                        )}
                      >
                        <p className="text-[14px] font-black leading-none">{t.val}</p>
                        <p className="text-[8px] font-black uppercase mt-1 tracking-tighter">{t.label}</p>
                      </button>
                    ))}
                  </div>
                </div>

                <div className="overflow-x-auto max-h-[500px] custom-scrollbar">
                  <table className="w-full text-left border-collapse">
                    <thead className="sticky top-0 z-30">
                      <tr className="bg-[#1C2333] text-white text-[10px] font-black uppercase tracking-widest">
                        <th className="p-4 w-[50px] text-center border-r border-white/5">Sel</th>
                        <th className="p-4 w-[100px] border-r border-white/5">DSP ID</th>
                        <th className="p-4 w-[160px] border-r border-white/5">AWB Number</th>
                        <th className="p-4 w-[160px] border-r border-white/5">Client</th>
                        <th className="p-4 w-[160px] border-r border-white/5">Order ID</th>
                        <th className="p-4 border-r border-white/5">Remark</th>
                        <th className="p-4 w-[140px]">FE Name</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredRows.map((row, idx) => {
                        const isAWBDupe = duplicateAWBs.has(row.awb);
                        return (
                          <tr key={row.id} className={cn(
                            "border-b border-slate-100 hover:bg-blue-50/50 transition-colors group",
                            row.selected ? "bg-blue-50" : ""
                          )}>
                            <td className="p-3 text-center">
                              <input type="checkbox" checked={row.selected} onChange={() => {
                                setSessions(prev => prev.map(s => {
                                  if (s.id === selectedSessionId) {
                                    return { ...s, data: s.data.map(r => r.id === row.id ? { ...r, selected: !r.selected } : r) };
                                  }
                                  return s;
                                }));
                              }} className="w-4 h-4 accent-[#1565C0] cursor-pointer" />
                            </td>
                            <td className="p-3 font-mono text-[11px] font-bold text-slate-400">{idx === 0 ? row.dspId : ""}</td>
                            <td 
                              className={cn(
                                "p-3 font-mono text-[13px] font-black tracking-wider cursor-pointer hover:underline flex items-center gap-2",
                                isAWBDupe ? "text-red-600" : "text-[#1565C0]"
                              )}
                              onClick={() => copyToClipboard(row.awb, "AWB")}
                            >
                              {row.awb}
                            </td>
                            <td className="p-3 text-[11px] font-bold text-[#374151] truncate max-w-[160px]">{row.client}</td>
                            <td className="p-3 text-[11px] text-slate-500 font-medium truncate max-w-[160px]">{row.orderId}</td>
                            <td className="p-3">
                              <span className="inline-flex px-3 py-1 rounded-full text-[9px] font-black uppercase tracking-widest border bg-amber-50 text-amber-700 border-amber-200">
                                {row.remark || "N/A"}
                              </span>
                            </td>
                            <td className="p-3 text-[11px] font-black text-[#64748B] truncate">{row.feName}</td>
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
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 animate-in fade-in duration-300">
            {/* Left Side: Upload & Processing */}
            <div className="lg:col-span-7 space-y-6">
              <div className="bg-white rounded-[14px] p-6 shadow-sm border border-slate-200 space-y-5">
                <h2 className="text-[14px] font-[800] text-[#1C2333] flex items-center gap-2"><FileSpreadsheet className="w-5 h-5 text-[#2E7D32]" /> Official Remark Replacer</h2>
                <div 
                  className={cn(
                    "border-2 border-dashed rounded-2xl p-10 text-center transition-all relative group overflow-hidden",
                    isProcessing ? "bg-slate-50 border-slate-200 opacity-50 cursor-not-allowed" : "border-[#CBD5E1] bg-[#F8FAFC] hover:border-[#1976D2] hover:bg-[#E3F2FD] cursor-pointer"
                  )}
                >
                  <input 
                    type="file" 
                    disabled={isProcessing} 
                    onChange={handleReplacerUpload} 
                    className="absolute inset-0 opacity-0 cursor-pointer z-20" 
                    accept=".xlsx,.xls,.csv,.tsv,.ods" 
                  />
                  {isProcessing ? (
                    <div className="flex flex-col items-center gap-3">
                       <div className="w-10 h-10 border-[3px] border-green-200 border-t-green-600 rounded-full animate-spin" />
                       <p className="text-[12px] font-[700] text-green-600 uppercase tracking-widest">Replacing Remarks...</p>
                    </div>
                  ) : (
                    <>
                      <div className="w-12 h-12 bg-blue-500/10 rounded-full flex items-center justify-center mx-auto mb-4 group-hover:scale-110 transition-transform">
                        <Download className="w-6 h-6 text-[#1565C0]" />
                      </div>
                      <div>
                        <p className="text-[15px] font-[700] text-[#111827]">Upload EOD Rejection File</p>
                        <p className="text-[11px] text-[#64748B] mt-1 tracking-wide">Excel, CSV, TSV supported</p>
                        {uploadError && (
                          <p className="mt-3 text-[11px] text-red-600 font-bold bg-red-50 p-2 rounded-lg border border-red-100 flex items-center justify-center gap-2">
                            <AlertTriangle className="w-4 h-4" /> {uploadError}
                          </p>
                        )}
                      </div>
                    </>
                  )}
                </div>

                {replacerMeta && (
                  <div className="grid grid-cols-3 gap-4">
                    <div className="bg-[#1C2333] p-5 rounded-2xl text-center shadow-lg border-b-4 border-black/20">
                       <p className="text-3xl font-black text-white">{replacerMeta.total}</p>
                       <p className="text-[9px] font-black text-slate-400 uppercase tracking-[0.2em] mt-1">Total Packets</p>
                    </div>
                    <div className="bg-[#2E7D32] p-5 rounded-2xl text-center shadow-lg border-b-4 border-black/20">
                       <p className="text-3xl font-black text-white">{replacerMeta.replaced}</p>
                       <p className="text-[9px] font-black text-[#6EE7A6] uppercase tracking-[0.2em] mt-1">Remarks Replaced</p>
                    </div>
                    <div className="bg-[#F57F17] p-5 rounded-2xl text-center shadow-lg border-b-4 border-black/20">
                       <p className="text-3xl font-black text-white">{replacerMeta.missing}</p>
                       <p className="text-[9px] font-black text-[#FFFDE7] uppercase tracking-[0.2em] mt-1">No Mapping Found</p>
                    </div>
                  </div>
                )}
              </div>

              {replacerData.length > 0 && (
                <div className="bg-white rounded-[14px] shadow-sm border border-slate-200 overflow-hidden animate-in slide-in-from-bottom-5 duration-500">
                  <div className="p-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
                    <div className="flex gap-3">
                      <button onClick={downloadReplacerExcel} className="bg-[#2E7D32] hover:bg-[#1B5E20] text-white px-6 py-2.5 rounded-lg font-[700] text-[12px] uppercase tracking-widest flex items-center gap-2 shadow-sm transition-all hover:-translate-y-0.5"><Download className="w-4 h-4" /> Download Replaced Excel</button>
                      <button onClick={copyReplacerTable} className="bg-[#1565C0] hover:bg-[#0D47A1] text-white px-6 py-2.5 rounded-lg font-[700] text-[12px] uppercase tracking-widest flex items-center gap-2 shadow-sm transition-all hover:-translate-y-0.5"><Copy className="w-4 h-4" /> Copy Table</button>
                    </div>
                    <button onClick={() => { setReplacerData([]); setReplacerMeta(null); }} className="text-[10px] font-black text-red-500 uppercase tracking-widest hover:underline px-4">Clear All</button>
                  </div>
                  
                  <div className="overflow-x-auto max-h-[500px] custom-scrollbar">
                    <table className="w-full text-left">
                      <thead className="sticky top-0 z-30">
                        <tr className="bg-[#1C2333] text-white text-[10px] font-black uppercase tracking-widest">
                          <th className="p-4 w-[50px] text-center border-r border-white/5">#</th>
                          <th className="p-4 w-[160px] border-r border-white/5">AWB Number</th>
                          <th className="p-4 border-r border-white/5">Updated Remark</th>
                          <th className="p-4 w-[180px]">Original Remark</th>
                        </tr>
                      </thead>
                      <tbody>
                        {replacerData.map((row, idx) => (
                          <tr key={idx} className={cn(
                            "border-b border-slate-100 transition-colors group",
                            row.__isReplaced ? "bg-[#F0FDF4]" : "bg-[#FFFDE7]"
                          )}>
                            <td className="p-3 text-[10px] font-black text-slate-400 text-center">{idx + 1}</td>
                            <td className="p-3 font-mono text-[13px] font-black text-[#1565C0] cursor-pointer hover:underline" onClick={() => copyToClipboard(row[replacerMeta!.awbKey], "AWB")}>
                              {row[replacerMeta!.awbKey]}
                            </td>
                            <td className="p-3">
                              <span className={cn(
                                "inline-flex px-3 py-1.5 rounded-lg text-[10px] font-[700] uppercase tracking-wider",
                                row.__isReplaced ? "text-[#2E7D32]" : "text-[#D97706]"
                              )}>
                                {row[replacerMeta!.remarkKey]}
                              </span>
                            </td>
                            <td className="p-3 text-[10px] text-slate-400 italic font-medium truncate max-w-[180px]">
                              {row.__oldRemark}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>

            {/* Right Side: Mapping Reference */}
            <div className="lg:col-span-5">
              <div className="bg-white rounded-[14px] shadow-sm border border-slate-200 overflow-hidden sticky top-[130px]">
                <div className="bg-[#2E7D32] p-4 text-white">
                  <h3 className="text-[12px] font-[800] uppercase tracking-widest flex items-center gap-2">
                    <CheckCircle2 className="w-5 h-5" /> Built-in Remark Mapping
                  </h3>
                  <p className="text-[10px] text-green-100 mt-1 font-medium">Automatic NSL to Official Delhivery Remark Conversion</p>
                </div>
                <div className="p-4 overflow-y-auto max-h-[70vh] space-y-3 custom-scrollbar">
                  {Object.entries(REMARK_MAPPING).map(([nsl, official]) => (
                    <div key={nsl} className="p-4 bg-[#F8FAFC] rounded-xl border border-slate-100 space-y-3 hover:border-green-200 transition-all">
                      <div className="flex flex-col gap-2">
                        <span className="text-[8px] font-black text-slate-400 uppercase tracking-widest">NSL REMARK</span>
                        <div className="bg-amber-100 text-amber-800 px-3 py-1.5 rounded-lg text-[10px] font-[700] border border-amber-200 inline-block w-fit">
                          {nsl}
                        </div>
                      </div>
                      <div className="flex items-center justify-center">
                        <ArrowRight className="w-4 h-4 text-green-500" />
                      </div>
                      <div className="flex flex-col gap-2">
                        <span className="text-[8px] font-black text-slate-400 uppercase tracking-widest">OFFICIAL REMARK</span>
                        <div className="bg-green-600 text-white px-3 py-1.5 rounded-lg text-[10px] font-[700] border border-green-700 inline-block w-fit shadow-md">
                          {official}
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        )}
      </main>

      <style jsx global>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=IBM+Plex+Mono:wght@500;600;700&display=swap');
        
        body { font-family: 'Inter', sans-serif; background: #F0F4FA; }
        
        .custom-scrollbar::-webkit-scrollbar { width: 8px; height: 8px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #F1F5F9; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 10px; border: 3px solid #F1F5F9; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #94A3B8; }

        input[type="number"]::-webkit-inner-spin-button,
        input[type="number"]::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }
        
        ::selection { background: #BFDBFE; color: #1E40AF; }

        @keyframes pulse-dot {
          0%, 100% { transform: scale(1); opacity: 1; }
          50% { transform: scale(1.2); opacity: 0.7; }
        }
      `}</style>
    </div>
  );
}

