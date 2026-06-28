
"use client";

import React, { useState, useEffect, useMemo, useCallback } from "react";
import { 
  Truck, 
  Trash2, 
  Download, 
  Copy, 
  X, 
  User,
  Hash,
  CheckCircle2,
  FileSpreadsheet,
  Layers,
  Check,
  ArrowRight,
  Info,
  Calendar,
  Filter
} from "lucide-react";
import * as XLSX from "xlsx";
import { cn } from "@/lib/utils";

/**
 * @fileOverview Delhivery POD Management Tool v12.0 (Professional Logistics Edition)
 * Optimized for Palam Vihar RPC. Features Robust Remark Replacer & Filtered Export.
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
  headers: string[];
}

export default function PODTool() {
  const [activeTab, setActiveTab] = useState<"eod" | "remark">("eod");
  const [sessions, setSessions] = useState<Session[]>([]);
  const [selectedSessionId, setSelectedSessionId] = useState<string | null>(null);
  const [statusFilter, setStatusFilter] = useState<string>("all");
  const [activeRemarkChip, setActiveRemarkChip] = useState<string | null>(null);
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
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) setSessions(parsed);
      } catch (e) {
        localStorage.removeItem(STORAGE_KEY);
      }
    }
  }, []);

  useEffect(() => {
    if (isMounted) {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(sessions));
    }
  }, [sessions, isMounted]);

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
      showToast("FE Name and DSP ID are required!", "err");
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
        const wb = XLSX.read(bstr, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(ws);
        
        if (!Array.isArray(rawData) || rawData.length === 0) throw new Error("File empty or invalid");

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

        if (!remarkKey) throw new Error("Remark column not found in this file. Please upload the correct EOD rejection file.");
        if (!awbKey) throw new Error("AWB column not found. Please upload the correct EOD rejection file.");

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

  const getVisibleRows = useCallback(() => {
    const currentSession = sessions.find(s => s.id === selectedSessionId);
    if (!currentSession) return [];
    
    let rows = currentSession.data;
    if (statusFilter === 'pending') rows = rows.filter(r => r.status === 'pending');
    else if (statusFilter === 'dispatched') rows = rows.filter(r => r.status === 'dispatched' || r.status === 'dispatch');
    else if (statusFilter === 'rto') rows = rows.filter(r => r.status === 'rto');
    else if (statusFilter === 'dto') rows = rows.filter(r => r.status === 'dto' || r.status === 'delivered');

    if (activeRemarkChip !== null) {
      rows = rows.filter(r => (r.remark || '') === activeRemarkChip);
    }
    
    return rows;
  }, [sessions, selectedSessionId, statusFilter, activeRemarkChip]);

  const currentSession = useMemo(() => sessions.find(s => s.id === selectedSessionId) || null, [sessions, selectedSessionId]);
  const filteredRows = useMemo(() => getVisibleRows(), [getVisibleRows]);

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

  const pendingRemarkStats = useMemo(() => {
    if (!currentSession) return [];
    const pending = currentSession.data.filter(r => r.status === 'pending');
    const counts: Record<string, number> = {};
    pending.forEach(r => {
      const rem = r.remark || '';
      counts[rem] = (counts[rem] || 0) + 1;
    });
    return Object.entries(counts)
      .map(([text, count]) => ({ text, count }))
      .sort((a, b) => b.count - a.count);
  }, [currentSession]);

  const downloadExcel = () => {
    const rows = getVisibleRows();
    if (!rows.length || !currentSession) return;

    const header = ['Date', 'DSP ID', 'AWB Number', 'Client', 'Order ID', 'Remark', 'FE Name'];
    const excelData = rows.map((r, i) => [
      r.date, 
      i === 0 ? currentSession.dspId : "", 
      { v: r.awb, t: 's', z: '@' }, 
      r.client, 
      r.orderId, 
      r.remark, 
      r.feName
    ]);
    
    const ws = XLSX.utils.aoa_to_sheet([header, ...excelData]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "POD Report");
    
    const filterName = activeRemarkChip ? `_${activeRemarkChip.replace(/\s+/g, '')}` : '';
    const filename = `POD_${currentSession.dspId}_${currentSession.date}_${statusFilter.toUpperCase()}${filterName}.xlsx`;
    XLSX.writeFile(wb, filename);
    showToast(`Downloaded ${rows.length} rows - ${statusFilter.toUpperCase()}`, "ok");
  };

  const downloadReplacerExcel = () => {
    if (!replacerData.length || !replacerMeta) return;
    
    const exportData = replacerData.map(row => {
      const { __isReplaced, __oldRemark, ...rest } = row;
      const cleanRow: any = {};
      Object.keys(rest).forEach(k => {
        if (k === replacerMeta.remarkKey) {
          cleanRow[k] = rest[k];
        } else if (/waybill|awb|awbnumber|waybillno/i.test(k.replace(/[\s_-]/g, ""))) {
          cleanRow[k] = { v: String(rest[k]), t: 's', z: '@' };
        } else {
          cleanRow[k] = rest[k];
        }
      });
      return cleanRow;
    });

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Replaced Remarks");
    
    const date = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `EOD_Official_Remarks_${date}.xlsx`);
  };

  const copyTable = () => {
    const rows = getVisibleRows();
    if (!rows.length || !currentSession) return;
    const text = rows.map((r, i) => `${r.date}\t${i === 0 ? currentSession.dspId : ""}\t${r.awb}\t${r.client}\t${r.orderId}\t${r.remark}\t${r.feName}`).join("\n");
    copyToClipboard(text, "Table Data Copied");
  };

  const clearAllSessions = () => {
    if (confirm("Are you sure you want to clear all data? This will freshly reset the tool.")) {
      setSessions([]);
      setSelectedSessionId(null);
      setReplacerData([]);
      setReplacerMeta(null);
      localStorage.removeItem(STORAGE_KEY);
      showToast("System reset successfully!", "ok");
    }
  };

  const stats = useMemo(() => {
    if (!currentSession) return { total: 0, pending: 0, dispatched: 0, rto: 0, dto: 0 };
    return {
      total: currentSession.data.length,
      pending: currentSession.data.filter(r => r.status === 'pending').length,
      dispatched: currentSession.data.filter(r => r.status === 'dispatched' || r.status === 'dispatch').length,
      rto: currentSession.data.filter(r => r.status === 'rto').length,
      dto: currentSession.data.filter(r => r.status === 'dto' || r.status === 'delivered').length
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
            <h1 className="text-[15px] font-[800] tracking-tight leading-none uppercase">POD Management Tool</h1>
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
            <div className="bg-[#F9A825]/20 px-3 py-1 rounded-full border border-[#F9A825]/30 text-[9px] font-black text-[#F9A825] uppercase">
              {currentSession.data.length} Rows
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
            onClick={() => { setActiveTab(tab.id as any); setStatusFilter("all"); setActiveRemarkChip(null); }}
            className={cn(
              "py-3.5 text-[11px] font-black uppercase tracking-[0.08em] transition-all relative outline-none",
              activeTab === tab.id ? "text-white" : "text-[#94A3B8] hover:text-slate-300"
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
            {/* Setup Section */}
            <div className="bg-white rounded-[14px] p-6 shadow-sm border border-slate-200 space-y-5">
              <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-[#64748B] uppercase tracking-[0.8px] flex items-center gap-2"><Hash className="w-3.5 h-3.5" /> DSP Number</label>
                  <input 
                    type="number" 
                    value={setupData.dspId} 
                    onChange={(e) => setSetupData({...setupData, dspId: e.target.value})}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 text-[13px] font-bold text-[#111827] outline-none focus:border-[#1565C0] transition-all" 
                    placeholder="Enter DSP No..." 
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-[#64748B] uppercase tracking-[0.8px] flex items-center gap-2"><User className="w-3.5 h-3.5" /> FE / Biker Name</label>
                  <input 
                    type="text" 
                    value={setupData.feName} 
                    onChange={(e) => setSetupData({...setupData, feName: e.target.value})}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 text-[13px] font-bold text-[#111827] outline-none focus:border-[#1565C0] transition-all" 
                    placeholder="Enter FE Name..." 
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-[#64748B] uppercase tracking-[0.8px] flex items-center gap-2"><Calendar className="w-3.5 h-3.5" /> Operation Date</label>
                  <input 
                    type="date" 
                    value={setupData.date} 
                    onChange={(e) => setSetupData({...setupData, date: e.target.value})}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 text-[13px] font-bold text-[#111827] outline-none focus:border-[#1565C0]" 
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
                        <p className="text-[11px] text-[#64748B] mt-1 tracking-wide">Excel or CSV files accepted</p>
                      </div>
                    </>
                  )}
                </div>
              </div>
            </div>

            {/* Sessions Grid */}
            {sessions.length > 0 && (
              <div className="space-y-4">
                <div className="flex items-center justify-between px-1">
                  <h2 className="text-[10px] font-black text-[#64748B] uppercase tracking-[0.2em] flex items-center gap-2"><Layers className="w-4 h-4" /> Recent Operation Sessions</h2>
                  <button onClick={clearAllSessions} className="text-[10px] font-black text-red-500 uppercase tracking-widest hover:underline transition-all">
                    Clear All Sessions
                  </button>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-4 gap-4">
                  {sessions.map(s => (
                    <div 
                      key={s.id} 
                      onClick={() => { setSelectedSessionId(s.id); setStatusFilter("all"); setActiveRemarkChip(null); }}
                      className={cn(
                        "bg-white p-4 rounded-xl border-l-4 transition-all cursor-pointer relative group flex flex-col justify-between shadow-sm",
                        selectedSessionId === s.id ? "border-[#1565C0] bg-blue-50/30" : "border-slate-200 hover:border-blue-300"
                      )}
                    >
                      <button 
                        onClick={(e) => { e.stopPropagation(); setSessions(prev => prev.filter(x => x.id !== s.id)); if(selectedSessionId === s.id) setSelectedSessionId(null); }}
                        className="absolute top-3 right-3 opacity-0 group-hover:opacity-100 p-1 text-red-400 hover:text-red-600 transition-opacity"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                      <div>
                        <p className="text-[11px] font-[800] truncate text-[#1C2333] uppercase">{s.feName}</p>
                        <p className="text-[9px] font-mono text-slate-400 mt-0.5">DSP: {s.dspId} · {s.date}</p>
                      </div>
                      <div className="mt-4 grid grid-cols-4 gap-1">
                        {[
                          { val: s.data.length, color: 'bg-slate-100 text-slate-600' },
                          { val: s.data.filter(r => r.status === 'pending').length, color: 'bg-amber-100 text-amber-700' },
                          { val: s.data.filter(r => r.status === 'rto').length, color: 'bg-red-100 text-red-700' },
                          { val: s.data.filter(r => r.status === 'dto').length, color: 'bg-green-100 text-green-700' }
                        ].map((stat, i) => (
                          <div key={i} className={cn("text-[10px] font-black py-1 rounded text-center", stat.color)}>{stat.val}</div>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Module 1 Dashboard */}
            {currentSession && (
              <div className="space-y-6">
                <div className="bg-white rounded-[14px] shadow-sm border border-slate-200 overflow-hidden">
                  <div className="p-4 bg-slate-50 border-b border-slate-200 flex flex-wrap gap-4 items-center justify-between">
                    <div className="flex gap-3">
                      <button onClick={downloadExcel} className="bg-[#2E7D32] hover:bg-[#1B5E20] text-white px-5 py-2.5 rounded-lg font-[700] text-[12px] uppercase tracking-widest flex items-center gap-2 shadow-sm transition-all hover:-translate-y-0.5"><Download className="w-4 h-4" /> Download Excel</button>
                      <button onClick={copyTable} className="bg-[#1565C0] hover:bg-[#0D47A1] text-white px-5 py-2.5 rounded-lg font-[700] text-[12px] uppercase tracking-widest flex items-center gap-2 shadow-sm transition-all hover:-translate-y-0.5"><Copy className="w-4 h-4" /> Copy Table</button>
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
                          onClick={() => { setStatusFilter(t.id); setActiveRemarkChip(null); }}
                          className={cn(
                            "px-4 py-2 rounded-xl border transition-all text-center min-w-[70px]",
                            statusFilter === t.id ? `bg-[#1C2333] border-[#1C2333] text-white shadow-lg` : `border-slate-200 text-[#64748B] hover:bg-white`
                          )}
                        >
                          <p className="text-[14px] font-black leading-none">{t.val}</p>
                          <p className="text-[8px] font-black uppercase mt-1 tracking-tighter">{t.label}</p>
                        </button>
                      ))}
                    </div>
                  </div>

                  {/* Pending Remark Breakdown */}
                  {statusFilter === 'pending' && pendingRemarkStats.length > 0 && (
                    <div className="p-5 border-b border-slate-100 animate-in fade-in slide-in-from-top-2 duration-300">
                      <div className="flex items-center justify-between mb-4">
                        <div>
                          <h3 className="text-[11px] font-black text-[#1C2333] uppercase tracking-wider flex items-center gap-2">
                             <Filter className="w-3.5 h-3.5 text-blue-500" /> Remark Breakdown — Pending
                          </h3>
                          <p className="text-[9px] text-slate-400 font-bold uppercase mt-0.5">
                            {activeRemarkChip ? `Showing: ${activeRemarkChip || 'No Remark'}` : 'Click any remark chip to filter'}
                          </p>
                        </div>
                        {activeRemarkChip !== null && (
                          <button onClick={() => setActiveRemarkChip(null)} className="text-[10px] font-black text-blue-600 uppercase flex items-center gap-1 hover:underline">
                            ← All Pending
                          </button>
                        )}
                      </div>
                      <div className="flex flex-wrap gap-2">
                        {pendingRemarkStats.map((stat, i) => {
                          const isReject = /reject|intact|content|barcode/i.test(stat.text);
                          const isActive = activeRemarkChip === stat.text;
                          return (
                            <button 
                              key={i} 
                              onClick={() => setActiveRemarkChip(isActive ? null : stat.text)}
                              className={cn(
                                "flex items-center gap-2 px-3 py-1.5 rounded-lg border text-[10px] font-bold uppercase tracking-wide transition-all",
                                isReject ? "bg-red-50 border-red-100 text-red-600 hover:bg-red-100" : 
                                isActive ? "bg-[#1565C0] border-[#1565C0] text-white shadow-md" : 
                                "bg-slate-50 border-slate-200 text-slate-600 hover:border-blue-300 hover:text-blue-600"
                              )}
                            >
                              <span>{stat.text || "No Remark"}</span>
                              <span className={cn(
                                "px-1.5 rounded-full text-[9px]",
                                isActive ? "bg-white/20" : "bg-slate-200 text-slate-500"
                              )}>{stat.count}</span>
                            </button>
                          );
                        })}
                      </div>
                    </div>
                  )}

                  <div className="overflow-x-auto max-h-[500px] custom-scrollbar">
                    <table className="w-full text-left">
                      <thead className="sticky top-0 z-30">
                        <tr className="bg-[#1C2333] text-white text-[10px] font-black uppercase tracking-widest">
                          <th className="p-4 w-[100px] border-r border-white/5">DSP ID</th>
                          <th className="p-4 w-[160px] border-r border-white/5">AWB Number</th>
                          <th className="p-4 w-[160px] border-r border-white/5">Client</th>
                          <th className="p-4 w-[160px] border-r border-white/5">Order ID</th>
                          <th className="p-4 border-r border-white/5">Remark</th>
                          <th className="p-4 w-[140px]">FE Name</th>
                          <th className="p-4 w-[60px] text-center">Action</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredRows.map((row, idx) => {
                          const isAWBDupe = duplicateAWBs.has(row.awb);
                          return (
                            <tr key={row.id} className="border-b border-slate-100 hover:bg-blue-50/50 transition-colors">
                              <td className="p-3 font-mono text-[11px] font-bold text-slate-400">{idx === 0 ? currentSession.dspId : ""}</td>
                              <td 
                                className={cn(
                                  "p-3 font-mono text-[13px] font-black tracking-wider cursor-pointer hover:underline",
                                  isAWBDupe ? "text-red-600" : "text-[#1565C0]"
                                )}
                                onClick={() => copyToClipboard(row.awb, "AWB")}
                              >
                                {row.awb}
                              </td>
                              <td className="p-3 text-[11px] font-bold text-[#374151] truncate max-w-[160px]">{row.client}</td>
                              <td className="p-3 text-[11px] text-slate-500 font-medium truncate max-w-[160px]">{row.orderId}</td>
                              <td className="p-3">
                                <span className="text-[11px] font-[600] text-slate-600">
                                  {row.remark || "N/A"}
                                </span>
                              </td>
                              <td className="p-3 text-[11px] font-black text-[#64748B] truncate">{row.feName}</td>
                              <td className="p-3 text-center">
                                <button onClick={() => {
                                  setSessions(prev => prev.map(s => {
                                    if (s.id === selectedSessionId) {
                                      return { ...s, data: s.data.filter(r => r.id !== row.id) };
                                    }
                                    return s;
                                  }));
                                }} className="p-1.5 text-red-400 hover:bg-red-50 hover:text-red-600 rounded-lg transition-all">
                                  <Trash2 className="w-4 h-4" />
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}
          </div>
        ) : (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 animate-in fade-in duration-300">
            {/* Left Side: Upload & Preview */}
            <div className="lg:col-span-7 space-y-6">
              <div className="bg-white rounded-[14px] p-6 shadow-sm border border-slate-200 space-y-5">
                <h2 className="text-[14px] font-[800] text-[#1C2333] flex items-center gap-2"><FileSpreadsheet className="w-5 h-5 text-[#2E7D32]" /> Remark Replacer Dashboard</h2>
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
                        <p className="text-[11px] text-[#64748B] mt-1 tracking-wide">Excel or CSV formats supported</p>
                        {uploadError && (
                          <div className="mt-4 p-3 bg-red-50 border border-red-100 rounded-xl">
                            <p className="text-[11px] text-red-600 font-bold flex items-center justify-center gap-2">
                              <Info className="w-4 h-4" /> {uploadError}
                            </p>
                          </div>
                        )}
                      </div>
                    </>
                  )}
                </div>

                {replacerMeta && (
                  <div className="grid grid-cols-3 gap-4">
                    <div className="bg-white p-4 rounded-xl border border-slate-200 text-center shadow-sm">
                       <p className="text-2xl font-black text-[#1C2333]">{replacerMeta.total}</p>
                       <p className="text-[9px] font-black text-[#64748B] uppercase tracking-[0.2em] mt-1">Total Packets</p>
                    </div>
                    <div className="bg-white p-4 rounded-xl border border-slate-200 text-center shadow-sm">
                       <p className="text-2xl font-black text-[#2E7D32]">{replacerMeta.replaced}</p>
                       <p className="text-[9px] font-black text-[#64748B] uppercase tracking-[0.2em] mt-1">Remarks Replaced</p>
                    </div>
                    <div className="bg-white p-4 rounded-xl border border-slate-200 text-center shadow-sm">
                       <p className="text-2xl font-black text-[#F57F17]">{replacerMeta.missing}</p>
                       <p className="text-[9px] font-black text-[#64748B] uppercase tracking-[0.2em] mt-1">No Mapping Found</p>
                    </div>
                  </div>
                )}
              </div>

              {replacerData.length > 0 && (
                <div className="bg-white rounded-[14px] shadow-sm border border-slate-200 overflow-hidden">
                  <div className="p-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
                    <div className="flex gap-3">
                      <button onClick={downloadReplacerExcel} className="bg-[#2E7D32] hover:bg-[#1B5E20] text-white px-6 py-2.5 rounded-lg font-[700] text-[12px] uppercase tracking-widest flex items-center gap-2 shadow-sm transition-all hover:-translate-y-0.5"><Download className="w-4 h-4" /> Download Official Excel</button>
                    </div>
                    <button onClick={() => { setReplacerData([]); setReplacerMeta(null); }} className="text-[10px] font-black text-red-500 uppercase tracking-widest hover:underline px-4">Clear Preview</button>
                  </div>
                  
                  <div className="overflow-x-auto max-h-[500px] custom-scrollbar">
                    <table className="w-full text-left border-collapse">
                      <thead className="sticky top-0 z-30">
                        <tr className="bg-[#1C2333] text-white text-[10px] font-black uppercase tracking-widest">
                          {replacerMeta?.headers.map((h, i) => (
                            <th key={i} className="p-4 border-r border-white/5 whitespace-nowrap">{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {replacerData.map((row, idx) => (
                          <tr key={idx} className={cn(
                            "border-b border-slate-100 transition-colors",
                            row.__isReplaced ? "bg-[#F0FDF4]" : "bg-[#FFFDE7]"
                          )}>
                            {replacerMeta?.headers.map((h, i) => {
                              const isRemarkCol = h === replacerMeta.remarkKey;
                              const isAWBCol = /waybill|awb|awbnumber|waybillno/i.test(h.replace(/[\s_-]/g, ""));
                              return (
                                <td key={i} className={cn(
                                  "p-3 text-[11px] font-medium border-r border-slate-100",
                                  isAWBCol ? "font-mono text-[#1565C0] font-black" : "text-[#374151]",
                                  isRemarkCol && row.__isReplaced ? "text-[#2E7D32] font-black" : 
                                  isRemarkCol && !row.__isReplaced ? "text-[#D97706] font-black" : ""
                                )}>
                                  {row[h]}
                                </td>
                              );
                            })}
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
                  <p className="text-[10px] text-green-100 mt-1 font-medium">Auto-replace NSL Remarks with Official Delhivery Remarks</p>
                </div>
                <div className="p-4 overflow-y-auto max-h-[70vh] space-y-3 custom-scrollbar">
                  {Object.entries(REMARK_MAPPING).map(([nsl, official]) => (
                    <div key={nsl} className="p-4 bg-[#F8FAFC] rounded-xl border border-slate-100 space-y-3 hover:border-green-200 transition-all">
                      <div className="flex flex-col gap-2">
                        <span className="text-[8px] font-black text-[#64748B] uppercase tracking-widest">NSL REMARK</span>
                        <div className="bg-amber-100 text-amber-800 px-3 py-1.5 rounded-lg text-[10px] font-[700] border border-amber-200 inline-block w-fit">
                          {nsl}
                        </div>
                      </div>
                      <div className="flex items-center justify-center">
                        <ArrowRight className="w-4 h-4 text-green-500" />
                      </div>
                      <div className="flex flex-col gap-2">
                        <span className="text-[8px] font-black text-[#64748B] uppercase tracking-widest">OFFICIAL REMARK</span>
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
        
        body { font-family: 'Inter', sans-serif; background: #F0F4FA; color: #374151; }
        
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
