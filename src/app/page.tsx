
"use client";

import React, { useState, useEffect, useMemo, useCallback } from "react";
import { 
  Truck, 
  Trash2, 
  Download, 
  Copy, 
  X, 
  CheckCircle2, 
  FileSpreadsheet, 
  Loader2, 
  Search,
  Star,
  Plus
} from "lucide-react";
import * as XLSX from "xlsx";
import { cn } from "@/lib/utils";

/**
 * @fileOverview Delhivery POD Management Tool - Enterprise Edition (Pixel Perfect)
 * Optimized for Palam Vihar RPC. Built for Ashu.
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
  isIntact?: boolean;
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
  const [activeRemarkChip, setActiveRemarkChip] = useState<string | null>(null);
  const [setupData, setSetupData] = useState({ feName: "", dspId: "", date: "" });
  const [isMounted, setIsMounted] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");
  
  const [replacerData, setReplacerData] = useState<any[]>([]);
  const [replacerMeta, setReplacerMeta] = useState<{headers: string[], remarkKey: string} | null>(null);

  useEffect(() => {
    setIsMounted(true);
    setSetupData(prev => ({ ...prev, date: new Date().toISOString().split('T')[0] }));
  }, []);

  const showToast = useCallback((msg: string, type: 'ok' | 'err' | 'info') => {
    if (typeof document === 'undefined') return;
    const toast = document.createElement('div');
    toast.className = `fixed bottom-10 left-1/2 -translate-x-1/2 px-6 py-3 rounded-xl text-[12px] font-bold z-[300] shadow-2xl transition-all duration-300 transform scale-95 opacity-0 animate-in fade-in slide-in-from-bottom-5 border ${
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

  const currentSession = useMemo(() => sessions.find(s => s.id === selectedSessionId) || null, [sessions, selectedSessionId]);

  const getVisibleRows = useCallback(() => {
    if (!currentSession) return [];
    let rows = currentSession.data;
    
    if (statusFilter === 'pending') rows = rows.filter(r => r.status === 'pending');
    else if (statusFilter === 'dispatched') rows = rows.filter(r => r.status === 'dispatched' || r.status === 'dispatch');
    else if (statusFilter === 'rto') rows = rows.filter(r => r.status === 'rto');
    else if (statusFilter === 'dto') rows = rows.filter(r => r.status === 'dto' || r.status === 'delivered');

    if (activeRemarkChip !== null) {
      rows = rows.filter(r => (r.remark || '') === activeRemarkChip);
    }
    
    if (searchTerm) {
      const s = searchTerm.toLowerCase();
      rows = rows.filter(r => r.awb.includes(s) || r.client.toLowerCase().includes(s) || r.orderId.toLowerCase().includes(s));
    }
    
    return rows;
  }, [currentSession, statusFilter, activeRemarkChip, searchTerm]);

  const filteredRows = useMemo(() => getVisibleRows(), [getVisibleRows]);

  const stats = useMemo(() => {
    if (!currentSession) return { total: 0, pending: 0, dispatched: 0, rto: 0, dto: 0, intact: 0 };
    return {
      total: currentSession.data.length,
      pending: currentSession.data.filter(r => r.status === 'pending').length,
      dispatched: currentSession.data.filter(r => r.status === 'dispatched' || r.status === 'dispatch').length,
      rto: currentSession.data.filter(r => r.status === 'rto').length,
      dto: currentSession.data.filter(r => r.status === 'dto' || r.status === 'delivered').length,
      intact: currentSession.data.filter(r => r.isIntact).length
    };
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
      setTimeout(() => {
        try {
          const bstr = evt.target?.result;
          const wb = XLSX.read(bstr, { type: 'binary' });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rawData = XLSX.utils.sheet_to_json(ws);
          
          if (!Array.isArray(rawData) || rawData.length === 0) throw new Error("File empty");

          const parsedRows: PODRow[] = rawData.map((row: any) => {
            const keys = Object.keys(row);
            const findVal = (regex: RegExp) => {
              const key = keys.find(k => regex.test(k.toLowerCase().replace(/[\s_-]/g, "")));
              return key ? row[key] : "";
            };
            const awb = fixAWB(findVal(/waybill|awb|awbnumber|waybillno/));
            const statusRaw = String(findVal(/status|currentstatus|current status/)).toLowerCase().trim();
            const status = STATUS_MAP[statusRaw] || "unknown";
            const remark = String(findVal(/remark|remarks|remark1/)).trim();
            const isIntact = /reject|intact|content|barcode/i.test(remark);

            return {
              id: crypto.randomUUID(),
              awb,
              client: String(findVal(/client|clientname/)),
              orderId: String(findVal(/order|orderid|orderno/)),
              status,
              remark,
              feName: setupData.feName,
              dspId: setupData.dspId,
              date: setupData.date,
              selected: false,
              isIntact
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
          }
        } catch (err) {
          showToast("Error parsing file!", "err");
        } finally {
          setIsProcessing(false);
        }
      }, 0);
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  const handleClearAllSessions = () => {
    setSessions([]);
    setSelectedSessionId(null);
    showToast("System Reset: All sessions cleared", "info");
  };

  const handleDeleteRow = useCallback((rowId: string) => {
    setSessions(prev => prev.map(s => {
      if (s.id === selectedSessionId) {
        return { ...s, data: s.data.filter(r => r.id !== rowId) };
      }
      return s;
    }));
    showToast("Row deleted", "info");
  }, [selectedSessionId, showToast]);

  const downloadExcel = () => {
    const rows = filteredRows;
    if (!rows.length || !currentSession) return;

    const header = ['Date', 'DSP ID', 'AWB Number', 'Client', 'Order ID', 'Remark', 'FE Name'];
    const excelData = rows.map((r, i) => [
      r.date, 
      i === 0 ? currentSession.dspId : "", 
      { v: String(r.awb), t: 's', z: '@' }, 
      r.client, 
      r.orderId, 
      r.remark, 
      r.feName
    ]);
    
    const ws = XLSX.utils.aoa_to_sheet([header, ...excelData]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "POD Report");
    
    const chipName = activeRemarkChip ? `_${activeRemarkChip.replace(/\s+/g, '')}` : '';
    const filename = `POD_${currentSession.dspId}_${currentSession.date}_${statusFilter.toUpperCase()}${chipName}.xlsx`;
    XLSX.writeFile(wb, filename);
    showToast(`Downloaded ${rows.length} rows — ${statusFilter.toUpperCase()}`, "ok");
  };

  const copyTable = () => {
    const rows = filteredRows;
    if (!rows.length || !currentSession) return;
    const text = rows.map((r, i) => `${r.date}\t${i === 0 ? currentSession.dspId : ""}\t${r.awb}\t${r.client}\t${r.orderId}\t${r.remark}\t${r.feName}`).join("\n");
    navigator.clipboard.writeText(text).then(() => showToast(`Copied ${rows.length} rows`, "ok"));
  };

  const handleReplacerUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    const reader = new FileReader();
    reader.onload = (evt) => {
      setTimeout(() => {
        try {
          const bstr = evt.target?.result;
          const wb = XLSX.read(bstr, { type: 'binary' });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rawData = XLSX.utils.sheet_to_json(ws, { defval: "" });
          
          if (!Array.isArray(rawData) || rawData.length === 0) throw new Error("Empty file");

          const headers = Object.keys(rawData[0]);
          const remarkKey = headers.find(h => /remark|nsl/i.test(h)) || "";
          const awbKey = headers.find(h => /awb|waybill/i.test(h)) || "";

          if (!remarkKey || !awbKey) throw new Error("Missing columns");

          const processed = rawData.map(row => {
            const original = String(row[remarkKey] || "").trim();
            let replaced = original;
            let isReplaced = false;

            for (const [key, val] of Object.entries(REMARK_MAPPING)) {
              if (original.toLowerCase() === key.toLowerCase() || original.toLowerCase().includes(key.toLowerCase())) {
                replaced = val;
                isReplaced = true;
                break;
              }
            }

            return { ...row, [remarkKey]: replaced, __isReplaced: isReplaced, [awbKey]: fixAWB(row[awbKey]) };
          });

          setReplacerData(processed);
          setReplacerMeta({ headers, remarkKey });
          showToast("File processed successfully!", "ok");
        } catch (err) {
          showToast("Error processing file!", "err");
        } finally {
          setIsProcessing(false);
        }
      }, 0);
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  const downloadReplacerExcel = () => {
    if (!replacerData.length || !replacerMeta) return;
    const cleanData = replacerData.map(r => {
      const { __isReplaced, ...rest } = r;
      const awbKey = replacerMeta.headers.find(h => /awb|waybill/i.test(h)) || "";
      if (awbKey) {
        rest[awbKey] = { v: String(rest[awbKey]), t: 's', z: '@' };
      }
      return rest;
    });
    const ws = XLSX.utils.json_to_sheet(cleanData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Replaced Remarks");
    XLSX.writeFile(wb, `EOD_Official_Remarks_${new Date().toISOString().split('T')[0]}.xlsx`);
    showToast("Official Excel Downloaded", "ok");
  };

  const groupedPendingRows = useMemo(() => {
    if (statusFilter !== 'pending' || !filteredRows.length) return null;
    const groups: Record<string, PODRow[]> = {};
    filteredRows.forEach(r => {
      const rem = r.remark || "No Remark";
      if (!groups[rem]) groups[rem] = [];
      groups[rem].push(r);
    });
    return Object.entries(groups).sort((a, b) => b[1].length - a[1].length);
  }, [statusFilter, filteredRows]);

  if (!isMounted) return null;

  return (
    <div className="min-h-screen bg-[#F0F4FA] font-body text-[#374151] select-auto">
      {/* Sticky Header Wrapper */}
      <div className="fixed top-0 left-0 w-full z-[200]">
        <div className="h-[3px] w-full bg-gradient-to-r from-[#1565C0] via-[#F9A825] via-[#2E7D32] to-[#D32F2F]" />
        
        <header className="h-[58px] bg-[#1C2333] px-6 flex items-center justify-between text-white shadow-lg">
          <div className="flex items-center gap-3">
            <div className="bg-gradient-to-br from-[#1565C0] to-blue-700 p-1.5 rounded-xl shadow-lg border border-white/10">
              <Truck className="w-5 h-5 text-white" />
            </div>
            <div className="flex flex-col">
              <h1 className="text-[15px] font-[800] tracking-tight leading-none uppercase">POD Management Tool</h1>
              <p className="text-[9px] text-[#94A3B8] font-bold mt-1 uppercase tracking-wider">
                Delhivery · Palam Vihar RPC · <span className="text-[#F9A825] font-black italic">By Ashu</span>
              </p>
            </div>
          </div>

          <div className="hidden lg:flex flex-col items-center">
            <span className="text-[9px] font-black text-[#6B8CAE] uppercase tracking-[0.25em] mb-0.5">FIELD OPERATIONS DASHBOARD</span>
            <span className="text-[10px] font-bold text-white/40 uppercase tracking-widest">Manage · Track · Export</span>
          </div>

          <div className="flex items-center gap-4">
            <div className="flex items-center gap-2 bg-[#2E7D32]/20 px-3 py-1 rounded-full border border-[#2E7D32]/30">
              <div className="w-1.5 h-1.5 bg-[#2E7D32] rounded-full animate-pulse shadow-[0_0_8px_#2E7D32]" />
              <span className="text-[9px] font-black text-[#6EE7A6] uppercase tracking-widest">Live</span>
            </div>
            {currentSession && (
              <div className="bg-[#F9A825]/20 px-3 py-1 rounded-full border border-[#F9A825]/30 text-[9px] font-mono font-black text-[#F9A825] uppercase">
                {currentSession.data.length} ROWS
              </div>
            )}
          </div>
        </header>

        <nav className="bg-[#1C2333] px-6 flex gap-8 border-t border-white/5 shadow-md">
          {[
            { id: "eod", label: "Daily EOD Rejection" },
            { id: "remark", label: "EOD Rejection Remark" }
          ].map(tab => (
            <button 
              key={tab.id}
              onClick={() => { setActiveTab(tab.id as any); setStatusFilter("all"); setActiveRemarkChip(null); }}
              className={cn(
                "py-3 text-[11px] font-black uppercase tracking-[0.08em] transition-all relative outline-none",
                activeTab === tab.id ? "text-white" : "text-[#94A3B8] hover:text-slate-300"
              )}
            >
              {tab.label}
              {activeTab === tab.id && <div className="absolute bottom-0 left-0 w-full h-[3px] bg-[#F9A825] rounded-t-full shadow-lg" />}
            </button>
          ))}
        </nav>
      </div>

      <main className="pt-[130px] p-6 max-w-[1360px] mx-auto space-y-6">
        {activeTab === "eod" ? (
          <div className="space-y-6">
            {/* Session Setup Section */}
            <div className="bg-white rounded-[14px] p-6 shadow-sm border border-[#E2E8F0]">
              <div className="text-[10px] font-black text-[#64748B] uppercase tracking-[0.15em] mb-4">SESSION SETUP</div>
              <div className="flex flex-col md:flex-row gap-4 mb-6">
                <div className="flex-1 space-y-1.5">
                  <label className="text-[10px] font-black text-[#64748B] uppercase tracking-[0.8px]">DSP ID</label>
                  <input type="number" value={setupData.dspId} onChange={(e) => setSetupData({...setupData, dspId: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 text-[13px] font-bold outline-none focus:border-primary transition-all" placeholder="Enter ID" />
                </div>
                <div className="flex-1 space-y-1.5">
                  <label className="text-[10px] font-black text-[#64748B] uppercase tracking-[0.8px]">FIELD EXECUTIVE (FE) NAME</label>
                  <input type="text" value={setupData.feName} onChange={(e) => setSetupData({...setupData, feName: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 text-[13px] font-bold outline-none focus:border-primary transition-all" placeholder="Enter Name" />
                </div>
                <div className="flex-[0.5] space-y-1.5">
                  <label className="text-[10px] font-black text-[#64748B] uppercase tracking-[0.8px]">DATE</label>
                  <input type="date" value={setupData.date} onChange={(e) => setSetupData({...setupData, date: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 text-[13px] font-bold" />
                </div>
              </div>
              
              <div className="border-2 border-dashed border-[#CBD5E1] rounded-2xl p-8 text-center bg-[#F8FAFC] hover:border-primary hover:bg-[#E3F2FD] transition-all cursor-pointer relative">
                <input type="file" disabled={isProcessing} onChange={handleFileUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                {isProcessing ? <Loader2 className="w-10 h-10 text-primary animate-spin mx-auto" /> : 
                  <div className="space-y-2">
                    <Download className="w-8 h-8 text-primary/40 mx-auto" />
                    <p className="text-[14px] font-[600] text-[#111827]">Drop Delhivery export file here, or click to upload</p>
                    <p className="text-[10px] text-[#64748B] uppercase tracking-widest">Data will be saved automatically for current session</p>
                  </div>
                }
              </div>
            </div>

            {/* Sessions Grid */}
            {sessions.length > 0 && (
              <div className="space-y-4">
                <div className="flex items-center justify-between">
                  <h2 className="text-[10px] font-black text-[#1C2333] uppercase tracking-[0.2em]">ALL FE SESSIONS</h2>
                  <button onClick={handleClearAllSessions} className="text-[10px] font-black text-[#D32F2F] uppercase tracking-widest flex items-center gap-2 hover:underline pr-2">
                    <Trash2 className="w-3.5 h-3.5" /> Clear All Sessions
                  </button>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-4 gap-4">
                  {sessions.map(s => (
                    <div 
                      key={s.id} 
                      onClick={() => setSelectedSessionId(s.id)}
                      className={cn(
                        "bg-white p-4 rounded-xl border-l-[3.5px] shadow-sm cursor-pointer relative transition-all",
                        selectedSessionId === s.id ? "border-l-[#1565C0] bg-[#F0F7FF] ring-1 ring-[#1565C0]/20" : "border-l-slate-300 border border-slate-200"
                      )}
                    >
                      <button onClick={(e) => { e.stopPropagation(); setSessions(prev => prev.filter(x => x.id !== s.id)); if(selectedSessionId === s.id) setSelectedSessionId(null); }} className="absolute top-3 right-3 text-slate-300 hover:text-red-500 transition-colors"><X className="w-4 h-4" /></button>
                      <p className="text-[15px] font-[800] text-[#1C2333] uppercase truncate pr-6">{s.feName}</p>
                      <p className="text-[10px] font-mono font-bold text-slate-400 mt-1">{s.dspId} — {s.date}</p>
                      <div className="flex flex-wrap gap-1.5 mt-3">
                        <span className="px-1.5 py-0.5 rounded bg-slate-100 text-[10px] font-black text-slate-600">{s.data.length} total</span>
                        <span className="px-1.5 py-0.5 rounded bg-amber-50 text-[10px] font-black text-[#F9A825]">{s.data.filter(r => r.status === 'pending').length} pending</span>
                        <span className="px-1.5 py-0.5 rounded bg-red-50 text-[10px] font-black text-[#D32F2F]">{s.data.filter(r => r.status === 'rto').length} RTO</span>
                        <span className="px-1.5 py-0.5 rounded bg-green-50 text-[10px] font-black text-[#2E7D32]">{s.data.filter(r => r.status === 'dto' || r.status === 'delivered').length} DTO</span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {currentSession && (
              <div className="space-y-6">
                <div>
                   <p className="text-[10px] font-black text-[#1565C0] uppercase tracking-[0.2em] mb-3 ml-1">
                     CURRENT: <span className="text-[#1C2333]">{currentSession.feName.toUpperCase()} — {currentSession.dspId}</span>
                   </p>
                   {/* Colourfull Status Tabs */}
                   <div className="bg-white rounded-[14px] shadow-sm border border-slate-200 overflow-hidden flex h-[90px]">
                     {[
                       { id: 'all', label: 'All', val: stats.total, color: 'text-[#1565C0]', active: 'bg-white border-[#1565C0]' },
                       { id: 'pending', label: 'Pending', val: stats.pending, color: 'text-[#F9A825]', active: 'bg-[#FFFDE7] border-[#F9A825]' },
                       { id: 'dispatched', label: 'Dispatch', val: stats.dispatched, color: 'text-[#1565C0]', active: 'bg-white border-[#1565C0]' },
                       { id: 'rto', label: 'RTO', val: stats.rto, color: 'text-[#D32F2F]', active: 'bg-[#FFF5F5] border-[#D32F2F]' },
                       { id: 'dto', label: 'DTO', val: stats.dto, color: 'text-[#2E7D32]', active: 'bg-[#F0FDF4] border-[#2E7D32]' }
                     ].map((t, i) => (
                       <button 
                         key={t.id} 
                         onClick={() => { setStatusFilter(t.id); setActiveRemarkChip(null); }}
                         className={cn(
                           "flex-1 flex flex-col items-center justify-center transition-all border-b-[4px] relative",
                           i !== 4 && "border-r border-slate-100",
                           statusFilter === t.id ? t.active : "border-transparent bg-white hover:bg-slate-50"
                         )}
                       >
                         <span className={cn("text-[26px] font-[900] leading-none mb-1", statusFilter === t.id ? t.color : 'text-slate-400')}>{t.val}</span>
                         <span className={cn("text-[9px] font-black uppercase tracking-[0.1em]", statusFilter === t.id ? t.color : 'text-slate-500')}>{t.label}</span>
                       </button>
                     ))}
                   </div>
                </div>

                {/* Remark Breakdown Logic */}
                {statusFilter === 'pending' && pendingRemarkStats.length > 0 && (
                  <div className="bg-white rounded-[14px] p-6 shadow-sm border border-slate-200 space-y-4">
                    <div>
                      <h3 className="text-[10px] font-black text-[#1C2333] uppercase tracking-wider">REMARK BREAKDOWN — PENDING</h3>
                      <p className="text-[9px] text-slate-400 font-bold uppercase mt-0.5">Click any remark chip to filter</p>
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
                              isActive 
                                ? "bg-[#1565C0] text-white border-[#1565C0] shadow-md" 
                                : isReject 
                                  ? "bg-[#FFF5F5] border-[#FFCDD2] text-[#D32F2F] hover:bg-red-100"
                                  : "bg-[#F3F4F6] border-[#E5E7EB] text-[#374151] hover:border-[#1565C0] hover:text-[#1565C0]"
                            )}
                          >
                            <span>{stat.text || "No Remark"}</span>
                            <span className={cn("px-1.5 rounded-full text-[9px]", isActive ? "bg-white/20 text-white" : "bg-white text-slate-500 shadow-sm")}>{stat.count}</span>
                          </button>
                        );
                      })}
                    </div>
                  </div>
                )}

                <div className="flex items-center justify-between">
                  <div className="flex gap-2">
                    <button onClick={downloadExcel} className="bg-[#388E3C] hover:bg-[#2E7D32] text-white px-5 py-2.5 rounded-lg font-[700] text-[12px] uppercase tracking-widest flex items-center gap-2 shadow-sm transition-all transform hover:-translate-y-0.5"><Download className="w-4 h-4" /> Download Excel</button>
                    <button onClick={copyTable} className="bg-[#1976D2] hover:bg-[#1565C0] text-white px-5 py-2.5 rounded-lg font-[700] text-[12px] uppercase tracking-widest flex items-center gap-2 shadow-sm transition-all transform hover:-translate-y-0.5"><Copy className="w-4 h-4" /> Copy Table</button>
                  </div>
                  <button onClick={() => { setSessions(prev => prev.map(s => s.id === selectedSessionId ? {...s, data: []} : s)); }} className="text-[11px] font-black text-slate-400 hover:text-red-500 uppercase tracking-widest px-4 outline-none">Clear Session</button>
                </div>

                <div className="bg-white rounded-[14px] shadow-sm border border-slate-200 overflow-hidden">
                  <div className="p-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
                    <div className="flex items-center gap-2">
                      <span className="text-[11px] font-black text-[#1C2333] uppercase tracking-widest">{statusFilter.toUpperCase()} SHIPMENTS</span>
                      <span className="px-2 py-0.5 bg-[#CBD5E1] rounded text-[9px] font-black text-[#475569] uppercase">{filteredRows.length} ROWS</span>
                    </div>
                    <div className="flex items-center gap-4">
                      <div className="relative">
                        <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                        <input type="text" placeholder="Search AWB..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="bg-white border border-slate-200 rounded-full pl-9 pr-4 py-1.5 text-[11px] font-bold outline-none focus:border-primary w-[220px] shadow-inner" />
                      </div>
                    </div>
                  </div>

                  <div className="overflow-x-auto max-h-[700px] custom-scrollbar">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="bg-[#1C2333] text-white">
                          <th className="p-4 w-[40px] text-center"><input type="checkbox" /></th>
                          <th className="p-4 w-[40px] text-center">X</th>
                          <th className="p-4 w-[80px] text-[9.5px] uppercase font-[700] tracking-widest">DSP ID</th>
                          <th className="p-4 w-[160px] text-[9.5px] uppercase font-[700] tracking-widest">AWB Number</th>
                          <th className="p-4 w-[140px] text-[9.5px] uppercase font-[700] tracking-widest">Client</th>
                          <th className="p-4 w-[140px] text-[9.5px] uppercase font-[700] tracking-widest">Order ID</th>
                          <th className="p-4 text-[9.5px] uppercase font-[700] tracking-widest">Remark</th>
                          <th className="p-4 w-[120px] text-[9.5px] uppercase font-[700] tracking-widest">FE Name</th>
                        </tr>
                      </thead>
                      <tbody>
                        {statusFilter === 'pending' && groupedPendingRows ? (
                          groupedPendingRows.map(([remark, rows]) => (
                            <React.Fragment key={remark}>
                              <tr className="bg-gradient-to-r from-[#0D1B2E] to-[#1A2F4A] border-y border-white/5">
                                <td colSpan={8} className="p-3 px-5">
                                  <div className="flex items-center gap-3">
                                    <Star className="w-3.5 h-3.5 text-[#F9A825] fill-current" />
                                    <span className="text-[10px] font-black text-white uppercase tracking-[0.1em]">{remark}</span>
                                    <span className="px-2 py-0.5 bg-[#F9A825] text-white rounded text-[9px] font-black uppercase shadow-sm">{rows.length} PKT</span>
                                  </div>
                                </td>
                              </tr>
                              {rows.map((row, idx) => <DataRow key={row.id} row={row} idx={idx} isFirstInGroup={idx === 0} onDelete={handleDeleteRow} />)}
                            </React.Fragment>
                          ))
                        ) : (
                          <>
                            {currentSession && (
                              <tr className="bg-gradient-to-r from-[#0D1B2E] to-[#1A2F4A] border-y border-white/5">
                                <td colSpan={8} className="p-3 px-5">
                                  <div className="flex items-center justify-between">
                                    <div className="flex items-center gap-3">
                                      <span className="text-[14px] font-mono font-black text-[#F9A825]">{currentSession.dspId}</span>
                                      <span className="px-2 py-0.5 bg-[#F9A825] text-white rounded text-[9px] font-black uppercase shadow-sm">{filteredRows.length} PKT</span>
                                    </div>
                                    <div className="flex items-center gap-4 text-[#6B8CAE] text-[10px] font-bold uppercase tracking-widest">
                                      <span>{currentSession.feName}</span>
                                      <span className="w-1.5 h-1.5 bg-[#6B8CAE] rounded-full opacity-30" />
                                      <span>{currentSession.date}</span>
                                    </div>
                                  </div>
                                </td>
                              </tr>
                            )}
                            {filteredRows.map((row, idx) => <DataRow key={row.id} row={row} idx={idx} isFirstInGroup={idx === 0} onDelete={handleDeleteRow} />)}
                          </>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}
          </div>
        ) : (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
            <div className="lg:col-span-7 space-y-6">
              <div className="bg-white rounded-[14px] p-6 shadow-sm border border-slate-200">
                <h2 className="text-[14px] font-[800] text-[#1C2333] flex items-center gap-2 mb-6"><FileSpreadsheet className="w-5 h-5 text-[#2E7D32]" /> Remark Replacer Dashboard</h2>
                <div className="border-2 border-dashed border-slate-200 rounded-2xl p-10 text-center bg-[#F8FAFC] hover:border-primary hover:bg-[#E3F2FD] transition-all cursor-pointer relative">
                  <input type="file" disabled={isProcessing} onChange={handleReplacerUpload} className="absolute inset-0 opacity-0 cursor-pointer z-20" />
                  {isProcessing ? <Loader2 className="w-10 h-10 text-green-600 animate-spin mx-auto" /> : 
                    <div className="space-y-3">
                      <Download className="w-10 h-10 text-primary/40 mx-auto" />
                      <p className="text-[15px] font-[700] text-[#111827]">Upload EOD Rejection File</p>
                      <p className="text-[11px] text-[#64748B]">Keep original structure, only replace remarks</p>
                    </div>
                  }
                </div>
              </div>

              {replacerData.length > 0 && (
                <div className="bg-white rounded-[14px] shadow-sm border border-slate-200 overflow-hidden">
                  <div className="p-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
                    <button onClick={downloadReplacerExcel} className="bg-[#2E7D32] hover:bg-[#1B5E20] text-white px-6 py-2.5 rounded-lg font-[700] text-[12px] uppercase tracking-widest flex items-center gap-2 shadow-sm transition-all transform hover:-translate-y-0.5"><Download className="w-4 h-4" /> Download Official Excel</button>
                    <button onClick={() => setReplacerData([])} className="text-[10px] font-black text-red-500 uppercase tracking-widest px-4">Clear Preview</button>
                  </div>
                  <div className="overflow-x-auto max-h-[600px] custom-scrollbar">
                    <table className="w-full text-left border-collapse text-[11px]">
                      <thead className="sticky top-0 z-30">
                        <tr className="bg-[#1C2333] text-white">
                          {replacerMeta?.headers.map((h, i) => <th key={i} className="p-4 font-black uppercase tracking-widest border-r border-white/5 whitespace-nowrap">{h}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {replacerData.map((row, idx) => (
                          <tr key={idx} className={cn("border-b border-slate-100", row.__isReplaced ? "bg-[#F0FDF4]" : "bg-[#FFFDE7]")}>
                            {replacerMeta?.headers.map((h, i) => (
                              <td key={i} className={cn("p-3 whitespace-nowrap", h === replacerMeta.remarkKey && row.__isReplaced ? "text-[#2E7D32] font-black" : h === replacerMeta.remarkKey ? "text-[#D97706] font-black" : "text-slate-600")}>
                                {row[h]}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>

            <div className="lg:col-span-5">
              <div className="bg-white rounded-[14px] shadow-sm border border-slate-200 overflow-hidden sticky top-[130px]">
                <div className="bg-[#2E7D32] p-4 text-white shadow-md">
                  <h3 className="text-[11px] font-[800] uppercase tracking-widest flex items-center gap-2">Built-in Remark Mapping</h3>
                  <p className="text-[9px] text-green-100 mt-0.5">NSL Remarks → Official Remarks</p>
                </div>
                <div className="p-4 overflow-y-auto max-h-[70vh] space-y-3 custom-scrollbar">
                  {Object.entries(REMARK_MAPPING).map(([nsl, official]) => (
                    <div key={nsl} className="p-3 bg-slate-50 rounded-xl border border-slate-100 space-y-2 hover:bg-white transition-colors">
                      <div className="flex flex-col gap-1">
                        <span className="text-[8px] font-black text-slate-400 uppercase">NSL Remark</span>
                        <span className="bg-[#FFFDE7] text-[#D97706] border border-[#FDE68A] px-2 py-1 rounded-md text-[10px] font-bold">{nsl}</span>
                      </div>
                      <div className="flex justify-center text-green-600"><CheckCircle2 className="w-4 h-4" /></div>
                      <div className="flex flex-col gap-1">
                        <span className="text-[8px] font-black text-slate-400 uppercase">Official Remark</span>
                        <span className="bg-[#F0FDF4] text-[#2E7D32] border border-green-200 px-2 py-1 rounded-md text-[10px] font-bold">{official}</span>
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
        .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #F1F5F9; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 10px; }
        .select-auto { user-select: auto !important; }
      `}</style>
    </div>
  );
}

function DataRow({ row, idx, isFirstInGroup, onDelete }: { row: PODRow, idx: number, isFirstInGroup: boolean, onDelete: (id: string) => void }) {
  return (
    <tr className={cn("border-b border-slate-100 transition-colors group", row.isIntact ? "bg-[#FFF5F5]" : "hover:bg-[#F0F7FF]")}>
      <td className="p-3 text-center"><input type="checkbox" className="accent-[#1565C0]" /></td>
      <td className="p-3 text-center">
        <button onClick={() => onDelete(row.id)} className="text-slate-300 hover:text-red-500 transition-colors"><Trash2 className="w-3.5 h-3.5 mx-auto" /></button>
      </td>
      <td className="p-3 font-mono text-[11.5px] font-bold text-slate-400">{isFirstInGroup ? row.dspId : ""}</td>
      <td onClick={() => navigator.clipboard.writeText(row.awb)} className="p-3 font-mono text-[11.5px] font-[600] text-[#1565C0] cursor-pointer hover:underline">{row.awb}</td>
      <td className="p-3 text-[11px] font-[500] text-[#374151] truncate max-w-[140px]">{row.client}</td>
      <td className="p-3 text-[11px] font-[500] text-[#374151] truncate max-w-[140px]">{row.orderId}</td>
      <td className="p-3">
        <span className={cn(
          "px-2 py-1 rounded-md text-[10px] font-black uppercase tracking-tight border shadow-sm",
          row.isIntact ? "bg-[#FFF5F5] text-[#D32F2F] border-[#FFCDD2]" : "bg-[#FFFDE7] text-[#D97706] border-[#FDE68A]"
        )}>
          {row.remark || "N/A"}
        </span>
      </td>
      <td className="p-3 text-[11px] font-bold text-slate-400">{row.feName}</td>
    </tr>
  );
}
