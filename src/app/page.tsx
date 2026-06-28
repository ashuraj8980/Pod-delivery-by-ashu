
"use client";

import React, { useState, useEffect, useMemo, useCallback } from "react";
import { 
  Truck, 
  Trash2, 
  Download, 
  Copy, 
  X, 
  FileSpreadsheet, 
  Loader2, 
  Search,
  Star,
  AlertCircle,
  ArrowRight,
  CheckCircle2,
  Info
} from "lucide-react";
import * as XLSX from "xlsx";
import { cn } from "@/lib/utils";

/**
 * @fileOverview Delhivery POD Management Tool - Palam Vihar RPC Edition
 * Optimized for Ashu. 
 * Features: Zero-Freeze Remark Replacer, Strict Column Preservation, Text-Format Excel Exports.
 * RESET ON RELOAD: No persistent storage.
 */

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

const STATUS_MAP: Record<string, string> = {
  "pending": "pending",
  "dispatched": "dispatched",
  "dispatch": "dispatched",
  "rto": "rto",
  "dto": "dto",
  "delivered": "dto"
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
  const [uploadError, setUploadError] = useState<string | null>(null);
  
  // Remark Replacer State
  const [replacerData, setReplacerData] = useState<any[]>([]);
  const [replacerMeta, setReplacerMeta] = useState<{headers: string[], remarkKey: string, awbKey: string, dspKey: string} | null>(null);
  const [replacerStats, setReplacerStats] = useState({ total: 0, replaced: 0, noMapping: 0 });

  useEffect(() => {
    setIsMounted(true);
    setSetupData(prev => ({ ...prev, date: new Date().toISOString().split('T')[0] }));
  }, []);

  const showToast = useCallback((msg: string, type: 'ok' | 'err' | 'info') => {
    if (typeof document === 'undefined') return;
    const toast = document.createElement('div');
    toast.className = `fixed bottom-10 left-1/2 -translate-x-1/2 px-6 py-3 rounded-xl text-[12px] font-bold z-[500] shadow-2xl transition-all duration-300 transform scale-95 opacity-0 animate-in fade-in slide-in-from-bottom-5 border ${
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

  const fixValueToString = (val: any) => {
    if (val === null || val === undefined) return "";
    let str = String(val).trim().replace(/['",]/g, ""); 
    // Handle scientific notation (e.g. 5.5E+13)
    if (/^[\d.]+[eE][+\-]?\d+$/.test(str)) {
      try {
        str = BigInt(Math.round(Number(val))).toString();
      } catch (e) {
        str = String(val);
      }
    }
    return str.replace(/\.0$/, "");
  };

  // MODULE 1: DAILY EOD REJECTION UPLOAD
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!setupData.feName || !setupData.dspId) {
      showToast("FE Name and DSP ID are required!", "err");
      e.target.value = "";
      return;
    }
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    setUploadError(null);
    const reader = new FileReader();
    
    reader.onload = (evt) => {
      setTimeout(() => {
        try {
          const data = new Uint8Array(evt.target?.result as ArrayBuffer);
          const wb = XLSX.read(data, { type: 'array' });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rawData = XLSX.utils.sheet_to_json(ws);
          
          if (!Array.isArray(rawData) || rawData.length === 0) throw new Error("Empty file content");

          const parsedRows: PODRow[] = rawData.map((row: any) => {
            const keys = Object.keys(row);
            const findVal = (regex: RegExp) => {
              const key = keys.find(k => regex.test(k.toLowerCase().replace(/[\s_-]/g, "")));
              return key ? row[key] : "";
            };
            const awb = fixValueToString(findVal(/waybill|awb|awbnumber|waybillno|awbno/));
            const statusRaw = String(findVal(/status|currentstatus|current status/)).toLowerCase().trim();
            const status = STATUS_MAP[statusRaw] || "unknown";
            const remark = String(findVal(/remark|remarks|remark1|nsl/)).trim();
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

          if (parsedRows.length === 0) throw new Error("No valid shipments found.");

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
        } catch (err: any) {
          setUploadError(err.message);
          showToast("Import failed!", "err");
        } finally {
          setIsProcessing(false);
        }
      }, 50);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  // MODULE 2: REMARK REPLACER (ULTRA-FIXED)
  const handleReplacerUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    setUploadError(null);
    
    const reader = new FileReader();
    reader.onload = (evt) => {
      // Background processing to prevent freeze
      setTimeout(() => {
        try {
          const data = new Uint8Array(evt.target?.result as ArrayBuffer);
          const wb = XLSX.read(data, { 
            type: 'array', 
            cellText: true, 
            cellDates: true 
          });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rawData = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
          
          if (!Array.isArray(rawData) || rawData.length === 0) {
            throw new Error("This file appears to be empty.");
          }

          const headers = Object.keys(rawData[0]);
          
          // Strict Requirement Check
          const remarkKey = headers.find(h => h === "Remarks Of NSL") || "";
          const awbKey = headers.find(h => h === "Awb") || headers.find(h => /awb|waybill/i.test(h)) || "";
          const dspKey = headers.find(h => h === "DSP No") || "";

          if (!remarkKey) {
            throw new Error("This file does not have Remarks Of NSL column. Please upload the correct Delhivery EOD rejection sheet.");
          }

          let replacedCount = 0;
          let noMappingCount = 0;

          const processed = rawData.map(row => {
            const originalRemark = String(row[remarkKey] || "").trim();
            let finalRemark = originalRemark;
            let isReplaced = false;

            // Mapping Logic: Exact Case-Insensitive first
            const mappingEntry = Object.entries(REMARK_MAPPING).find(([key]) => 
              key.toLowerCase() === originalRemark.toLowerCase()
            );

            if (mappingEntry) {
              finalRemark = mappingEntry[1];
              isReplaced = true;
            } else {
              // Partial Match
              const partialEntry = Object.entries(REMARK_MAPPING).find(([key]) => 
                originalRemark.toLowerCase().includes(key.toLowerCase())
              );
              if (partialEntry) {
                finalRemark = partialEntry[1];
                isReplaced = true;
              }
            }

            if (isReplaced) replacedCount++;
            else noMappingCount++;

            // Clean AWB and DSP
            const cleanAwb = fixValueToString(row[awbKey]);
            const cleanDsp = fixValueToString(row[dspKey]);
            
            return { 
              ...row, 
              [remarkKey]: finalRemark,
              [awbKey]: cleanAwb,
              [dspKey]: cleanDsp,
              __isReplaced: isReplaced, 
              __id: crypto.randomUUID()
            };
          });

          setReplacerData(processed);
          setReplacerMeta({ headers, remarkKey, awbKey, dspKey });
          setReplacerStats({ total: processed.length, replaced: replacedCount, noMapping: noMappingCount });
          showToast(`Processed ${processed.length} rows — ${replacedCount} replaced`, "ok");
        } catch (err: any) {
          setUploadError(err.message);
          showToast("Upload Failed", "err");
        } finally {
          setIsProcessing(false);
        }
      }, 0);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const currentSession = useMemo(() => sessions.find(s => s.id === selectedSessionId) || null, [sessions, selectedSessionId]);

  const filteredRows = useMemo(() => {
    if (!currentSession) return [];
    let rows = currentSession.data;
    if (statusFilter !== 'all') {
      if (statusFilter === 'pending') rows = rows.filter(r => r.status === 'pending');
      else if (statusFilter === 'dispatched') rows = rows.filter(r => r.status === 'dispatched' || r.status === 'dispatch');
      else if (statusFilter === 'rto') rows = rows.filter(r => r.status === 'rto');
      else if (statusFilter === 'dto') rows = rows.filter(r => r.status === 'dto' || r.status === 'delivered');
    }
    if (activeRemarkChip) rows = rows.filter(r => r.remark === activeRemarkChip);
    if (searchTerm) {
      const s = searchTerm.toLowerCase();
      rows = rows.filter(r => r.awb.includes(s) || r.client.toLowerCase().includes(s) || r.orderId.toLowerCase().includes(s));
    }
    return rows;
  }, [currentSession, statusFilter, activeRemarkChip, searchTerm]);

  const stats = useMemo(() => {
    if (!currentSession) return { total: 0, pending: 0, dispatched: 0, rto: 0, dto: 0 };
    return {
      total: currentSession.data.length,
      pending: currentSession.data.filter(r => r.status === 'pending').length,
      dispatched: currentSession.data.filter(r => r.status === 'dispatched' || r.status === 'dispatch').length,
      rto: currentSession.data.filter(r => r.status === 'rto').length,
      dto: currentSession.data.filter(r => r.status === 'dto' || r.status === 'delivered').length,
    };
  }, [currentSession]);

  const pendingRemarkStats = useMemo(() => {
    if (!currentSession || statusFilter !== 'pending') return [];
    const pending = currentSession.data.filter(r => r.status === 'pending');
    const counts: Record<string, number> = {};
    pending.forEach(r => { counts[r.remark] = (counts[r.remark] || 0) + 1; });
    return Object.entries(counts).map(([text, count]) => ({ text, count })).sort((a, b) => b.count - a.count);
  }, [currentSession, statusFilter]);

  const downloadExcel = () => {
    if (!filteredRows.length || !currentSession) return;
    const header = ['Date', 'DSP ID', 'AWB Number', 'Client', 'Order ID', 'Remark', 'FE Name'];
    const excelData = filteredRows.map((r, i) => [r.date, i === 0 ? currentSession.dspId : "", { v: String(r.awb), t: 's', z: '@' }, r.client, r.orderId, r.remark, r.feName]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...excelData]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, `POD_${currentSession.dspId}_${statusFilter}.xlsx`);
    showToast(`Downloaded ${filteredRows.length} rows — ${statusFilter}`, "ok");
  };

  const copyTable = () => {
    if (!filteredRows.length || !currentSession) return;
    const text = filteredRows.map((r, i) => `${r.date}\t${i === 0 ? currentSession.dspId : ""}\t${r.awb}\t${r.client}\t${r.orderId}\t${r.remark}\t${r.feName}`).join("\n");
    navigator.clipboard.writeText(text).then(() => showToast(`Copied ${filteredRows.length} rows`, "ok"));
  };

  const downloadOfficialExcel = () => {
    if (!replacerData.length || !replacerMeta) return;
    
    // Prepare data for export
    const exportData = replacerData.map(r => {
      const { __isReplaced, __id, ...rest } = r;
      // Force AWB and DSP as Text
      if (replacerMeta.awbKey) rest[replacerMeta.awbKey] = { v: String(rest[replacerMeta.awbKey]), t: 's', z: '@' };
      if (replacerMeta.dspKey) rest[replacerMeta.dspKey] = { v: String(rest[replacerMeta.dspKey]), t: 's', z: '@' };
      return rest;
    });

    const ws = XLSX.utils.json_to_sheet(exportData);
    
    // Header Styling (Simulation through AOE)
    const range = XLSX.utils.decode_range(ws['!ref']!);
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const address = XLSX.utils.encode_col(C) + "1";
      if (!ws[address]) continue;
      ws[address].s = {
        fill: { fgColor: { rgb: "1C2333" } },
        font: { color: { rgb: "FFFFFF" }, bold: true, sz: 11, name: "Calibri" },
        alignment: { horizontal: "center" }
      };
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Remarks");
    const dateStr = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `EOD_Official_Remarks_${dateStr}.xlsx`);
    showToast("Official Excel Downloaded", "ok");
  };

  const groupedPendingRows = useMemo(() => {
    if (statusFilter !== 'pending') return null;
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
    <div className="min-h-screen bg-[#F0F4FA] font-body text-[#374151]">
      {/* GLOBAL HEADER (STICKY) */}
      <div className="fixed top-0 left-0 w-full z-[400] shadow-lg">
        <div className="h-[3px] w-full bg-gradient-to-r from-[#1565C0] via-[#F9A825] via-[#2E7D32] to-[#D32F2F]" />
        <header className="h-[58px] bg-[#1C2333] px-6 flex items-center justify-between text-white border-b border-white/5">
          <div className="flex items-center gap-3">
            <div className="bg-gradient-to-br from-[#1565C0] to-blue-700 p-1.5 rounded-xl border border-white/10 shadow-lg"><Truck className="w-5 h-5 text-white" /></div>
            <div className="flex flex-col">
              <h1 className="text-[15px] font-[800] tracking-tight uppercase">POD Management Tool</h1>
              <p className="text-[9px] text-[#94A3B8] font-bold mt-1 uppercase tracking-wider">Delhivery · Palam Vihar RPC · <span className="text-[#F9A825] font-black italic">By Ashu</span></p>
            </div>
          </div>
          <div className="hidden lg:flex flex-col items-center">
            <span className="text-[9px] font-black text-[#6B8CAE] uppercase tracking-[0.25em] mb-0.5">FIELD OPERATIONS DASHBOARD</span>
            <span className="text-[10px] font-bold text-white/40 uppercase tracking-widest">Manage · Track · Export</span>
          </div>
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-2 bg-[#2E7D32]/20 px-3 py-1 rounded-full border border-[#2E7D32]/30">
              <div className="w-1.5 h-1.5 bg-[#2E7D32] rounded-full animate-pulse shadow-[0_0_8px_#2E7D32]" />
              <span className="text-[9px] font-black text-[#6EE7A6] uppercase">Live</span>
            </div>
            {currentSession && <div className="bg-[#F9A825]/20 px-3 py-1 rounded-full text-[9px] font-mono font-black text-[#F9A825] uppercase tracking-wider">{currentSession.data.length} ROWS</div>}
          </div>
        </header>
        <nav className="bg-[#1C2333] px-6 flex gap-8">
          {[ { id: "eod", label: "Daily EOD Rejection" }, { id: "remark", label: "EOD Rejection Remark" } ].map(tab => (
            <button key={tab.id} onClick={() => { setActiveTab(tab.id as any); setStatusFilter("all"); }} className={cn("py-3 text-[11px] font-black uppercase tracking-widest transition-all relative", activeTab === tab.id ? "text-white" : "text-[#94A3B8] hover:text-white")}>
              {tab.label}
              {activeTab === tab.id && <div className="absolute bottom-0 left-0 w-full h-[3px] bg-[#F9A825] rounded-t-full shadow-lg" />}
            </button>
          ))}
        </nav>
      </div>

      <main className="pt-[130px] p-6 max-w-[1440px] mx-auto space-y-6">
        {activeTab === "eod" ? (
          <div className="space-y-6">
            {/* SESSION SETUP CARD */}
            <div className="bg-white rounded-[14px] p-6 shadow-sm border border-[#E2E8F0]">
              <div className="text-[10px] font-black text-[#64748B] uppercase tracking-[0.15em] mb-4">SESSION SETUP</div>
              <div className="flex flex-col md:flex-row gap-4 mb-6">
                <input type="number" value={setupData.dspId} onChange={(e) => setSetupData({...setupData, dspId: e.target.value})} className="flex-1 bg-slate-50 border rounded-xl px-4 py-2.5 text-[13px] font-bold outline-none ring-offset-2 focus:ring-2 focus:ring-primary/20" placeholder="DSP ID" />
                <input type="text" value={setupData.feName} onChange={(e) => setSetupData({...setupData, feName: e.target.value})} className="flex-1 bg-slate-50 border rounded-xl px-4 py-2.5 text-[13px] font-bold outline-none ring-offset-2 focus:ring-2 focus:ring-primary/20" placeholder="FE Name" />
                <input type="date" value={setupData.date} onChange={(e) => setSetupData({...setupData, date: e.target.value})} className="flex-[0.5] bg-slate-50 border rounded-xl px-4 py-2.5 text-[13px] font-bold outline-none" />
              </div>
              
              <div className={cn("border-2 border-dashed rounded-2xl p-10 text-center transition-all cursor-pointer relative", uploadError ? "border-red-400 bg-red-50" : "border-[#CBD5E1] bg-[#F8FAFC] hover:border-primary hover:bg-slate-50", isProcessing && "opacity-80 pointer-events-none")}>
                <input type="file" disabled={isProcessing} onChange={handleFileUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                {isProcessing ? (
                  <div className="flex flex-col items-center gap-3"><Loader2 className="w-10 h-10 text-primary animate-spin" /><p className="text-[11px] font-black uppercase text-primary tracking-widest">Processing Session...</p></div>
                ) : uploadError ? (
                  <div className="space-y-3"><AlertCircle className="w-10 h-10 text-red-500 mx-auto" /><p className="text-[14px] font-[700] text-red-600 uppercase">Invalid Session File</p><p className="text-[10px] text-red-400 font-bold uppercase">{uploadError}</p><button onClick={() => setUploadError(null)} className="bg-red-600 text-white px-4 py-1.5 rounded-lg text-[10px] font-black uppercase">Try Again</button></div>
                ) : (
                  <div className="space-y-2"><Download className="w-10 h-10 text-primary/40 mx-auto" /><p className="text-[14px] font-[700] text-[#111827]">Drop Delhivery export file here, or click to upload</p><p className="text-[10px] text-[#64748B] uppercase tracking-widest">Fresh Start Enabled: Reset on Reload</p></div>
                )}
              </div>
            </div>

            {/* SAVED SESSIONS GRID */}
            {sessions.length > 0 && (
              <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-4 gap-4">
                {sessions.map(s => (
                  <div key={s.id} onClick={() => setSelectedSessionId(s.id)} className={cn("bg-white p-4 rounded-xl border-l-[3px] shadow-sm cursor-pointer relative transition-all group", selectedSessionId === s.id ? "border-l-[#1565C0] bg-[#F0F7FF] ring-1 ring-[#1565C0]/20" : "border-l-slate-300 border border-slate-200")}>
                    <button onClick={(e) => { e.stopPropagation(); setSessions(prev => prev.filter(x => x.id !== s.id)); if(selectedSessionId === s.id) setSelectedSessionId(null); }} className="absolute top-3 right-3 text-slate-300 hover:text-red-500 transition-colors opacity-0 group-hover:opacity-100"><X className="w-4 h-4" /></button>
                    <p className="text-[15px] font-[800] text-[#1C2333] uppercase truncate pr-6">{s.feName}</p>
                    <p className="text-[10px] font-mono font-bold text-slate-400 mt-1">{s.dspId} — {s.date}</p>
                    <div className="flex flex-wrap gap-1.5 mt-3">
                      <span className="px-1.5 py-0.5 rounded bg-slate-100 text-[9px] font-black text-slate-600 uppercase">{s.data.length} TOTAL</span>
                      <span className="px-1.5 py-0.5 rounded bg-amber-50 text-[9px] font-black text-[#F9A825] uppercase">{s.data.filter(r => r.status === 'pending').length} PENDING</span>
                      <span className="px-1.5 py-0.5 rounded bg-red-50 text-[9px] font-black text-[#D32F2F] uppercase">{s.data.filter(r => r.status === 'rto').length} RTO</span>
                      <span className="px-1.5 py-0.5 rounded bg-green-50 text-[9px] font-black text-[#2E7D32] uppercase">{s.data.filter(r => r.status === 'dto').length} DTO</span>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {/* DATA VIEW AREA */}
            {currentSession && (
              <div className="space-y-6">
                <div className="text-[10px] font-black text-[#64748B] uppercase tracking-widest pl-1">CURRENT: {currentSession.feName} — {currentSession.dspId}</div>
                
                {/* COLOURFUL STATUS TABS */}
                <div className="bg-white rounded-[14px] shadow-sm border border-slate-200 overflow-hidden flex h-[100px]">
                  {[
                    { id: 'all', label: 'All', val: stats.total, color: 'text-[#1565C0]', activeBg: 'bg-white', activeBorder: 'border-[#1565C0]' },
                    { id: 'pending', label: 'Pending', val: stats.pending, color: 'text-[#F9A825]', activeBg: 'bg-[#FFFDE7]', activeBorder: 'border-[#F9A825]' },
                    { id: 'dispatched', label: 'Dispatch', val: stats.dispatched, color: 'text-[#1565C0]', activeBg: 'bg-white', activeBorder: 'border-[#1565C0]' },
                    { id: 'rto', label: 'RTO', val: stats.rto, color: 'text-[#D32F2F]', activeBg: 'bg-[#FFF5F5]', activeBorder: 'border-[#D32F2F]' },
                    { id: 'dto', label: 'DTO', val: stats.dto, color: 'text-[#2E7D32]', activeBg: 'bg-[#F0FDF4]', activeBorder: 'border-[#2E7D32]' }
                  ].map((t, i) => (
                    <button key={t.id} onClick={() => { setStatusFilter(t.id); setActiveRemarkChip(null); }} className={cn("flex-1 flex flex-col items-center justify-center transition-all border-b-[4px] border-r relative", i === 4 && "border-r-0", statusFilter === t.id ? `${t.activeBg} ${t.activeBorder}` : "bg-white")}>
                      <span className={cn("text-[26px] font-[900] leading-none mb-1", t.color)}>{t.val}</span>
                      <span className={cn("text-[9px] font-black uppercase tracking-widest", statusFilter === t.id ? t.color : "text-slate-500")}>{t.label}</span>
                    </button>
                  ))}
                </div>

                {/* PENDING REMARK BREAKDOWN */}
                {statusFilter === 'pending' && pendingRemarkStats.length > 0 && (
                  <div className="bg-white rounded-[14px] p-6 shadow-sm border border-slate-200 space-y-4 animate-in slide-in-from-top-2">
                    <div className="flex flex-col gap-1">
                      <h3 className="text-[11px] font-black text-[#1C2333] uppercase tracking-wider">REMARK BREAKDOWN — PENDING</h3>
                      <p className="text-[9px] font-bold text-slate-400 uppercase">Click any remark chip to filter the table</p>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      {pendingRemarkStats.map((stat, i) => (
                        <button key={i} onClick={() => setActiveRemarkChip(activeRemarkChip === stat.text ? null : stat.text)} className={cn("flex items-center gap-2 px-3 py-1.5 rounded-lg border text-[10px] font-bold uppercase transition-all shadow-sm", activeRemarkChip === stat.text ? "bg-[#1565C0] text-white border-[#1565C0]" : "bg-[#F3F4F6] border-[#E5E7EB] text-[#374151] hover:bg-slate-200 hover:border-primary")}>
                          <span>{stat.text || "No Remark"}</span>
                          <span className={cn("px-1.5 py-0.5 rounded-full text-[9px] font-black", activeRemarkChip === stat.text ? "bg-white text-primary" : "bg-white text-slate-500")}>{stat.count}</span>
                        </button>
                      ))}
                    </div>
                  </div>
                )}

                {/* ACTION BAR */}
                <div className="flex items-center justify-between">
                  <div className="flex gap-2">
                    <button onClick={downloadExcel} className="bg-[#388E3C] hover:bg-[#2E7D32] text-white px-5 py-2.5 rounded-lg font-[700] text-[12px] uppercase tracking-widest flex items-center gap-2 shadow-sm transition-transform active:scale-95"><Download className="w-4 h-4" /> Download Excel</button>
                    <button onClick={copyTable} className="bg-[#1976D2] hover:bg-[#1565C0] text-white px-5 py-2.5 rounded-lg font-[700] text-[12px] uppercase tracking-widest flex items-center gap-2 shadow-sm transition-transform active:scale-95"><Copy className="w-4 h-4" /> Copy Table</button>
                  </div>
                  <button onClick={() => setSessions([])} className="bg-red-500 hover:bg-red-600 text-white px-4 py-2 rounded-lg text-[10px] font-black uppercase tracking-widest transition-all">Clear All Sessions</button>
                </div>

                {/* MAIN SHIPMENTS TABLE */}
                <div className="bg-white rounded-[14px] shadow-sm border border-slate-200 overflow-hidden">
                  <div className="p-4 border-b bg-slate-50 flex items-center justify-between">
                    <div className="flex items-center gap-2"><span className="text-[11px] font-black uppercase text-[#1C2333]">{statusFilter.toUpperCase()} SHIPMENTS</span><span className="px-2 py-0.5 bg-[#CBD5E1] rounded text-[9px] font-black uppercase">{filteredRows.length} ROWS</span></div>
                    <div className="flex items-center gap-4">
                      <div className="relative"><Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" /><input type="text" placeholder="Search AWB..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="bg-white border rounded-full pl-9 pr-4 py-1.5 text-[11px] font-bold outline-none w-[240px] focus:ring-2 focus:ring-primary/10" /></div>
                      <input type="checkbox" className="w-4 h-4 accent-primary" />
                    </div>
                  </div>
                  <div className="overflow-x-auto max-h-[700px] custom-scrollbar">
                    <table className="w-full text-left border-collapse">
                      <thead className="sticky top-0 z-30">
                        <tr className="bg-[#1C2333] text-white">
                          <th className="p-4 w-[50px] text-center"></th>
                          <th className="p-4 w-[50px] text-center">X</th>
                          <th className="p-4 w-[100px] text-[9.5px] uppercase font-[700]">DSP ID</th>
                          <th className="p-4 w-[180px] text-[9.5px] uppercase font-[700]">AWB Number</th>
                          <th className="p-4 w-[160px] text-[9.5px] uppercase font-[700]">Client</th>
                          <th className="p-4 w-[160px] text-[9.5px] uppercase font-[700]">Order ID</th>
                          <th className="p-4 text-[9.5px] uppercase font-[700]">Remark</th>
                          <th className="p-4 w-[140px] text-[9.5px] uppercase font-[700]">FE Name</th>
                        </tr>
                      </thead>
                      <tbody>
                        {statusFilter === 'pending' && groupedPendingRows ? (
                          groupedPendingRows.map(([remark, rows]) => (
                            <React.Fragment key={remark}>
                              <tr className="bg-gradient-to-r from-[#0D1B2E] to-[#1A2F4A] border-y border-white/5">
                                <td colSpan={8} className="p-3 px-5"><div className="flex items-center gap-3"><Star className="w-3.5 h-3.5 text-[#F9A825] fill-current" /><span className="text-[10px] font-black text-white uppercase tracking-[0.1em]">{remark}</span><span className="px-2 py-0.5 bg-[#F9A825] text-white rounded text-[9px] font-black uppercase tracking-widest">{rows.length} PKT</span></div></td>
                              </tr>
                              {rows.map((row, idx) => <DataRow key={row.id} row={row} idx={idx} isFirstInGroup={idx === 0} onDelete={(id) => setSessions(prev => prev.map(s => s.id === selectedSessionId ? {...s, data: s.data.filter(r => r.id !== id)} : s))} />)}
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
                                      <span className="px-2 py-0.5 bg-[#F9A825] text-white rounded text-[9px] font-black uppercase tracking-widest">{filteredRows.length} PKT</span>
                                    </div>
                                    <div className="text-[#6B8CAE] text-[10px] font-bold uppercase tracking-[0.2em]">{currentSession.feName} · {currentSession.date}</div>
                                  </div>
                                </td>
                              </tr>
                            )}
                            {filteredRows.map((row, idx) => <DataRow key={row.id} row={row} idx={idx} isFirstInGroup={idx === 0} onDelete={(id) => setSessions(prev => prev.map(s => s.id === selectedSessionId ? {...s, data: s.data.filter(r => r.id !== id)} : s))} />)}
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
          /* MODULE 2: REMARK REPLACER (ULTRA-FIXED) */
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 animate-in fade-in duration-500">
            <div className="lg:col-span-8 space-y-6">
              <div className="bg-white rounded-[14px] p-6 shadow-sm border border-slate-200">
                <div className="flex items-center justify-between mb-6">
                  <h2 className="text-[14px] font-[800] text-[#1C2333] flex items-center gap-2"><FileSpreadsheet className="w-5 h-5 text-[#2E7D32]" /> Remark Replacer Dashboard</h2>
                  {replacerData.length > 0 && (
                    <div className="flex items-center gap-3">
                      <div className="flex items-center gap-2 px-3 py-1 bg-slate-100 rounded-full border border-slate-200"><span className="text-[10px] font-black text-slate-500 uppercase">Total:</span><span className="text-[11px] font-bold">{replacerStats.total}</span></div>
                      <div className="flex items-center gap-2 px-3 py-1 bg-green-50 rounded-full border border-green-100"><span className="text-[10px] font-black text-green-600 uppercase">Replaced:</span><span className="text-[11px] font-bold text-green-700">{replacerStats.replaced}</span></div>
                      <div className="flex items-center gap-2 px-3 py-1 bg-amber-50 rounded-full border border-amber-100"><span className="text-[10px] font-black text-amber-600 uppercase">Original:</span><span className="text-[11px] font-bold text-amber-700">{replacerStats.noMapping}</span></div>
                    </div>
                  )}
                </div>

                <div className={cn("border-2 border-dashed rounded-2xl p-10 text-center transition-all cursor-pointer relative", uploadError ? "border-red-400 bg-red-50" : "border-slate-200 bg-[#F8FAFC] hover:border-primary", isProcessing && "opacity-80 pointer-events-none")}>
                  <input type="file" disabled={isProcessing} onChange={handleReplacerUpload} className="absolute inset-0 opacity-0 cursor-pointer z-20" />
                  {isProcessing ? (
                    <div className="flex flex-col items-center gap-3"><Loader2 className="w-10 h-10 text-primary animate-spin" /><p className="text-[11px] font-black uppercase text-primary tracking-widest">Processing file...</p></div>
                  ) : uploadError ? (
                    <div className="space-y-3">
                      <AlertCircle className="w-10 h-10 text-red-500 mx-auto" />
                      <p className="text-[13px] font-[700] text-red-600 uppercase max-w-[400px] mx-auto leading-relaxed">{uploadError}</p>
                      <button onClick={() => setUploadError(null)} className="bg-red-600 text-white px-6 py-2 rounded-lg text-[10px] font-black uppercase mt-4">Try Again</button>
                    </div>
                  ) : (
                    <div className="space-y-3"><Download className="w-10 h-10 text-primary/40 mx-auto" /><p className="text-[15px] font-[700] text-[#111827]">Drop Official Delhivery EOD Sheet here, or click to upload</p><p className="text-[11px] text-[#64748B] uppercase tracking-widest">Supports strict column mapping for Ashu</p></div>
                  )}
                </div>
              </div>

              {replacerData.length > 0 && (
                <div className="bg-white rounded-[14px] shadow-sm border border-slate-200 overflow-hidden">
                  <div className="p-4 bg-slate-50 border-b flex items-center justify-between">
                    <div className="flex gap-2">
                      <button onClick={downloadOfficialExcel} className="bg-[#1C2333] hover:bg-slate-800 text-white px-6 py-2.5 rounded-lg font-[700] text-[12px] uppercase tracking-widest flex items-center gap-2 shadow-sm transition-all"><Download className="w-4 h-4" /> Download Official Excel</button>
                      <button onClick={() => {
                        const headers = replacerMeta?.headers.join('\t') || "";
                        const rows = replacerData.map(r => {
                          const { __isReplaced, __id, ...rest } = r;
                          return Object.values(rest).join('\t');
                        }).join('\n');
                        navigator.clipboard.writeText(headers + '\n' + rows).then(() => showToast("Copied full table", "ok"));
                      }} className="bg-[#1565C0] hover:bg-blue-700 text-white px-6 py-2.5 rounded-lg font-[700] text-[12px] uppercase tracking-widest flex items-center gap-2 shadow-sm transition-all"><Copy className="w-4 h-4" /> Copy Full Table</button>
                    </div>
                    <button onClick={() => { setReplacerData([]); setReplacerMeta(null); }} className="text-[10px] font-black text-red-500 uppercase tracking-widest px-4 hover:underline">Clear Data</button>
                  </div>
                  <div className="overflow-x-auto max-h-[600px] custom-scrollbar">
                    <table className="w-full text-left border-collapse text-[11px]">
                      <thead className="sticky top-0 z-30">
                        <tr className="bg-[#1C2333] text-white">
                          <th className="p-4 w-[40px] text-center">X</th>
                          {replacerMeta?.headers.map((h, i) => <th key={i} className="p-4 font-black uppercase border-r border-white/5 whitespace-nowrap">{h}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {replacerData.map((row, idx) => (
                          <tr key={row.__id} className={cn("border-b transition-colors", row.__isReplaced ? "bg-[#F0FDF4]" : "bg-[#FFFDE7]")}>
                            <td className="p-3 text-center"><button onClick={() => setReplacerData(prev => prev.filter(r => r.__id !== row.__id))} className="text-slate-300 hover:text-red-500 transition-colors"><Trash2 className="w-3.5 h-3.5 mx-auto" /></button></td>
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

            <div className="lg:col-span-4">
              <div className="bg-white rounded-[14px] shadow-sm border border-slate-200 overflow-hidden sticky top-[130px]">
                <div className="bg-[#2E7D32] p-4 text-white shadow-md">
                  <h3 className="text-[11px] font-[800] uppercase tracking-widest flex items-center gap-2">REMARK MAPPING LIST</h3>
                  <p className="text-[9px] text-green-100 mt-0.5">NSL Remark → Official Replacement</p>
                </div>
                <div className="p-4 overflow-y-auto max-h-[70vh] space-y-3 custom-scrollbar">
                  {Object.entries(REMARK_MAPPING).map(([nsl, official]) => (
                    <div key={nsl} className="p-4 bg-slate-50 rounded-2xl border border-slate-100 space-y-3 hover:bg-white shadow-sm transition-colors group">
                      <div className="flex flex-col gap-1"><span className="text-[8px] font-black text-slate-400 uppercase">NSL REMARK</span><span className="bg-[#FFFDE7] text-[#D97706] border border-[#FDE68A] px-3 py-1.5 rounded-xl text-[10px] font-bold">{nsl}</span></div>
                      <div className="flex justify-center"><ArrowRight className="w-3.5 h-3.5 text-green-600 group-hover:scale-125 transition-transform" /></div>
                      <div className="flex flex-col gap-1"><span className="text-[8px] font-black text-slate-400 uppercase">OFFICIAL REMARK</span><span className="bg-[#F0FDF4] text-[#2E7D32] border border-green-200 px-3 py-1.5 rounded-xl text-[10px] font-bold">{official}</span></div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        )}
      </main>

      <style jsx global>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;700;900&family=IBM+Plex+Mono:wght@600&display=swap');
        body { font-family: 'Inter', sans-serif; }
        .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #F1F5F9; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 10px; }
        .font-mono { font-family: 'IBM+Plex+Mono', monospace; }
      `}</style>
    </div>
  );
}

function DataRow({ row, idx, isFirstInGroup, onDelete }: { row: PODRow, idx: number, isFirstInGroup: boolean, onDelete: (id: string) => void }) {
  return (
    <tr className={cn("border-b transition-colors group", row.isIntact ? "bg-[#FFF5F5]" : "hover:bg-[#F0F7FF]")}>
      <td className="p-3 text-center"><input type="checkbox" className="w-4 h-4 accent-[#1565C0]" /></td>
      <td className="p-3 text-center"><button onClick={() => onDelete(row.id)} className="text-slate-300 hover:text-red-500 transition-colors"><Trash2 className="w-3.5 h-3.5 mx-auto" /></button></td>
      <td className="p-3 font-mono text-[11px] font-bold text-slate-400">{isFirstInGroup ? row.dspId : ""}</td>
      <td onClick={() => {
        navigator.clipboard.writeText(row.awb);
        // Simple visual feedback could be added here
      }} className="p-3 font-mono text-[11.5px] font-[600] text-[#1565C0] cursor-pointer hover:underline transition-all tracking-tight">{row.awb}</td>
      <td className="p-3 text-[11px] text-[#374151] truncate max-w-[140px] font-medium">{row.client}</td>
      <td className="p-3 text-[11px] text-[#374151] truncate max-w-[140px] font-medium">{row.orderId}</td>
      <td className="p-3">
        <span className={cn("px-2.5 py-1.5 rounded-xl text-[10px] font-black uppercase tracking-tight border shadow-sm inline-block", row.isIntact ? "bg-[#FFF5F5] text-[#D32F2F] border-[#FFCDD2]" : "bg-[#FFFDE7] text-[#D97706] border-[#FDE68A]")}>
          {row.remark || "N/A"}
        </span>
      </td>
      <td className="p-3 text-[11px] font-bold text-slate-400 uppercase">{row.feName}</td>
    </tr>
  );
}
