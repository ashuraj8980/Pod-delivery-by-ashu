
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
  AlertCircle,
  ArrowRight,
  Info,
  BarChart3,
  Settings,
  Database,
  Calendar,
  Layers
} from "lucide-react";
import * as XLSX from "xlsx";
import { cn } from "@/lib/utils";

/**
 * @fileOverview Delhivery POD Management Tool - Palam Vihar RPC Edition
 * Professional Enterprise Dashboard with 100% AWB Precision for WPS Office.
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

/**
 * Global Date Formatter: YYYY-MM-DD -> DD-MM-YYYY
 */
const formatDate = (val: any): string => {
  if (!val) return "";
  const str = String(val).trim();
  if (/^\d{2}-\d{2}-\d{4}$/.test(str)) return str;
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    const [y, m, d] = str.split('-');
    return `${d}-${m}-${y}`;
  }
  const d = new Date(str);
  if (!isNaN(d.getTime())) {
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    return `${day}-${month}-${year}`;
  }
  return str;
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
  
  const [replacerData, setReplacerData] = useState<any[]>([]);
  const [replacerMeta, setReplacerMeta] = useState<{headers: string[], remarkKey: string, awbKey: string} | null>(null);
  const [replacerStats, setReplacerStats] = useState({ total: 0, replaced: 0, noMapping: 0 });

  // Persistence: Load on Mount
  useEffect(() => {
    setIsMounted(true);
    setSetupData(prev => ({ ...prev, date: new Date().toISOString().split('T')[0] }));
    const saved = localStorage.getItem('pod_sessions_rpc_hd_v2');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) setSessions(parsed);
      } catch (e) {
        console.error("Failed to load sessions", e);
      }
    }
  }, []);

  // Persistence: Save on Change
  useEffect(() => {
    if (isMounted) {
      localStorage.setItem('pod_sessions_rpc_hd_v2', JSON.stringify(sessions));
    }
  }, [sessions, isMounted]);

  const showToast = useCallback((msg: string, type: 'ok' | 'err' | 'info') => {
    if (typeof document === 'undefined') return;
    const toast = document.createElement('div');
    toast.className = `fixed bottom-8 left-1/2 -translate-x-1/2 px-6 py-3 rounded-2xl text-[13px] font-bold z-[500] shadow-2xl transition-all duration-300 transform scale-95 opacity-0 animate-in fade-in slide-in-from-bottom-4 border ${
      type === 'ok' ? 'bg-[#0F172A] text-emerald-400 border-emerald-500/20' : 
      type === 'err' ? 'bg-[#0F172A] text-rose-400 border-rose-500/20' : 
      'bg-[#0F172A] text-blue-400 border-blue-500/20'
    }`;
    toast.style.opacity = '1';
    toast.style.transform = 'translateX(-50%) scale(1)';
    toast.innerHTML = `<span class="flex items-center gap-2">${msg}</span>`;
    document.body.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateX(-50%) scale(0.95)';
      setTimeout(() => toast.remove(), 300);
    }, 3000);
  }, []);

  const fixValueToString = (val: any) => {
    if (val === null || val === undefined) return "";
    if (typeof val === 'number') {
      return val.toFixed(0);
    }
    return String(val).trim();
  };

  /**
   * Professional Copy Function with Dual MIME Types (Text + HTML)
   * This forces WPS Office/Excel to treat digits as Text via mso-number-format.
   */
  const copyDataProfessional = useCallback(async (rows: any[], headers: string[], isSingle = false) => {
    if (!rows.length) return;

    // Plain Text Version: Prefix AWB with ' to trick basic spreadsheet interpreters
    const plainText = rows.map(r => 
      headers.map(h => {
        const val = String(r[h] || "").trim().replace(/[\r\n\t]+/g, " ");
        if (h.toLowerCase().includes('awb')) return `'${val}`;
        return val;
      }).join("\t")
    ).join("\n");

    // HTML Table Version: Use Microsoft-specific CSS to force TEXT formatting
    const rowsHtml = rows.map(r => {
      const cells = headers.map(h => {
        const val = String(r[h] || "").trim().replace(/[\r\n\t]+/g, " ");
        const style = h.toLowerCase().includes('awb') || h.toLowerCase().includes('dsp') || h.toLowerCase().includes('order') 
          ? 'style=\'mso-number-format:"\\@"\'' 
          : '';
        return `<td ${style}>${val}</td>`;
      }).join("");
      return `<tr>${cells}</tr>`;
    }).join("");

    const htmlTable = `<html><head><meta charset="utf-8"><style>table{border-collapse:collapse;}td{white-space:nowrap;padding:2px 4px;font-family:sans-serif;font-size:10pt;}</style></head><body><table border="1"><tbody>${rowsHtml}</tbody></table></body></html>`;

    try {
      const blobs: Record<string, Blob> = {
        'text/plain': new Blob([plainText], { type: 'text/plain' }),
        'text/html': new Blob([htmlTable], { type: 'text/html' })
      };
      
      await navigator.clipboard.write([new ClipboardItem(blobs)]);
      return true;
    } catch (err) {
      console.error("Clipboard Error:", err);
      // Fallback to basic copy
      const textArea = document.createElement("textarea");
      textArea.value = plainText;
      document.body.appendChild(textArea);
      textArea.select();
      document.execCommand("copy");
      document.body.removeChild(textArea);
      return true;
    }
  }, []);

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
          const wb = XLSX.read(data, { type: 'array', cellDates: true, cellText: true, raw: true });
          const ws = wb.Sheets[wb.SheetNames[0]];
          
          const rawData = XLSX.utils.sheet_to_json(ws, { raw: true, defval: "" });
          
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
              orderId: fixValueToString(findVal(/order|orderid|orderno/)),
              status,
              remark: remark || "No Remark",
              feName: setupData.feName,
              dspId: fixValueToString(setupData.dspId),
              date: formatDate(setupData.date),
              selected: false,
              isIntact
            };
          }).filter(row => row.awb.length >= 3 && row.status !== "unknown");

          if (parsedRows.length === 0) throw new Error("No valid shipments found.");

          const existingSessionIndex = sessions.findIndex(s => s.dspId === setupData.dspId);
          
          if (existingSessionIndex > -1) {
            setSessions(prev => {
              const next = [...prev];
              next[existingSessionIndex] = { 
                ...next[existingSessionIndex], 
                feName: setupData.feName, 
                data: parsedRows, 
                date: formatDate(setupData.date), 
                timestamp: Date.now() 
              };
              return next;
            });
            setSelectedSessionId(sessions[existingSessionIndex].id);
            showToast(`Session updated for DSP ${setupData.dspId}`, "ok");
          } else {
            const newSession: Session = {
              id: crypto.randomUUID(),
              feName: setupData.feName,
              dspId: setupData.dspId,
              date: formatDate(setupData.date),
              data: parsedRows,
              timestamp: Date.now()
            };
            setSessions(prev => [newSession, ...prev]);
            setSelectedSessionId(newSession.id);
            showToast(`Imported ${parsedRows.length} rows!`, "ok");
          }
        } catch (err: any) {
          setUploadError(err.message);
          showToast("Import failed!", "err");
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

  const remarkCountsInStatus = useMemo(() => {
    if (!currentSession) return {};
    let baseData = currentSession.data;
    if (statusFilter !== 'all') {
      if (statusFilter === 'pending') baseData = baseData.filter(r => r.status === 'pending');
      else if (statusFilter === 'dispatched') baseData = baseData.filter(r => r.status === 'dispatched' || r.status === 'dispatch');
      else if (statusFilter === 'rto') baseData = baseData.filter(r => r.status === 'rto');
      else if (statusFilter === 'dto') baseData = baseData.filter(r => r.status === 'dto' || r.status === 'delivered');
    }
    const counts: Record<string, number> = {};
    baseData.forEach(r => {
      const rem = r.remark || "No Remark";
      counts[rem] = (counts[rem] || 0) + 1;
    });
    return counts;
  }, [currentSession, statusFilter]);

  const sortedRemarks = useMemo(() => {
    return Object.entries(remarkCountsInStatus).sort((a, b) => b[1] - a[1]);
  }, [remarkCountsInStatus]);

  const handleCopyTable = useCallback(async (isShortcut = false) => {
    if (!filteredRows.length || !currentSession) return;
    
    const headers = ['Date', 'DSP ID', 'Awb', 'Client', 'Order ID', 'Remark', 'FE Name'];
    const exportRows = filteredRows.map((r, i) => ({
      'Date': formatDate(r.date),
      'DSP ID': i === 0 ? r.dspId : "",
      'Awb': r.awb,
      'Client': r.client,
      'Order ID': r.orderId,
      'Remark': r.remark,
      'FE Name': r.feName
    }));

    const success = await copyDataProfessional(exportRows, headers);
    if (success) {
      showToast(isShortcut ? `Shortcut — Copied ${filteredRows.length} rows` : `Copied ${filteredRows.length} rows to clipboard`, "ok");
    }
  }, [filteredRows, currentSession, copyDataProfessional, showToast]);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      const isCopy = (e.ctrlKey && (e.key === 't' || e.key === 'T')) || 
                     (e.ctrlKey && e.shiftKey && (e.key === 'c' || e.key === 'C'));
      
      if (isCopy) {
        e.preventDefault(); 
        handleCopyTable(true);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [handleCopyTable]);

  const downloadExcel = () => {
    if (!filteredRows.length || !currentSession) return;
    const header = ['Date', 'DSP ID', 'Awb', 'Client', 'Order ID', 'Remark', 'FE Name'];
    const excelData = filteredRows.map((r, i) => [
      formatDate(r.date), 
      { v: i === 0 ? String(r.dspId) : "", t: 's', z: '@' }, 
      { v: String(r.awb), t: 's', z: '@' }, 
      r.client, 
      { v: String(r.orderId), t: 's', z: '@' }, 
      r.remark, 
      r.feName
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...excelData]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, `POD_${currentSession.dspId}.xlsx`);
    showToast(`Downloaded ${filteredRows.length} rows`, "ok");
  };

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

  if (!isMounted) return null;

  return (
    <div className="min-h-screen bg-[#F1F5F9] text-sharp">
      {/* Delhivery Enterprise Header */}
      <header className="fixed top-0 left-0 w-full z-[400] bg-[#0F172A] shadow-lg border-b border-white/5">
        <div className="h-1 w-full bg-gradient-to-r from-[#2563EB] via-[#F59E0B] to-[#EF4444]" />
        <div className="px-6 h-[64px] flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="bg-blue-600/20 p-2 rounded-lg border border-blue-500/20 shadow-inner">
              <Truck className="w-6 h-6 text-blue-400" />
            </div>
            <div>
              <h1 className="text-[17px] font-extrabold text-white tracking-tight leading-none mb-1">POD Management Tool</h1>
              <div className="flex items-center gap-2">
                <span className="text-[10px] bg-white/10 px-1.5 py-0.5 rounded text-slate-300 font-bold uppercase tracking-widest">Palam Vihar RPC</span>
                <span className="text-[10px] font-bold text-[#F59E0B] tracking-wider">BY ASHU</span>
              </div>
            </div>
          </div>
          <div className="flex h-full gap-8">
            {[ { id: "eod", label: "Overview & Export" }, { id: "remark", label: "Remark Replacer" } ].map(tab => (
              <button 
                key={tab.id} 
                onClick={() => { setActiveTab(tab.id as any); setStatusFilter("all"); setActiveRemarkChip(null); }} 
                className={cn(
                  "px-1 py-4 text-[13px] font-bold transition-all relative flex items-center",
                  activeTab === tab.id ? "text-white opacity-100" : "text-slate-400 opacity-70 hover:opacity-100"
                )}
              >
                {tab.label}
                {activeTab === tab.id && <div className="absolute bottom-0 left-0 w-full h-[3px] bg-blue-500 shadow-[0_-2px_8px_rgba(59,130,246,0.6)]" />}
              </button>
            ))}
          </div>
        </div>
      </header>

      <main className="pt-[90px] px-6 pb-20 max-w-[1520px] mx-auto space-y-6">
        {activeTab === "eod" ? (
          <div className="space-y-6">
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
              {/* Import Section */}
              <div className="lg:col-span-4">
                <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-sm">
                  <div className="flex items-center justify-between mb-5">
                    <p className="text-[12px] font-bold text-slate-900 uppercase tracking-widest flex items-center gap-2">
                      <Database className="w-4 h-4 text-blue-600" /> 
                      Session Setup
                    </p>
                    <Settings className="w-4 h-4 text-slate-300" />
                  </div>
                  
                  <div className="space-y-4">
                    <div className="grid grid-cols-2 gap-4">
                      <div className="space-y-1">
                        <label className="text-[11px] font-bold text-slate-400 uppercase">DSP ID</label>
                        <input type="number" value={setupData.dspId} onChange={(e) => setSetupData({...setupData, dspId: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-lg px-4 h-[42px] text-[14px] font-bold text-slate-900 outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 transition-all" placeholder="DSP" />
                      </div>
                      <div className="space-y-1">
                        <label className="text-[11px] font-bold text-slate-400 uppercase">Date</label>
                        <input type="date" value={setupData.date} onChange={(e) => setSetupData({...setupData, date: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-lg px-4 h-[42px] text-[14px] font-bold text-slate-900 outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 transition-all" />
                      </div>
                    </div>
                    
                    <div className="space-y-1">
                      <label className="text-[11px] font-bold text-slate-400 uppercase">FE Name</label>
                      <input type="text" value={setupData.feName} onChange={(e) => setSetupData({...setupData, feName: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-lg px-4 h-[42px] text-[14px] font-bold text-slate-900 outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 transition-all" placeholder="Enter name" />
                    </div>

                    <div className={cn("border-2 border-dashed rounded-xl p-6 text-center transition-all cursor-pointer relative mt-2", uploadError ? "border-rose-200 bg-rose-50" : "border-slate-200 bg-slate-50 hover:border-blue-500 hover:bg-white", isProcessing && "opacity-80 pointer-events-none")}>
                      <input type="file" disabled={isProcessing} onChange={handleFileUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                      {isProcessing ? (
                        <div className="flex flex-col items-center gap-2"><Loader2 className="w-6 h-6 text-blue-600 animate-spin" /><p className="text-[12px] font-bold text-blue-600">Importing...</p></div>
                      ) : uploadError ? (
                        <div className="space-y-2"><AlertCircle className="w-6 h-6 text-rose-500 mx-auto" /><p className="text-[12px] font-bold text-rose-600">{uploadError}</p></div>
                      ) : (
                        <div className="space-y-2"><Download className="w-6 h-6 text-slate-300 mx-auto" /><div className="space-y-1"><p className="text-[13px] font-bold text-slate-700">Drop Delhivery CSV File</p><p className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">or click to browse</p></div></div>
                      )}
                    </div>
                  </div>
                </div>
              </div>

              {/* Sessions Grid */}
              <div className="lg:col-span-8">
                <div className="space-y-4">
                  <div className="flex items-center justify-between px-1">
                    <p className="text-[12px] font-bold text-slate-900 uppercase tracking-widest flex items-center gap-2">
                      <Layers className="w-4 h-4 text-blue-600" />
                      Active Dispatch Sessions
                    </p>
                    {sessions.length > 0 && (
                      <button onClick={() => { if(confirm("Clear all history?")) { setSessions([]); setSelectedSessionId(null); } }} className="text-[11px] font-bold text-rose-600 hover:text-rose-700 transition-colors flex items-center gap-1.5"><Trash2 className="w-3.5 h-3.5" /> Clear History</button>
                    )}
                  </div>
                  
                  <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
                    {sessions.length > 0 ? sessions.map(s => {
                      const isActive = selectedSessionId === s.id;
                      const sTotal = s.data.length;
                      const sPending = s.data.filter(r => r.status === 'pending').length;
                      const sRto = s.data.filter(r => r.status === 'rto').length;
                      const sDto = s.data.filter(r => r.status === 'dto' || r.status === 'delivered').length;

                      return (
                        <div 
                          key={s.id} 
                          onClick={() => { setSelectedSessionId(s.id); setStatusFilter('all'); setActiveRemarkChip(null); }} 
                          className={cn(
                            "bg-white p-4 rounded-xl border-2 cursor-pointer relative transition-all duration-300 group shadow-sm", 
                            isActive ? "border-blue-500 ring-4 ring-blue-500/5" : "border-slate-100 hover:border-slate-300"
                          )}
                        >
                          <button onClick={(e) => { e.stopPropagation(); setSessions(prev => prev.filter(x => x.id !== s.id)); if(selectedSessionId === s.id) setSelectedSessionId(null); }} className="absolute top-4 right-4 text-slate-300 hover:text-rose-600 transition-colors opacity-0 group-hover:opacity-100"><X className="w-4 h-4" /></button>
                          <p className="text-[15px] font-extrabold text-slate-900 truncate pr-8 mb-1">{s.feName}</p>
                          <div className="flex items-center gap-2 mb-4">
                            <span className="text-[10px] font-bold text-slate-400 bg-slate-100 px-1.5 py-0.5 rounded uppercase tracking-wider">{s.dspId}</span>
                            <span className="text-[10px] font-bold text-slate-400 flex items-center gap-1"><Calendar className="w-3 h-3" /> {formatDate(s.date)}</span>
                          </div>
                          
                          <div className="grid grid-cols-2 gap-2">
                            <div className="bg-slate-50 px-2 py-1.5 rounded-lg border border-slate-100">
                              <p className="text-[11px] font-bold text-slate-800">{sTotal} pkt</p>
                            </div>
                            <div className="bg-amber-50 px-2 py-1.5 rounded-lg border border-amber-100">
                              <p className="text-[11px] font-bold text-amber-700">{sPending} pending</p>
                            </div>
                            <div className="bg-rose-50 px-2 py-1.5 rounded-lg border border-rose-100">
                              <p className="text-[11px] font-bold text-rose-700">{sRto} RTO</p>
                            </div>
                            <div className="bg-emerald-50 px-2 py-1.5 rounded-lg border border-emerald-100">
                              <p className="text-[11px] font-bold text-emerald-700">{sDto} DTO</p>
                            </div>
                          </div>
                        </div>
                      );
                    }) : (
                      <div className="col-span-full border border-dashed border-slate-200 rounded-xl flex flex-col items-center justify-center py-10 opacity-60 bg-white">
                        <FileSpreadsheet className="w-10 h-10 text-slate-200 mb-2" />
                        <p className="text-[12px] font-bold text-slate-400 uppercase tracking-widest">No history found</p>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </div>

            {currentSession && (
              <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
                {/* Dashboard Tiles (Delhivery Style) */}
                <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
                  {[
                    { id: 'all', label: 'All', val: stats.total, color: 'text-slate-900', bg: 'bg-white', icon: Layers },
                    { id: 'pending', label: 'Pending', val: stats.pending, color: 'text-amber-600', bg: 'bg-white', icon: Search },
                    { id: 'dispatched', label: 'Dispatch', val: stats.dispatched, color: 'text-blue-600', bg: 'bg-white', icon: Truck },
                    { id: 'rto', label: 'RTO', val: stats.rto, color: 'text-rose-600', bg: 'bg-white', icon: AlertCircle },
                    { id: 'dto', label: 'DTO', val: stats.dto, color: 'text-emerald-600', bg: 'bg-white', icon: Info }
                  ].map((t) => (
                    <button 
                      key={t.id} 
                      onClick={() => { setStatusFilter(t.id); setActiveRemarkChip(null); }} 
                      className={cn(
                        "p-5 rounded-xl border transition-all relative flex flex-col shadow-sm group", 
                        statusFilter === t.id ? "border-blue-500 bg-white ring-4 ring-blue-500/5" : "border-slate-200 bg-white hover:border-slate-300"
                      )}
                    >
                      <div className="flex items-center justify-between mb-2">
                        <span className={cn("text-[11px] font-black uppercase tracking-widest", statusFilter === t.id ? t.color : "text-slate-400")}>{t.label}</span>
                        <t.icon className={cn("w-4 h-4", statusFilter === t.id ? t.color : "text-slate-300")} />
                      </div>
                      <span className={cn("text-[32px] font-black tracking-tighter leading-none mb-1", t.color)}>{t.val}</span>
                      {statusFilter === t.id && <div className="absolute bottom-0 left-0 w-full h-1 bg-current" />}
                    </button>
                  ))}
                </div>

                {/* Remark Breakdown - Only for Active Status */}
                {sortedRemarks.length > 0 && (
                  <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-sm">
                    <div className="flex items-center justify-between mb-4">
                      <div className="space-y-0.5">
                        <h3 className="text-[12px] font-bold text-slate-900 flex items-center gap-2 uppercase tracking-widest">
                          <BarChart3 className="w-4 h-4 text-blue-600" />
                          Remark Breakdown — {statusFilter === 'all' ? 'Overall' : statusFilter.toUpperCase()}
                        </h3>
                        <p className="text-[10px] font-semibold text-slate-400 uppercase tracking-widest">Click any remark chip to filter</p>
                      </div>
                      {activeRemarkChip && (
                        <button onClick={() => setActiveRemarkChip(null)} className="flex items-center gap-2 px-3 py-1.5 bg-slate-900 text-white rounded-lg text-[10px] font-bold uppercase tracking-widest transition-all btn-hover">All {statusFilter}</button>
                      )}
                    </div>
                    <div className="flex flex-wrap gap-2">
                      {sortedRemarks.map(([rem, count]) => {
                        const isRed = /reject|intact/i.test(rem);
                        const isActive = activeRemarkChip === rem;
                        return (
                          <button
                            key={rem}
                            onClick={() => setActiveRemarkChip(isActive ? null : rem)}
                            className={cn(
                              "group px-4 py-2 rounded-lg flex items-center gap-3 text-[13px] font-semibold transition-all border",
                              isActive 
                                ? "bg-blue-600 text-white border-blue-700 shadow-lg scale-[1.02] z-10" 
                                : isRed 
                                  ? "bg-rose-50 text-rose-700 border-rose-100 hover:bg-rose-100/50" 
                                  : "bg-slate-50 text-slate-700 border-slate-100 hover:border-slate-300 hover:bg-white"
                            )}
                          >
                            <span className="truncate max-w-[280px]">{rem}</span>
                            <span className={cn(
                              "min-w-[22px] h-[22px] rounded flex items-center justify-center text-[11px] font-black shadow-inner",
                              isActive ? "bg-white/20 text-white" : isRed ? "bg-rose-600 text-white" : "bg-slate-900 text-white"
                            )}>
                              {count}
                            </span>
                          </button>
                        );
                      })}
                    </div>
                  </div>
                )}

                {/* Table Controls */}
                <div className="flex items-center justify-between gap-4 py-1">
                  <div className="flex gap-3">
                    <button onClick={downloadExcel} className="h-10 px-6 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-[13px] font-bold flex items-center gap-2 shadow-sm transition-all btn-hover"><Download className="w-4 h-4" /> Download Excel</button>
                    <button onClick={() => handleCopyTable(false)} className="h-10 px-6 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-[13px] font-bold flex items-center gap-2 shadow-sm transition-all btn-hover"><Copy className="w-4 h-4" /> Copy Table</button>
                  </div>
                  <div className="relative group">
                    <Search className="w-4 h-4 absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400 group-focus-within:text-blue-500 transition-colors" />
                    <input type="text" placeholder="Search waybill, client or order..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="bg-white border border-slate-200 rounded-lg pl-10 pr-4 h-10 text-[13px] font-bold text-slate-900 outline-none w-[320px] focus:border-blue-500 focus:ring-4 focus:ring-blue-500/5 transition-all shadow-sm" />
                  </div>
                </div>

                {/* Data Table Container */}
                <div className="bg-white rounded-xl overflow-hidden border-2 border-orange-500/40 shadow-xl">
                  <div className="overflow-x-auto custom-scrollbar">
                    <table className="w-full text-left border-collapse table-fixed text-[11px]">
                      <thead className="bg-[#0F172A] text-white">
                        <tr className="h-11">
                          <th style={{ width: '32px' }} className="px-2 text-center"><input type="checkbox" className="w-3.5 h-3.5 rounded" /></th>
                          <th style={{ width: '28px' }} className="px-1 text-center"><Trash2 className="w-3.5 h-3.5 mx-auto opacity-30" /></th>
                          <th style={{ width: '80px' }} className="px-3 font-bold uppercase tracking-widest opacity-80">DSP ID</th>
                          <th style={{ width: '130px' }} className="px-3 font-bold uppercase tracking-widest opacity-80">AWB Number</th>
                          <th style={{ width: '110px' }} className="px-3 font-bold uppercase tracking-widest opacity-80">Client</th>
                          <th style={{ width: '110px' }} className="px-3 font-bold uppercase tracking-widest opacity-80">Order ID</th>
                          <th className="px-3 font-bold uppercase tracking-widest opacity-80">Remark</th>
                          <th style={{ width: '80px' }} className="px-3 font-bold uppercase tracking-widest opacity-80">FE Name</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr className="bg-slate-800 text-white h-8 border-b border-white/5">
                          <td colSpan={8} className="px-3">
                            <div className="flex items-center gap-3">
                              <span className="text-[12px] font-black text-amber-400">{currentSession.dspId}</span>
                              <span className="w-px h-2.5 bg-white/20" />
                              <span className="text-[10px] font-bold text-blue-400 bg-blue-400/10 px-1.5 py-0.5 rounded-full">{filteredRows.length} pkt</span>
                              <span className="ml-auto text-[10px] text-slate-400 font-bold uppercase tracking-widest">{currentSession.feName} · {formatDate(currentSession.date)}</span>
                            </div>
                          </td>
                        </tr>
                        {filteredRows.length > 0 ? filteredRows.map((row) => (
                          <tr key={row.id} className={cn("h-10 border-b border-orange-100 transition-colors group", row.isIntact ? "bg-rose-50/30" : "hover:bg-blue-50/50")}>
                            <td className="px-2 text-center"><input type="checkbox" checked={row.selected} onChange={() => {}} className="w-3.5 h-3.5 rounded" /></td>
                            <td className="px-1 text-center"><button onClick={() => setSessions(prev => prev.map(s => s.id === selectedSessionId ? {...s, data: s.data.filter(r => r.id !== row.id)} : s))} className="text-slate-300 hover:text-rose-600 transition-colors"><Trash2 className="w-3.5 h-3.5 mx-auto" /></button></td>
                            <td className="px-3 font-bold text-slate-900 truncate">{row.dspId}</td>
                            <td 
                              onClick={async () => {
                                const val = String(row.awb);
                                const plainText = `'${val}`;
                                const htmlTable = `<html><body><table><tr><td style='mso-number-format:"\\@"'>${val}</td></tr></table></body></html>`;
                                await navigator.clipboard.write([new ClipboardItem({ 'text/html': new Blob([htmlTable], { type: 'text/html' }), 'text/plain': new Blob([plainText], { type: 'text/plain' }) })]);
                                showToast(`AWB ${val} copied`, 'ok');
                              }}
                              className="px-3 font-black text-blue-700 cursor-pointer hover:underline tracking-tight truncate" 
                              style={{ fontFamily: '"IBM Plex Mono", monospace', fontSize: '11.5px' }}
                            >
                              {row.awb}
                            </td>
                            <td className="px-3 font-bold text-slate-700 truncate">{row.client}</td>
                            <td className="px-3 font-bold text-slate-700 truncate">{row.orderId}</td>
                            <td className="px-3">
                              <span className={cn(
                                "px-2 py-0.5 rounded text-[10px] font-semibold border whitespace-nowrap overflow-hidden text-ellipsis max-w-[280px] inline-block", 
                                row.isIntact ? "bg-rose-50 text-rose-700 border-rose-200" : "bg-amber-50 text-amber-700 border-amber-200"
                              )}>
                                {row.remark}
                              </span>
                            </td>
                            <td className="px-3 font-bold text-slate-500 truncate">{row.feName}</td>
                          </tr>
                        )) : (
                          <tr>
                            <td colSpan={8} className="h-32 text-center bg-white">
                              <p className="text-[12px] font-bold text-slate-300 uppercase tracking-widest">No data matching current filters</p>
                            </td>
                          </tr>
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
            {/* Remark Replacer Section */}
            <div className="lg:col-span-8 space-y-6">
              <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between mb-6">
                  <h2 className="text-[17px] font-bold text-slate-900 flex items-center gap-3 uppercase tracking-tight">
                    <FileSpreadsheet className="w-5 h-5 text-emerald-600" /> 
                    Remark Replacer
                  </h2>
                  {replacerData.length > 0 && (
                    <div className="flex gap-3">
                      <div className="px-3 py-1 bg-slate-100 rounded-lg text-[10px] font-bold text-slate-600 uppercase tracking-wider">TOTAL: {replacerStats.total}</div>
                      <div className="px-3 py-1 bg-emerald-100 rounded-lg text-[10px] font-bold text-emerald-800 uppercase tracking-wider">REPLACED: {replacerStats.replaced}</div>
                    </div>
                  )}
                </div>
                <div className={cn("border-2 border-dashed rounded-xl p-10 text-center transition-all cursor-pointer relative", uploadError ? "border-rose-200 bg-rose-50" : "border-slate-200 bg-slate-50 hover:border-blue-500 hover:bg-white", isProcessing && "opacity-80 pointer-events-none")}>
                  <input type="file" disabled={isProcessing} onChange={handleReplacerUpload} className="absolute inset-0 opacity-0 cursor-pointer z-20" />
                  {isProcessing ? (
                    <div className="flex flex-col items-center gap-2"><Loader2 className="w-8 h-8 text-blue-600 animate-spin" /><p className="text-[12px] font-bold text-blue-600">Analyzing...</p></div>
                  ) : uploadError ? (
                    <div className="space-y-3"><AlertCircle className="w-8 h-8 text-rose-500 mx-auto" /><p className="text-[12px] font-bold text-rose-600">{uploadError}</p></div>
                  ) : (
                    <div className="space-y-2"><Download className="w-8 h-8 text-slate-200 mx-auto" /><div className="space-y-1"><p className="text-[15px] font-bold text-slate-900">Drop Master EOD Sheet Here</p><p className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">Auto replacement engine</p></div></div>
                  )}
                </div>
              </div>
              
              {replacerData.length > 0 && (
                <div className="bg-white rounded-xl shadow-xl border-2 border-emerald-500/20 overflow-hidden animate-in fade-in slide-in-from-bottom-4 duration-500">
                  <div className="p-4 bg-slate-50 border-b flex items-center justify-between">
                    <div className="flex gap-3">
                      <button onClick={downloadOfficialExcel} className="h-10 px-6 bg-slate-900 text-white rounded-lg text-[13px] font-bold flex items-center gap-2 transition-all btn-hover"><Download className="w-4 h-4" /> Download Result</button>
                      <button onClick={() => copyDataProfessional(replacerData, replacerMeta?.headers || [])} className="h-10 px-6 bg-blue-600 text-white rounded-lg text-[13px] font-bold flex items-center gap-2 transition-all btn-hover"><Copy className="w-4 h-4" /> Copy Rows</button>
                    </div>
                    <button onClick={() => { setReplacerData([]); setReplacerMeta(null); }} className="text-[11px] font-bold text-rose-600 hover:text-rose-700 uppercase tracking-widest px-4">Discard</button>
                  </div>
                  <div className="overflow-x-auto max-h-[580px] custom-scrollbar">
                    <table className="w-full text-left border-collapse table-fixed text-[11px]">
                      <thead className="sticky top-0 z-30 bg-[#0F172A] text-white">
                        <tr className="h-10">
                          <th style={{ width: '32px' }} className="px-2 text-center"><Trash2 className="w-3.5 h-3.5 mx-auto opacity-30" /></th>
                          {replacerMeta?.headers.map((h, i) => <th key={i} className="px-3 font-bold uppercase tracking-widest opacity-80 whitespace-nowrap">{h}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {replacerData.map((row) => (
                          <tr key={row.__id} className={cn("h-10 border-b border-emerald-50 transition-colors", row.__isReplaced ? "bg-emerald-50/40" : "bg-white hover:bg-slate-50")}>
                            <td className="px-2 text-center"><button onClick={() => setReplacerData(prev => prev.filter(r => r.__id !== row.__id))} className="text-slate-300 hover:text-rose-500 transition-colors"><Trash2 className="w-3.5 h-3.5 mx-auto" /></button></td>
                            {replacerMeta?.headers.map((h, i) => (
                              <td key={i} className={cn("px-3 whitespace-nowrap font-bold truncate", h === "Remarks Of NSL" && row.__isReplaced ? "text-emerald-700" : h === "Remarks Of NSL" ? "text-amber-700" : "text-slate-900")}>
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
              <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden sticky top-28">
                <div className="bg-emerald-600 p-4 text-white">
                  <h3 className="text-[11px] font-bold uppercase tracking-widest flex items-center gap-2"><Info className="w-4 h-4" /> Replacer Matrix</h3>
                </div>
                <div className="p-4 overflow-y-auto max-h-[64vh] space-y-3 custom-scrollbar bg-slate-50/50">
                  {Object.entries(REMARK_MAPPING).map(([nsl, official]) => (
                    <div key={nsl} className="p-4 bg-white rounded-lg border border-slate-200 space-y-3 hover:border-emerald-400 transition-all group">
                      <div className="space-y-1">
                        <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest group-hover:text-emerald-500 transition-colors">Original Input</p>
                        <p className="text-[13px] font-semibold text-slate-800 leading-tight">{nsl}</p>
                      </div>
                      <div className="space-y-1">
                        <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">Final Output</p>
                        <p className="text-[13px] font-bold text-emerald-700 leading-tight">{official}</p>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

