
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
  Database
} from "lucide-react";
import * as XLSX from "xlsx";
import { cn } from "@/lib/utils";

/**
 * @fileOverview Delhivery POD Management Tool - Palam Vihar RPC Edition
 * Reverted to original screenshot layout with HD clarity and AWB precision.
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
 * Global Date Formatter: DD-MM-YYYY
 */
const formatDate = (val: any): string => {
  if (!val) return "";
  const str = String(val).trim();
  // If already DD-MM-YYYY, return as is
  if (/^\d{2}-\d{2}-\d{4}$/.test(str)) return str;
  // If YYYY-MM-DD
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
    const saved = localStorage.getItem('pod_sessions_v2_final');
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
      localStorage.setItem('pod_sessions_v2_final', JSON.stringify(sessions));
    }
  }, [sessions, isMounted]);

  const showToast = useCallback((msg: string, type: 'ok' | 'err' | 'info') => {
    if (typeof document === 'undefined') return;
    const toast = document.createElement('div');
    toast.className = `fixed bottom-8 left-1/2 -translate-x-1/2 px-6 py-3 rounded-xl text-[13px] font-bold z-[500] shadow-2xl transition-all duration-300 transform scale-95 opacity-0 animate-in fade-in slide-in-from-bottom-4 border ${
      type === 'ok' ? 'bg-emerald-900 text-emerald-100 border-emerald-500/20' : 
      type === 'err' ? 'bg-rose-900 text-rose-100 border-rose-500/20' : 
      'bg-slate-900 text-white border-white/10'
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
    return String(val).trim().replace(/['",]/g, ""); 
  };

  const copyDataProfessional = useCallback(async (rows: any[], headers: string[]) => {
    if (!rows.length) return;

    // Plain text format with apostrophe prefix for WPS/Excel
    const plainText = rows.map(r => 
      headers.map(h => {
        const val = String(r[h] || "").trim();
        if (h === 'Awb' || h === 'AWB Number') return `'${val}`;
        return val;
      }).join("\t")
    ).join("\n");

    // HTML table format for high precision
    const rowsHtml = rows.map(r => {
      const cells = headers.map(h => {
        const val = String(r[h] || "").trim();
        const style = (h === 'Awb' || h === 'AWB Number' || h === 'DSP ID' || h === 'Order ID') 
          ? 'style=\'mso-number-format:"\\@"\'' 
          : '';
        return `<td ${style}>${val}</td>`;
      }).join("");
      return `<tr>${cells}</tr>`;
    }).join("");

    const htmlTable = `<html><body><table border="1"><tbody>${rowsHtml}</tbody></table></body></html>`;

    try {
      const blobs: Record<string, Blob> = {
        'text/plain': new Blob([plainText], { type: 'text/plain' }),
        'text/html': new Blob([htmlTable], { type: 'text/html' })
      };
      await navigator.clipboard.write([new ClipboardItem(blobs)]);
      return true;
    } catch (err) {
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
      showToast("FE Name and DSP ID required!", "err");
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
          const wb = XLSX.read(data, { type: 'array', cellDates: true, cellText: false, raw: true });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rawData = XLSX.utils.sheet_to_json(ws, { raw: true, defval: "" });
          
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
              isIntact
            };
          }).filter(row => row.awb.length >= 3 && row.status !== "unknown");

          if (parsedRows.length === 0) throw new Error("No valid shipments found.");

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

  const handleReplacerUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsProcessing(true);
    const reader = new FileReader();
    reader.onload = (evt) => {
      setTimeout(() => {
        try {
          const data = new Uint8Array(evt.target?.result as ArrayBuffer);
          const wb = XLSX.read(data, { type: 'array', raw: true });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rawData = XLSX.utils.sheet_to_json(ws, { defval: "", raw: true });
          
          const allHeaders = Object.keys(rawData[0]);
          const remarkKey = allHeaders.find(h => h.trim() === "Remarks Of NSL") || "";
          const targetHeaders = ["Date", "DSP No", "Awb", "Client Name", "Order- No", "Remarks Of NSL", "Fe Name"];
          
          const processed = rawData.map(row => {
            const originalRemark = String(row[remarkKey] || "").trim();
            const mappingEntry = Object.entries(REMARK_MAPPING).find(([key]) => 
              key.toLowerCase() === originalRemark.toLowerCase() || originalRemark.toLowerCase().includes(key.toLowerCase())
            );
            const cleanRow: any = { __id: crypto.randomUUID(), __isReplaced: !!mappingEntry };
            targetHeaders.forEach(h => {
              if (h === "Remarks Of NSL") cleanRow[h] = mappingEntry ? mappingEntry[1] : originalRemark;
              else if (h === "Date") cleanRow[h] = formatDate(row[h]);
              else if (h === "Awb" || h === "DSP No" || h === "Order- No") cleanRow[h] = fixValueToString(row[h]);
              else {
                const ik = allHeaders.find(k => k.trim().toLowerCase() === h.toLowerCase());
                cleanRow[h] = ik ? row[ik] : "";
              }
            });
            return cleanRow;
          });

          setReplacerData(processed);
          setReplacerMeta({ headers: targetHeaders, remarkKey, awbKey: "Awb" });
          showToast(`Processed ${processed.length} rows`, "ok");
        } catch (err: any) {
          showToast("Failed to process", "err");
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

  const remarkBreakdown = useMemo(() => {
    if (!currentSession || statusFilter !== 'pending') return [];
    const counts: Record<string, number> = {};
    currentSession.data.filter(r => r.status === 'pending').forEach(r => {
      const rem = r.remark || "No Remark";
      counts[rem] = (counts[rem] || 0) + 1;
    });
    return Object.entries(counts).sort((a, b) => b[1] - a[1]);
  }, [currentSession, statusFilter]);

  const handleCopyTable = useCallback(async () => {
    if (!filteredRows.length || !currentSession) return;
    const headers = ['Date', 'DSP ID', 'AWB Number', 'Client', 'Order ID', 'Remark', 'FE Name'];
    const exportRows = filteredRows.map((r, i) => ({
      'Date': formatDate(r.date),
      'DSP ID': i === 0 ? r.dspId : "",
      'AWB Number': r.awb,
      'Client': r.client,
      'Order ID': r.orderId,
      'Remark': r.remark,
      'FE Name': r.feName
    }));
    const success = await copyDataProfessional(exportRows, headers);
    if (success) showToast(`Copied ${filteredRows.length} rows`, "ok");
  }, [filteredRows, currentSession, copyDataProfessional, showToast]);

  const downloadExcel = () => {
    if (!filteredRows.length || !currentSession) return;
    const header = ['Date', 'DSP ID', 'AWB Number', 'Client', 'Order ID', 'Remark', 'FE Name'];
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
    XLSX.writeFile(wb, `POD_Report_${currentSession.dspId}.xlsx`);
    showToast("Downloaded Excel", "ok");
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
    <div className="min-h-screen bg-[#F8FAFC]">
      {/* HEADER - Minimal Dark Look */}
      <header className="bg-[#0F172A] border-b border-white/5 px-6 h-14 flex items-center justify-between shadow-lg sticky top-0 z-[500]">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center text-white shadow-lg">
            <Truck className="w-5 h-5" />
          </div>
          <div>
            <h1 className="text-sm font-black text-white tracking-tight uppercase">POD Tool</h1>
            <p className="text-[10px] text-blue-400 font-bold uppercase tracking-widest leading-none">Palam Vihar RPC</p>
          </div>
        </div>
        <div className="flex gap-8 h-full">
          {[ {id:"eod", label:"Daily EOD Rejection"}, {id:"remark", label:"Remark Replacer"} ].map(tab => (
            <button 
              key={tab.id}
              onClick={() => { setActiveTab(tab.id as any); setStatusFilter("all"); setActiveRemarkChip(null); }}
              className={cn(
                "h-full px-1 text-[13px] font-bold transition-all relative border-b-2",
                activeTab === tab.id ? "text-white border-blue-500" : "text-slate-400 border-transparent hover:text-white"
              )}
            >
              {tab.label}
            </button>
          ))}
        </div>
        <div className="flex items-center gap-3">
          <span className="text-[10px] font-black text-amber-400 tracking-widest uppercase">By Ashu</span>
        </div>
      </header>

      <main className="max-w-[1440px] mx-auto p-6 space-y-6">
        {activeTab === "eod" ? (
          <>
            {/* SESSION SETUP SECTION */}
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6 overflow-hidden relative">
              <div className="flex items-center gap-2 mb-6">
                <Settings className="w-4 h-4 text-blue-600" />
                <h2 className="text-[11px] font-black text-slate-900 uppercase tracking-widest">Session Setup</h2>
              </div>
              
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
                <div className="space-y-1.5">
                  <label className="text-[11px] font-bold text-slate-500 uppercase tracking-wide">DSP ID</label>
                  <input type="text" value={setupData.dspId} onChange={(e) => setSetupData({...setupData, dspId: e.target.value})} className="w-full bg-[#F9FAFB] border-[1.5px] border-[#D1D5DB] rounded-lg px-3.5 h-[42px] text-[14px] font-bold text-slate-900 outline-none focus:border-[#1976D2] focus:ring-4 focus:ring-blue-500/5" placeholder="Enter DSP ID" />
                </div>
                <div className="space-y-1.5">
                  <label className="text-[11px] font-bold text-slate-500 uppercase tracking-wide">FE Name</label>
                  <input type="text" value={setupData.feName} onChange={(e) => setSetupData({...setupData, feName: e.target.value})} className="w-full bg-[#F9FAFB] border-[1.5px] border-[#D1D5DB] rounded-lg px-3.5 h-[42px] text-[14px] font-bold text-slate-900 outline-none focus:border-[#1976D2] focus:ring-4 focus:ring-blue-500/5" placeholder="Enter Full Name" />
                </div>
                <div className="space-y-1.5">
                  <label className="text-[11px] font-bold text-slate-500 uppercase tracking-wide">Date</label>
                  <input type="date" value={setupData.date} onChange={(e) => setSetupData({...setupData, date: e.target.value})} className="w-full bg-[#F9FAFB] border-[1.5px] border-[#D1D5DB] rounded-lg px-3.5 h-[42px] text-[14px] font-bold text-slate-900 outline-none focus:border-[#1976D2] focus:ring-4 focus:ring-blue-500/5" />
                </div>
              </div>

              <div className={cn("border-2 border-dashed rounded-2xl p-10 text-center transition-all cursor-pointer relative group", isProcessing ? "bg-slate-50 opacity-80" : "bg-slate-50 hover:bg-white hover:border-blue-400")}>
                <input type="file" disabled={isProcessing} onChange={handleFileUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                {isProcessing ? (
                  <div className="flex flex-col items-center gap-3"><Loader2 className="w-10 h-10 text-blue-600 animate-spin" /><p className="text-sm font-black text-blue-600 uppercase">Processing File...</p></div>
                ) : (
                  <div className="space-y-4">
                    <Download className="w-10 h-10 text-slate-300 mx-auto group-hover:scale-110 transition-transform" />
                    <div>
                      <p className="text-sm font-black text-slate-900">Drop Delhivery CSV here, or click to upload</p>
                      <p className="text-[10px] text-slate-400 font-bold uppercase mt-1 tracking-widest">Ready for processing</p>
                    </div>
                  </div>
                )}
              </div>
            </div>

            {/* SESSIONS LIST */}
            <div className="space-y-3">
              <div className="flex items-center justify-between px-1">
                <h3 className="text-[11px] font-black text-slate-900 uppercase tracking-widest">Active Dispatch Sessions</h3>
                {sessions.length > 0 && (
                  <button onClick={() => { if(confirm("Clear history?")) setSessions([]); }} className="text-[10px] font-black text-rose-600 uppercase tracking-widest hover:underline">Clear All</button>
                )}
              </div>
              <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-4 xl:grid-cols-5 gap-4">
                {sessions.map(s => {
                  const sTotal = s.data.length;
                  const sPending = s.data.filter(r => r.status === 'pending').length;
                  const sRto = s.data.filter(r => r.status === 'rto').length;
                  const sDto = s.data.filter(r => r.status === 'dto' || r.status === 'delivered').length;
                  return (
                    <div 
                      key={s.id}
                      onClick={() => { setSelectedSessionId(s.id); setStatusFilter('all'); setActiveRemarkChip(null); }}
                      className={cn(
                        "bg-white p-4 rounded-xl border-[1.5px] cursor-pointer relative transition-all group",
                        selectedSessionId === s.id ? "border-blue-500 shadow-md ring-4 ring-blue-500/5" : "border-slate-200 hover:border-blue-200"
                      )}
                    >
                      <button onClick={(e) => { e.stopPropagation(); setSessions(prev => prev.filter(x => x.id !== s.id)); }} className="absolute top-2 right-2 text-slate-300 hover:text-rose-600 transition-colors opacity-0 group-hover:opacity-100"><X className="w-4 h-4" /></button>
                      <p className="text-[15px] font-black text-slate-900 truncate pr-6">{s.feName}</p>
                      <p className="text-[12px] font-bold text-slate-400 mb-3">{s.dspId} — {s.date}</p>
                      <div className="flex flex-wrap gap-1.5">
                        <span className="px-2 py-0.5 bg-slate-100 rounded text-[10px] font-black text-slate-600">{sTotal} pkt</span>
                        <span className="px-2 py-0.5 bg-amber-100 rounded text-[10px] font-black text-amber-700">{sPending} pending</span>
                        <span className="px-2 py-0.5 bg-rose-100 rounded text-[10px] font-black text-rose-700">{sRto} RTO</span>
                        <span className="px-2 py-0.5 bg-emerald-100 rounded text-[10px] font-black text-emerald-700">{sDto} DTO</span>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            {currentSession && (
              <div className="space-y-6">
                {/* STATUS TABS - Large Numbers Style */}
                <div className="bg-white rounded-2xl border-[1.5px] border-slate-200 shadow-xl overflow-hidden flex divide-x divide-slate-100">
                  {[
                    { id: 'all', label: 'All', val: stats.total, color: 'text-blue-600' },
                    { id: 'pending', label: 'Pending', val: stats.pending, color: 'text-amber-600' },
                    { id: 'dispatched', label: 'Dispatch', val: stats.dispatched, color: 'text-blue-600' },
                    { id: 'rto', label: 'RTO', val: stats.rto, color: 'text-rose-600' },
                    { id: 'dto', label: 'DTO', val: stats.dto, color: 'text-emerald-600' }
                  ].map((t) => (
                    <button 
                      key={t.id}
                      onClick={() => { setStatusFilter(t.id); setActiveRemarkChip(null); }}
                      className={cn(
                        "flex-1 py-5 flex flex-col items-center justify-center transition-all relative group h-[100px]",
                        statusFilter === t.id ? "bg-slate-50" : "hover:bg-slate-50/50"
                      )}
                    >
                      <span className={cn("text-[28px] font-black leading-none mb-1 tracking-tighter", t.color)}>{t.val}</span>
                      <span className={cn("text-[11px] font-black uppercase tracking-[0.2em]", statusFilter === t.id ? t.color : "text-slate-400")}>{t.label}</span>
                      {statusFilter === t.id && <div className="absolute bottom-0 left-0 w-full h-[3px] bg-blue-600" />}
                    </button>
                  ))}
                </div>

                {/* REMARK BREAKDOWN CARD */}
                {statusFilter === 'pending' && remarkBreakdown.length > 0 && (
                  <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-6 animate-in fade-in slide-in-from-top-4 duration-500">
                    <div className="flex items-center justify-between mb-5">
                      <div>
                        <h4 className="text-[12px] font-black text-slate-900 uppercase tracking-widest flex items-center gap-2">
                          <BarChart3 className="w-4 h-4 text-amber-500" />
                          Remark Breakdown — Pending
                        </h4>
                        <p className="text-[10px] font-bold text-slate-400 uppercase mt-1 tracking-widest">Click any remark chip to filter</p>
                      </div>
                      {activeRemarkChip && (
                        <button onClick={() => setActiveRemarkChip(null)} className="h-8 px-4 bg-slate-900 text-white rounded-lg text-[10px] font-black uppercase tracking-widest">All Pending</button>
                      )}
                    </div>
                    <div className="flex flex-wrap gap-3">
                      {remarkBreakdown.map(([rem, count]) => (
                        <button 
                          key={rem}
                          onClick={() => setActiveRemarkChip(activeRemarkChip === rem ? null : rem)}
                          className={cn(
                            "px-4 py-2.5 rounded-lg border flex items-center gap-3 transition-all",
                            activeRemarkChip === rem 
                              ? "bg-blue-600 border-blue-600 text-white shadow-lg" 
                              : "bg-slate-50 border-slate-200 text-slate-800 hover:border-blue-400"
                          )}
                        >
                          <span className="text-[13px] font-semibold">{rem}</span>
                          <span className={cn("text-[11px] font-black min-w-[20px] h-[20px] rounded flex items-center justify-center", activeRemarkChip === rem ? "bg-white/20" : "bg-slate-900 text-white")}>{count}</span>
                        </button>
                      ))}
                    </div>
                  </div>
                )}

                {/* TABLE CONTROLS */}
                <div className="flex items-center justify-between gap-4">
                  <div className="flex gap-2">
                    <button onClick={downloadExcel} className="h-10 px-5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-[13px] font-black flex items-center gap-2 shadow-lg transition-all"><Download className="w-4 h-4" /> Download Excel</button>
                    <button onClick={handleCopyTable} className="h-10 px-5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-[13px] font-black flex items-center gap-2 shadow-lg transition-all"><Copy className="w-4 h-4" /> Copy Table</button>
                  </div>
                  <div className="relative">
                    <Search className="w-4 h-4 absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" />
                    <input type="text" placeholder="Search waybill, client or order..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="bg-white border-[1.5px] border-slate-200 rounded-lg pl-10 pr-4 h-10 text-[13px] font-bold text-slate-900 outline-none w-[320px] focus:border-blue-500 shadow-sm" />
                  </div>
                </div>

                {/* DATA TABLE */}
                <div className="bg-white rounded-xl border-[1.5px] border-orange-500 shadow-2xl overflow-hidden">
                  <div className="overflow-x-auto custom-scrollbar">
                    <table className="w-full text-left border-collapse table-fixed" style={{ tableLayout: 'fixed' }}>
                      <thead className="bg-[#0F172A] text-white">
                        <tr className="h-11">
                          <th style={{ width: '32px' }} className="px-2 text-center"><input type="checkbox" className="w-3.5 h-3.5" /></th>
                          <th style={{ width: '28px' }} className="px-1 text-center"><Trash2 className="w-3.5 h-3.5 opacity-50 mx-auto" /></th>
                          <th style={{ width: '80px' }} className="px-2 text-[11px] font-black uppercase">DSP ID</th>
                          <th style={{ width: '130px' }} className="px-2 text-[11px] font-black uppercase">AWB Number</th>
                          <th style={{ width: '110px' }} className="px-2 text-[11px] font-black uppercase">Client</th>
                          <th style={{ width: '110px' }} className="px-2 text-[11px] font-black uppercase">Order ID</th>
                          <th className="px-2 text-[11px] font-black uppercase">Remark</th>
                          <th style={{ width: '80px' }} className="px-2 text-[11px] font-black uppercase">FE Name</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr className="bg-slate-800 text-white h-8 border-b border-white/10">
                          <td colSpan={8} className="px-2">
                            <div className="flex items-center gap-3">
                              <span className="text-[12px] font-black text-amber-400 uppercase">{currentSession.dspId}</span>
                              <span className="w-px h-3 bg-white/20" />
                              <span className="text-[10px] font-bold bg-white/10 px-1.5 rounded uppercase tracking-widest">{filteredRows.length} pkt</span>
                              <span className="ml-auto text-[10px] text-slate-400 font-bold uppercase">{currentSession.feName} — {formatDate(currentSession.date)}</span>
                            </div>
                          </td>
                        </tr>
                        {filteredRows.length > 0 ? filteredRows.map((row) => (
                          <tr key={row.id} className={cn("h-11 border-b border-orange-100 transition-colors group", row.isIntact ? "bg-rose-50/30" : "hover:bg-blue-50/30")}>
                            <td className="px-2 text-center"><input type="checkbox" className="w-3 h-3" /></td>
                            <td className="px-1 text-center"><button onClick={() => setSessions(prev => prev.map(s => s.id === selectedSessionId ? {...s, data: s.data.filter(r => r.id !== row.id)} : s))} className="text-slate-300 hover:text-rose-600 transition-colors"><Trash2 className="w-3.5 h-3.5 mx-auto" /></button></td>
                            <td className="px-2 text-[11px] font-black text-slate-900 truncate">{row.dspId}</td>
                            <td className="px-2 text-[12px] font-black text-blue-700 font-mono tracking-tighter truncate">{row.awb}</td>
                            <td className="px-2 text-[11px] font-bold text-slate-900 truncate">{row.client}</td>
                            <td className="px-2 text-[11px] font-bold text-slate-900 truncate">{row.orderId}</td>
                            <td className="px-2">
                              <span className={cn(
                                "px-2 py-0.5 rounded-md text-[11px] font-semibold border shadow-sm truncate inline-block max-w-full",
                                row.isIntact ? "bg-rose-50 text-rose-700 border-rose-200" : "bg-amber-50 text-amber-700 border-amber-200"
                              )}>
                                {row.remark}
                              </span>
                            </td>
                            <td className="px-2 text-[11px] font-bold text-slate-500 truncate">{row.feName}</td>
                          </tr>
                        )) : (
                          <tr><td colSpan={8} className="h-32 text-center text-[12px] font-black text-slate-300 uppercase tracking-widest">No matching shipments</td></tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}
          </>
        ) : (
          /* REMARK REPLACER SECTION */
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
            <div className="lg:col-span-8 space-y-6">
              <div className="bg-white rounded-2xl p-8 border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between mb-8">
                  <h2 className="text-lg font-black text-slate-900 uppercase tracking-tighter flex items-center gap-3">
                    <FileSpreadsheet className="w-6 h-6 text-emerald-600" /> Remark Replacer
                  </h2>
                </div>
                <div className="border-2 border-dashed rounded-2xl p-12 text-center transition-all cursor-pointer relative bg-slate-50 hover:bg-white hover:border-blue-500">
                  <input type="file" onChange={handleReplacerUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                  <div className="space-y-4">
                    <Download className="w-10 h-10 text-slate-300 mx-auto" />
                    <div>
                      <p className="text-base font-black text-slate-900">Drop Master EOD Sheet Here</p>
                      <p className="text-[11px] text-slate-400 font-bold uppercase tracking-widest mt-1">Automatic remark replacement engine</p>
                    </div>
                  </div>
                </div>
              </div>

              {replacerData.length > 0 && (
                <div className="bg-white rounded-2xl border-2 border-emerald-500/20 shadow-xl overflow-hidden">
                  <div className="p-4 bg-slate-50 border-b flex items-center justify-between">
                    <div className="flex gap-2">
                      <button onClick={() => {}} className="h-10 px-6 bg-slate-900 text-white rounded-xl text-[13px] font-black uppercase">Download Report</button>
                    </div>
                    <button onClick={() => setReplacerData([])} className="text-[11px] font-black text-rose-600 uppercase">Discard</button>
                  </div>
                  <div className="overflow-x-auto max-h-[600px] custom-scrollbar">
                    <table className="w-full text-left border-collapse table-fixed text-[12px]">
                      <thead className="sticky top-0 bg-[#0F172A] text-white">
                        <tr className="h-12">
                          <th style={{ width: '40px' }} className="px-4 text-center">X</th>
                          {replacerMeta?.headers.map((h, i) => <th key={i} className="px-4 font-black text-[10px] uppercase tracking-widest">{h}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {replacerData.map((row) => (
                          <tr key={row.__id} className={cn("h-11 border-b transition-colors", row.__isReplaced ? "bg-emerald-50/40" : "bg-amber-50/40")}>
                            <td className="px-4 text-center"><button onClick={() => setReplacerData(prev => prev.filter(r => r.__id !== row.__id))} className="text-slate-300 hover:text-rose-600">×</button></td>
                            {replacerMeta?.headers.map((h, i) => (
                              <td key={i} className={cn("px-4 font-bold truncate", h === "Remarks Of NSL" && row.__isReplaced ? "text-emerald-700" : h === "Remarks Of NSL" ? "text-amber-700" : "text-slate-900")}>
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
              <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden sticky top-20">
                <div className="bg-emerald-600 p-4 text-white font-black uppercase text-[12px] tracking-widest">Replacer Matrix</div>
                <div className="p-4 space-y-4 max-h-[70vh] overflow-y-auto custom-scrollbar">
                  {Object.entries(REMARK_MAPPING).map(([nsl, off]) => (
                    <div key={nsl} className="p-4 bg-slate-50 rounded-xl border border-slate-100 space-y-2">
                      <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest">Original</p>
                      <p className="text-[13px] font-bold text-amber-700 leading-tight">{nsl}</p>
                      <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest pt-2">Replacement</p>
                      <p className="text-[13px] font-black text-emerald-800 leading-tight">{off}</p>
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
