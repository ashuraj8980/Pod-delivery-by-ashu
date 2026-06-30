
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
  Filter,
  BarChart3,
  CheckCircle2,
  Settings,
  Database
} from "lucide-react";
import * as XLSX from "xlsx";
import { cn } from "@/lib/utils";

/**
 * @fileOverview Delhivery POD Management Tool - Palam Vihar RPC Edition
 * Professional HD UI with 100% AWB Precision for WPS Office.
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
    const saved = localStorage.getItem('pod_sessions_rpc_hd');
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
      localStorage.setItem('pod_sessions_rpc_hd', JSON.stringify(sessions));
    }
  }, [sessions, isMounted]);

  const showToast = useCallback((msg: string, type: 'ok' | 'err' | 'info') => {
    if (typeof document === 'undefined') return;
    const toast = document.createElement('div');
    toast.className = `fixed bottom-8 left-1/2 -translate-x-1/2 px-6 py-3 rounded-2xl text-[13px] font-bold z-[500] shadow-2xl transition-all duration-300 transform scale-95 opacity-0 animate-in fade-in slide-in-from-bottom-4 border ${
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

  const copyDataProfessional = useCallback(async (rows: any[], headers: string[], isSingle = false) => {
    if (!rows.length) return;

    const plainText = rows.map(r => 
      headers.map(h => {
        const val = String(r[h] || "").trim().replace(/[\r\n\t]+/g, " ");
        if (h.toLowerCase().includes('awb')) return `'${val}`;
        return val;
      }).join("\t")
    ).join("\n");

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

  const handleReplacerUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    setUploadError(null);
    
    const reader = new FileReader();
    reader.onload = (evt) => {
      setTimeout(() => {
        try {
          const data = new Uint8Array(evt.target?.result as ArrayBuffer);
          const wb = XLSX.read(data, { type: 'array', cellText: true, cellDates: true, raw: true });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rawData = XLSX.utils.sheet_to_json(ws, { defval: "", raw: true });
          
          if (!Array.isArray(rawData) || rawData.length === 0) throw new Error("Empty file.");

          const allHeaders = Object.keys(rawData[0]);
          const remarkKey = allHeaders.find(h => h.trim() === "Remarks Of NSL") || "";
          const awbKey = allHeaders.find(h => h.trim() === "Awb") || "";

          if (!remarkKey) throw new Error("This file does not have 'Remarks Of NSL' column.");

          const targetHeaders = ["Date", "DSP No", "Awb", "Client Name", "Order- No", "Remarks Of NSL", "Fe Name"];
          let replacedCount = 0;
          let noMappingCount = 0;

          const processed = rawData.map(row => {
            const originalRemark = String(row[remarkKey] || "").trim();
            let finalRemark = originalRemark;
            let isReplaced = false;

            const mappingEntry = Object.entries(REMARK_MAPPING).find(([key]) => 
              key.toLowerCase() === originalRemark.toLowerCase() || originalRemark.toLowerCase().includes(key.toLowerCase())
            );

            if (mappingEntry) {
              finalRemark = mappingEntry[1];
              isReplaced = true;
              replacedCount++;
            } else {
              noMappingCount++;
            }

            const cleanRow: any = { __isReplaced: isReplaced, __id: crypto.randomUUID() };
            targetHeaders.forEach(h => {
              if (h === "Remarks Of NSL") {
                cleanRow[h] = finalRemark;
              } else if (h === "Awb") {
                cleanRow[h] = fixValueToString(row[awbKey]);
              } else if (h === "DSP No") {
                cleanRow[h] = fixValueToString(row["DSP No"]);
              } else if (h === "Order- No") {
                cleanRow[h] = fixValueToString(row["Order- No"]);
              } else if (h === "Date") {
                cleanRow[h] = formatDate(row["Date"]);
              } else {
                const inputKey = allHeaders.find(key => key.trim().toLowerCase() === h.toLowerCase());
                cleanRow[h] = inputKey ? row[inputKey] : "";
              }
            });
            return cleanRow;
          });

          setReplacerData(processed);
          setReplacerMeta({ headers: targetHeaders, remarkKey, awbKey });
          setReplacerStats({ total: processed.length, replaced: replacedCount, noMapping: noMappingCount });
          showToast(`Processed ${processed.length} rows`, "ok");
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

  const downloadOfficialExcel = () => {
    if (!replacerData.length || !replacerMeta) return;
    const headers = replacerMeta.headers; 
    const ws = XLSX.utils.json_to_sheet(replacerData.map(r => {
      const row: any = {};
      headers.forEach(h => {
        if (h === "Awb" || h === "DSP No" || h === "Order- No") {
          row[h] = { v: String(r[h]), t: 's', z: '@' };
        } else if (h === "Date") {
          row[h] = formatDate(r[h]);
        } else {
          row[h] = r[h];
        }
      });
      return row;
    }), { header: headers });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Remarks");
    XLSX.writeFile(wb, `EOD_Remarks_${new Date().toISOString().split('T')[0]}.xlsx`);
    showToast("Excel Downloaded", "ok");
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
      <header className="fixed top-0 left-0 w-full z-[400] bg-[#0F172A] shadow-xl">
        <div className="h-[2px] w-full bg-gradient-to-r from-blue-600 via-amber-400 to-rose-500" />
        <div className="px-6 h-[68px] flex items-center justify-between text-white">
          <div className="flex items-center gap-5">
            <div className="bg-blue-600/10 p-2.5 rounded-xl border border-blue-500/20 shadow-inner">
              <Truck className="w-6 h-6 text-blue-400" />
            </div>
            <div className="space-y-0.5">
              <h1 className="text-[18px] font-extrabold leading-none tracking-tight text-white drop-shadow-sm">POD Management Tool</h1>
              <div className="flex items-center gap-2">
                <span className="text-[10px] bg-white/10 px-1.5 py-0.5 rounded text-slate-300 font-bold tracking-widest uppercase">Palam Vihar RPC</span>
                <span className="text-[10px] font-extrabold text-amber-400 tracking-wider">BY ASHU</span>
              </div>
            </div>
          </div>
          <div className="flex gap-10 h-full">
            {[ { id: "eod", label: "Overview & Export" }, { id: "remark", label: "Remark Replacer" } ].map(tab => (
              <button 
                key={tab.id} 
                onClick={() => { setActiveTab(tab.id as any); setStatusFilter("all"); setActiveRemarkChip(null); }} 
                className={cn(
                  "px-2 py-4 text-[13px] font-bold transition-all relative h-full flex items-center",
                  activeTab === tab.id ? "text-white opacity-100" : "text-slate-400 opacity-70 hover:opacity-100"
                )}
              >
                {tab.label}
                {activeTab === tab.id && <div className="absolute bottom-0 left-0 w-full h-[3px] bg-blue-500 shadow-[0_-2px_8px_rgba(59,130,246,0.5)]" />}
              </button>
            ))}
          </div>
        </div>
      </header>

      <main className="pt-[100px] px-6 pb-20 max-w-[1480px] mx-auto space-y-8">
        {activeTab === "eod" ? (
          <div className="space-y-8">
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
              <div className="lg:col-span-4">
                <div className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm hover:shadow-md transition-all duration-300">
                  <div className="flex items-center justify-between mb-6">
                    <p className="text-[12px] font-black text-slate-900 uppercase tracking-widest flex items-center gap-2">
                      <Database className="w-4 h-4 text-blue-600" /> 
                      Import Data
                    </p>
                    <Settings className="w-4 h-4 text-slate-300" />
                  </div>
                  
                  <div className="space-y-5">
                    <div className="grid grid-cols-2 gap-4">
                      <div className="space-y-1.5">
                        <label className="text-[11px] font-bold text-slate-500 uppercase tracking-wide">DSP ID</label>
                        <input type="number" value={setupData.dspId} onChange={(e) => setSetupData({...setupData, dspId: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 h-[44px] text-[14px] font-bold text-slate-900 outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 transition-all" placeholder="ID" />
                      </div>
                      <div className="space-y-1.5">
                        <label className="text-[11px] font-bold text-slate-500 uppercase tracking-wide">Date</label>
                        <input type="date" value={setupData.date} onChange={(e) => setSetupData({...setupData, date: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 h-[44px] text-[14px] font-bold text-slate-900 outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 transition-all" />
                      </div>
                    </div>
                    
                    <div className="space-y-1.5">
                      <label className="text-[11px] font-bold text-slate-500 uppercase tracking-wide">FE Name</label>
                      <input type="text" value={setupData.feName} onChange={(e) => setSetupData({...setupData, feName: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 h-[44px] text-[14px] font-bold text-slate-900 outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 transition-all" placeholder="Enter Full Name" />
                    </div>

                    <div className={cn("border-2 border-dashed rounded-2xl p-8 text-center transition-all cursor-pointer relative group mt-2", uploadError ? "border-rose-200 bg-rose-50" : "border-slate-200 bg-slate-50 hover:border-blue-400 hover:bg-white", isProcessing && "opacity-80 pointer-events-none")}>
                      <input type="file" disabled={isProcessing} onChange={handleFileUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                      {isProcessing ? (
                        <div className="flex flex-col items-center gap-3"><Loader2 className="w-8 h-8 text-blue-600 animate-spin" /><p className="text-[14px] font-bold text-blue-600">Processing...</p></div>
                      ) : uploadError ? (
                        <div className="space-y-3"><AlertCircle className="w-8 h-8 text-rose-500 mx-auto" /><p className="text-[13px] font-bold text-rose-600">{uploadError}</p><button onClick={() => setUploadError(null)} className="bg-rose-600 text-white px-5 py-2 rounded-lg text-[11px] font-bold uppercase">Try Again</button></div>
                      ) : (
                        <div className="space-y-3"><Download className="w-8 h-8 text-slate-300 mx-auto group-hover:scale-110 group-hover:text-blue-500 transition-all duration-300" /><div className="space-y-1"><p className="text-[13px] font-bold text-slate-800">Drop Delhivery CSV File</p><p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest">or click to browse local files</p></div></div>
                      )}
                    </div>
                  </div>
                </div>
              </div>

              <div className="lg:col-span-8">
                <div className="space-y-4 h-full flex flex-col">
                  <div className="flex items-center justify-between px-1">
                    <p className="text-[12px] font-black text-slate-900 uppercase tracking-[0.1em]">Active Dispatch Sessions</p>
                    {sessions.length > 0 && (
                      <button onClick={() => { if(confirm("Clear all history?")) { setSessions([]); setSelectedSessionId(null); } }} className="text-[10px] font-black text-rose-600 hover:text-rose-700 transition-colors uppercase tracking-widest flex items-center gap-1.5"><Trash2 className="w-3.5 h-3.5" /> Clear All History</button>
                    )}
                  </div>
                  
                  <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4 flex-1">
                    {sessions.length > 0 ? sessions.map(s => {
                      const sTotal = s.data.length;
                      const sPending = s.data.filter(r => r.status === 'pending').length;
                      const sRto = s.data.filter(r => r.status === 'rto').length;
                      const sDto = s.data.filter(r => r.status === 'dto' || r.status === 'delivered').length;
                      const isActive = selectedSessionId === s.id;
                      
                      return (
                        <div 
                          key={s.id} 
                          onClick={() => { setSelectedSessionId(s.id); setStatusFilter('all'); setActiveRemarkChip(null); }} 
                          className={cn(
                            "bg-white p-5 rounded-2xl border-2 cursor-pointer relative transition-all duration-300 group flex flex-col justify-between shadow-sm", 
                            isActive ? "border-blue-500 ring-4 ring-blue-500/5 shadow-blue-100/50" : "border-slate-100 hover:border-slate-200 hover:shadow-md"
                          )}
                        >
                          <button onClick={(e) => { e.stopPropagation(); setSessions(prev => prev.filter(x => x.id !== s.id)); if(selectedSessionId === s.id) setSelectedSessionId(null); }} className="absolute top-4 right-4 text-slate-300 hover:text-rose-600 transition-colors opacity-0 group-hover:opacity-100"><X className="w-4 h-4" /></button>
                          <div>
                            <p className="text-[15px] font-extrabold text-slate-900 truncate pr-8 leading-tight">{s.feName}</p>
                            <p className="text-[10px] font-bold text-slate-400 mt-1 uppercase tracking-widest">{s.dspId} · {formatDate(s.date)}</p>
                          </div>
                          
                          <div className="mt-5 grid grid-cols-2 gap-2">
                            <div className="bg-slate-50 px-3 py-1.5 rounded-lg border border-slate-100">
                              <p className="text-[9px] font-bold text-slate-400 uppercase tracking-tighter">Packets</p>
                              <p className="text-[13px] font-black text-slate-800">{sTotal}</p>
                            </div>
                            <div className="bg-amber-50 px-3 py-1.5 rounded-lg border border-amber-100">
                              <p className="text-[9px] font-bold text-amber-500 uppercase tracking-tighter">Pending</p>
                              <p className="text-[13px] font-black text-amber-700">{sPending}</p>
                            </div>
                            <div className="bg-rose-50 px-3 py-1.5 rounded-lg border border-rose-100">
                              <p className="text-[9px] font-bold text-rose-500 uppercase tracking-tighter">RTO</p>
                              <p className="text-[13px] font-black text-rose-700">{sRto}</p>
                            </div>
                            <div className="bg-emerald-50 px-3 py-1.5 rounded-lg border border-emerald-100">
                              <p className="text-[9px] font-bold text-emerald-500 uppercase tracking-tighter">DTO</p>
                              <p className="text-[13px] font-black text-emerald-700">{sDto}</p>
                            </div>
                          </div>
                        </div>
                      );
                    }) : (
                      <div className="col-span-full border-2 border-dashed border-slate-100 rounded-2xl flex flex-col items-center justify-center py-12 opacity-50">
                        <FileSpreadsheet className="w-10 h-10 text-slate-200 mb-2" />
                        <p className="text-[12px] font-bold text-slate-400 uppercase tracking-[0.2em]">No history found</p>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </div>

            {currentSession && (
              <div className="space-y-8 animate-in fade-in slide-in-from-bottom-6 duration-500">
                <div className="bg-white rounded-3xl border border-slate-200 shadow-xl overflow-hidden flex h-[100px]">
                  {[
                    { id: 'all', label: 'Overview', val: stats.total, color: 'text-slate-900', bg: 'bg-white' },
                    { id: 'pending', label: 'Pending', val: stats.pending, color: 'text-amber-600', bg: 'bg-amber-50/20' },
                    { id: 'dispatched', label: 'Dispatched', val: stats.dispatched, color: 'text-blue-600', bg: 'bg-blue-50/20' },
                    { id: 'rto', label: 'RTO (Rejection)', val: stats.rto, color: 'text-rose-600', bg: 'bg-rose-50/20' },
                    { id: 'dto', label: 'DTO (Success)', val: stats.dto, color: 'text-emerald-600', bg: 'bg-emerald-50/20' }
                  ].map((t, i) => (
                    <button 
                      key={t.id} 
                      onClick={() => { setStatusFilter(t.id); setActiveRemarkChip(null); }} 
                      className={cn(
                        "flex-1 flex flex-col items-center justify-center transition-all border-r border-slate-100 last:border-r-0 relative group", 
                        statusFilter === t.id ? t.bg : "bg-white hover:bg-slate-50/50"
                      )}
                    >
                      <span className={cn("text-[32px] font-black leading-none mb-1 tracking-tighter transition-transform group-active:scale-95", t.color)}>{t.val}</span>
                      <span className={cn("text-[11px] font-black uppercase tracking-[0.2em]", statusFilter === t.id ? t.color : "text-slate-400")}>{t.label}</span>
                      {statusFilter === t.id && <div className="absolute bottom-0 left-0 w-full h-1.5 bg-current" />}
                    </button>
                  ))}
                </div>

                {sortedRemarks.length > 0 && (
                  <div className="bg-white rounded-3xl p-8 border border-slate-200 shadow-xl">
                    <div className="flex items-center justify-between mb-8">
                      <div className="space-y-1">
                        <h3 className="text-[13px] font-black text-slate-900 flex items-center gap-2 uppercase tracking-widest">
                          <BarChart3 className="w-5 h-5 text-blue-600" />
                          Remark Analysis — {statusFilter.toUpperCase()}
                        </h3>
                        <p className="text-[11px] font-bold text-slate-400 uppercase tracking-widest opacity-60">Drill down into specific rejection reasons</p>
                      </div>
                      {activeRemarkChip && (
                        <button onClick={() => setActiveRemarkChip(null)} className="flex items-center gap-2 px-4 py-2 bg-slate-900 text-white rounded-xl text-[10px] font-black uppercase tracking-widest transition-all active:scale-95 shadow-lg shadow-slate-200">Reset View</button>
                      )}
                    </div>
                    <div className="flex flex-wrap gap-3">
                      {sortedRemarks.map(([rem, count]) => {
                        const isRed = /reject|intact/i.test(rem);
                        const isActive = activeRemarkChip === rem;
                        return (
                          <button
                            key={rem}
                            onClick={() => setActiveRemarkChip(isActive ? null : rem)}
                            className={cn(
                              "group px-5 py-3 rounded-2xl flex items-center gap-4 text-[13px] font-bold transition-all border-2",
                              isActive 
                                ? "bg-blue-600 text-white border-blue-700 shadow-xl shadow-blue-200 scale-[1.03] z-10" 
                                : isRed 
                                  ? "bg-rose-50 text-rose-700 border-rose-100 hover:border-rose-200 hover:bg-rose-100/50" 
                                  : "bg-slate-50 text-slate-800 border-slate-100 hover:border-slate-300 hover:bg-white"
                            )}
                          >
                            <span className="truncate max-w-[320px]">{rem}</span>
                            <span className={cn(
                              "min-w-[26px] h-[26px] rounded-lg flex items-center justify-center text-[11px] font-black shadow-inner",
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

                <div className="flex items-center justify-between gap-4 py-2">
                  <div className="flex gap-4">
                    <button onClick={downloadExcel} className="h-12 px-8 bg-emerald-600 hover:bg-emerald-700 text-white rounded-2xl text-[14px] font-black flex items-center gap-3 shadow-lg shadow-emerald-100 transition-all btn-hover"><Download className="w-5 h-5" /> Export Excel</button>
                    <button onClick={() => handleCopyTable(false)} className="h-12 px-8 bg-blue-600 hover:bg-blue-700 text-white rounded-2xl text-[14px] font-black flex items-center gap-3 shadow-lg shadow-blue-100 transition-all btn-hover"><Copy className="w-5 h-5" /> Copy Data</button>
                  </div>
                  <div className="relative group">
                    <Search className="w-5 h-5 absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 group-focus-within:text-blue-500 transition-colors" />
                    <input type="text" placeholder="Search waybill, client or order..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="bg-white border-2 border-slate-100 rounded-2xl pl-12 pr-6 h-12 text-[14px] font-bold text-slate-900 outline-none w-[360px] focus:border-blue-500 focus:ring-4 focus:ring-blue-500/5 transition-all shadow-sm" />
                  </div>
                </div>

                <div className="bg-white rounded-[2.5rem] overflow-hidden border-2 border-orange-500/30 shadow-2xl">
                  <div className="overflow-x-auto custom-scrollbar">
                    <table className="w-full text-left border-collapse table-fixed">
                      <thead className="bg-[#0F172A] text-white">
                        <tr className="h-12">
                          <th style={{ width: '40px' }} className="px-4 text-center"><input type="checkbox" className="w-4 h-4 rounded-md" /></th>
                          <th style={{ width: '36px' }} className="px-2 text-center"><Trash2 className="w-4 h-4 mx-auto opacity-50" /></th>
                          <th style={{ width: '90px' }} className="px-4 text-[11px] font-black uppercase tracking-widest opacity-80">DSP ID</th>
                          <th style={{ width: '150px' }} className="px-4 text-[11px] font-black uppercase tracking-widest opacity-80">AWB Number</th>
                          <th style={{ width: '130px' }} className="px-4 text-[11px] font-black uppercase tracking-widest opacity-80">Client</th>
                          <th style={{ width: '130px' }} className="px-4 text-[11px] font-black uppercase tracking-widest opacity-80">Order ID</th>
                          <th className="px-4 text-[11px] font-black uppercase tracking-widest opacity-80">Final Remark</th>
                          <th style={{ width: '100px' }} className="px-4 text-[11px] font-black uppercase tracking-widest opacity-80">FE Name</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr className="bg-slate-800 text-white h-10 border-b border-white/5">
                          <td colSpan={8} className="px-4">
                            <div className="flex items-center gap-3">
                              <span className="text-[13px] font-black text-amber-400 tracking-tighter">{currentSession.dspId}</span>
                              <span className="w-[1.5px] h-3 bg-white/20" />
                              <span className="text-[11px] font-black text-blue-400 uppercase tracking-widest">{filteredRows.length} shipments found</span>
                              <span className="ml-auto text-[10px] text-slate-400 font-bold uppercase tracking-widest">{currentSession.feName} · {formatDate(currentSession.date)}</span>
                            </div>
                          </td>
                        </tr>
                        {filteredRows.length > 0 ? filteredRows.map((row) => (
                          <tr key={row.id} className={cn("h-12 border-b border-orange-100 transition-colors group", row.isIntact ? "bg-rose-50/40" : "hover:bg-blue-50/50")}>
                            <td className="px-4 text-center"><input type="checkbox" checked={row.selected} onChange={() => {}} className="w-4 h-4 rounded-md" /></td>
                            <td className="px-2 text-center"><button onClick={() => setSessions(prev => prev.map(s => s.id === selectedSessionId ? {...s, data: s.data.filter(r => r.id !== row.id)} : s))} className="text-slate-300 hover:text-rose-600 transition-colors"><Trash2 className="w-4 h-4 mx-auto" /></button></td>
                            <td className="px-4 text-[12px] font-black text-slate-900 truncate">{row.dspId}</td>
                            <td 
                              onClick={async () => {
                                const val = String(row.awb);
                                const plainText = `'${val}`;
                                const htmlTable = `<html><body><table><tr><td style='mso-number-format:"\\@"'>${val}</td></tr></table></body></html>`;
                                await navigator.clipboard.write([new ClipboardItem({ 'text/html': new Blob([htmlTable], { type: 'text/html' }), 'text/plain': new Blob([plainText], { type: 'text/plain' }) })]);
                                showToast(`AWB ${val} copied`, 'ok');
                              }}
                              className="px-4 text-[13px] font-black text-blue-700 cursor-pointer hover:underline tracking-tight truncate" 
                              style={{ fontFamily: '"IBM Plex Mono", monospace' }}
                            >
                              {row.awb}
                            </td>
                            <td className="px-4 text-[12px] font-bold text-slate-800 truncate">{row.client}</td>
                            <td className="px-4 text-[12px] font-bold text-slate-800 truncate">{row.orderId}</td>
                            <td className="px-4">
                              <span className={cn(
                                "px-3 py-1 rounded-xl text-[11px] font-black border whitespace-nowrap overflow-hidden text-ellipsis max-w-[240px] inline-block shadow-sm", 
                                row.isIntact ? "bg-rose-50 text-rose-700 border-rose-200" : "bg-amber-50 text-amber-700 border-amber-200"
                              )}>
                                {row.remark || "No Remark"}
                              </span>
                            </td>
                            <td className="px-4 text-[12px] font-bold text-slate-500 truncate">{row.feName}</td>
                          </tr>
                        )) : (
                          <tr>
                            <td colSpan={8} className="h-32 text-center">
                              <p className="text-[12px] font-black text-slate-300 uppercase tracking-widest">No data matching current filters</p>
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
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
            <div className="lg:col-span-8 space-y-8">
              <div className="bg-white rounded-3xl p-8 border border-slate-200 shadow-xl">
                <div className="flex items-center justify-between mb-8">
                  <h2 className="text-[18px] font-black text-slate-900 flex items-center gap-3 uppercase tracking-tighter">
                    <FileSpreadsheet className="w-6 h-6 text-emerald-600" /> 
                    Remark Replacer
                  </h2>
                  {replacerData.length > 0 && (
                    <div className="flex gap-3">
                      <div className="px-4 py-1.5 bg-slate-100 rounded-xl text-[11px] font-black text-slate-700">TOTAL: {replacerStats.total}</div>
                      <div className="px-4 py-1.5 bg-emerald-100 rounded-xl text-[11px] font-black text-emerald-800">REPLACED: {replacerStats.replaced}</div>
                    </div>
                  )}
                </div>
                <div className={cn("border-2 border-dashed rounded-3xl p-12 text-center transition-all cursor-pointer relative group", uploadError ? "border-rose-200 bg-rose-50" : "border-slate-200 bg-slate-50 hover:border-blue-500 hover:bg-white", isProcessing && "opacity-80 pointer-events-none")}>
                  <input type="file" disabled={isProcessing} onChange={handleReplacerUpload} className="absolute inset-0 opacity-0 cursor-pointer z-20" />
                  {isProcessing ? (
                    <div className="flex flex-col items-center gap-4"><Loader2 className="w-10 h-10 text-blue-600 animate-spin" /><p className="text-[14px] font-black text-blue-600 uppercase tracking-widest">Analyzing Excel File...</p></div>
                  ) : uploadError ? (
                    <div className="space-y-4"><AlertCircle className="w-10 h-10 text-rose-500 mx-auto" /><p className="text-[14px] font-black text-rose-600">{uploadError}</p><button onClick={() => setUploadError(null)} className="bg-rose-600 text-white px-6 py-3 rounded-2xl text-[12px] font-black uppercase shadow-lg shadow-rose-100">Try Again</button></div>
                  ) : (
                    <div className="space-y-4"><Download className="w-10 h-10 text-slate-200 mx-auto group-hover:scale-110 group-hover:text-blue-500 transition-all duration-500" /><div className="space-y-1"><p className="text-[16px] font-black text-slate-900">Drop Master EOD Sheet Here</p><p className="text-[11px] text-slate-400 font-bold uppercase tracking-[0.3em]">Automatic remark replacement engine</p></div></div>
                  )}
                </div>
              </div>
              
              {replacerData.length > 0 && (
                <div className="bg-white rounded-[2.5rem] shadow-2xl border-2 border-emerald-500/20 overflow-hidden animate-in fade-in slide-in-from-bottom-8 duration-700">
                  <div className="p-6 bg-slate-50/50 border-b flex items-center justify-between">
                    <div className="flex gap-4">
                      <button onClick={downloadOfficialExcel} className="h-12 px-8 bg-slate-900 text-white rounded-2xl text-[14px] font-black flex items-center gap-3 transition-all btn-hover shadow-xl shadow-slate-200"><Download className="w-5 h-5" /> Download Report</button>
                      <button onClick={() => copyDataProfessional(replacerData, replacerMeta?.headers || [])} className="h-12 px-8 bg-blue-600 text-white rounded-2xl text-[14px] font-black flex items-center gap-3 transition-all btn-hover shadow-xl shadow-blue-200"><Copy className="w-5 h-5" /> Copy Result</button>
                    </div>
                    <button onClick={() => { setReplacerData([]); setReplacerMeta(null); }} className="text-[11px] font-black text-rose-600 px-6 hover:text-rose-700 uppercase tracking-widest">Discard</button>
                  </div>
                  <div className="overflow-x-auto max-h-[640px] custom-scrollbar">
                    <table className="w-full text-left border-collapse table-fixed text-[12px]">
                      <thead className="sticky top-0 z-30 bg-[#0F172A] text-white">
                        <tr className="h-12">
                          <th style={{ width: '40px' }} className="px-4 text-center"><Trash2 className="w-4 h-4 mx-auto" /></th>
                          {replacerMeta?.headers.map((h, i) => <th key={i} className="px-4 font-black whitespace-nowrap text-[10px] uppercase tracking-widest opacity-80">{h}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {replacerData.map((row) => (
                          <tr key={row.__id} className={cn("h-12 border-b border-emerald-50 transition-colors", row.__isReplaced ? "bg-emerald-50/30" : "bg-amber-50/30")}>
                            <td className="px-4 text-center"><button onClick={() => setReplacerData(prev => prev.filter(r => r.__id !== row.__id))} className="text-slate-300 hover:text-rose-500 transition-colors"><Trash2 className="w-4 h-4 mx-auto" /></button></td>
                            {replacerMeta?.headers.map((h, i) => (
                              <td key={i} className={cn("px-4 whitespace-nowrap font-bold truncate", h === "Remarks Of NSL" && row.__isReplaced ? "text-emerald-700" : h === "Remarks Of NSL" ? "text-amber-700" : "text-slate-900")}>
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
              <div className="bg-white rounded-3xl shadow-xl border border-slate-200 overflow-hidden sticky top-28">
                <div className="bg-emerald-600 p-5 text-white flex items-center justify-between">
                  <h3 className="text-[12px] font-black uppercase tracking-widest flex items-center gap-2"><Info className="w-5 h-5" /> Replacer Matrix</h3>
                </div>
                <div className="p-6 overflow-y-auto max-h-[68vh] space-y-4 custom-scrollbar">
                  {Object.entries(REMARK_MAPPING).map(([nsl, official]) => (
                    <div key={nsl} className="p-5 bg-slate-50/80 rounded-2xl border border-slate-100 space-y-4 hover:border-emerald-300 hover:bg-white transition-all duration-300">
                      <div className="space-y-1.5">
                        <p className="text-[9px] font-black text-slate-400 uppercase tracking-[0.2em]">Original Input</p>
                        <p className="bg-amber-50/50 text-amber-700 border border-amber-200 px-3 py-2 rounded-xl text-[13px] font-bold leading-tight">{nsl}</p>
                      </div>
                      <div className="flex justify-center"><ArrowRight className="w-4 h-4 text-emerald-400" /></div>
                      <div className="space-y-1.5">
                        <p className="text-[9px] font-black text-slate-400 uppercase tracking-[0.2em]">Final Output</p>
                        <p className="bg-emerald-50 text-emerald-800 border border-emerald-200 px-3 py-2 rounded-xl text-[13px] font-black leading-tight">{official}</p>
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
