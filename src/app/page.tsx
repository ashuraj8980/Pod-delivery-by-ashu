
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
  CheckCircle2
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
  // If already in DD-MM-YYYY format, return it
  if (/^\d{2}-\d{2}-\d{4}$/.test(str)) return str;
  // If in YYYY-MM-DD format
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    const [y, m, d] = str.split('-');
    return `${d}-${m}-${y}`;
  }
  // Try JS Date parsing as fallback
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
    const saved = localStorage.getItem('pod_sessions_rpc');
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
      localStorage.setItem('pod_sessions_rpc', JSON.stringify(sessions));
    }
  }, [sessions, isMounted]);

  const showToast = useCallback((msg: string, type: 'ok' | 'err' | 'info') => {
    if (typeof document === 'undefined') return;
    const toast = document.createElement('div');
    toast.className = `fixed bottom-8 left-1/2 -translate-x-1/2 px-6 py-3 rounded-xl text-[13px] font-bold z-[500] shadow-xl transition-all duration-300 transform scale-95 opacity-0 animate-in fade-in slide-in-from-bottom-4 border ${
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

  /**
   * Universal Compatibility Clipboard Logic
   * Uses dual MIME types (plain and HTML) to protect AWB digits in WPS Office.
   */
  const copyDataProfessional = useCallback(async (rows: any[], headers: string[], isSingle = false) => {
    if (!rows.length) return;

    // 1. Build Plain Text (with apostrophe prefix for WPS plain-text paste fallback)
    const plainText = rows.map(r => 
      headers.map(h => {
        const val = String(r[h] || "").trim().replace(/[\r\n\t]+/g, " ");
        // Prefix AWB with apostrophe for spreadsheet text interpretation
        if (h.toLowerCase().includes('awb')) return `'${val}`;
        return val;
      }).join("\t")
    ).join("\n");

    // 2. Build HTML Table (Professional styling for WPS/Excel)
    const rowsHtml = rows.map(r => {
      const cells = headers.map(h => {
        const val = String(r[h] || "").trim().replace(/[\r\n\t]+/g, " ");
        // Apply mso-number-format:"\@" to treat digits as text explicitly
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
      // Fallback for browsers that don't support ClipboardItem (rare)
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
      showToast(isShortcut ? `Shortcut — Copied ${filteredRows.length} rows to WPS` : `Copied ${filteredRows.length} rows — paste in WPS with Ctrl+V`, "ok");
    }
  }, [filteredRows, currentSession, copyDataProfessional, showToast]);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Support both Ctrl+T and Ctrl+Shift+C as requested
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
      <header className="fixed top-0 left-0 w-full z-[400] bg-[#0F172A] shadow-lg">
        <div className="h-1 w-full bg-gradient-to-r from-blue-600 via-amber-400 to-emerald-500" />
        <div className="px-6 h-[64px] flex items-center justify-between text-white">
          <div className="flex items-center gap-4">
            <div className="bg-blue-600 p-2 rounded-lg shadow-lg"><Truck className="w-5 h-5 text-white" /></div>
            <div>
              <h1 className="text-[16px] font-extrabold leading-none tracking-tight">POD Tool</h1>
              <p className="text-[11px] text-slate-400 font-bold mt-1 tracking-wider uppercase">Palam Vihar RPC · <span className="text-amber-400">By Ashu</span></p>
            </div>
          </div>
          <div className="flex gap-8 h-full">
            {[ { id: "eod", label: "Daily EOD Rejection" }, { id: "remark", label: "EOD Rejection Remark" } ].map(tab => (
              <button key={tab.id} onClick={() => { setActiveTab(tab.id as any); setStatusFilter("all"); setActiveRemarkChip(null); }} className={cn("px-2 py-4 text-[13px] font-semibold transition-all relative h-full", activeTab === tab.id ? "text-white" : "text-slate-400 hover:text-white")}>
                {tab.label}
                {activeTab === tab.id && <div className="absolute bottom-0 left-0 w-full h-1 bg-amber-400" />}
              </button>
            ))}
          </div>
        </div>
      </header>

      <main className="pt-24 px-6 pb-20 max-w-[1440px] mx-auto space-y-6">
        {activeTab === "eod" ? (
          <div className="space-y-6">
            <div className="bg-white rounded-xl p-6 border-[1.5px] border-[#E2E8F0]">
              <p className="text-[12px] font-bold text-slate-500 uppercase tracking-widest mb-4 flex items-center gap-2"><Info className="w-4 h-4 text-blue-600" /> Session Setup</p>
              <div className="flex flex-col md:flex-row gap-4 mb-6">
                <div className="flex-1 space-y-1.5">
                  <label className="text-[12px] font-medium text-slate-600 px-1">DSP ID</label>
                  <input type="number" value={setupData.dspId} onChange={(e) => setSetupData({...setupData, dspId: e.target.value})} className="w-full bg-[#F9FAFB] border-[1.5px] border-[#D1D5DB] rounded-lg px-4 h-[42px] text-[14px] font-bold text-[#111827] outline-none focus:border-[#1976D2] focus:bg-white focus:ring-[3px] focus:ring-[#1976D2]/10 transition-all" placeholder="Enter DSP ID" />
                </div>
                <div className="flex-1 space-y-1.5">
                  <label className="text-[12px] font-medium text-slate-600 px-1">FE Name</label>
                  <input type="text" value={setupData.feName} onChange={(e) => setSetupData({...setupData, feName: e.target.value})} className="w-full bg-[#F9FAFB] border-[1.5px] border-[#D1D5DB] rounded-lg px-4 h-[42px] text-[14px] font-bold text-[#111827] outline-none focus:border-[#1976D2] focus:bg-white focus:ring-[3px] focus:ring-[#1976D2]/10 transition-all" placeholder="Enter FE Name" />
                </div>
                <div className="flex-[0.5] space-y-1.5">
                  <label className="text-[12px] font-medium text-slate-600 px-1">Date</label>
                  <input type="date" value={setupData.date} onChange={(e) => setSetupData({...setupData, date: e.target.value})} className="w-full bg-[#F9FAFB] border-[1.5px] border-[#D1D5DB] rounded-lg px-4 h-[42px] text-[14px] font-bold text-[#111827] outline-none focus:border-[#1976D2] focus:bg-white focus:ring-[3px] focus:ring-[#1976D2]/10 transition-all" />
                </div>
              </div>
              <div className={cn("border-[2px] border-dashed rounded-xl p-10 text-center transition-all cursor-pointer relative group", uploadError ? "border-rose-300 bg-rose-50" : "border-slate-300 bg-slate-50 hover:border-blue-500 hover:bg-white", isProcessing && "opacity-80 pointer-events-none")}>
                <input type="file" disabled={isProcessing} onChange={handleFileUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                {isProcessing ? (
                  <div className="flex flex-col items-center gap-3"><Loader2 className="w-8 h-8 text-blue-600 animate-spin" /><p className="text-[14px] font-bold text-blue-600">Processing file...</p></div>
                ) : uploadError ? (
                  <div className="space-y-3"><AlertCircle className="w-8 h-8 text-rose-500 mx-auto" /><p className="text-[14px] font-bold text-rose-600">{uploadError}</p><button onClick={() => setUploadError(null)} className="bg-rose-600 text-white px-5 py-2 rounded-lg text-[13px] font-bold uppercase">Try Again</button></div>
                ) : (
                  <div className="space-y-2"><Download className="w-8 h-8 text-slate-400 mx-auto group-hover:scale-110 transition-transform" /><p className="text-[14px] font-bold text-slate-800">Drop Delhivery CSV here, or click to upload</p><p className="text-[11px] text-slate-500 font-semibold">Ready for processing</p></div>
                )}
              </div>
            </div>

            {sessions.length > 0 && (
              <div className="space-y-3">
                <div className="flex items-center justify-between px-1">
                  <p className="text-[12px] font-bold text-slate-500 uppercase tracking-widest">Active Sessions</p>
                  <button 
                    onClick={() => { if(confirm("Clear all saved sessions?")) { setSessions([]); setSelectedSessionId(null); } }} 
                    className="text-[11px] font-bold text-rose-500 hover:underline transition-all"
                  >
                    Clear All Sessions
                  </button>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                  {sessions.map(s => {
                    const sTotal = s.data.length;
                    const sPending = s.data.filter(r => r.status === 'pending').length;
                    const sRto = s.data.filter(r => r.status === 'rto').length;
                    const sDto = s.data.filter(r => r.status === 'dto' || r.status === 'delivered').length;
                    
                    return (
                      <div key={s.id} onClick={() => { setSelectedSessionId(s.id); setStatusFilter('all'); setActiveRemarkChip(null); }} className={cn("bg-white p-4 rounded-xl border-l-[4px] cursor-pointer relative transition-all group border-[1.5px] border-[#E2E8F0]", selectedSessionId === s.id ? "border-l-blue-600 ring-4 ring-blue-500/5" : "border-l-slate-300 opacity-70 hover:opacity-100")}>
                        <button onClick={(e) => { e.stopPropagation(); setSessions(prev => prev.filter(x => x.id !== s.id)); if(selectedSessionId === s.id) setSelectedSessionId(null); }} className="absolute top-3 right-3 text-slate-300 hover:text-rose-500 transition-colors opacity-0 group-hover:opacity-100"><X className="w-4 h-4" /></button>
                        <p className="text-[15px] font-bold text-slate-900 truncate pr-8">{s.feName}</p>
                        <p className="text-[12px] font-medium text-slate-500 mt-1 uppercase tracking-tight">{s.dspId} — {formatDate(s.date)}</p>
                        <div className="mt-3 flex flex-wrap gap-1.5">
                          <span className="px-1.5 py-0.5 bg-slate-100 rounded text-[10px] font-bold text-slate-600 uppercase tracking-tight whitespace-nowrap">{sTotal} pkt</span>
                          <span className="px-1.5 py-0.5 bg-amber-100 rounded text-[10px] font-bold text-amber-700 uppercase tracking-tight whitespace-nowrap">{sPending} pending</span>
                          <span className="px-1.5 py-0.5 bg-rose-100 rounded text-[10px] font-bold text-rose-700 uppercase tracking-tight whitespace-nowrap">{sRto} RTO</span>
                          <span className="px-1.5 py-0.5 bg-emerald-100 rounded text-[10px] font-bold text-emerald-700 uppercase tracking-tight whitespace-nowrap">{sDto} DTO</span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {currentSession && (
              <div className="space-y-6">
                <div className="bg-white rounded-xl border-[1.5px] border-[#E2E8F0] overflow-hidden flex h-[80px]">
                  {[
                    { id: 'all', label: 'All', val: stats.total, color: 'text-blue-700' },
                    { id: 'pending', label: 'Pending', val: stats.pending, color: 'text-amber-600' },
                    { id: 'dispatched', label: 'Dispatch', val: stats.dispatched, color: 'text-blue-700' },
                    { id: 'rto', label: 'RTO', val: stats.rto, color: 'text-rose-600' },
                    { id: 'dto', label: 'DTO', val: stats.dto, color: 'text-emerald-600' }
                  ].map((t, i) => (
                    <button key={t.id} onClick={() => { setStatusFilter(t.id); setActiveRemarkChip(null); }} className={cn("flex-1 flex flex-col items-center justify-center transition-all border-b-[4px] border-r", i === 4 && "border-r-0", statusFilter === t.id ? "bg-white border-blue-600" : "bg-slate-50/30 border-transparent hover:bg-white")}>
                      <span className={cn("text-[28px] font-extrabold leading-none mb-1", t.color)}>{t.val}</span>
                      <span className={cn("text-[11px] font-bold uppercase tracking-widest", statusFilter === t.id ? t.color : "text-slate-400")}>{t.label}</span>
                    </button>
                  ))}
                </div>

                {sortedRemarks.length > 0 && (
                  <div className="bg-white rounded-xl p-6 border-[1.5px] border-[#E2E8F0] shadow-sm">
                    <div className="flex items-center justify-between mb-6">
                      <div className="space-y-1">
                        <h3 className="text-[12px] font-bold text-slate-800 flex items-center gap-2">
                          <BarChart3 className="w-4 h-4 text-blue-600" />
                          Remark Breakdown — {statusFilter === 'all' ? 'All' : statusFilter === 'rto' ? 'RTO' : statusFilter === 'dto' ? 'DTO' : statusFilter.charAt(0).toUpperCase() + statusFilter.slice(1)}
                        </h3>
                        <p className="text-[10px] font-medium text-slate-400 uppercase tracking-tight">Click any remark chip to filter</p>
                      </div>
                      {activeRemarkChip && (
                        <button 
                          onClick={() => setActiveRemarkChip(null)} 
                          className="flex items-center gap-2 px-3 py-1 bg-slate-900 text-white rounded-lg text-[10px] font-bold uppercase transition-all active:scale-95"
                        >
                          All {statusFilter === 'all' ? '' : statusFilter.toUpperCase()}
                        </button>
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
                              "group px-4 py-2 rounded-lg flex items-center gap-3 text-[13px] font-semibold transition-all border-[1.5px]",
                              isActive 
                                ? "bg-blue-600 text-white border-blue-700 shadow-md scale-[1.02]" 
                                : isRed 
                                  ? "bg-rose-50 text-rose-700 border-rose-200 hover:bg-rose-100" 
                                  : "bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100 hover:border-slate-300"
                            )}
                          >
                            <span className="truncate max-w-[280px]">{rem}</span>
                            <span className={cn(
                              "min-w-[22px] h-[22px] rounded-md flex items-center justify-center text-[10px] font-extrabold shadow-sm",
                              isActive ? "bg-white text-blue-600" : isRed ? "bg-rose-600 text-white" : "bg-white text-slate-600"
                            )}>
                              {count}
                            </span>
                          </button>
                        );
                      })}
                    </div>
                  </div>
                )}

                <div className="flex items-center justify-between gap-4">
                  <div className="flex gap-3">
                    <button onClick={downloadExcel} className="h-10 px-6 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-[13px] font-bold flex items-center gap-2 shadow-sm transition-all"><Download className="w-4 h-4" /> Download Excel</button>
                    <button onClick={() => handleCopyTable(false)} className="h-10 px-6 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-[13px] font-bold flex items-center gap-2 shadow-sm transition-all"><Copy className="w-4 h-4" /> Copy Table</button>
                  </div>
                  <div className="relative"><Search className="w-4 h-4 absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" /><input type="text" placeholder="Search by AWB / Client..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="bg-white border-[1.5px] border-[#D1D5DB] rounded-lg pl-10 pr-4 h-10 text-[14px] font-semibold outline-none w-[300px] focus:border-[#1976D2] transition-all" /></div>
                </div>

                <div className="bg-white rounded-xl overflow-hidden border-[1.5px] border-[#F97316]">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse table-fixed">
                      <thead className="bg-[#0F172A] text-white">
                        <tr className="h-10">
                          <th style={{ width: '32px' }} className="px-2 text-center"><input type="checkbox" className="w-3.5 h-3.5" /></th>
                          <th style={{ width: '28px' }} className="px-2 text-center"><Trash2 className="w-3.5 h-3.5 mx-auto" /></th>
                          <th style={{ width: '80px' }} className="px-2 text-[11px] font-bold">DSP ID</th>
                          <th style={{ width: '130px' }} className="px-2 text-[11px] font-bold">AWB Number</th>
                          <th style={{ width: '110px' }} className="px-2 text-[11px] font-bold">Client</th>
                          <th style={{ width: '110px' }} className="px-2 text-[11px] font-bold">Order ID</th>
                          <th className="px-2 text-[11px] font-bold">Remark</th>
                          <th style={{ width: '80px' }} className="px-2 text-[11px] font-bold">FE Name</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr className="bg-slate-800 text-white h-8">
                          <td colSpan={8} className="px-3">
                            <div className="flex items-center gap-2">
                              <span className="text-[12px] font-bold text-amber-400">{currentSession.dspId}</span>
                              <span className="px-1.5 py-0.5 bg-blue-600 text-white rounded text-[10px] font-bold">{filteredRows.length} pkt</span>
                              <span className="text-[10px] text-slate-400 font-bold ml-auto">{currentSession.feName} · {formatDate(currentSession.date)}</span>
                            </div>
                          </td>
                        </tr>
                        {filteredRows.map((row) => (
                          <tr key={row.id} className={cn("h-10 border-b border-[#FED7AA] transition-colors group", row.isIntact ? "bg-rose-50/50" : "hover:bg-blue-50/30")}>
                            <td className="px-2 text-center"><input type="checkbox" checked={row.selected} onChange={() => {}} className="w-3.5 h-3.5" /></td>
                            <td className="px-2 text-center"><button onClick={() => setSessions(prev => prev.map(s => s.id === selectedSessionId ? {...s, data: s.data.filter(r => r.id !== row.id)} : s))} className="text-slate-300 hover:text-rose-500 transition-colors"><Trash2 className="w-3.5 h-3.5 mx-auto" /></button></td>
                            <td className="px-2 text-[11px] font-bold text-slate-900 truncate">{row.dspId}</td>
                            <td 
                              onClick={async () => {
                                // Single AWB Click precision fix for WPS
                                const val = String(row.awb);
                                const plainText = `'${val}`;
                                const htmlTable = `<html><body><table><tr><td style='mso-number-format:"\\@"'>${val}</td></tr></table></body></html>`;
                                const htmlBlob = new Blob([htmlTable], { type: 'text/html' });
                                const textBlob = new Blob([plainText], { type: 'text/plain' });
                                await navigator.clipboard.write([new ClipboardItem({ 'text/html': htmlBlob, 'text/plain': textBlob })]);
                                showToast(`AWB ${val} copied`, 'ok');
                              }}
                              className="px-2 text-[12px] font-bold text-blue-700 cursor-pointer hover:underline tracking-tighter truncate" 
                              style={{ fontFamily: '"IBM Plex Mono", monospace' }}
                            >
                              {row.awb}
                            </td>
                            <td className="px-2 text-[11px] font-bold text-slate-800 truncate">{row.client}</td>
                            <td className="px-2 text-[11px] font-bold text-slate-800 truncate">{row.orderId}</td>
                            <td className="px-2">
                              <span className={cn(
                                "px-2 py-0.5 rounded-lg text-[11px] font-semibold border whitespace-nowrap overflow-hidden text-ellipsis max-w-[200px] inline-block", 
                                row.isIntact ? "bg-rose-50 text-rose-700 border-rose-100" : "bg-amber-50 text-amber-700 border-amber-100"
                              )}>
                                {row.remark || "No Remark"}
                              </span>
                            </td>
                            <td className="px-2 text-[11px] font-bold text-slate-500 truncate">{row.feName}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}
          </div>
        ) : (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
            <div className="lg:col-span-8 space-y-6">
              <div className="bg-white rounded-xl p-6 border-[1.5px] border-[#E2E8F0] shadow-sm">
                <div className="flex items-center justify-between mb-6">
                  <h2 className="text-[16px] font-extrabold text-slate-900 flex items-center gap-2"><FileSpreadsheet className="w-5 h-5 text-emerald-600" /> Remark Replacer</h2>
                  {replacerData.length > 0 && (
                    <div className="flex gap-2">
                      <div className="px-3 py-1 bg-slate-100 rounded-lg text-[11px] font-bold text-slate-600">Total: {replacerStats.total}</div>
                      <div className="px-3 py-1 bg-emerald-50 rounded-lg text-[11px] font-bold text-emerald-700">Replaced: {replacerStats.replaced}</div>
                      <div className="px-3 py-1 bg-amber-50 rounded-lg text-[11px] font-bold text-amber-700">Original: {replacerStats.noMapping}</div>
                    </div>
                  )}
                </div>
                <div className={cn("border-[2px] border-dashed rounded-xl p-10 text-center transition-all cursor-pointer relative group", uploadError ? "border-rose-300 bg-rose-50" : "border-slate-300 bg-slate-50 hover:border-blue-500 hover:bg-white", isProcessing && "opacity-80 pointer-events-none")}>
                  <input type="file" disabled={isProcessing} onChange={handleReplacerUpload} className="absolute inset-0 opacity-0 cursor-pointer z-20" />
                  {isProcessing ? (
                    <div className="flex flex-col items-center gap-3"><Loader2 className="w-8 h-8 text-blue-600 animate-spin" /><p className="text-[14px] font-bold text-blue-600">Processing file...</p></div>
                  ) : uploadError ? (
                    <div className="space-y-3"><AlertCircle className="w-8 h-8 text-rose-500 mx-auto" /><p className="text-[14px] font-bold text-rose-600">{uploadError}</p><button onClick={() => setUploadError(null)} className="bg-rose-600 text-white px-5 py-2 rounded-lg text-[13px] font-bold uppercase mt-2">Try Again</button></div>
                  ) : (
                    <div className="space-y-2"><Download className="w-8 h-8 text-slate-400 mx-auto group-hover:scale-110 transition-transform" /><p className="text-[14px] font-bold text-slate-800">Drop EOD Sheet here, or click to upload</p><p className="text-[11px] text-slate-500 font-semibold">Mapping active</p></div>
                  )}
                </div>
              </div>
              {replacerData.length > 0 && (
                <div className="bg-white rounded-xl shadow-sm border-[1.5px] border-[#F97316] overflow-hidden">
                  <div className="p-4 bg-slate-50 border-b flex items-center justify-between">
                    <div className="flex gap-3">
                      <button onClick={downloadOfficialExcel} className="h-10 px-6 bg-slate-900 text-white rounded-lg text-[13px] font-bold flex items-center gap-2 transition-all active:scale-95"><Download className="w-4 h-4" /> Download Excel</button>
                      <button onClick={() => copyDataProfessional(replacerData, replacerMeta?.headers || [])} className="h-10 px-6 bg-blue-600 text-white rounded-lg text-[13px] font-bold flex items-center gap-2 transition-all active:scale-95"><Copy className="w-4 h-4" /> Copy Data</button>
                    </div>
                    <button onClick={() => { setReplacerData([]); setReplacerMeta(null); }} className="text-[11px] font-bold text-rose-500 px-4 hover:underline">Clear</button>
                  </div>
                  <div className="overflow-x-auto max-h-[700px] custom-scrollbar">
                    <table className="w-full text-left border-collapse table-fixed text-[11px]">
                      <thead className="sticky top-0 z-30 bg-[#0F172A] text-white">
                        <tr className="h-10">
                          <th style={{ width: '40px' }} className="px-2 text-center"><Trash2 className="w-3.5 h-3.5 mx-auto" /></th>
                          {replacerMeta?.headers.map((h, i) => <th key={i} className="px-2 font-bold whitespace-nowrap text-[10px]">{h}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {replacerData.map((row) => (
                          <tr key={row.__id} className={cn("h-10 border-b border-[#FED7AA] transition-colors", row.__isReplaced ? "bg-[#F0FDF4]" : "bg-[#FFFDE7]")}>
                            <td className="px-2 text-center"><button onClick={() => setReplacerData(prev => prev.filter(r => r.__id !== row.__id))} className="text-slate-300 hover:text-rose-500 transition-colors"><Trash2 className="w-3.5 h-3.5 mx-auto" /></button></td>
                            {replacerMeta?.headers.map((h, i) => (
                              <td key={i} className={cn("px-2 whitespace-nowrap font-bold truncate", h === "Remarks Of NSL" && row.__isReplaced ? "text-emerald-700" : h === "Remarks Of NSL" ? "text-amber-700" : "text-slate-800")}>
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
              <div className="bg-white rounded-xl shadow-sm border-[1.5px] border-[#E2E8F0] overflow-hidden sticky top-24">
                <div className="bg-emerald-600 p-4 text-white">
                  <h3 className="text-[11px] font-extrabold uppercase tracking-widest flex items-center gap-2"><Info className="w-4 h-4" /> Remark Mapping Guide</h3>
                </div>
                <div className="p-4 overflow-y-auto max-h-[70vh] space-y-3 custom-scrollbar">
                  {Object.entries(REMARK_MAPPING).map(([nsl, official]) => (
                    <div key={nsl} className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3 hover:border-emerald-300 transition-all">
                      <div className="space-y-1"><p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Input Remark</p><p className="bg-amber-50 text-amber-700 border border-amber-100 px-3 py-1.5 rounded-lg text-[12px] font-bold leading-tight">{nsl}</p></div>
                      <div className="flex justify-center"><ArrowRight className="w-3 h-3 text-emerald-500" /></div>
                      <div className="space-y-1"><p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Auto Replaced</p><p className="bg-emerald-50 text-emerald-700 border border-emerald-100 px-3 py-1.5 rounded-lg text-[12px] font-bold leading-tight">{official}</p></div>
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
