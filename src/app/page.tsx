
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
 * SESSION LOGIC: Uniqueness based on DSP ID only.
 * RESET ON RELOAD: No persistent storage.
 * CLIPBOARD: Dual-format (HTML+Text) for WPS precision without visible quote prefixes.
 * SHORTCUT: Ctrl+T for instant copy.
 * UI: High-Contrast HD Clarity Mode.
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
    toast.className = `fixed bottom-10 left-1/2 -translate-x-1/2 px-7 py-3.5 rounded-2xl text-[13px] font-[900] z-[500] shadow-2xl transition-all duration-300 transform scale-95 opacity-0 animate-in fade-in slide-in-from-bottom-5 border ${
      type === 'ok' ? 'bg-[#052E0F] text-[#6EE7A6] border-[#6EE7A6]/20' : 
      type === 'err' ? 'bg-[#2D0808] text-[#FCA5A5] border-[#FCA5A5]/20' : 
      'bg-[#1C2333] text-white border-white/20'
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
    if (/^[\d.]+[eE][+\-]?\d+$/.test(str)) {
      try {
        str = BigInt(Math.round(Number(val))).toString();
      } catch (e) {
        str = String(val);
      }
    }
    return str.replace(/\.0$/, "");
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

          const existingSession = sessions.find(s => s.dspId === setupData.dspId);
          
          if (existingSession) {
            setSessions(prev => prev.map(s => 
              s.dspId === setupData.dspId 
              ? { ...s, feName: setupData.feName, data: parsedRows, date: setupData.date, timestamp: Date.now() } 
              : s
            ));
            setSelectedSessionId(existingSession.id);
            showToast(`Session updated for DSP ${setupData.dspId}`, "ok");
          } else {
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
          const wb = XLSX.read(data, { type: 'array', cellText: true, cellDates: true });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rawData = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
          
          if (!Array.isArray(rawData) || rawData.length === 0) throw new Error("Empty file.");

          const allHeaders = Object.keys(rawData[0]);
          const remarkKey = allHeaders.find(h => /Remarks Of NSL/i.test(h.trim())) || "";
          const awbKey = allHeaders.find(h => /Awb/i.test(h.trim())) || "";
          const dspKey = allHeaders.find(h => /DSP No/i.test(h.trim())) || "";

          if (!remarkKey) throw new Error("Remarks Of NSL column not found.");

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

            const cleanAwb = fixValueToString(row[awbKey]);
            const cleanDsp = fixValueToString(row[dspKey]);
            
            const cleanRow: any = { __isReplaced: isReplaced, __id: crypto.randomUUID() };
            targetHeaders.forEach(h => {
              if (h === "Remarks Of NSL") cleanRow[h] = finalRemark;
              else if (h === "Awb") cleanRow[h] = cleanAwb;
              else if (h === "DSP No") cleanRow[h] = cleanDsp;
              else {
                const inputKey = allHeaders.find(key => key.trim().toLowerCase() === h.toLowerCase());
                cleanRow[h] = inputKey ? row[inputKey] : "";
              }
            });
            return cleanRow;
          });

          setReplacerData(processed);
          setReplacerMeta({ headers: targetHeaders, remarkKey, awbKey, dspKey });
          setReplacerStats({ total: processed.length, replaced: replacedCount, noMapping: noMappingCount });
          showToast(`Processed ${processed.length} rows`, "ok");
        } catch (err: any) {
          setUploadError(err.message);
          showToast("Upload Failed", "err");
        } finally {
          setIsProcessing(false);
        }
      }, 50);
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

  const copySingleAwb = useCallback(async (awb: string) => {
    const cleanAwb = String(awb).trim();
    const htmlTable = `<html><head><meta charset="utf-8"><style>td{mso-number-format:"\\@";}</style></head><body><table><tr><td>${cleanAwb}</td></tr></table></body></html>`;
    try {
      const textBlob = new Blob([cleanAwb], { type: 'text/plain' });
      const htmlBlob = new Blob([htmlTable], { type: 'text/html' });
      const data = [new ClipboardItem({ 'text/plain': textBlob, 'text/html': htmlBlob })];
      await navigator.clipboard.write(data);
      showToast(`AWB ${cleanAwb} Copied (Precise)`, 'ok');
    } catch (err) {
      await navigator.clipboard.writeText(cleanAwb);
      showToast(`AWB ${cleanAwb} Copied`, 'ok');
    }
  }, [showToast]);

  const copyTableToClipboard = useCallback(async (isShortcut = false) => {
    if (!filteredRows.length || !currentSession) return;
    const cleanValue = (val: any) => String(val || "").replace(/[\r\n]+/g, " ").trim();
    const clipboardText = filteredRows.map((r, i) => {
      const dspVal = i === 0 ? currentSession.dspId : "";
      return [cleanValue(r.date), cleanValue(dspVal), cleanValue(r.awb), cleanValue(r.client), cleanValue(r.orderId), cleanValue(r.remark), cleanValue(r.feName)].join("\t");
    }).join("\n");

    const rowsHtml = filteredRows.map((r, i) => {
      const dspVal = i === 0 ? currentSession.dspId : "";
      return `<tr><td>${cleanValue(r.date)}</td><td>${cleanValue(dspVal)}</td><td>${cleanValue(r.awb)}</td><td>${cleanValue(r.client)}</td><td>${cleanValue(r.orderId)}</td><td>${cleanValue(r.remark)}</td><td>${cleanValue(r.feName)}</td></tr>`;
    }).join("");
    const htmlTable = `<html><head><meta charset="utf-8"><style>td{mso-number-format:"\\@";white-space:nowrap;vertical-align:middle;padding:2px;font-family:Calibri,sans-serif;font-size:11pt;}</style></head><body><table>${rowsHtml}</table></body></html>`;

    try {
      const textBlob = new Blob([clipboardText], { type: 'text/plain' });
      const htmlBlob = new Blob([htmlTable], { type: 'text/html' });
      const data = [new ClipboardItem({ 'text/plain': textBlob, 'text/html': htmlBlob })];
      await navigator.clipboard.write(data);
      showToast(isShortcut ? `Ctrl+T — Copied ${filteredRows.length} rows` : `Copied ${filteredRows.length} rows — paste in WPS`, "ok");
    } catch (err) {
      showToast("Copy failed!", "err");
    }
  }, [filteredRows, currentSession, showToast]);

  const copyRemarkDataToClipboard = async () => {
    if (!replacerData.length || !replacerMeta) return;
    const headers = replacerMeta.headers;
    const cleanValue = (val: any) => String(val || "").replace(/[\r\n]+/g, " ").trim();
    const clipboardText = replacerData.map(r => headers.map(h => cleanValue(r[h])).join("\t")).join("\n");
    const rowsHtml = replacerData.map(r => `<tr>${headers.map(h => `<td>${cleanValue(r[h])}</td>`).join("")}</tr>`).join("");
    const htmlTable = `<html><head><meta charset="utf-8"><style>td{mso-number-format:"\\@";white-space:nowrap;vertical-align:middle;padding:2px;font-family:Calibri,sans-serif;font-size:11pt;}</style></head><body><table>${rowsHtml}</table></body></html>`;
    try {
      const textBlob = new Blob([clipboardText], { type: 'text/plain' });
      const htmlBlob = new Blob([htmlTable], { type: 'text/html' });
      const data = [new ClipboardItem({ 'text/plain': textBlob, 'text/html': htmlBlob })];
      await navigator.clipboard.write(data);
      showToast(`Copied ${replacerData.length} rows (Text Precision Mode)`, "ok");
    } catch (err) {
      showToast("Copy failed!", "err");
    }
  };

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.ctrlKey && (e.key === 't' || e.key === 'T')) {
        e.preventDefault(); 
        copyTableToClipboard(true);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [copyTableToClipboard]);

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

  const downloadExcel = () => {
    if (!filteredRows.length || !currentSession) return;
    const header = ['Date', 'DSP ID', 'AWB Number', 'Client', 'Order ID', 'Remark', 'FE Name'];
    const excelData = filteredRows.map((r) => [
      r.date, 
      { v: String(currentSession.dspId), t: 's', z: '@' }, 
      { v: String(r.awb), t: 's', z: '@' }, 
      r.client, 
      r.orderId, 
      r.remark, 
      r.feName
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...excelData]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, `POD_${currentSession.dspId}_${statusFilter}.xlsx`);
    showToast(`Downloaded ${filteredRows.length} rows`, "ok");
  };

  const downloadOfficialExcel = () => {
    if (!replacerData.length || !replacerMeta) return;
    const exportData = replacerData.map(r => {
      const row: any = {};
      replacerMeta.headers.forEach(h => {
        let val = r[h];
        if (h === "Awb" || h === "DSP No") row[h] = { v: String(val), t: 's', z: '@' };
        else row[h] = val;
      });
      return row;
    });
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Remarks");
    XLSX.writeFile(wb, `EOD_Official_Remarks_${new Date().toISOString().split('T')[0]}.xlsx`);
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
    <div className="min-h-screen bg-[#F4F7FA] font-body text-slate-900 antialiased">
      <div className="fixed top-0 left-0 w-full z-[400] shadow-2xl">
        <div className="h-[4px] w-full bg-gradient-to-r from-blue-700 via-amber-500 via-green-600 to-red-600" />
        <header className="h-[64px] bg-[#0F172A] px-8 flex items-center justify-between text-white">
          <div className="flex items-center gap-4">
            <div className="bg-blue-600 p-2 rounded-2xl shadow-xl"><Truck className="w-6 h-6 text-white" /></div>
            <div className="flex flex-col">
              <h1 className="text-[17px] font-[900] tracking-tight uppercase leading-none">POD Management Tool</h1>
              <p className="text-[10px] text-slate-400 font-bold mt-1.5 uppercase tracking-widest">Palam Vihar RPC · <span className="text-amber-400 font-black italic">By Ashu</span></p>
            </div>
          </div>
          <div className="flex items-center gap-6">
            <div className="flex items-center gap-2.5 bg-green-500/20 px-4 py-1.5 rounded-full border border-green-500/30">
              <div className="w-2 h-2 bg-green-500 rounded-full animate-pulse shadow-[0_0_10px_#22c55e]" />
              <span className="text-[10px] font-[900] text-green-400 uppercase tracking-tighter">HD LIVE</span>
            </div>
            {currentSession && <div className="bg-amber-500/20 px-4 py-1.5 rounded-full text-[10px] font-mono font-black text-amber-500 uppercase tracking-widest">{currentSession.data.length} ROWS</div>}
          </div>
        </header>
        <nav className="bg-[#0F172A] px-8 flex gap-10 border-t border-white/5">
          {[ { id: "eod", label: "Daily EOD Rejection" }, { id: "remark", label: "EOD Rejection Remark" } ].map(tab => (
            <button key={tab.id} onClick={() => { setActiveTab(tab.id as any); setStatusFilter("all"); }} className={cn("py-4 text-[12px] font-[900] uppercase tracking-[0.2em] transition-all relative", activeTab === tab.id ? "text-white" : "text-slate-400 hover:text-white")}>
              {tab.label}
              {activeTab === tab.id && <div className="absolute bottom-0 left-0 w-full h-[4px] bg-amber-500 rounded-t-full shadow-[0_-2px_15px_#f59e0b]" />}
            </button>
          ))}
        </nav>
      </div>

      <main className="pt-[150px] p-8 max-w-[1500px] mx-auto space-y-8 pb-20">
        {activeTab === "eod" ? (
          <div className="space-y-8 animate-in fade-in duration-700">
            <div className="bg-white rounded-[24px] p-8 shadow-xl border border-slate-200">
              <div className="text-[11px] font-[900] text-slate-500 uppercase tracking-[0.2em] mb-6 flex items-center gap-2">
                <Info className="w-4 h-4 text-blue-600" /> SESSION SETUP
              </div>
              <div className="flex flex-col md:flex-row gap-5 mb-8">
                <input type="number" value={setupData.dspId} onChange={(e) => setSetupData({...setupData, dspId: e.target.value})} className="flex-1 bg-slate-50 border-2 border-slate-100 rounded-2xl px-5 py-3.5 text-[14px] font-black outline-none focus:border-blue-500 transition-all shadow-inner" placeholder="Enter DSP ID" />
                <input type="text" value={setupData.feName} onChange={(e) => setSetupData({...setupData, feName: e.target.value})} className="flex-1 bg-slate-50 border-2 border-slate-100 rounded-2xl px-5 py-3.5 text-[14px] font-black outline-none focus:border-blue-500 transition-all shadow-inner" placeholder="Enter FE Name" />
                <input type="date" value={setupData.date} onChange={(e) => setSetupData({...setupData, date: e.target.value})} className="flex-[0.5] bg-slate-50 border-2 border-slate-100 rounded-2xl px-5 py-3.5 text-[14px] font-black outline-none focus:border-blue-500 transition-all shadow-inner" />
              </div>
              
              <div className={cn("border-[3px] border-dashed rounded-[30px] p-16 text-center transition-all cursor-pointer relative group", uploadError ? "border-red-500 bg-red-50/50" : "border-slate-200 bg-slate-50/50 hover:border-blue-600 hover:bg-white", isProcessing && "opacity-80 pointer-events-none")}>
                <input type="file" disabled={isProcessing} onChange={handleFileUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                {isProcessing ? (
                  <div className="flex flex-col items-center gap-4"><Loader2 className="w-12 h-12 text-blue-600 animate-spin" /><p className="text-[13px] font-[900] uppercase text-blue-600 tracking-[0.3em]">Analyzing Session...</p></div>
                ) : uploadError ? (
                  <div className="space-y-4"><AlertCircle className="w-12 h-12 text-red-500 mx-auto" /><p className="text-[16px] font-[900] text-red-600 uppercase">Invalid Session File</p><button onClick={() => setUploadError(null)} className="bg-red-600 text-white px-6 py-2 rounded-xl text-[11px] font-black uppercase shadow-lg">Try Again</button></div>
                ) : (
                  <div className="space-y-3"><Download className="w-12 h-12 text-blue-600/30 mx-auto group-hover:scale-110 transition-transform" /><p className="text-[17px] font-[900] text-slate-900">Drop Delhivery CSV here, or click to upload</p><p className="text-[11px] text-slate-500 uppercase tracking-widest font-bold">HD PRECISION MODE ENABLED</p></div>
                )}
              </div>
            </div>

            {sessions.length > 0 && (
              <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-4 gap-5">
                {sessions.map(s => (
                  <div key={s.id} onClick={() => setSelectedSessionId(s.id)} className={cn("bg-white p-5 rounded-[22px] border-l-[6px] shadow-lg cursor-pointer relative transition-all group", selectedSessionId === s.id ? "border-l-blue-600 ring-2 ring-blue-600/10 scale-[1.02]" : "border-l-slate-300 border border-slate-200 opacity-70 hover:opacity-100")}>
                    <button onClick={(e) => { e.stopPropagation(); setSessions(prev => prev.filter(x => x.id !== s.id)); if(selectedSessionId === s.id) setSelectedSessionId(null); }} className="absolute top-4 right-4 text-slate-300 hover:text-red-500 transition-colors opacity-0 group-hover:opacity-100"><X className="w-5 h-5" /></button>
                    <p className="text-[16px] font-[900] text-slate-900 uppercase truncate pr-8">{s.feName}</p>
                    <p className="text-[11px] font-mono font-black text-slate-500 mt-2 uppercase">{s.dspId} — {s.date}</p>
                    <div className="mt-4"><span className="px-2.5 py-1 rounded-lg bg-slate-100 text-[10px] font-[900] text-slate-600 uppercase">{s.data.length} TOTAL ROWS</span></div>
                  </div>
                ))}
              </div>
            )}

            {currentSession && (
              <div className="space-y-8 animate-in slide-in-from-bottom-5 duration-700">
                <div className="bg-white rounded-[24px] shadow-xl border border-slate-200 overflow-hidden flex h-[110px]">
                  {[
                    { id: 'all', label: 'All', val: stats.total, color: 'text-blue-700', activeBg: 'bg-white', activeBorder: 'border-blue-700' },
                    { id: 'pending', label: 'Pending', val: stats.pending, color: 'text-amber-600', activeBg: 'bg-amber-50/50', activeBorder: 'border-amber-500' },
                    { id: 'dispatched', label: 'Dispatch', val: stats.dispatched, color: 'text-blue-700', activeBg: 'bg-white', activeBorder: 'border-blue-700' },
                    { id: 'rto', label: 'RTO', val: stats.rto, color: 'text-red-600', activeBg: 'bg-red-50/50', activeBorder: 'border-red-500' },
                    { id: 'dto', label: 'DTO', val: stats.dto, color: 'text-green-600', activeBg: 'bg-green-50/50', activeBorder: 'border-green-500' }
                  ].map((t, i) => (
                    <button key={t.id} onClick={() => { setStatusFilter(t.id); setActiveRemarkChip(null); }} className={cn("flex-1 flex flex-col items-center justify-center transition-all border-b-[6px] border-r relative", i === 4 && "border-r-0", statusFilter === t.id ? `${t.activeBg} ${t.activeBorder}` : "bg-white border-transparent")}>
                      <span className={cn("text-[32px] font-[900] leading-none mb-1.5 tabular-nums", t.color)}>{t.val}</span>
                      <span className={cn("text-[10px] font-[900] uppercase tracking-[0.2em]", statusFilter === t.id ? t.color : "text-slate-400")}>{t.label}</span>
                    </button>
                  ))}
                </div>

                <div className="flex items-center justify-between px-2">
                  <div className="flex gap-4">
                    <button onClick={downloadExcel} className="bg-green-700 hover:bg-green-800 text-white px-7 py-3 rounded-2xl font-[900] text-[13px] uppercase tracking-widest flex items-center gap-3 shadow-xl transition-all active:scale-95"><Download className="w-5 h-5" /> Download Excel</button>
                    <button onClick={() => copyTableToClipboard(false)} className="bg-blue-700 hover:bg-blue-800 text-white px-7 py-3 rounded-2xl font-[900] text-[13px] uppercase tracking-widest flex items-center gap-3 shadow-xl transition-all active:scale-95"><Copy className="w-5 h-5" /> Copy Table</button>
                  </div>
                  <button onClick={() => setSessions([])} className="bg-red-500/10 hover:bg-red-500 text-red-600 hover:text-white px-5 py-2.5 rounded-xl text-[11px] font-[900] uppercase tracking-widest transition-all border border-red-500/20">Reset All Sessions</button>
                </div>

                <div className="bg-white rounded-[26px] shadow-2xl border border-slate-200 overflow-hidden">
                  <div className="p-5 border-b bg-slate-50/80 flex items-center justify-between">
                    <div className="flex items-center gap-4"><span className="text-[12px] font-[900] uppercase text-slate-900 tracking-wider">{statusFilter.toUpperCase()} SHIPMENTS</span><span className="px-3 py-1 bg-slate-200 rounded-lg text-[10px] font-black uppercase text-slate-700">{filteredRows.length} RECORDS FOUND</span></div>
                    <div className="relative"><Search className="w-4 h-4 absolute left-4 top-1/2 -translate-y-1/2 text-slate-400" /><input type="text" placeholder="Search by AWB / Order ID..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="bg-white border-2 border-slate-100 rounded-full pl-12 pr-6 py-2.5 text-[12px] font-[900] outline-none w-[320px] focus:border-blue-500 shadow-inner" /></div>
                  </div>
                  <div className="overflow-x-auto max-h-[800px] custom-scrollbar">
                    <table className="w-full text-left border-collapse">
                      <thead className="sticky top-0 z-30">
                        <tr className="bg-[#0F172A] text-white">
                          <th className="p-5 w-[60px] text-center"></th>
                          <th className="p-5 w-[60px] text-center"><X className="w-4 h-4 mx-auto" /></th>
                          <th className="p-5 w-[110px] text-[10px] uppercase font-[900] tracking-widest">DSP ID</th>
                          <th className="p-5 w-[200px] text-[10px] uppercase font-[900] tracking-widest">AWB Number</th>
                          <th className="p-5 w-[180px] text-[10px] uppercase font-[900] tracking-widest">Client</th>
                          <th className="p-5 w-[180px] text-[10px] uppercase font-[900] tracking-widest">Order ID</th>
                          <th className="p-5 text-[10px] uppercase font-[900] tracking-widest">NSL Remark</th>
                          <th className="p-5 w-[150px] text-[10px] uppercase font-[900] tracking-widest">FE Name</th>
                        </tr>
                      </thead>
                      <tbody>
                        {statusFilter === 'pending' && groupedPendingRows ? (
                          groupedPendingRows.map(([remark, rows]) => (
                            <React.Fragment key={remark}>
                              <tr className="bg-gradient-to-r from-slate-900 to-slate-800 border-y border-white/5">
                                <td colSpan={8} className="p-4 px-6"><div className="flex items-center gap-4"><Star className="w-4 h-4 text-amber-500 fill-current" /><span className="text-[11px] font-[900] text-white uppercase tracking-[0.2em]">{remark}</span><span className="px-3 py-1 bg-amber-500 text-white rounded-lg text-[10px] font-black uppercase tracking-widest">{rows.length} PKT</span></div></td>
                              </tr>
                              {rows.map((row, idx) => <DataRow key={row.id} row={row} idx={idx} isFirstInGroup={idx === 0} onCopyAwb={copySingleAwb} onDelete={(id) => setSessions(prev => prev.map(s => s.id === selectedSessionId ? {...s, data: s.data.filter(r => r.id !== id)} : s))} />)}
                            </React.Fragment>
                          ))
                        ) : (
                          <>
                            {currentSession && (
                              <tr className="bg-gradient-to-r from-slate-900 to-slate-800 border-y border-white/5">
                                <td colSpan={8} className="p-4 px-6">
                                  <div className="flex items-center justify-between">
                                    <div className="flex items-center gap-4">
                                      <span className="text-[16px] font-mono font-black text-amber-500">{currentSession.dspId}</span>
                                      <span className="px-3 py-1 bg-amber-500 text-white rounded-lg text-[10px] font-black uppercase tracking-widest">{filteredRows.length} PKT</span>
                                    </div>
                                    <div className="text-slate-400 text-[11px] font-black uppercase tracking-[0.3em]">{currentSession.feName} · {currentSession.date}</div>
                                  </div>
                                </td>
                              </tr>
                            )}
                            {filteredRows.map((row, idx) => <DataRow key={row.id} row={row} idx={idx} isFirstInGroup={idx === 0} onCopyAwb={copySingleAwb} onDelete={(id) => setSessions(prev => prev.map(s => s.id === selectedSessionId ? {...s, data: s.data.filter(r => r.id !== id)} : s))} />)}
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
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 animate-in slide-in-from-right-10 duration-700">
            <div className="lg:col-span-8 space-y-8">
              <div className="bg-white rounded-[24px] p-8 shadow-xl border border-slate-200">
                <div className="flex items-center justify-between mb-8">
                  <h2 className="text-[17px] font-[900] text-slate-900 flex items-center gap-3"><FileSpreadsheet className="w-6 h-6 text-green-700" /> Remark Replacer Dashboard</h2>
                  {replacerData.length > 0 && (
                    <div className="flex items-center gap-4">
                      <div className="px-4 py-1.5 bg-slate-100 rounded-2xl border border-slate-200 text-[11px] font-[900] uppercase text-slate-600">Total: {replacerStats.total}</div>
                      <div className="px-4 py-1.5 bg-green-50 rounded-2xl border border-green-200 text-[11px] font-[900] uppercase text-green-700">Replaced: {replacerStats.replaced}</div>
                      <div className="px-4 py-1.5 bg-amber-50 rounded-2xl border border-amber-200 text-[11px] font-[900] uppercase text-amber-700">Original: {replacerStats.noMapping}</div>
                    </div>
                  )}
                </div>

                <div className={cn("border-[3px] border-dashed rounded-[30px] p-16 text-center transition-all cursor-pointer relative group", uploadError ? "border-red-500 bg-red-50/50" : "border-slate-200 bg-slate-50/50 hover:border-blue-600 hover:bg-white", isProcessing && "opacity-80 pointer-events-none")}>
                  <input type="file" disabled={isProcessing} onChange={handleReplacerUpload} className="absolute inset-0 opacity-0 cursor-pointer z-20" />
                  {isProcessing ? (
                    <div className="flex flex-col items-center gap-4"><Loader2 className="w-12 h-12 text-blue-600 animate-spin" /><p className="text-[13px] font-[900] uppercase text-blue-600 tracking-[0.3em]">Processing file...</p></div>
                  ) : uploadError ? (
                    <div className="space-y-4">
                      <AlertCircle className="w-12 h-12 text-red-500 mx-auto" />
                      <p className="text-[15px] font-[900] text-red-600 uppercase">{uploadError}</p>
                      <button onClick={() => setUploadError(null)} className="bg-red-600 text-white px-8 py-2.5 rounded-xl text-[11px] font-black uppercase mt-4 shadow-lg">Try Again</button>
                    </div>
                  ) : (
                    <div className="space-y-3"><Download className="w-12 h-12 text-blue-600/30 mx-auto group-hover:scale-110 transition-transform" /><p className="text-[17px] font-[900] text-slate-900">Drop Delhivery EOD Sheet here, or click to upload</p><p className="text-[11px] text-slate-500 uppercase tracking-widest font-bold">AUTO-MAPPING ACTIVE</p></div>
                  )}
                </div>
              </div>

              {replacerData.length > 0 && (
                <div className="bg-white rounded-[26px] shadow-2xl border border-slate-200 overflow-hidden">
                  <div className="p-5 bg-slate-50/80 border-b flex items-center justify-between">
                    <div className="flex gap-4">
                      <button onClick={downloadOfficialExcel} className="bg-[#0F172A] hover:bg-slate-800 text-white px-7 py-3 rounded-2xl font-[900] text-[13px] uppercase tracking-widest flex items-center gap-3 shadow-xl transition-all"><Download className="w-5 h-5" /> Download Official Excel</button>
                      <button onClick={copyRemarkDataToClipboard} className="bg-blue-700 hover:bg-blue-800 text-white px-7 py-3 rounded-2xl font-[900] text-[13px] uppercase tracking-widest flex items-center gap-3 shadow-xl transition-all"><Copy className="w-5 h-5" /> Copy Data</button>
                    </div>
                    <button onClick={() => { setReplacerData([]); setReplacerMeta(null); }} className="text-[11px] font-black text-red-500 uppercase tracking-widest px-5 hover:underline decoration-2">Clear Data</button>
                  </div>
                  <div className="overflow-x-auto max-h-[700px] custom-scrollbar">
                    <table className="w-full text-left border-collapse text-[12px]">
                      <thead className="sticky top-0 z-30">
                        <tr className="bg-[#0F172A] text-white">
                          <th className="p-5 w-[50px] text-center"><Trash2 className="w-4 h-4 mx-auto" /></th>
                          {replacerMeta?.headers.map((h, i) => <th key={i} className="p-5 font-[900] uppercase border-r border-white/5 whitespace-nowrap tracking-wider">{h}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {replacerData.map((row) => (
                          <tr key={row.__id} className={cn("border-b transition-colors", row.__isReplaced ? "bg-green-50/50" : "bg-amber-50/50")}>
                            <td className="p-4 text-center"><button onClick={() => setReplacerData(prev => prev.filter(r => r.__id !== row.__id))} className="text-slate-300 hover:text-red-500 transition-colors"><Trash2 className="w-4 h-4 mx-auto" /></button></td>
                            {replacerMeta?.headers.map((h, i) => (
                              <td key={i} className={cn("p-4 whitespace-nowrap font-bold", h === "Remarks Of NSL" && row.__isReplaced ? "text-green-700 font-black" : h === "Remarks Of NSL" ? "text-amber-700 font-black" : "text-slate-800")}>
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
              <div className="bg-white rounded-[24px] shadow-2xl border border-slate-200 overflow-hidden sticky top-[150px]">
                <div className="bg-green-700 p-5 text-white shadow-xl">
                  <h3 className="text-[12px] font-[900] uppercase tracking-[0.2em] flex items-center gap-3"><Info className="w-5 h-5" /> REMARK MAPPING GUIDE</h3>
                </div>
                <div className="p-5 overflow-y-auto max-h-[75vh] space-y-4 custom-scrollbar bg-slate-50/50">
                  {Object.entries(REMARK_MAPPING).map(([nsl, official]) => (
                    <div key={nsl} className="p-5 bg-white rounded-2xl border-2 border-slate-100 space-y-4 hover:border-green-500/30 shadow-sm transition-all group">
                      <div className="flex flex-col gap-2"><span className="text-[9px] font-[900] text-slate-400 uppercase tracking-widest">INPUT REMARK</span><span className="bg-amber-50 text-amber-700 border border-amber-200 px-4 py-2 rounded-xl text-[11px] font-black leading-tight shadow-sm">{nsl}</span></div>
                      <div className="flex justify-center"><ArrowRight className="w-4 h-4 text-green-600" /></div>
                      <div className="flex flex-col gap-2"><span className="text-[9px] font-[900] text-slate-400 uppercase tracking-widest">AUTO REPLACED TO</span><span className="bg-green-50 text-green-700 border border-green-200 px-4 py-2 rounded-xl text-[11px] font-black leading-tight shadow-sm">{official}</span></div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        )}
      </main>

      <style jsx global>{`
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;800&family=IBM+Plex+Mono:wght@700&display=swap');
        body { font-family: 'Plus Jakarta Sans', sans-serif; -webkit-font-smoothing: antialiased; }
        .custom-scrollbar::-webkit-scrollbar { width: 8px; height: 8px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #F1F5F9; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 10px; border: 2px solid #F1F5F9; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #94A3B8; }
      `}</style>
    </div>
  );
}

function DataRow({ row, idx, isFirstInGroup, onCopyAwb, onDelete }: { 
  row: PODRow, 
  idx: number, 
  isFirstInGroup: boolean, 
  onCopyAwb: (awb: string) => void,
  onDelete: (id: string) => void 
}) {
  return (
    <tr className={cn("border-b transition-all group", row.isIntact ? "bg-red-50/30" : "hover:bg-blue-50/50")}>
      <td className="p-4 text-center"><input type="checkbox" className="w-5 h-5 accent-blue-700 rounded-lg cursor-pointer" /></td>
      <td className="p-4 text-center"><button onClick={() => onDelete(row.id)} className="text-slate-300 hover:text-red-500 transition-colors"><Trash2 className="w-4 h-4 mx-auto" /></button></td>
      <td className="p-4 font-mono text-[12px] font-[900] text-slate-900">{isFirstInGroup ? row.dspId : ""}</td>
      <td onClick={() => onCopyAwb(row.awb)} className="p-4 font-mono text-[13px] font-[900] text-blue-700 cursor-pointer hover:underline underline-offset-4 decoration-2 transition-all tracking-tighter">{row.awb}</td>
      <td className="p-4 text-[12px] text-slate-800 truncate max-w-[160px] font-black">{row.client}</td>
      <td className="p-4 text-[12px] text-slate-800 truncate max-w-[160px] font-black">{row.orderId}</td>
      <td className="p-4">
        <span className={cn(
          "px-3 py-1.5 rounded-xl text-[11px] font-[900] uppercase tracking-tighter border-2 shadow-sm inline-block whitespace-nowrap overflow-hidden text-overflow-ellipsis max-w-[200px] align-middle", 
          row.isIntact ? "bg-red-50 text-red-700 border-red-200" : "bg-amber-50 text-amber-700 border-amber-200"
        )}>
          {row.remark || "NO REMARK"}
        </span>
      </td>
      <td className="p-4 text-[11px] font-black text-slate-500 uppercase tracking-wider">{row.feName}</td>
    </tr>
  );
}
