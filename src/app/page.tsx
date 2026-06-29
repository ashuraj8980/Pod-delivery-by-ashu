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
  Filter
} from "lucide-react";
import * as XLSX from "xlsx";
import { cn } from "@/lib/utils";

/**
 * @fileOverview Delhivery POD Management Tool - Palam Vihar RPC Edition
 * High-Definition UI Redesign with Professional Readability.
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
  
  const [replacerData, setReplacerData] = useState<any[]>([]);
  const [replacerMeta, setReplacerMeta] = useState<{headers: string[], remarkKey: string, awbKey: string} | null>(null);
  const [replacerStats, setReplacerStats] = useState({ total: 0, replaced: 0, noMapping: 0 });

  useEffect(() => {
    setIsMounted(true);
    setSetupData(prev => ({ ...prev, date: new Date().toISOString().split('T')[0] }));
  }, []);

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

  const copyDataProfessional = useCallback(async (rows: any[], headers: string[]) => {
    if (!rows.length) return;
    
    const cleanValue = (val: any) => String(val || "").replace(/[\r\n\t]+/g, " ").trim();
    const plainText = rows.map(r => headers.map(h => cleanValue(r[h])).join("\t")).join("\n");

    const rowsHtml = rows.map(r => {
      const cells = headers.map(h => {
        const val = cleanValue(r[h]);
        const style = h === "Awb" || h === "DSP No" || h === "Order- No" || h === "Awb Number" || h === "DSP ID" || h === "Order ID" ? 'style=\'mso-number-format:"\\@"\'' : '';
        return `<td ${style}>${val}</td>`;
      }).join("");
      return `<tr>${cells}</tr>`;
    }).join("");

    const htmlTable = `<html><body><table border="1"><tbody>${rowsHtml}</tbody></table></body></html>`;

    try {
      const textBlob = new Blob([plainText], { type: 'text/plain' });
      const htmlBlob = new Blob([htmlTable], { type: 'text/html' });
      const data = [new ClipboardItem({ 'text/plain': textBlob, 'text/html': htmlBlob })];
      await navigator.clipboard.write(data);
      return true;
    } catch (err) {
      await navigator.clipboard.writeText(plainText);
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
              remark: remark || "NO REMARK",
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
            showToast(`Session loaded for DSP ${setupData.dspId}`, "ok");
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
          const wb = XLSX.read(data, { type: 'array', cellText: true, cellDates: true });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rawData = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
          
          if (!Array.isArray(rawData) || rawData.length === 0) throw new Error("Empty file.");

          const allHeaders = Object.keys(rawData[0]);
          const remarkKey = allHeaders.find(h => h.trim() === "Remarks Of NSL") || "";
          const awbKey = allHeaders.find(h => h.trim() === "Awb") || "";

          if (!remarkKey) throw new Error("This file does not have Remarks Of NSL column. Please upload the correct Delhivery EOD rejection sheet.");

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
            const cleanDsp = fixValueToString(row["DSP No"]);
            
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

  // Unique remarks for the current status filter
  const uniqueRemarksInStatus = useMemo(() => {
    if (!currentSession) return [];
    let baseData = currentSession.data;
    if (statusFilter !== 'all') {
      if (statusFilter === 'pending') baseData = baseData.filter(r => r.status === 'pending');
      else if (statusFilter === 'dispatched') baseData = baseData.filter(r => r.status === 'dispatched' || r.status === 'dispatch');
      else if (statusFilter === 'rto') baseData = baseData.filter(r => r.status === 'rto');
      else if (statusFilter === 'dto') baseData = baseData.filter(r => r.status === 'dto' || r.status === 'delivered');
    }
    const remarks = baseData.map(r => r.remark);
    return Array.from(new Set(remarks)).filter(Boolean).sort();
  }, [currentSession, statusFilter]);

  const handleCopyTable = useCallback(async (isShortcut = false) => {
    if (!filteredRows.length || !currentSession) return;
    
    const headers = ['Date', 'DSP ID', 'Awb', 'Client', 'Order ID', 'Remark', 'FE Name'];
    const exportRows = filteredRows.map((r, i) => ({
      'Date': r.date,
      'DSP ID': i === 0 ? currentSession.dspId : "",
      'Awb': r.awb,
      'Client': r.client,
      'Order ID': r.orderId,
      'Remark': r.remark,
      'FE Name': r.feName
    }));

    const success = await copyDataProfessional(exportRows, headers);
    if (success) {
      showToast(isShortcut ? `Ctrl+T — Copied ${filteredRows.length} rows` : `Copied ${filteredRows.length} rows to clipboard`, "ok");
    }
  }, [filteredRows, currentSession, copyDataProfessional, showToast]);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.ctrlKey && (e.key === 't' || e.key === 'T')) {
        e.preventDefault(); 
        handleCopyTable(true);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [handleCopyTable]);

  const copySingleAwb = useCallback(async (awb: string) => {
    const success = await copyDataProfessional([{ "Awb": awb }], ["Awb"]);
    if (success) showToast(`AWB ${awb} Copied`, 'ok');
  }, [copyDataProfessional, showToast]);

  const downloadExcel = () => {
    if (!filteredRows.length || !currentSession) return;
    const header = ['Date', 'DSP ID', 'AWB Number', 'Client', 'Order ID', 'Remark', 'FE Name'];
    const excelData = filteredRows.map((r, i) => [
      r.date, 
      { v: i === 0 ? String(currentSession.dspId) : "", t: 's', z: '@' }, 
      { v: String(r.awb), t: 's', z: '@' }, 
      r.client, 
      r.orderId, 
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
        } else {
          row[h] = r[h];
        }
      });
      return row;
    }), { header: headers });

    const range = XLSX.utils.decode_range(ws['!ref']!);
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const address = XLSX.utils.encode_col(C) + "1";
      if (!ws[address]) continue;
      ws[address].s = {
        fill: { fgColor: { rgb: "1C2333" } },
        font: { color: { rgb: "FFFFFF" }, bold: true, sz: 11, name: "Calibri" },
        alignment: { horizontal: "center", vertical: "center" }
      };
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Remarks");
    XLSX.writeFile(wb, `EOD_Official_Remarks_${new Date().toISOString().split('T')[0]}.xlsx`);
    showToast("Official Excel Downloaded", "ok");
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
              <h1 className="text-[16px] font-extrabold uppercase tracking-tight leading-none">POD Tool</h1>
              <p className="text-[11px] text-slate-400 font-bold mt-1 uppercase tracking-wider">Palam Vihar RPC · <span className="text-amber-400">By Ashu</span></p>
            </div>
          </div>
          <div className="flex gap-8 h-full">
            {[ { id: "eod", label: "Daily EOD Rejection" }, { id: "remark", label: "EOD Rejection Remark" } ].map(tab => (
              <button key={tab.id} onClick={() => { setActiveTab(tab.id as any); setStatusFilter("all"); setActiveRemarkChip(null); }} className={cn("px-2 py-4 text-[13px] font-semibold uppercase tracking-widest transition-all relative h-full", activeTab === tab.id ? "text-white" : "text-slate-400 hover:text-white")}>
                {tab.label}
                {activeTab === tab.id && <div className="absolute bottom-0 left-0 w-full h-1 bg-amber-400 shadow-[0_-2px_10px_#fbbf24]" />}
              </button>
            ))}
          </div>
        </div>
      </header>

      <main className="pt-24 px-6 pb-20 max-w-[1440px] mx-auto space-y-6">
        {activeTab === "eod" ? (
          <div className="space-y-6">
            <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-sm">
              <p className="text-[12px] font-bold text-slate-500 uppercase tracking-widest mb-4 flex items-center gap-2"><Info className="w-4 h-4 text-blue-600" /> Session Setup</p>
              <div className="flex flex-col md:flex-row gap-4 mb-6">
                <div className="flex-1 space-y-1.5">
                  <label className="text-[12px] font-semibold text-slate-600 px-1 uppercase tracking-tight">DSP ID</label>
                  <input type="number" value={setupData.dspId} onChange={(e) => setSetupData({...setupData, dspId: e.target.value})} className="w-full bg-[#F9FAFB] border-[1.5px] border-[#D1D5DB] rounded-lg px-4 h-[42px] text-[14px] font-bold text-[#111827] outline-none focus:border-[#1976D2] focus:bg-white focus:ring-[3px] focus:ring-[#1976D2]/10 transition-all" placeholder="Enter DSP ID" />
                </div>
                <div className="flex-1 space-y-1.5">
                  <label className="text-[12px] font-semibold text-slate-600 px-1 uppercase tracking-tight">FE Name</label>
                  <input type="text" value={setupData.feName} onChange={(e) => setSetupData({...setupData, feName: e.target.value})} className="w-full bg-[#F9FAFB] border-[1.5px] border-[#D1D5DB] rounded-lg px-4 h-[42px] text-[14px] font-bold text-[#111827] outline-none focus:border-[#1976D2] focus:bg-white focus:ring-[3px] focus:ring-[#1976D2]/10 transition-all" placeholder="Enter FE Name" />
                </div>
                <div className="flex-[0.5] space-y-1.5">
                  <label className="text-[12px] font-semibold text-slate-600 px-1 uppercase tracking-tight">Date</label>
                  <input type="date" value={setupData.date} onChange={(e) => setSetupData({...setupData, date: e.target.value})} className="w-full bg-[#F9FAFB] border-[1.5px] border-[#D1D5DB] rounded-lg px-4 h-[42px] text-[14px] font-bold text-[#111827] outline-none focus:border-[#1976D2] focus:bg-white focus:ring-[3px] focus:ring-[#1976D2]/10 transition-all" />
                </div>
              </div>
              
              <div className={cn("border-[2px] border-dashed rounded-xl p-10 text-center transition-all cursor-pointer relative group", uploadError ? "border-rose-300 bg-rose-50" : "border-slate-300 bg-slate-50 hover:border-blue-500 hover:bg-white", isProcessing && "opacity-80 pointer-events-none")}>
                <input type="file" disabled={isProcessing} onChange={handleFileUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                {isProcessing ? (
                  <div className="flex flex-col items-center gap-3"><Loader2 className="w-8 h-8 text-blue-600 animate-spin" /><p className="text-[14px] font-bold text-blue-600 uppercase tracking-widest">Processing File...</p></div>
                ) : uploadError ? (
                  <div className="space-y-3"><AlertCircle className="w-8 h-8 text-rose-500 mx-auto" /><p className="text-[14px] font-bold text-rose-600">{uploadError}</p><button onClick={() => setUploadError(null)} className="bg-rose-600 text-white px-5 py-2 rounded-lg text-[13px] font-bold uppercase">Try Again</button></div>
                ) : (
                  <div className="space-y-2"><Download className="w-8 h-8 text-slate-400 mx-auto group-hover:scale-110 transition-transform" /><p className="text-[14px] font-bold text-slate-800">Drop Delhivery CSV here, or click to upload</p><p className="text-[11px] text-slate-500 uppercase tracking-widest font-semibold">Ready for processing</p></div>
                )}
              </div>
            </div>

            {sessions.length > 0 && (
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                {sessions.map(s => (
                  <div key={s.id} onClick={() => { setSelectedSessionId(s.id); setStatusFilter('all'); setActiveRemarkChip(null); }} className={cn("bg-white p-4 rounded-xl border-l-[4px] shadow-sm cursor-pointer relative transition-all group border border-slate-200", selectedSessionId === s.id ? "border-l-blue-600 ring-4 ring-blue-500/5" : "border-l-slate-300 opacity-70 hover:opacity-100")}>
                    <button onClick={(e) => { e.stopPropagation(); setSessions(prev => prev.filter(x => x.id !== s.id)); if(selectedSessionId === s.id) setSelectedSessionId(null); }} className="absolute top-3 right-3 text-slate-300 hover:text-rose-500 transition-colors opacity-0 group-hover:opacity-100"><X className="w-4 h-4" /></button>
                    <p className="text-[15px] font-bold text-slate-900 truncate pr-8">{s.feName}</p>
                    <p className="text-[12px] font-semibold text-slate-500 mt-1 uppercase tracking-tight">{s.dspId} — {s.date}</p>
                    <div className="mt-3"><span className="px-2 py-0.5 bg-slate-100 rounded text-[11px] font-bold text-slate-600 uppercase tracking-tight">{s.data.length} Rows</span></div>
                  </div>
                ))}
              </div>
            )}

            {currentSession && (
              <div className="space-y-6">
                <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex h-[80px]">
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

                {/* Remark Selection Chips */}
                {uniqueRemarksInStatus.length > 0 && (
                  <div className="flex flex-wrap gap-2 py-2 px-1">
                    {uniqueRemarksInStatus.map(rem => (
                      <button
                        key={rem}
                        onClick={() => setActiveRemarkChip(activeRemarkChip === rem ? null : rem)}
                        className={cn(
                          "px-4 py-1.5 rounded-lg text-[13px] font-bold border transition-all flex items-center gap-2",
                          activeRemarkChip === rem 
                            ? "bg-blue-600 text-white border-blue-700 shadow-md scale-105" 
                            : "bg-white text-slate-700 border-slate-200 hover:border-blue-300 hover:bg-blue-50/30"
                        )}
                      >
                        <Filter className={cn("w-3.5 h-3.5", activeRemarkChip === rem ? "text-white" : "text-blue-500")} />
                        {rem}
                      </button>
                    ))}
                  </div>
                )}

                <div className="flex items-center justify-between gap-4">
                  <div className="flex gap-3">
                    <button onClick={downloadExcel} className="h-10 px-6 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-[13px] font-bold uppercase tracking-wider flex items-center gap-2 shadow-sm transition-all"><Download className="w-4 h-4" /> Download Excel</button>
                    <button onClick={() => handleCopyTable(false)} className="h-10 px-6 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-[13px] font-bold uppercase tracking-wider flex items-center gap-2 shadow-sm transition-all"><Copy className="w-4 h-4" /> Copy Table</button>
                  </div>
                  <div className="relative"><Search className="w-4 h-4 absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" /><input type="text" placeholder="Search by AWB / Client..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="bg-white border-[1.5px] border-slate-200 rounded-lg pl-10 pr-4 h-10 text-[14px] font-semibold outline-none w-[300px] focus:border-blue-500 transition-all shadow-sm" /></div>
                </div>

                <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                  <div className="overflow-x-auto max-h-[800px] custom-scrollbar">
                    <table className="w-full text-left border-collapse">
                      <thead className="sticky top-0 z-30 bg-[#0F172A] text-white">
                        <tr className="h-11">
                          <th className="px-4 w-[50px] text-center"><X className="w-3.5 h-3.5 mx-auto" /></th>
                          <th className="px-4 w-[110px] text-[11px] font-bold uppercase tracking-widest">DSP ID</th>
                          <th className="px-4 w-[180px] text-[11px] font-bold uppercase tracking-widest">AWB Number</th>
                          <th className="px-4 text-[11px] font-bold uppercase tracking-widest">Client</th>
                          <th className="px-4 text-[11px] font-bold uppercase tracking-widest">Order ID</th>
                          <th className="px-4 text-[11px] font-bold uppercase tracking-widest">Remark</th>
                          <th className="px-4 w-[140px] text-[11px] font-bold uppercase tracking-widest">FE Name</th>
                        </tr>
                      </thead>
                      <tbody>
                        {currentSession && (
                          <tr className="bg-slate-800 text-white h-10">
                            <td colSpan={7} className="px-4">
                              <div className="flex items-center gap-3">
                                <span className="text-[14px] font-bold text-amber-400">{currentSession.dspId}</span>
                                <span className="px-2 py-0.5 bg-blue-600 text-white rounded text-[11px] font-bold uppercase tracking-tighter">{filteredRows.length} Pkt</span>
                                <span className="text-[12px] text-slate-400 font-bold uppercase ml-auto">{currentSession.feName} · {currentSession.date}</span>
                              </div>
                            </td>
                          </tr>
                        )}
                        {filteredRows.map((row) => (
                          <tr key={row.id} className={cn("h-11 border-b border-[#F1F5F9] transition-colors group", row.isIntact ? "bg-rose-50/30" : "hover:bg-blue-50/30")}>
                            <td className="px-4 text-center"><button onClick={() => setSessions(prev => prev.map(s => s.id === selectedSessionId ? {...s, data: s.data.filter(r => r.id !== row.id)} : s))} className="text-slate-300 hover:text-rose-500 transition-colors"><Trash2 className="w-4 h-4 mx-auto" /></button></td>
                            <td className="px-4 text-[13px] font-bold text-slate-900">{row.dspId}</td>
                            <td onClick={() => copySingleAwb(row.awb)} className="px-4 text-[12px] font-bold text-blue-700 cursor-pointer hover:underline transition-all tracking-tight" style={{ fontFamily: '"IBM Plex Mono", monospace' }}>{row.awb}</td>
                            <td className="px-4 text-[13px] font-bold text-slate-800 truncate max-w-[150px]">{row.client}</td>
                            <td className="px-4 text-[13px] font-bold text-slate-800 truncate max-w-[150px]">{row.orderId}</td>
                            <td className="px-4">
                              <span className={cn(
                                "px-3 py-1 rounded-lg text-[13px] font-bold uppercase border whitespace-nowrap overflow-hidden text-ellipsis max-w-[200px] inline-block", 
                                row.isIntact ? "bg-rose-50 text-rose-700 border-rose-100" : "bg-amber-50 text-amber-700 border-amber-100"
                              )}>
                                {row.remark || "NO REMARK"}
                              </span>
                            </td>
                            <td className="px-4 text-[12px] font-bold text-slate-500 uppercase">{row.feName}</td>
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
              <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between mb-6">
                  <h2 className="text-[16px] font-extrabold text-slate-900 flex items-center gap-2"><FileSpreadsheet className="w-5 h-5 text-emerald-600" /> Remark Replacer</h2>
                  {replacerData.length > 0 && (
                    <div className="flex gap-2">
                      <div className="px-3 py-1 bg-slate-100 rounded-lg text-[11px] font-bold text-slate-600 uppercase">Total: {replacerStats.total}</div>
                      <div className="px-3 py-1 bg-emerald-50 rounded-lg text-[11px] font-bold text-emerald-700 uppercase">Replaced: {replacerStats.replaced}</div>
                      <div className="px-3 py-1 bg-amber-50 rounded-lg text-[11px] font-bold text-amber-700 uppercase">Original: {replacerStats.noMapping}</div>
                    </div>
                  )}
                </div>

                <div className={cn("border-[2px] border-dashed rounded-xl p-10 text-center transition-all cursor-pointer relative group", uploadError ? "border-rose-300 bg-rose-50" : "border-slate-300 bg-slate-50 hover:border-blue-500 hover:bg-white", isProcessing && "opacity-80 pointer-events-none")}>
                  <input type="file" disabled={isProcessing} onChange={handleReplacerUpload} className="absolute inset-0 opacity-0 cursor-pointer z-20" />
                  {isProcessing ? (
                    <div className="flex flex-col items-center gap-3"><Loader2 className="w-8 h-8 text-blue-600 animate-spin" /><p className="text-[14px] font-bold text-blue-600 uppercase tracking-widest">Processing file...</p></div>
                  ) : uploadError ? (
                    <div className="space-y-3"><AlertCircle className="w-8 h-8 text-rose-500 mx-auto" /><p className="text-[14px] font-bold text-rose-600">{uploadError}</p><button onClick={() => setUploadError(null)} className="bg-rose-600 text-white px-5 py-2 rounded-lg text-[13px] font-bold uppercase mt-2">Try Again</button></div>
                  ) : (
                    <div className="space-y-2"><Download className="w-8 h-8 text-slate-400 mx-auto group-hover:scale-110 transition-transform" /><p className="text-[14px] font-bold text-slate-800">Drop EOD Sheet here, or click to upload</p><p className="text-[11px] text-slate-500 uppercase tracking-widest font-semibold">Mapping active</p></div>
                  )}
                </div>
              </div>

              {replacerData.length > 0 && (
                <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                  <div className="p-4 bg-slate-50 border-b flex items-center justify-between">
                    <div className="flex gap-3">
                      <button onClick={downloadOfficialExcel} className="h-10 px-6 bg-slate-900 text-white rounded-lg text-[13px] font-bold uppercase tracking-wider flex items-center gap-2"><Download className="w-4 h-4" /> Download Excel</button>
                      <button onClick={() => copyDataProfessional(replacerData, replacerMeta?.headers || [])} className="h-10 px-6 bg-blue-600 text-white rounded-lg text-[13px] font-bold uppercase tracking-wider flex items-center gap-2"><Copy className="w-4 h-4" /> Copy Data</button>
                    </div>
                    <button onClick={() => { setReplacerData([]); setReplacerMeta(null); }} className="text-[11px] font-bold text-rose-500 uppercase px-4 hover:underline tracking-widest">Clear</button>
                  </div>
                  <div className="overflow-x-auto max-h-[700px] custom-scrollbar">
                    <table className="w-full text-left border-collapse text-[13px]">
                      <thead className="sticky top-0 z-30 bg-[#0F172A] text-white">
                        <tr className="h-11">
                          <th className="px-4 w-[50px] text-center"><Trash2 className="w-4 h-4 mx-auto" /></th>
                          {replacerMeta?.headers.map((h, i) => <th key={i} className="px-4 font-bold uppercase border-r border-white/5 whitespace-nowrap tracking-wider text-[11px]">{h}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {replacerData.map((row) => (
                          <tr key={row.__id} className={cn("h-11 border-b border-[#F1F5F9] transition-colors", row.__isReplaced ? "bg-[#F0FDF4]" : "bg-[#FFFDE7]")}>
                            <td className="px-4 text-center"><button onClick={() => setReplacerData(prev => prev.filter(r => r.__id !== row.__id))} className="text-slate-300 hover:text-rose-500 transition-colors"><Trash2 className="w-4 h-4 mx-auto" /></button></td>
                            {replacerMeta?.headers.map((h, i) => (
                              <td key={i} className={cn("px-4 whitespace-nowrap font-bold", h === "Remarks Of NSL" && row.__isReplaced ? "text-emerald-700" : h === "Remarks Of NSL" ? "text-amber-700" : "text-slate-800")}>
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
              <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden sticky top-24">
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

      <style jsx global>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@700&display=swap');
      `}</style>
    </div>
  );
}
