"use client";

import React, { useState, useEffect, useMemo } from "react";
import { 
  Truck, 
  Trash2, 
  Download, 
  Copy, 
  X, 
  AlertCircle, 
  Filter,
  Settings,
  ChevronRight,
  ChevronLeft,
  Calendar,
  User,
  Hash,
  ArrowUpDown
} from "lucide-react";
import * as XLSX from "xlsx";
import { cn } from "@/lib/utils";

/**
 * @fileOverview Delhivery POD Management Tool v1.1
 * Features: AWB Scientific Notation Fix, Filtered Export, Remark Breakdown.
 * Author: Ashu (ashuraj9771@gmail.com)
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

export default function PODTool() {
  const [activeTab, setActiveTab] = useState<"eod" | "remark">("eod");
  const [sessions, setSessions] = useState<Session[]>([]);
  const [currentSession, setCurrentSession] = useState<Session | null>(null);
  const [statusFilter, setStatusFilter] = useState<string>("all");
  const [remarkFilter, setRemarkFilter] = useState<string | null>(null);
  const [setupData, setSetupData] = useState({ feName: "", dspId: "", date: "" });
  const [showIntactModal, setShowIntactModal] = useState(false);
  const [isMounted, setIsMounted] = useState(false);
  
  // Module 2 State
  const [replacerData, setReplacerData] = useState<any[]>([]);
  const [replacerStats, setReplacerStats] = useState({ total: 0, replaced: 0, missing: 0 });

  useEffect(() => {
    setIsMounted(true);
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) setSessions(JSON.parse(saved));
    setSetupData(prev => ({ ...prev, date: new Date().toISOString().split('T')[0] }));
  }, []);

  useEffect(() => {
    if (isMounted) {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(sessions));
    }
  }, [sessions, isMounted]);

  // Keyboard Navigation
  useEffect(() => {
    if (!isMounted) return;
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.target instanceof HTMLInputElement || e.target instanceof HTMLTextAreaElement) return;
      
      if (e.key === "ArrowLeft") {
        const tabs = ["all", "pending", "dispatched", "rto", "dto"];
        const idx = tabs.indexOf(statusFilter);
        if (idx > 0) {
          setStatusFilter(tabs[idx - 1]);
          setRemarkFilter(null);
        }
      }
      if (e.key === "ArrowRight") {
        const tabs = ["all", "pending", "dispatched", "rto", "dto"];
        const idx = tabs.indexOf(statusFilter);
        if (idx < tabs.length - 1) {
          setStatusFilter(tabs[idx + 1]);
          setRemarkFilter(null);
        }
      }
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [statusFilter, isMounted]);

  const fixAWB = (val: any) => {
    let str = String(val).trim();
    if (/^[\d.]+[eE][+\-]?\d+$/.test(str)) {
      str = BigInt(Math.round(Number(val))).toString();
    }
    return str.replace(/\.0$/, "");
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!setupData.feName || !setupData.dspId) {
      showToast("Bhai, FE aur DSP Number enter karein!", "err");
      e.target.value = "";
      return;
    }

    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rawData = XLSX.utils.sheet_to_json(ws);

      const parsedRows: PODRow[] = rawData.map((row: any) => {
        const keys = Object.keys(row);
        const findVal = (regex: RegExp) => {
          const key = keys.find(k => regex.test(k.toLowerCase().replace(/[\s_-]/g, "")));
          return key ? row[key] : "";
        };

        const awb = fixAWB(findVal(/waybill|awb|awbnumber/));
        const statusRaw = String(findVal(/status|currentstatus/)).toLowerCase().trim();
        const status = STATUS_MAP[statusRaw] || "unknown";

        return {
          awb,
          client: String(findVal(/client|clientname/)),
          orderId: String(findVal(/order|orderid/)),
          status,
          remark: String(findVal(/remark|remarks/)),
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
        setCurrentSession(newSession);
        setSessions(prev => [newSession, ...prev]);
        showToast(`Imported ${parsedRows.length} rows!`, "ok");
      } else {
        showToast("No valid rows found!", "err");
      }
    };
    reader.readAsBinaryString(file);
  };

  const filteredRows = useMemo(() => {
    if (!currentSession) return [];
    let rows = currentSession.data;
    if (statusFilter !== "all") {
      rows = rows.filter(r => r.status === statusFilter);
    }
    if (remarkFilter !== null) {
      rows = rows.filter(r => (r.remark || "") === remarkFilter);
    }
    return rows;
  }, [currentSession, statusFilter, remarkFilter]);

  const copyTable = (dataRows: PODRow[]) => {
    if (dataRows.length === 0) {
      showToast("No data to copy!", "err");
      return;
    }
    // No headers, tab separated, dsp only on first row
    const text = dataRows.map((r, i) => {
      const dsp = i === 0 ? r.dspId : "";
      return `${r.date}\t${dsp}\t${r.awb}\t${r.client}\t${r.orderId}\t${r.remark}\t${r.feName}`;
    }).join("\n");
    
    navigator.clipboard.writeText(text).then(() => {
      showToast("Copied — Paste in Excel", "info");
    }).catch(() => {
      showToast("Copy failed!", "err");
    });
  };

  const downloadExcel = (dataRows: PODRow[], session: Session) => {
    if (dataRows.length === 0) {
      showToast("No data to download!", "err");
      return;
    }
    const header = ['Date', 'DSP ID', 'AWB Number', 'Client', 'Order ID', 'Remark', 'FE Name'];
    const rows = dataRows.map((r, i) => [
      r.date,
      i === 0 ? session.dspId : "",
      { v: r.awb, t: 's' },
      r.client,
      r.orderId,
      r.remark,
      r.feName
    ]);

    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    // Set widths
    ws['!cols'] = [{ wch: 13 }, { wch: 12 }, { wch: 26 }, { wch: 20 }, { wch: 20 }, { wch: 36 }, { wch: 18 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "POD_Report");
    XLSX.writeFile(wb, `Delhivery_POD_${session.feName}_${statusFilter}.xlsx`);
    showToast("Downloading Filtered Excel...", "ok");
  };

  const handleReplacerUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const rawData = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);

      let replaced = 0;
      let missing = 0;

      const processed = rawData.map((row: any, idx) => {
        const findVal = (regex: RegExp) => {
          const key = Object.keys(row).find(k => regex.test(k.toLowerCase().replace(/[\s_-]/g, "")));
          return key ? row[key] : "";
        };

        const oldRemark = String(findVal(/remark/));
        let official = REMARK_MAPPING[oldRemark];
        if (!official) {
          const key = Object.keys(REMARK_MAPPING).find(k => oldRemark.toLowerCase().includes(k.toLowerCase()));
          if (key) official = REMARK_MAPPING[key];
        }

        if (official) replaced++; else missing++;

        return {
          id: idx + 1,
          date: String(findVal(/date/)),
          dsp: String(findVal(/dsp/)),
          awb: fixAWB(findVal(/awb/)),
          client: String(findVal(/client/)),
          oldRemark,
          officialRemark: official || oldRemark,
          isReplaced: !!official,
          feName: String(findVal(/fe/))
        };
      });

      setReplacerData(processed);
      setReplacerStats({ total: processed.length, replaced, missing });
      showToast(`Processed ${processed.length} rows`, "ok");
    };
    reader.readAsBinaryString(file);
  };

  const showToast = (msg: string, type: 'ok' | 'err' | 'info') => {
    if (!isMounted) return;
    const toast = document.createElement('div');
    toast.className = `fixed bottom-10 left-1/2 -translate-x-1/2 px-6 py-3 rounded-2xl text-xs font-bold z-[100] shadow-2xl animate-in slide-in-from-bottom-2 duration-300 ${
      type === 'ok' ? 'bg-[#052E0F] text-[#6EE7A6]' : 
      type === 'err' ? 'bg-[#2D0808] text-[#FCA5A5]' : 
      'bg-[#1C2333] text-[#93C5FD]'
    }`;
    toast.innerText = msg;
    document.body.appendChild(toast);
    setTimeout(() => {
      toast.classList.add('animate-out', 'fade-out', 'duration-500');
      setTimeout(() => toast.remove(), 500);
    }, 2800);
  };

  const stats = useMemo(() => {
    if (!currentSession) return { total: 0, pending: 0, dispatched: 0, rto: 0, dto: 0, intact: 0 };
    return {
      total: currentSession.data.length,
      pending: currentSession.data.filter(r => r.status === 'pending').length,
      dispatched: currentSession.data.filter(r => r.status === 'dispatched').length,
      rto: currentSession.data.filter(r => r.status === 'rto').length,
      dto: currentSession.data.filter(r => r.status === 'dto').length,
      intact: currentSession.data.filter(r => r.status === 'pending' && (r.remark.toLowerCase().includes('intact') || r.remark.toLowerCase().includes('reject but package'))).length
    };
  }, [currentSession]);

  const uniqueRemarks = useMemo(() => {
    if (!currentSession || statusFilter !== "pending") return [];
    const counts: Record<string, number> = {};
    currentSession.data.filter(r => r.status === 'pending').forEach(r => {
      const rem = r.remark || "";
      counts[rem] = (counts[rem] || 0) + 1;
    });
    return Object.entries(counts).sort((a, b) => b[1] - a[1]);
  }, [currentSession, statusFilter]);

  if (!isMounted) return null;

  return (
    <div className="min-h-screen bg-[#F0F4FA] font-body text-[#1C2333] select-none">
      <div className="h-[3px] w-full bg-gradient-to-r from-[#1565C0] via-[#F9A825] via-[#2E7D32] to-[#D32F2F]" />
      
      <header className="h-[58px] bg-[#1C2333] px-6 flex items-center justify-between text-white shadow-xl relative z-20">
        <div className="flex items-center gap-3">
          <div className="bg-gradient-to-br from-blue-400 to-blue-600 p-1.5 rounded-lg">
            <Truck className="w-6 h-6 text-white" />
          </div>
          <div className="flex flex-col">
            <h1 className="text-[15px] font-bold tracking-tight">POD Management Tool</h1>
            <p className="text-[10px] text-slate-400 font-medium">
              Delhivery · Palam Vihar RPC · <span className="text-[#F9A825] font-bold italic">By Ashu</span>
            </p>
          </div>
        </div>
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2 bg-green-500/10 px-3 py-1 rounded-full border border-green-500/20">
            <div className="w-1.5 h-1.5 bg-green-500 rounded-full animate-pulse shadow-[0_0_8px_rgba(34,197,94,0.6)]" />
            <span className="text-[9px] font-black text-green-400 tracking-widest uppercase">Live</span>
          </div>
          <div className="bg-[#F9A825] px-3 py-1 rounded-lg text-[#1C2333] font-code text-[11px] font-black">
            {currentSession?.data.length || 0} ROWS
          </div>
        </div>
      </header>

      <nav className="bg-[#1C2333] px-6 flex gap-10">
        <button 
          onClick={() => setActiveTab("eod")}
          className={cn(
            "py-3 text-[11px] font-bold uppercase tracking-widest transition-all relative",
            activeTab === "eod" ? "text-white" : "text-slate-500 hover:text-slate-300"
          )}
        >
          Daily EOD Rejection
          {activeTab === "eod" && <div className="absolute bottom-0 left-0 w-full h-[3px] bg-[#F9A825] rounded-t-full shadow-[0_-2px_8px_rgba(249,168,37,0.4)]" />}
        </button>
        <button 
          onClick={() => setActiveTab("remark")}
          className={cn(
            "py-3 text-[11px] font-bold uppercase tracking-widest transition-all relative",
            activeTab === "remark" ? "text-white" : "text-slate-500 hover:text-slate-300"
          )}
        >
          EOD Rejection Remark
          {activeTab === "remark" && <div className="absolute bottom-0 left-0 w-full h-[3px] bg-[#F9A825] rounded-t-full" />}
        </button>
      </nav>

      <main className="p-6 max-w-[1400px] mx-auto space-y-6">
        {activeTab === "eod" ? (
          <>
            <div className="bg-white rounded-[1.25rem] p-6 shadow-sm border border-[#E2E8F0] space-y-6">
              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1">
                  <label className="text-[9px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-1"><Hash className="w-3 h-3" /> DSP ID (Required)</label>
                  <input 
                    type="number" 
                    value={setupData.dspId} 
                    onChange={(e) => setSetupData({...setupData, dspId: e.target.value})}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-xs font-code font-bold outline-none focus:ring-2 ring-blue-500/10" 
                    placeholder="Enter DSP No..." 
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-[9px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-1"><User className="w-3 h-3" /> FE / Biker Name</label>
                  <input 
                    type="text" 
                    value={setupData.feName} 
                    onChange={(e) => setSetupData({...setupData, feName: e.target.value})}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-xs font-bold outline-none" 
                    placeholder="Enter FE Name..." 
                  />
                </div>
              </div>

              <div 
                className={cn(
                  "border-2 border-dashed rounded-2xl p-8 text-center transition-all relative group",
                  (!setupData.feName || !setupData.dspId) ? "bg-slate-50 border-slate-100 opacity-50 cursor-not-allowed" : "border-blue-200 hover:border-blue-500 hover:bg-blue-50/30 cursor-pointer"
                )}
              >
                <input 
                  type="file" 
                  disabled={!setupData.feName || !setupData.dspId} 
                  onChange={handleFileUpload} 
                  className="absolute inset-0 opacity-0 cursor-pointer disabled:cursor-not-allowed" 
                  accept=".xlsx,.xls,.csv" 
                />
                <div className="space-y-3">
                  <Download className={cn("w-10 h-10 mx-auto transition-colors", (!setupData.feName || !setupData.dspId) ? "text-slate-300" : "text-blue-400 group-hover:text-blue-600")} />
                  <div>
                    <p className="text-[11px] font-bold text-slate-600">
                      {(!setupData.feName || !setupData.dspId) ? "Bhai, pehle DSP ID aur FE Name fill karein!" : "Drop Delhivery export file here or click to upload"}
                    </p>
                  </div>
                </div>
              </div>
            </div>

            {currentSession && (
              <div className="grid grid-cols-5 gap-4">
                {[
                  { id: 'all', label: 'All', val: stats.total, col: 'blue' },
                  { id: 'pending', label: 'Pending', val: stats.pending, col: 'amber' },
                  { id: 'dispatched', label: 'Dispatch', val: stats.dispatched, col: 'blue' },
                  { id: 'rto', label: 'RTO', val: stats.rto, col: 'red' },
                  { id: 'dto', label: 'DTO', val: stats.dto, col: 'green' }
                ].map(t => (
                  <button 
                    key={t.id} 
                    onClick={() => { setStatusFilter(t.id); setRemarkFilter(null); }}
                    className={cn(
                      "bg-white p-5 rounded-[1.5rem] shadow-sm border-b-4 transition-all text-center",
                      statusFilter === t.id ? `border-${t.col}-500 bg-${t.col}-50/30 ring-4 ring-${t.col}-500/5` : "border-transparent border-slate-100"
                    )}
                  >
                    <p className={cn("text-3xl font-[800] leading-none", `text-${t.col}-600`)}>{t.val}</p>
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-[0.2em] mt-2">{t.label}</p>
                  </button>
                ))}
              </div>
            )}

            {statusFilter === "pending" && uniqueRemarks.length > 0 && (
              <div className="bg-white rounded-[2rem] p-6 border border-slate-200 shadow-sm">
                <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-4">REMARK BREAKDOWN — PENDING</p>
                <div className="flex flex-wrap gap-3">
                  {uniqueRemarks.map(([rem, count]) => (
                    <button 
                      key={rem} 
                      onClick={() => setRemarkFilter(rem)}
                      className={cn(
                        "px-5 py-2.5 rounded-full text-[10px] font-black uppercase tracking-widest border-2 transition-all flex items-center gap-2",
                        remarkFilter === rem ? "bg-blue-600 border-blue-600 text-white" : 
                        (rem.toLowerCase().includes('intact') || rem.toLowerCase().includes('reject but package')) ? "bg-red-50 border-red-100 text-red-600" :
                        "bg-slate-50 border-slate-200 text-slate-500"
                      )}
                    >
                      {rem || "No Remark"}
                      <span className="px-2 py-0.5 rounded-md bg-black/10 text-[9px]">{count}</span>
                    </button>
                  ))}
                  {remarkFilter !== null && (
                    <button onClick={() => setRemarkFilter(null)} className="text-[10px] font-black text-red-500 hover:underline uppercase">← All Pending</button>
                  )}
                </div>
              </div>
            )}

            {currentSession && (
              <div className="bg-white rounded-[2rem] shadow-2xl border border-slate-200 overflow-hidden">
                <div className="p-6 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
                  <div className="flex gap-3">
                    <button onClick={() => downloadExcel(filteredRows, currentSession)} className="bg-[#2E7D32] text-white px-6 py-3 rounded-2xl font-black text-[10px] uppercase tracking-widest flex items-center gap-2"><Download className="w-4 h-4" /> Download {statusFilter.toUpperCase()} Excel</button>
                    <button onClick={() => copyTable(filteredRows)} className="bg-[#1565C0] text-white px-6 py-3 rounded-2xl font-black text-[10px] uppercase tracking-widest flex items-center gap-2"><Copy className="w-4 h-4" /> Copy {statusFilter.toUpperCase()} Table</button>
                  </div>
                </div>
                
                <div className="overflow-x-auto max-h-[500px]">
                  <table className="w-full text-left">
                    <thead className="sticky top-0 z-30">
                      <tr className="bg-[#1C2333] text-white text-[10px] font-black uppercase tracking-[0.2em]">
                        <th className="p-4 w-[105px]">DSP ID</th>
                        <th className="p-4 w-[155px]">AWB Number</th>
                        <th className="p-4 w-[120px]">Client</th>
                        <th className="p-4 w-[120px]">Order ID</th>
                        <th className="p-4">Remark</th>
                        <th className="p-4 w-[115px]">FE Name</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredRows.map((row, idx) => (
                        <tr key={row.awb} className={cn(
                          "border-b border-slate-50 hover:bg-slate-50 transition-colors",
                          (row.remark.toLowerCase().includes("intact") || row.remark.toLowerCase().includes("reject but package")) && "bg-red-50/30"
                        )}>
                          <td className="p-4 font-code text-xs font-bold text-slate-400">{idx === 0 ? row.dspId : ""}</td>
                          <td className="p-4 font-code text-[11.5px] font-black text-[#1565C0] tracking-wider">{row.awb}</td>
                          <td className="p-4 text-[11px] font-bold text-slate-700">{row.client}</td>
                          <td className="p-4 text-[11px] text-slate-500">{row.orderId}</td>
                          <td className="p-4">
                            <span className={cn(
                              "inline-flex px-3 py-1 rounded-full text-[9px] font-black uppercase tracking-widest border",
                              (row.remark.toLowerCase().includes("intact") || row.remark.toLowerCase().includes("reject but package")) ? "bg-red-100 text-red-600 border-red-200" : "bg-[#FFFDE7] text-amber-600 border-amber-200"
                            )}>
                              {row.remark || "No Remark"}
                            </span>
                          </td>
                          <td className="p-4 text-[11px] font-bold text-slate-500">{row.feName}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </>
        ) : (
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            <div className="bg-white rounded-[2rem] p-8 border border-slate-200 shadow-sm space-y-8">
              <h2 className="text-lg font-black tracking-tight flex items-center gap-2"><AlertCircle className="w-5 h-5 text-green-600" /> EOD Remark Replacer</h2>
              <div className="border-4 border-dashed border-green-50 rounded-[2.5rem] p-12 text-center hover:border-green-200 transition-all cursor-pointer relative">
                <input type="file" onChange={handleReplacerUpload} className="absolute inset-0 opacity-0 cursor-pointer" accept=".xlsx,.xls,.csv" />
                <Download className="w-12 h-12 text-green-200 mx-auto" />
                <p className="text-sm font-black text-slate-700 mt-4">Upload EOD Export File</p>
              </div>
            </div>

            <div className="bg-[#1C2333] rounded-[2.5rem] p-8 text-white shadow-2xl">
              <h3 className="text-sm font-black text-yellow-400 uppercase tracking-[0.3em] mb-8">Official Mapping Reference</h3>
              <div className="overflow-y-auto max-h-[450px] space-y-2">
                {Object.entries(REMARK_MAPPING).map(([old, official]) => (
                  <div key={old} className="p-3 bg-white/5 rounded-xl text-[10px]">
                    <p className="text-slate-500 italic">{old}</p>
                    <p className="text-green-400 font-bold mt-1">→ {official}</p>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}
      </main>

      {stats.intact > 0 && activeTab === 'eod' && (
        <button 
          onClick={() => setShowIntactModal(true)}
          className="fixed bottom-10 right-10 bg-red-600 text-white px-8 py-5 rounded-full shadow-2xl flex items-center gap-3 z-40 border-b-4 border-red-800"
        >
          <AlertCircle className="w-6 h-6 fill-current" />
          <div className="flex flex-col items-start">
            <span className="text-[10px] font-black uppercase tracking-widest leading-none">Intact Alert</span>
            <span className="text-2xl font-black leading-none mt-1">{stats.intact} PKTS</span>
          </div>
        </button>
      )}

      {showIntactModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-6 bg-[#1C2333]/90 backdrop-blur-lg" onClick={() => setShowIntactModal(false)}>
          <div className="bg-white w-full max-w-4xl rounded-[3rem] shadow-2xl overflow-hidden" onClick={e => e.stopPropagation()}>
            <div className="bg-red-600 p-10 text-white">
              <h2 className="text-3xl font-black uppercase">Intact Packet Summary</h2>
            </div>
            <div className="p-10">
              <div className="grid grid-cols-2 gap-4">
                <button 
                  onClick={() => copyTable(currentSession?.data.filter(r => r.status === 'pending' && (r.remark.toLowerCase().includes('intact') || r.remark.toLowerCase().includes('reject but package'))) || [])} 
                  className="bg-slate-900 text-white p-6 rounded-3xl font-black text-[11px] uppercase"
                >
                  Copy Intact AWBs
                </button>
                <button 
                  onClick={() => downloadExcel(currentSession?.data.filter(r => r.status === 'pending' && (r.remark.toLowerCase().includes('intact') || r.remark.toLowerCase().includes('reject but package'))) || [], currentSession!)} 
                  className="bg-red-600 text-white p-6 rounded-3xl font-black text-[11px] uppercase"
                >
                  Download Intact Excel
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      <style jsx global>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500;600&display=swap');
        body { font-family: 'Inter', sans-serif; }
        .font-code { font-family: 'IBM Plex Mono', monospace; }
        ::-webkit-scrollbar { width: 6px; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 10px; }
      `}</style>
    </div>
  );
}
