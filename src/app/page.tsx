
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
  Settings,
  Database,
  BarChart3,
  Calendar,
  User,
  Hash,
  CheckCircle2,
  PackageSearch,
  ChevronDown,
  Check,
  AlertCircle
} from "lucide-react";
import * as XLSX from "xlsx";
import { cn } from "@/lib/utils";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  Popover,
  PopoverContent,
  PopoverTrigger,
} from "@/components/ui/popover";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Checkbox } from "@/components/ui/checkbox";

/**
 * @fileOverview Delhivery POD Management Tool - Palam Vihar RPC Edition
 * Restored with Remark Chips, Colorful UI, Client Grouping, and Non-Exportable Return Address.
 * All text converted to Normal Case as per user request.
 * Removed LocalStorage persistence - App starts fresh on every reload.
 */

const REMARK_MAPPING: Record<string, string> = {
  "Incomplete address & contact details": "Return address not found (Need to new contact number)",
  "On Hold. Recipient unable to Accept Delivery": "Client out of station (Receive the shipment after 3 days)",
  "On Hold. Recipient unable to Accept Delivery for 5 days": "Client out of station (Receive the shipment after 3 days)",
  "Not Attempted": "Not attempted to client",
  "Seller/CWH permanently closed": "Seller/CWH permanently closed",
  "Recipient unavailable.Establishment closed": "Client office found close",
  "Reject but package intact": "Client not share OTP",
  "Reject - RID not found": "Not traced in client system",
  "Barcode/QR mismatch": "Client rejected due to barcode/QR mismatch",
  "Content mismatch/missing - package tampered": "Client rejected due to content mismatch",
  "Short shipment": "Short shipment received by FE",
  "Recipient wants delivery at a different address": "Return address shifted (Need to new return address)"
};

const STATUS_MAP: Record<string, string> = {
  "pending": "pending",
  "dispatched": "dispatched",
  "dispatch": "dispatched",
  "rto": "rto",
  "dto": "dto",
  "delivered": "rto"
};

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

const normalizeAWB = (val: any): string => {
  if (val === null || val === undefined) return "";
  let s = String(val).trim();
  if (s.toLowerCase().includes('e+') || s.includes('.')) {
    const num = Number(s);
    if (!isNaN(num)) {
      s = String(Math.round(num));
    }
  }
  return s.replace('.0', '');
};

const isValidAWB = (val: any): boolean => {
  const s = normalizeAWB(val);
  return /^\d{8,}$/.test(s);
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
  returnAddress?: string;
  isIntact?: boolean;
}

interface OTPRow {
  id: string;
  awb: string;
  client: string;
  otpStatus: string;
  sessionStatus: string;
  returnAddress: string;
  isFTPL: boolean;
  isRTONotClosed: boolean;
  isDTONotClosed: boolean;
  isPendingNotClosed: boolean;
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
  const [activeTab, setActiveTab] = useState<"eod" | "remark" | "otp">("eod");
  const [sessions, setSessions] = useState<Session[]>([]);
  const [selectedSessionId, setSelectedSessionId] = useState<string | null>(null);
  const [statusFilter, setStatusFilter] = useState<string>("all");
  const [activeRemarkChip, setActiveRemarkChip] = useState<string | null>(null);
  const [selectedClients, setSelectedClients] = useState<string[]>([]);
  const [setupData, setSetupData] = useState({ feName: "", dspId: "", date: "" });
  const [isMounted, setIsMounted] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");
  const [selectedRowIds, setSelectedRowIds] = useState<Set<string>>(new Set());
  
  const [replacerData, setReplacerData] = useState<any[]>([]);
  const [replacerMeta, setReplacerMeta] = useState<{headers: string[], remarkKey: string} | null>(null);

  const [otpData, setOtpData] = useState<OTPRow[]>([]);
  const [otpStatusFilter, setOtpStatusFilter] = useState<string>("all");
  const [otpUploadError, setOtpUploadError] = useState<string | null>(null);

  // Initial mount - Starts fresh every time (no localStorage persistence)
  useEffect(() => {
    setIsMounted(true);
    setSetupData(prev => ({ ...prev, date: new Date().toISOString().split('T')[0] }));
  }, []);

  // Clear OTP data when session changes to maintain data integrity
  useEffect(() => {
    if (isMounted) {
      setOtpData([]);
      setOtpUploadError(null);
    }
  }, [selectedSessionId, isMounted]);

  // Reset remark chips when tab changes
  useEffect(() => {
    setActiveRemarkChip(null);
  }, [statusFilter]);

  const currentSession = useMemo(() => sessions.find(s => s.id === selectedSessionId) || null, [sessions, selectedSessionId]);

  const showToast = useCallback((msg: string, type: 'ok' | 'err') => {
    if (typeof document === 'undefined') return;
    const toast = document.createElement('div');
    toast.className = `fixed bottom-8 left-1/2 -translate-x-1/2 px-6 py-3 rounded-xl text-[13px] font-bold z-[1000] shadow-2xl transition-all duration-300 border ${
      type === 'ok' ? 'bg-slate-900 text-white border-white/10' : 'bg-rose-900 text-white border-rose-500/20'
    }`;
    toast.innerHTML = msg;
    document.body.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      setTimeout(() => toast.remove(), 300);
    }, 3000);
  }, []);

  const copyDataToClipboard = useCallback(async (rows: any[], headers: string[]) => {
    if (!rows.length) return;
    // Explicitly exclude Return Address from clipboard exports as per requirement
    const exportHeaders = headers.filter(h => h !== 'Return Address' && h !== 'Return_Address' && h !== 'return address');
    const plainText = rows.map(r => 
      exportHeaders.map(h => String(r[h] || "").trim()).join("\t")
    ).join("\n");
    const rowsHtml = rows.map(r => {
      const cells = exportHeaders.map(h => {
        const val = String(r[h] || "").trim();
        const style = h.toLowerCase().includes('awb') ? 'style=\'mso-number-format:"\\@"\'' : '';
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
    } catch (err) { return false; }
  }, []);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!setupData.feName || !setupData.dspId) {
      showToast("please enter dsp id and fe name first!", "err");
      e.target.value = "";
      return;
    }
    const file = e.target.files?.[0];
    if (!file) return;
    setIsProcessing(true);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array', raw: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(ws, { raw: true, defval: "" });
        const parsedRows: PODRow[] = rawData.map((row: any) => {
          const keys = Object.keys(row);
          const findVal = (regex: RegExp) => {
            const key = keys.find(k => regex.test(k.toLowerCase().replace(/[\s_-]/g, "")));
            return key ? row[key] : "";
          };
          const rawAwb = findVal(/waybill|awb|awbnumber/);
          if (!isValidAWB(rawAwb)) return null;
          const awb = normalizeAWB(rawAwb);
          const statusRaw = String(findVal(/status|currentstatus/)).toLowerCase().trim();
          const status = STATUS_MAP[statusRaw] || "unknown";
          const remark = String(findVal(/remark|remarks|nsl/)).trim();
          const returnAddress = String(findVal(/return_address|returnaddress/)).trim();
          return {
            id: crypto.randomUUID(),
            awb,
            client: String(findVal(/client|clientname/)),
            orderId: normalizeAWB(findVal(/order|orderid/)),
            status,
            remark: remark || "no remark",
            feName: setupData.feName,
            dspId: setupData.dspId,
            date: formatDate(setupData.date),
            returnAddress,
            isIntact: /reject|intact|barcode|content/i.test(remark)
          };
        }).filter((row): row is PODRow => row !== null && row.awb.length >= 8 && row.status !== "unknown");
        if (parsedRows.length === 0) throw new Error("no valid data found in eod report.");
        const newSessionId = crypto.randomUUID();
        const newSession: Session = { id: newSessionId, feName: setupData.feName, dspId: setupData.dspId, date: formatDate(setupData.date), data: parsedRows, timestamp: Date.now() };
        setSessions(prev => [newSession, ...prev]);
        setSelectedSessionId(newSessionId);
        showToast(`imported ${parsedRows.length} rows!`, "ok");
      } catch (err: any) { showToast(err.message || "failed to import file", "err"); } finally { setIsProcessing(false); }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const filteredRows = useMemo(() => {
    if (!currentSession) return [];
    let rows = currentSession.data;
    if (statusFilter !== 'all') rows = rows.filter(r => r.status === statusFilter);
    if (activeRemarkChip) rows = rows.filter(r => r.remark === activeRemarkChip);
    if (selectedClients.length > 0) rows = rows.filter(r => selectedClients.includes(r.client));
    if (searchTerm) {
      const s = searchTerm.toLowerCase();
      rows = rows.filter(r => r.awb.includes(s) || r.client.toLowerCase().includes(s));
    }
    return rows;
  }, [currentSession, statusFilter, activeRemarkChip, searchTerm, selectedClients]);

  const stats = useMemo(() => {
    if (!currentSession) return { total: 0, pending: 0, dispatched: 0, rto: 0, dto: 0 };
    return {
      total: currentSession.data.length,
      pending: currentSession.data.filter(r => r.status === 'pending').length,
      dispatched: currentSession.data.filter(r => r.status === 'dispatched').length,
      rto: currentSession.data.filter(r => r.status === 'rto').length,
      dto: currentSession.data.filter(r => r.status === 'dto').length,
    };
  }, [currentSession]);

  const handleCopyAWBOnly = async (rowsToCopy: any[]) => {
    if (!rowsToCopy.length) return;
    // Clean AWB copy without apostrophe as per user request
    const text = rowsToCopy.map(r => r.awb).join('\n');
    try {
      await navigator.clipboard.writeText(text);
      showToast(`copied ${rowsToCopy.length} awb`, "ok");
    } catch (err) { showToast("failed to copy awb", "err"); }
  };

  const handleCopyTable = useCallback(async (rowsToCopy: any[]) => {
    if (!rowsToCopy.length) return;
    const headers = ['date', 'dsp id', 'awb number', 'client', 'order id', 'remark', 'fe name'];
    const exportRows = rowsToCopy.map((r, i) => ({
      'date': formatDate(r.date),
      'dsp id': i === 0 ? r.dspId : "",
      'awb number': r.awb,
      'client': r.client,
      'order id': r.orderId,
      'remark': r.remark,
      'fe name': r.feName
    }));
    const success = await copyDataToClipboard(exportRows, headers);
    if (success) showToast(`copied ${rowsToCopy.length} rows to clipboard`, "ok");
  }, [copyDataToClipboard, showToast]);

  const downloadExcel = (rowsToDownload: any[]) => {
    if (!rowsToDownload.length) return;
    const header = ['date', 'dsp id', 'awb number', 'client', 'order id', 'remark', 'fe name'];
    const excelData = rowsToDownload.map((r, i) => [
      formatDate(r.date), 
      { v: i === 0 ? String(r.dspId) : "", t: 's' }, 
      { v: String(r.awb), t: 's' }, 
      r.client, 
      { v: String(r.orderId), t: 's' }, 
      r.remark, 
      r.feName
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...excelData]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "report");
    XLSX.writeFile(wb, `report_${currentSession?.dspId || 'export'}.xlsx`);
  };

  const handleReplacerUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsProcessing(true);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array', raw: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(ws, { raw: true, defval: "" });
        const headers = Object.keys(rawData[0] || {});
        const remarkKey = headers.find(h => /remark|nsl|reason/i.test(h)) || "";
        const replaced = rawData.map(row => {
          const original = String(row[remarkKey] || "").trim();
          const target = REMARK_MAPPING[original];
          return { ...row, [remarkKey]: target || original, __isReplaced: !!target, __id: crypto.randomUUID() };
        });
        setReplacerData(replaced);
        setReplacerMeta({ headers, remarkKey });
        showToast(`converted ${replaced.filter(r => r.__isReplaced).length} remarks!`, "ok");
      } catch (err) { showToast("replacer failed", "err"); } finally { setIsProcessing(false); }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const handleOTPFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (!selectedSessionId || !currentSession) {
      showToast("please select a session in daily eod rejection tab first.", "err");
      e.target.value = "";
      return;
    }
    setOtpUploadError(null);
    setIsProcessing(true);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array', raw: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(ws, { raw: true, defval: "" });
        
        const sessionMap = new Map<string, PODRow>();
        currentSession.data.forEach(r => sessionMap.set(r.awb, r));
        
        let matchCount = 0;
        const tempOtpData: OTPRow[] = [];
        const otpAWBs = new Set<string>();

        rawData.forEach((row: any) => {
          const keys = Object.keys(row);
          const findVal = (regex: RegExp) => {
            const key = keys.find(k => regex.test(k.toLowerCase().replace(/[\s_-]/g, "")));
            return key ? row[key] : "";
          };
          const rawAwb = findVal(/waybill|awb/i);
          if (!isValidAWB(rawAwb)) return;
          const awb = normalizeAWB(rawAwb);
          if (sessionMap.has(awb)) matchCount++;
          
          const client = String(findVal(/client/i)).trim();
          const otpStatusRaw = String(findVal(/status|currentstatus/i)).toLowerCase().trim();
          let otpStatus = 'unknown';
          if (otpStatusRaw.includes('dispatched') || otpStatusRaw.includes('dispatch')) otpStatus = 'dispatched';
          else if (otpStatusRaw.includes('rto')) otpStatus = 'rto';
          else if (otpStatusRaw.includes('dto')) otpStatus = 'dto';
          else if (otpStatusRaw.includes('pending')) otpStatus = 'pending';

          const sessionRow = sessionMap.get(awb);
          const csvStatus = sessionRow?.status || 'not found';
          
          tempOtpData.push({
            id: crypto.randomUUID(),
            awb, client, otpStatus, sessionStatus: csvStatus,
            returnAddress: sessionRow?.returnAddress || "",
            isFTPL: client.toUpperCase().includes('FTPL'),
            isRTONotClosed: otpStatus === 'dispatched' && csvStatus === 'rto',
            isDTONotClosed: otpStatus === 'dispatched' && csvStatus === 'dto',
            isPendingNotClosed: otpStatus === 'dispatched' && csvStatus === 'pending'
          });
          otpAWBs.add(awb);
        });

        if (matchCount === 0) {
          setOtpUploadError(`wrong file uploaded. no awb numbers match the current session. please upload the delhivery otp report for ${currentSession.feName} ${currentSession.dspId}.`);
          setIsProcessing(false);
          e.target.value = "";
          return;
        }

        // Add session-only pending rows
        currentSession.data.filter(r => r.status === 'pending' && !otpAWBs.has(r.awb)).forEach(sr => {
          tempOtpData.push({
            id: crypto.randomUUID(),
            awb: sr.awb, client: sr.client, otpStatus: 'not found', sessionStatus: 'pending',
            returnAddress: sr.returnAddress || "", isFTPL: sr.client.toUpperCase().includes('FTPL'),
            isRTONotClosed: false, isDTONotClosed: false, isPendingNotClosed: false
          });
        });

        setOtpData(tempOtpData);
        showToast(`imported ${tempOtpData.length} records.`, "ok");
      } catch (err) { showToast("failed to process otp report", "err"); } finally { setIsProcessing(false); }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const otpFilteredRows = useMemo(() => {
    let rows = otpData;
    if (otpStatusFilter !== 'all') {
      if (otpStatusFilter === 'dispatched') {
        // Genuiune dispatched + RTO/DTO not closed (Dual placement)
        rows = rows.filter(r => (r.otpStatus === 'dispatched' && !r.isPendingNotClosed));
      } else if (otpStatusFilter === 'rto') {
        // OTP RTO + RTO not closed
        rows = rows.filter(r => r.otpStatus === 'rto' || r.isRTONotClosed);
      } else if (otpStatusFilter === 'dto') {
        // OTP DTO + DTO not closed
        rows = rows.filter(r => r.otpStatus === 'dto' || r.isDTONotClosed);
      } else if (otpStatusFilter === 'pending') {
        // OTP Pending + Pending not closed + Session only pending
        rows = rows.filter(r => r.otpStatus === 'pending' || r.isPendingNotClosed || (r.otpStatus === 'not found' && r.sessionStatus === 'pending'));
      }
    }
    return rows;
  }, [otpData, otpStatusFilter]);

  const otpStats = useMemo(() => {
    return {
      total: otpData.length,
      dispatched: otpData.filter(r => (r.otpStatus === 'dispatched' && !r.isPendingNotClosed)).length,
      rto: otpData.filter(r => r.otpStatus === 'rto' || r.isRTONotClosed).length,
      dto: otpData.filter(r => r.otpStatus === 'dto' || r.isDTONotClosed).length,
      pending: otpData.filter(r => r.otpStatus === 'pending' || r.isPendingNotClosed || (r.otpStatus === 'not found' && r.sessionStatus === 'pending')).length,
    };
  }, [otpData]);

  const toggleRowSelection = (id: string) => {
    setSelectedRowIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const toggleGroupSelection = (rows: any[], checked: boolean) => {
    setSelectedRowIds(prev => {
      const next = new Set(prev);
      rows.forEach(r => { if (checked) next.add(r.id); else next.delete(r.id); });
      return next;
    });
  };

  if (!isMounted) return null;

  return (
    <div className="min-h-screen bg-[#f8fafc]">
      <header className="bg-[#0f172a] border-b border-white/5 px-6 h-14 flex items-center justify-between sticky top-0 z-[100] shadow-lg">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center text-white shadow-xl">
            <Truck className="w-5 h-5" />
          </div>
          <div>
            <h1 className="text-[16px] font-extrabold text-white tracking-tight leading-none">pod tool</h1>
            <p className="text-[10px] text-blue-400 font-bold tracking-widest leading-none mt-1">palam vihar rpc</p>
          </div>
        </div>
        <div className="flex gap-8 h-full">
          {['eod', 'remark', 'otp'].map(id => (
            <button key={id} onClick={() => setActiveTab(id as any)} className={cn("h-full px-1 text-[13px] font-semibold transition-all relative border-b-2", activeTab === id ? "text-white border-blue-500" : "text-slate-400 border-transparent hover:text-white")}>
              {id === 'eod' ? 'daily eod rejection' : id === 'remark' ? 'eod rejection remark' : 'otp dispatch check'}
            </button>
          ))}
        </div>
        <div className="text-[11px] font-black text-amber-400 tracking-widest">by ashu</div>
      </header>

      <main className="max-w-[1800px] mx-auto p-6 space-y-6">
        {activeTab === "eod" && (
          <>
            <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-6 space-y-6">
              <div className="flex items-center gap-2 mb-2">
                <Settings className="w-4 h-4 text-blue-600" />
                <h2 className="text-[13px] font-bold text-[#111827] tracking-tight">session setup</h2>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                <input type="text" value={setupData.dspId} onChange={e => setSetupData({...setupData, dspId: e.target.value.replace(/\D/g, '')})} className="bg-[#f9fafb] border-[1.5px] border-[#d1d5db] rounded-lg px-3.5 h-[42px] text-[14px] font-bold outline-none focus:border-blue-500" placeholder="dsp id" />
                <input type="text" value={setupData.feName} onChange={e => setSetupData({...setupData, feName: e.target.value})} className="bg-[#f9fafb] border-[1.5px] border-[#d1d5db] rounded-lg px-3.5 h-[42px] text-[14px] font-bold outline-none focus:border-blue-500" placeholder="fe name" />
                <input type="date" value={setupData.date} onChange={e => setSetupData({...setupData, date: e.target.value})} className="bg-[#f9fafb] border-[1.5px] border-[#d1d5db] rounded-lg px-3.5 h-[42px] text-[14px] font-bold outline-none focus:border-blue-500" />
              </div>
              <div className="border-2 border-dashed rounded-xl p-8 text-center bg-slate-50 hover:bg-white hover:border-blue-500 transition-all cursor-pointer relative">
                <input type="file" onChange={handleFileUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                <div className="space-y-3">
                  <FileSpreadsheet className="w-6 h-6 text-slate-400 mx-auto" />
                  <p className="text-sm font-black">import daily eod report</p>
                </div>
              </div>
            </div>

            {sessions.length > 0 && (
              <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm grid grid-cols-[repeat(auto-fill,minmax(250px,1fr))] gap-4">
                {sessions.map(s => (
                  <div key={s.id} onClick={() => setSelectedSessionId(s.id)} className={cn("p-4 border rounded-xl cursor-pointer transition-all", selectedSessionId === s.id ? "ring-2 ring-blue-500" : "hover:border-blue-300")}>
                    <p className="font-extrabold text-[14px] text-slate-900">{s.feName}</p>
                    <p className="text-[11px] text-slate-500 font-bold">{s.dspId} • {s.date}</p>
                    <div className="mt-2 flex gap-1">
                      <span className="text-[10px] font-black bg-slate-100 text-slate-600 px-2 py-0.5 rounded-lg border border-slate-200">{s.data.length} pkt</span>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {currentSession && (
              <div className="space-y-6">
                <div className="bg-white rounded-xl border shadow-sm flex divide-x overflow-hidden">
                  {[
                    {id: 'all', label: 'all', color: 'text-slate-900', val: stats.total},
                    {id: 'pending', label: 'pending', color: 'text-amber-600', val: stats.pending},
                    {id: 'dispatched', label: 'dispatched', color: 'text-blue-600', val: stats.dispatched},
                    {id: 'rto', label: 'rto', color: 'text-emerald-600', val: stats.rto},
                    {id: 'dto', label: 'dto', color: 'text-emerald-600', val: stats.dto}
                  ].map(t => (
                    <button key={t.id} onClick={() => setStatusFilter(t.id)} className={cn("flex-1 py-6 flex flex-col items-center group h-[100px] transition-all relative", statusFilter === t.id ? "bg-slate-50" : "hover:bg-slate-50/30")}>
                      <span className={cn("text-[32px] font-extrabold leading-none mb-1", t.color)}>{t.val}</span>
                      <span className="text-[13px] font-black">{t.label}</span>
                      {statusFilter === t.id && <div className={cn("absolute bottom-0 w-full h-[3px]", t.color.replace('text-', 'bg-'))} />}
                    </button>
                  ))}
                </div>

                {statusFilter === 'pending' && (
                  <div className="flex flex-wrap gap-2 animate-in fade-in slide-in-from-top-2">
                    {Array.from(new Set(currentSession.data.filter(r => r.status === 'pending').map(r => r.remark))).map(remark => (
                      <button key={`chip-${remark}`} onClick={() => setActiveRemarkChip(activeRemarkChip === remark ? null : remark)} className={cn("px-3 py-1.5 rounded-full text-[11px] font-black transition-all border shadow-sm tracking-tight", activeRemarkChip === remark ? "bg-blue-600 text-white border-blue-600" : "bg-white text-slate-600 border-slate-200 hover:border-blue-400")}>
                        {remark}
                      </button>
                    ))}
                  </div>
                )}

                <div className="flex gap-2 bg-white border rounded-xl p-1 shadow-sm items-center">
                  <button onClick={() => handleCopyAWBOnly(filteredRows)} className="h-9 px-4 bg-slate-900 text-white rounded-lg text-[11px] font-black tracking-wider flex items-center gap-2">
                    <Copy className="w-3.5 h-3.5" /> copy selected awb
                  </button>
                  <button onClick={() => handleCopyTable(filteredRows)} className="h-9 px-5 bg-blue-600 text-white rounded-lg text-[12px] font-black">copy table</button>
                  <button onClick={() => downloadExcel(filteredRows)} className="h-9 px-5 bg-emerald-600 text-white rounded-lg text-[12px] font-black">download excel</button>
                </div>

                <div className="bg-white rounded-xl border-[1.5px] border-slate-200 shadow-2xl overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-center border-collapse table-fixed">
                      <thead className="bg-[#0f172a] text-white h-11">
                        <tr key="main-header">
                          <th style={{width: '32px'}} className="px-2"><input type="checkbox" onChange={e => toggleGroupSelection(filteredRows, e.target.checked)} checked={filteredRows.length > 0 && filteredRows.every(r => selectedRowIds.has(r.id))} /></th>
                          <th style={{width: '80px'}} className="text-[11px] font-bold tracking-tight">dsp id</th>
                          <th style={{width: '140px'}} className="text-[11px] font-bold tracking-tight">awb number</th>
                          <th style={{width: '180px'}} className="text-[11px] font-bold tracking-tight">client</th>
                          <th style={{width: '120px'}} className="text-[11px] font-bold tracking-tight">order id</th>
                          <th style={{width: '180px'}} className="text-[11px] font-bold tracking-tight">remark</th>
                          <th style={{width: '300px'}} className="text-[11px] font-bold tracking-tight">return address</th>
                          <th style={{width: '100px'}} className="text-[11px] font-bold tracking-tight">fe name</th>
                        </tr>
                      </thead>
                      <tbody>
                        {Object.entries(filteredRows.reduce((acc: any, row) => {
                          if (!acc[row.client]) acc[row.client] = [];
                          acc[row.client].push(row);
                          return acc;
                        }, {})).map(([client, rows]: any) => (
                          <React.Fragment key={`group-frag-${client}`}>
                            <tr key={`group-banner-${client}`} className="bg-slate-800 text-white h-9">
                              <td colSpan={8} className="text-left px-4">
                                <div className="flex items-center justify-between">
                                  <span className="text-[10px] font-black tracking-[0.1em] text-amber-400">{client} — {rows.length} pkt</span>
                                  <button onClick={() => handleCopyAWBOnly(rows)} className="text-[9px] border border-white/20 px-2 py-0.5 rounded hover:bg-white/10 transition-all font-bold">copy awbs</button>
                                </div>
                              </td>
                            </tr>
                            {rows.map((row: any) => (
                              <tr key={`row-${row.id}`} className={cn("border-b hover:bg-blue-50/40 transition-colors", selectedRowIds.has(row.id) && "bg-blue-50/50")}>
                                <td className="px-2 py-2"><input type="checkbox" checked={selectedRowIds.has(row.id)} onChange={() => toggleRowSelection(row.id)} /></td>
                                <td className="px-2 py-2 text-[13px] font-bold text-slate-600">{row.dspId}</td>
                                <td className="px-2 py-2 text-[13px] font-bold font-mono text-blue-700 cursor-pointer hover:underline" onClick={() => { navigator.clipboard.writeText(row.awb); showToast("awb copied", "ok"); }}>{row.awb}</td>
                                <td className="px-2 py-2 text-[13px] font-semibold text-slate-800">{row.client}</td>
                                <td className="px-2 py-2 text-[13px] font-medium text-slate-500">{row.orderId}</td>
                                <td className="px-2 py-2">
                                  <span className={cn("px-2 py-0.5 rounded text-[10px] font-black border", row.isIntact ? "bg-rose-50 text-rose-700 border-rose-200" : "bg-amber-50 text-amber-700 border-amber-200")}>
                                    {row.remark}
                                  </span>
                                </td>
                                <td className="px-4 py-2 text-[12px] whitespace-normal break-words text-left min-w-[300px] font-medium leading-relaxed">{row.returnAddress}</td>
                                <td className="px-2 py-2 text-[13px] font-bold text-slate-700">{row.feName}</td>
                              </tr>
                            ))}
                          </React.Fragment>
                        ))}
                        {filteredRows.length === 0 && (
                          <tr key="no-data-row"><td colSpan={8} className="py-20 text-slate-400 font-bold tracking-widest text-xs">no records found for this filter</td></tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}
          </>
        )}

        {activeTab === "remark" && (
          <div className="grid grid-cols-12 gap-8">
            <div className="col-span-8 space-y-6">
              <div className="bg-white rounded-xl p-8 border shadow-sm text-center border-dashed border-2 hover:border-blue-500 cursor-pointer relative transition-all">
                <input type="file" onChange={handleReplacerUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                <FileSpreadsheet className="w-10 h-10 text-slate-300 mx-auto mb-4" />
                <p className="font-black text-slate-600 tracking-tight">drop eod sheet to convert remarks</p>
              </div>
              {replacerData.length > 0 && (
                <div className="bg-white rounded-xl border border-emerald-500/20 shadow-xl overflow-hidden">
                  <div className="overflow-x-auto max-h-[600px]">
                    <table className="w-full text-center border-collapse bg-white">
                      <thead className="sticky top-0 bg-[#0f172a] text-white h-12">
                        <tr key="replacer-head">
                          {replacerMeta?.headers.map((h, i) => <th key={`rep-h-${i}`} className="px-4 font-bold text-[10px] tracking-widest">{h}</th>)}
                          <th style={{width: '300px'}} key="rep-h-addr" className="px-4 font-bold text-[10px] tracking-widest">return address</th>
                        </tr>
                      </thead>
                      <tbody>
                        {replacerData.map(row => (
                          <tr key={row.__id} className={cn("h-11 border-b transition-colors", row.__isReplaced ? "bg-emerald-50/40" : "bg-amber-50/40")}>
                            {replacerMeta?.headers.map((h, i) => <td key={`rep-c-${i}`} className="px-4 font-semibold text-[13px] truncate max-w-[200px] font-medium">{row[h]}</td>)}
                            <td className="px-4 py-2 text-[12px] whitespace-normal break-words text-left min-w-[300px] font-medium">{row.Return_Address || row.ReturnAddress || ""}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
            <div className="col-span-4">
              <div className="bg-white rounded-xl border shadow-sm sticky top-20 overflow-hidden">
                <div className="bg-blue-600 p-4 text-white font-black text-[12px] tracking-widest">replacement matrix</div>
                <div className="p-4 space-y-4 max-h-[70vh] overflow-y-auto custom-scrollbar">
                  {Object.entries(REMARK_MAPPING).map(([nsl, off]) => (
                    <div key={`matrix-${nsl}`} className="p-4 bg-slate-50 rounded-xl border space-y-2 border-slate-200 group hover:border-blue-400 transition-all">
                      <p className="text-[12px] font-bold text-amber-700 tracking-tight">{nsl}</p>
                      <p className="text-[13px] font-black text-emerald-700">{off}</p>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        )}

        {activeTab === "otp" && (
          <div className="space-y-6">
            {!selectedSessionId ? (
              <div className="bg-white rounded-xl border p-20 text-center animate-in fade-in zoom-in">
                <AlertCircle className="w-12 h-12 text-amber-500 mx-auto mb-4" />
                <h3 className="text-lg font-black text-slate-900 mb-1">session not selected</h3>
                <p className="text-sm font-bold text-slate-400">please select a session in daily eod rejection tab first.</p>
              </div>
            ) : (
              <>
                <div className="bg-white border-[1.5px] border-slate-200 rounded-xl p-5 shadow-sm flex items-center justify-between">
                  <div className="flex items-center gap-4">
                    <div className="w-12 h-12 bg-slate-900 rounded-xl flex items-center justify-center text-white shadow-lg">
                      <User className="w-6 h-6" />
                    </div>
                    <div>
                      <p className="text-lg font-black text-slate-900 leading-tight">{currentSession.feName}</p>
                      <p className="text-[10px] font-black text-slate-400 tracking-widest">{currentSession.dspId} • {currentSession.date}</p>
                    </div>
                  </div>
                  <div className="flex gap-2">
                    <span className="px-3 py-1.5 rounded-lg bg-slate-100 text-[10px] font-black text-slate-600 border border-slate-200">total: {stats.total}</span>
                    <span className="px-3 py-1.5 rounded-lg bg-amber-50 text-[10px] font-black text-amber-600 border border-amber-200">pending: {stats.pending}</span>
                    <span className="px-3 py-1.5 rounded-lg bg-emerald-50 text-[10px] font-black text-emerald-600 border border-emerald-200">rto: {stats.rto}</span>
                    <span className="px-3 py-1.5 rounded-lg bg-emerald-50 text-[10px] font-black text-emerald-600 border border-emerald-200">dto: {stats.dto}</span>
                  </div>
                </div>

                <div className="bg-white rounded-xl border-[1.5px] border-dashed border-slate-300 p-12 text-center space-y-4 hover:border-blue-500 transition-all bg-slate-50/50">
                  <div className="max-w-xl mx-auto space-y-4">
                    <div className="w-16 h-16 bg-white rounded-2xl flex items-center justify-center text-blue-600 shadow-sm mx-auto border">
                      <Download className="w-8 h-8" />
                    </div>
                    <div>
                      <h2 className="text-xl font-black text-slate-900 tracking-tight">upload delhivery otp report</h2>
                      <p className="text-[12px] font-bold text-slate-400 tracking-wider">upload the default delhivery export file for this fe session.</p>
                    </div>
                    <div className="relative cursor-pointer group">
                      <input type="file" onChange={handleOTPFileUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                      <div className="h-14 bg-white border-2 border-slate-200 rounded-xl flex items-center justify-center font-black text-slate-600 group-hover:border-blue-500 group-hover:text-blue-600 transition-all">
                        select file
                      </div>
                    </div>
                    {otpUploadError && <p className="text-rose-600 font-black text-[12px] tracking-tight">{otpUploadError}</p>}
                  </div>
                </div>

                {otpData.length > 0 && (
                  <div className="space-y-6">
                    <div className="bg-white rounded-xl border shadow-sm flex divide-x overflow-hidden">
                      {[
                        {id: 'all', label: 'all', val: otpStats.total, color: 'text-slate-900'},
                        {id: 'dispatched', label: 'dispatched', val: otpStats.dispatched, color: 'text-rose-600'},
                        {id: 'rto', label: 'rto', val: otpStats.rto, color: 'text-emerald-600'},
                        {id: 'dto', label: 'dto', val: otpStats.dto, color: 'text-emerald-600'},
                        {id: 'pending', label: 'pending', val: otpStats.pending, color: 'text-amber-600'}
                      ].map(t => (
                        <button key={`otp-tab-${t.id}`} onClick={() => setOtpStatusFilter(t.id)} className={cn("flex-1 py-6 flex flex-col items-center group h-[110px] transition-all relative", otpStatusFilter === t.id ? "bg-slate-50" : "hover:bg-slate-50/30")}>
                          <span className={cn("text-[36px] font-black leading-none mb-1", t.color)}>{t.val}</span>
                          <span className={cn("text-[13px] font-black tracking-widest", otpStatusFilter === t.id ? t.color : "text-slate-400")}>{t.label}</span>
                          {otpStatusFilter === t.id && <div className={cn("absolute bottom-0 w-full h-[4px]", t.color.replace('text-', 'bg-'))} />}
                        </button>
                      ))}
                    </div>

                    <div className="flex gap-2">
                      <button onClick={() => handleCopyTable(otpFilteredRows)} className="h-9 px-6 bg-slate-900 text-white rounded-xl text-[12px] font-black tracking-widest flex items-center gap-2 hover:bg-slate-800 transition-all">
                        <Copy className="w-4 h-4" /> copy table
                      </button>
                    </div>

                    <div className="bg-white rounded-xl border-[1.5px] border-slate-200 shadow-2xl overflow-hidden">
                      <div className="overflow-x-auto">
                        <table className="w-full text-center border-collapse table-fixed">
                          <thead className="bg-[#0f172a] text-white h-11">
                            <tr key="otp-main-header">
                              <th style={{width: '32px'}} className="px-2"><input type="checkbox" /></th>
                              <th style={{width: '150px'}} className="text-[11px] font-bold tracking-tight">awb number</th>
                              <th style={{width: '200px'}} className="text-[11px] font-bold tracking-tight">client name</th>
                              <th style={{width: '180px'}} className="text-[11px] font-bold tracking-tight">otp status</th>
                              <th style={{width: '180px'}} className="text-[11px] font-bold tracking-tight">session status</th>
                              <th style={{width: '350px'}} className="text-[11px] font-bold tracking-tight">return address</th>
                            </tr>
                          </thead>
                          <tbody>
                            {Object.entries(otpFilteredRows.reduce((acc: any, row) => {
                              if (!acc[row.client]) acc[row.client] = [];
                              acc[row.client].push(row);
                              return acc;
                        }, {})).sort((a: any, b: any) => b[1].length - a[1].length).map(([client, rows]: any) => (
                              <React.Fragment key={`otp-frag-${client}`}>
                                <tr key={`otp-banner-${client}`} className="bg-slate-800 text-white h-9">
                                  <td colSpan={6} className="text-left px-4">
                                    <div className="flex items-center justify-between">
                                      <span className="text-[10px] font-black tracking-[0.1em] text-amber-400">{client} — {rows.length} pkt</span>
                                      <button onClick={() => handleCopyAWBOnly(rows)} className="text-[9px] border border-white/20 px-2 py-0.5 rounded hover:bg-white/10 transition-all font-bold">copy awbs</button>
                                    </div>
                                  </td>
                                </tr>
                                {rows.map((row: any) => (
                                  <tr key={`otp-row-${row.id}`} className={cn(
                                    "border-b transition-all hover:bg-slate-50", 
                                    (row.isRTONotClosed || row.isDTONotClosed || row.isPendingNotClosed) ? "bg-amber-50/30 border-l-[4px] border-l-amber-500" : "bg-white",
                                    row.isFTPL && (row.otpStatus === 'dispatched' ? "bg-rose-50/50" : "bg-emerald-50/50")
                                  )}>
                                    <td className="px-2 py-2"><input type="checkbox" /></td>
                                    <td className="px-4 py-2 text-[13px] font-mono font-black text-blue-700 cursor-pointer hover:underline" onClick={() => { navigator.clipboard.writeText(row.awb); showToast("awb copied", "ok"); }}>{row.awb}</td>
                                    <td className={cn("px-4 py-2 text-[13px] font-black tracking-tight", row.isFTPL && (row.otpStatus === 'dispatched' ? "text-rose-600" : "text-emerald-600"))}>{row.client}</td>
                                    <td className="px-4 py-2">
                                      <div className="flex flex-col items-center gap-1">
                                        <span className={cn(
                                          "px-2.5 py-0.5 rounded text-[10px] font-black border shadow-sm", 
                                          row.otpStatus === 'dispatched' ? "bg-rose-600 text-white border-rose-500" : 
                                          row.otpStatus === 'pending' ? "bg-amber-500 text-white border-amber-400" : 
                                          "bg-emerald-600 text-white border-emerald-500"
                                        )}>
                                          {row.otpStatus}
                                        </span>
                                        {(row.isRTONotClosed || row.isDTONotClosed || row.isPendingNotClosed) && (
                                          <span className="px-1.5 py-0.5 rounded bg-amber-100 text-amber-800 text-[9px] font-black border border-amber-300 shadow-sm">
                                            {row.isRTONotClosed ? 'rto — not closed' : row.isDTONotClosed ? 'dto — not closed' : 'pending — not closed'}
                                          </span>
                                        )}
                                      </div>
                                    </td>
                                    <td className="px-4 py-2">
                                      <span className={cn("px-2 py-0.5 rounded text-[10px] font-black border", row.sessionStatus === 'pending' ? "bg-amber-50 text-amber-700 border-amber-200" : row.sessionStatus === 'rto' ? "bg-emerald-50 text-emerald-700 border-emerald-200" : "bg-slate-50 text-slate-700 border-slate-200")}>
                                        {row.sessionStatus}
                                      </span>
                                    </td>
                                    <td className="px-4 py-2 text-[12px] whitespace-normal break-words text-left min-w-[350px] font-medium leading-relaxed">{row.returnAddress}</td>
                                  </tr>
                                ))}
                              </React.Fragment>
                            ))}
                            {otpFilteredRows.length === 0 && (
                              <tr key="no-otp-matches-row"><td colSpan={6} className="py-20 font-black text-slate-300 tracking-widest text-xs">no matching data in this filter</td></tr>
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                )}
              </>
            )}
          </div>
        )}
      </main>
    </div>
  );
}
