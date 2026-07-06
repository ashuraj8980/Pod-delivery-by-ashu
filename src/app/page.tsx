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
 * Restored with Colorful UI, Client Grouping, Dual Tab Placement, and Non-Exportable Return Address.
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
  logicCase: number; // 0: Normal, 1: RTO-NC, 2: DTO-NC, 3: P-NC
  hasMismatch: boolean;
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

  useEffect(() => {
    setIsMounted(true);
    setSetupData(prev => ({ ...prev, date: new Date().toISOString().split('T')[0] }));
    const saved = localStorage.getItem('pod_master_v1');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        setSessions(parsed);
        if (parsed.length > 0) setSelectedSessionId(parsed[0].id);
      } catch (e) {}
    }
  }, []);

  useEffect(() => {
    if (isMounted) {
      localStorage.setItem('pod_master_v1', JSON.stringify(sessions));
    }
  }, [sessions, isMounted]);

  // Reset OTP module when session changes
  useEffect(() => {
    if (isMounted) {
      setOtpData([]);
      setOtpUploadError(null);
      setSelectedRowIds(new Set());
    }
  }, [selectedSessionId, isMounted]);

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
    // Strictly exclude Return Address from exports as requested
    const exportHeaders = headers.filter(h => h !== 'Return Address');
    
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
    } catch (err) {
      return false;
    }
  }, []);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!setupData.feName || !setupData.dspId) {
      showToast("Enter DSP ID and FE Name first!", "err");
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
            remark: remark || "No Remark",
            feName: setupData.feName,
            dspId: setupData.dspId,
            date: formatDate(setupData.date),
            returnAddress,
            isIntact: /reject|intact|barcode|content/i.test(remark)
          };
        }).filter((row): row is PODRow => row !== null && row.awb.length >= 8 && row.status !== "unknown");
        
        if (parsedRows.length === 0) throw new Error("No valid data found.");
        
        const newSessionId = crypto.randomUUID();
        const newSession: Session = {
          id: newSessionId,
          feName: setupData.feName,
          dspId: setupData.dspId,
          date: formatDate(setupData.date),
          data: parsedRows,
          timestamp: Date.now()
        };
        
        setSessions(prev => [newSession, ...prev]);
        setSelectedSessionId(newSessionId);
        showToast(`Imported ${parsedRows.length} rows!`, "ok");
      } catch (err: any) {
        showToast(err.message || "Failed to import", "err");
      } finally {
        setIsProcessing(false);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const filteredRows = useMemo(() => {
    if (!currentSession) return [];
    let rows = currentSession.data;
    if (statusFilter !== 'all') {
      rows = rows.filter(r => r.status === statusFilter);
    }
    if (activeRemarkChip) rows = rows.filter(r => r.remark === activeRemarkChip);
    if (selectedClients.length > 0) {
      rows = rows.filter(r => selectedClients.includes(r.client));
    }
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

  const handleCopyAWBOnly = async () => {
    if (!filteredRows.length) return;
    const text = filteredRows.map(r => r.awb).join('\n');
    try {
      await navigator.clipboard.writeText(text);
      showToast(`Copied ${filteredRows.length} AWB`, "ok");
    } catch (err) {
      showToast("Failed to copy AWB", "err");
    }
  };

  const handleCopyAllAWB = async () => {
    if (!filteredRows.length) return;
    const text = filteredRows.map(r => r.awb).join('\n');
    try {
      await navigator.clipboard.writeText(text);
      showToast(`Copied ${filteredRows.length} AWB`, "ok");
    } catch (err) {
      showToast("Failed to copy AWBs", "err");
    }
  };

  const handleCopyTable = useCallback(async () => {
    if (!filteredRows.length) return;
    const headers = ['Date', 'DSP ID', 'AWB Number', 'Client', 'Order ID', 'Remark', 'FE Name', 'Return Address'];
    const exportRows = filteredRows.map((r, i) => ({
      'Date': formatDate(r.date),
      'DSP ID': i === 0 ? r.dspId : "",
      'AWB Number': r.awb,
      'Client': r.client,
      'Order ID': r.orderId,
      'Remark': r.remark,
      'FE Name': r.feName,
      'Return Address': r.returnAddress
    }));
    const success = await copyDataToClipboard(exportRows, headers);
    if (success) showToast(`Copied ${filteredRows.length} rows to clipboard`, "ok");
  }, [filteredRows, copyDataToClipboard, showToast]);

  const downloadExcel = () => {
    if (!filteredRows.length) return;
    const header = ['Date', 'DSP ID', 'AWB Number', 'Client', 'Order ID', 'Remark', 'FE Name'];
    const excelData = filteredRows.map((r, i) => [
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
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, `Report_${currentSession?.dspId || 'Export'}.xlsx`);
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
          return {
            ...row,
            [remarkKey]: target || original,
            __isReplaced: !!target,
            __id: crypto.randomUUID()
          };
        });
        
        setReplacerData(replaced);
        setReplacerMeta({ headers, remarkKey });
        showToast(`Converted ${replaced.filter(r => r.__isReplaced).length} remarks!`, "ok");
      } catch (err) {
        showToast("Replacer failed", "err");
      } finally {
        setIsProcessing(false);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const handleOTPFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (!selectedSessionId || !currentSession) {
      showToast("Select a session in Daily EOD Rejection tab first!", "err");
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

        const newOtpData: OTPRow[] = [];
        const processedAWBs = new Set<string>();
        let matchedCount = 0;

        rawData.forEach((row: any) => {
          const keys = Object.keys(row);
          const findVal = (regex: RegExp) => {
            const key = keys.find(k => regex.test(k.toLowerCase().replace(/[\s_-]/g, "")));
            return key ? row[key] : "";
          };
          const rawAwb = findVal(/waybill|awb/i);
          if (!isValidAWB(rawAwb)) return;

          const awb = normalizeAWB(rawAwb);
          if (sessionMap.has(awb)) matchedCount++;

          const client = String(findVal(/client/i)).trim();
          const rawOtpStatus = String(findVal(/status|currentstatus/i)).trim().toLowerCase();
          
          let otpStatus = 'unknown';
          if (rawOtpStatus.includes('dispatched') || rawOtpStatus.includes('dispatch')) otpStatus = 'dispatched';
          else if (rawOtpStatus.includes('rto')) otpStatus = 'rto';
          else if (rawOtpStatus.includes('dto')) otpStatus = 'dto';
          else if (rawOtpStatus.includes('pending')) otpStatus = 'pending';

          const sessionRow = sessionMap.get(awb);
          const csvStatus = sessionRow?.status || 'NOT FOUND';
          
          let lCase = 0; 
          if (otpStatus === 'dispatched') {
            if (csvStatus === 'rto') lCase = 1;
            else if (csvStatus === 'dto') lCase = 2;
            else if (csvStatus === 'pending') lCase = 3;
          }

          newOtpData.push({
            id: crypto.randomUUID(),
            awb,
            client,
            otpStatus,
            sessionStatus: csvStatus,
            returnAddress: sessionRow?.returnAddress || "",
            isFTPL: client.toUpperCase().includes('FTPL'),
            logicCase: lCase,
            hasMismatch: lCase !== 0
          });
          processedAWBs.add(awb);
        });

        if (matchedCount === 0) {
          setOtpUploadError(`Wrong file uploaded. No AWB numbers match for ${currentSession.feName} ${currentSession.dspId}.`);
          setIsProcessing(false);
          e.target.value = "";
          return;
        }

        // Add session-only pending rows
        currentSession.data.filter(r => r.status === 'pending' && !processedAWBs.has(r.awb)).forEach(sr => {
          newOtpData.push({
            id: crypto.randomUUID(),
            awb: sr.awb,
            client: sr.client,
            otpStatus: 'NOT FOUND',
            sessionStatus: 'pending',
            returnAddress: sr.returnAddress || "",
            isFTPL: sr.client.toUpperCase().includes('FTPL'),
            logicCase: 0,
            hasMismatch: true
          });
        });

        setOtpData(newOtpData);
        showToast(`Imported ${newOtpData.length} records.`, "ok");
      } catch (err) {
        showToast("Failed to process OTP Report", "err");
      } finally {
        setIsProcessing(false);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const otpFilteredRows = useMemo(() => {
    let rows = otpData;
    if (otpStatusFilter !== 'all') {
      if (otpStatusFilter === 'dispatched') {
        rows = rows.filter(r => 
          (r.otpStatus === 'dispatched' && (r.sessionStatus === 'dispatched' || r.sessionStatus === 'NOT FOUND')) ||
          r.logicCase === 1 || 
          r.logicCase === 2
        );
      } else if (otpStatusFilter === 'rto') {
        rows = rows.filter(r => r.otpStatus === 'rto' || r.logicCase === 1);
      } else if (otpStatusFilter === 'dto') {
        rows = rows.filter(r => r.otpStatus === 'dto' || r.logicCase === 2);
      } else if (otpStatusFilter === 'pending') {
        rows = rows.filter(r => 
          r.otpStatus === 'pending' || 
          r.logicCase === 3 || 
          (r.sessionStatus === 'pending' && r.otpStatus === 'NOT FOUND')
        );
      }
    }
    return rows;
  }, [otpData, otpStatusFilter]);

  const otpStats = useMemo(() => {
    return {
      total: otpData.length,
      dispatched: otpData.filter(r => (r.otpStatus === 'dispatched' && (r.sessionStatus === 'dispatched' || r.sessionStatus === 'NOT FOUND')) || r.logicCase === 1 || r.logicCase === 2).length,
      rto: otpData.filter(r => r.otpStatus === 'rto' || r.logicCase === 1).length,
      dto: otpData.filter(r => r.otpStatus === 'dto' || r.logicCase === 2).length,
      pending: otpData.filter(r => r.otpStatus === 'pending' || r.logicCase === 3 || (r.sessionStatus === 'pending' && r.otpStatus === 'NOT FOUND')).length,
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
      rows.forEach(r => {
        if (checked) next.add(r.id);
        else next.delete(r.id);
      });
      return next;
    });
  };

  if (!isMounted) return null;

  return (
    <div className="min-h-screen bg-[#F8FAFC]">
      <header className="bg-[#0F172A] border-b border-white/5 px-6 h-14 flex items-center justify-between sticky top-0 z-[100] shadow-lg">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center text-white shadow-xl">
            <Truck className="w-5 h-5" />
          </div>
          <div>
            <h1 className="text-[16px] font-extrabold text-white tracking-tight leading-none">POD Tool</h1>
            <p className="text-[10px] text-blue-400 font-bold uppercase tracking-widest leading-none mt-1">Palam Vihar RPC</p>
          </div>
        </div>
        <div className="flex gap-8 h-full">
          {['eod', 'remark', 'otp'].map(id => (
            <button 
              key={id}
              onClick={() => setActiveTab(id as any)}
              className={cn(
                "h-full px-1 text-[13px] font-semibold transition-all relative border-b-2 capitalize",
                activeTab === id ? "text-white border-blue-500" : "text-slate-400 border-transparent hover:text-white"
              )}
            >
              {id === 'eod' ? 'Daily EOD Rejection' : id === 'remark' ? 'EOD Rejection Remark' : 'OTP Dispatch Check'}
            </button>
          ))}
        </div>
        <div className="text-[11px] font-black text-amber-400 uppercase tracking-widest">By Ashu</div>
      </header>

      <main className="max-w-[1800px] mx-auto p-6 space-y-6">
        {activeTab === "eod" && (
          <>
            <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-6 space-y-6">
              <div className="flex items-center gap-2 mb-2">
                <Settings className="w-4 h-4 text-blue-600" />
                <h2 className="text-[13px] font-bold text-[#111827] tracking-tight">Session Setup</h2>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                <input type="text" value={setupData.dspId} onChange={e => setSetupData({...setupData, dspId: e.target.value.replace(/\D/g, '')})} className="bg-[#F9FAFB] border-[1.5px] border-[#D1D5DB] rounded-lg px-3.5 h-[42px] text-[14px] font-bold outline-none focus:border-blue-500" placeholder="DSP ID" />
                <input type="text" value={setupData.feName} onChange={e => setSetupData({...setupData, feName: e.target.value})} className="bg-[#F9FAFB] border-[1.5px] border-[#D1D5DB] rounded-lg px-3.5 h-[42px] text-[14px] font-bold outline-none focus:border-blue-500" placeholder="FE Name" />
                <input type="date" value={setupData.date} onChange={e => setSetupData({...setupData, date: e.target.value})} className="bg-[#F9FAFB] border-[1.5px] border-[#D1D5DB] rounded-lg px-3.5 h-[42px] text-[14px] font-bold outline-none focus:border-blue-500" />
              </div>
              <div className="border-2 border-dashed rounded-xl p-8 text-center bg-slate-50 hover:bg-white hover:border-blue-500 transition-all cursor-pointer relative">
                <input type="file" onChange={handleFileUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                <div className="space-y-3">
                  <FileSpreadsheet className="w-6 h-6 text-slate-400 mx-auto" />
                  <p className="text-sm font-black">Upload Daily EOD Report</p>
                </div>
              </div>
            </div>

            {sessions.length > 0 && (
              <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm grid grid-cols-[repeat(auto-fill,minmax(250px,1fr))] gap-4">
                {sessions.map(s => (
                  <div key={s.id} onClick={() => setSelectedSessionId(s.id)} className={cn("p-4 border rounded-xl cursor-pointer transition-all", selectedSessionId === s.id ? "ring-2 ring-blue-500" : "hover:border-blue-300")}>
                    <p className="font-extrabold text-[14px]">{s.feName}</p>
                    <p className="text-[11px] text-slate-500 uppercase">{s.dspId} • {s.date}</p>
                    <div className="mt-2 flex gap-1">
                      <span className="text-[10px] font-bold bg-slate-100 px-1.5 py-0.5 rounded">{s.data.length} pkt</span>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {currentSession && (
              <div className="space-y-6">
                <div className="bg-white rounded-xl border shadow-sm flex divide-x overflow-hidden">
                  {[
                    {id: 'all', label: 'All', color: 'text-slate-900'},
                    {id: 'pending', label: 'Pending', color: 'text-amber-600'},
                    {id: 'dispatched', label: 'Dispatched', color: 'text-blue-600'},
                    {id: 'rto', label: 'RTO', color: 'text-emerald-600'},
                    {id: 'dto', label: 'DTO', color: 'text-emerald-600'}
                  ].map(t => (
                    <button key={t.id} onClick={() => setStatusFilter(t.id)} className={cn("flex-1 py-6 flex flex-col items-center group h-[100px] transition-all relative", statusFilter === t.id ? "bg-slate-50" : "hover:bg-slate-50/30")}>
                      <span className={cn("text-[32px] font-extrabold leading-none mb-1 capitalize", t.color)}>{(stats as any)[t.id] || stats.total}</span>
                      <span className="text-[13px] font-bold uppercase">{t.label}</span>
                      {statusFilter === t.id && <div className={cn("absolute bottom-0 w-full h-[3px]", t.color.replace('text-', 'bg-'))} />}
                    </button>
                  ))}
                </div>

                <div className="flex gap-2 bg-white border rounded-xl p-1 shadow-sm items-center">
                  <button onClick={handleCopyAWBOnly} className="h-8 px-4 bg-slate-900 text-white rounded-lg text-[11px] font-black uppercase tracking-wider flex items-center gap-2">
                    <Copy className="w-3.5 h-3.5" /> Copy Selected AWB
                  </button>
                  <button onClick={handleCopyAllAWB} className="h-8 px-4 bg-indigo-600 text-white rounded-lg text-[11px] font-black uppercase tracking-wider flex items-center gap-2">
                    <Copy className="w-3.5 h-3.5" /> Copy All AWB
                  </button>
                  <button onClick={handleCopyTable} className="h-10 px-5 bg-blue-600 text-white rounded-lg text-[13px] font-bold">Copy Table</button>
                  <button onClick={downloadExcel} className="h-10 px-5 bg-emerald-600 text-white rounded-lg text-[13px] font-bold">Download Excel</button>
                </div>

                <div className="bg-white rounded-xl border-[1.5px] border-[#F97316] shadow-2xl overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-center border-collapse table-fixed">
                      <thead className="bg-[#0F172A] text-white h-11">
                        <tr>
                          <th style={{width: '32px'}} className="px-2"><input type="checkbox" onChange={e => toggleGroupSelection(filteredRows, e.target.checked)} checked={filteredRows.every(r => selectedRowIds.has(r.id))} /></th>
                          <th style={{width: '80px'}} className="text-[11px] font-bold">DSP ID</th>
                          <th style={{width: '130px'}} className="text-[11px] font-bold">AWB Number</th>
                          <th style={{width: '110px'}} className="text-[11px] font-bold">Client</th>
                          <th style={{width: '110px'}} className="text-[11px] font-bold">Order ID</th>
                          <th style={{width: '150px'}} className="text-[11px] font-bold">Remark</th>
                          <th style={{width: '250px'}} className="text-[11px] font-bold">Return Address</th>
                          <th style={{width: '80px'}} className="text-[11px] font-bold">FE Name</th>
                        </tr>
                      </thead>
                      <tbody>
                        {Object.entries(filteredRows.reduce((acc: any, row) => {
                          if (!acc[row.client]) acc[row.client] = [];
                          acc[row.client].push(row);
                          return acc;
                        }, {})).map(([client, rows]: any) => (
                          <React.Fragment key={client}>
                            <tr key={`group-${client}`} className="bg-slate-800 text-white h-8">
                              <td colSpan={8} className="text-left px-4 text-[10px] font-bold uppercase tracking-widest">{client} — {rows.length} PKT</td>
                            </tr>
                            {rows.map((row: any) => (
                              <tr key={row.id} className={cn("border-b hover:bg-blue-50/30", selectedRowIds.has(row.id) && "bg-blue-50")}>
                                <td className="px-2 py-2"><input type="checkbox" checked={selectedRowIds.has(row.id)} onChange={() => toggleRowSelection(row.id)} /></td>
                                <td className="px-2 py-2 text-[13px] font-bold">{row.dspId}</td>
                                <td className="px-2 py-2 text-[13px] font-bold font-mono cursor-pointer hover:underline" onClick={() => {
                                  navigator.clipboard.writeText(row.awb);
                                  showToast("AWB Copied", "ok");
                                }}>{row.awb}</td>
                                <td className="px-2 py-2 text-[13px] font-semibold text-blue-600">{row.client}</td>
                                <td className="px-2 py-2 text-[13px] font-medium">{row.orderId}</td>
                                <td className="px-2 py-2">
                                  <span className={cn("px-2 py-0.5 rounded text-[10px] font-bold uppercase border", row.isIntact ? "bg-rose-50 text-rose-700 border-rose-200" : "bg-amber-50 text-amber-700 border-amber-200")}>
                                    {row.remark}
                                  </span>
                                </td>
                                <td className="px-4 py-2 text-[12px] whitespace-normal break-words text-left min-w-[250px]">{row.returnAddress}</td>
                                <td className="px-2 py-2 text-[13px] font-bold">{row.feName}</td>
                              </tr>
                            ))}
                          </React.Fragment>
                        ))}
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
              <div className="bg-white rounded-xl p-8 border shadow-sm text-center border-dashed border-2 hover:border-blue-500 cursor-pointer relative">
                <input type="file" onChange={handleReplacerUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                <FileSpreadsheet className="w-10 h-10 text-slate-300 mx-auto mb-4" />
                <p className="font-black">Drop EOD Sheet to Convert Remarks</p>
              </div>
              {replacerData.length > 0 && (
                <div className="bg-white rounded-xl border border-emerald-500/20 shadow-xl overflow-hidden">
                  <div className="overflow-x-auto max-h-[600px]">
                    <table className="w-full text-center border-collapse bg-white">
                      <thead className="sticky top-0 bg-[#0F172A] text-white h-12">
                        <tr>
                          {replacerMeta?.headers.map((h, i) => <th key={i} className="px-4 font-bold text-[10px] uppercase tracking-widest">{h}</th>)}
                          <th style={{width: '250px'}} className="px-4 font-bold text-[10px] uppercase tracking-widest">Return Address</th>
                        </tr>
                      </thead>
                      <tbody>
                        {replacerData.map(row => (
                          <tr key={row.__id} className={cn("h-11 border-b", row.__isReplaced ? "bg-emerald-50/40" : "bg-amber-50/40")}>
                            {replacerMeta?.headers.map((h, i) => <td key={i} className="px-4 font-semibold truncate max-w-[200px]">{row[h]}</td>)}
                            <td className="px-4 py-2 text-[12px] whitespace-normal break-words text-left min-w-[250px]">{row.Return_Address || row.ReturnAddress || ""}</td>
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
                <div className="bg-blue-600 p-4 text-white font-black text-[12px] uppercase">Replacement Matrix</div>
                <div className="p-4 space-y-4 max-h-[70vh] overflow-y-auto">
                  {Object.entries(REMARK_MAPPING).map(([nsl, off]) => (
                    <div key={nsl} className="p-4 bg-slate-50 rounded-xl border space-y-2">
                      <p className="text-[13px] font-bold text-amber-700">{nsl}</p>
                      <p className="text-[13px] font-extrabold text-emerald-700">{off}</p>
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
              <div className="bg-white rounded-xl border p-8 text-center"><AlertCircle className="w-8 h-8 text-amber-600 mx-auto mb-3" /><p className="font-bold">Please select a session in Daily EOD Rejection tab first.</p></div>
            ) : (
              <div className="space-y-6">
                {/* Session Card Info at Top */}
                <div className="bg-white border rounded-xl p-4 shadow-sm flex items-center justify-between">
                  <div className="flex items-center gap-4">
                    <User className="w-5 h-5" />
                    <div>
                      <p className="font-black">{currentSession.feName}</p>
                      <p className="text-[10px] font-bold text-slate-400 uppercase">{currentSession.dspId} • {currentSession.date}</p>
                    </div>
                  </div>
                  <div className="flex gap-2">
                    <span className="px-2 py-1 rounded bg-slate-100 text-[10px] font-black uppercase">Total: {stats.total}</span>
                    <span className="px-2 py-1 rounded bg-amber-50 text-[10px] font-black uppercase text-amber-600">Pending: {stats.pending}</span>
                    <span className="px-2 py-1 rounded bg-emerald-50 text-[10px] font-black uppercase text-emerald-600">RTO: {stats.rto}</span>
                    <span className="px-2 py-1 rounded bg-emerald-50 text-[10px] font-black uppercase text-emerald-600">DTO: {stats.dto}</span>
                  </div>
                </div>

                {/* Upload Zone */}
                <div className="bg-white rounded-xl border p-8 text-center space-y-4">
                  <h2 className="text-xl font-extrabold">Upload Delhivery OTP Report</h2>
                  <p className="text-[12px] font-bold text-slate-400">Upload the default Delhivery export file for this FE session.</p>
                  <div className={cn("border-2 border-dashed rounded-xl p-10 bg-slate-50 hover:bg-white hover:border-blue-500 transition-all cursor-pointer relative mx-auto max-w-2xl", isProcessing && "opacity-50")}>
                    <input type="file" onChange={handleOTPFileUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                    <Download className="w-8 h-8 text-slate-300 mx-auto mb-2" />
                    <p className="text-sm font-black">Click to select OTP Report</p>
                  </div>
                  {otpUploadError && <p className="text-rose-600 font-bold text-[12px]">{otpUploadError}</p>}
                </div>

                {otpData.length > 0 && (
                  <div className="space-y-6">
                    <div className="bg-white rounded-xl border shadow-sm flex divide-x overflow-hidden">
                      {[
                        {id: 'all', label: 'All', val: otpStats.total, color: 'text-blue-600'},
                        {id: 'dispatched', label: 'Dispatched', val: otpStats.dispatched, color: 'text-[#DC2626]'},
                        {id: 'rto', label: 'RTO', val: otpStats.rto, color: 'text-[#16A34A]'},
                        {id: 'dto', label: 'DTO', val: otpStats.dto, color: 'text-[#16A34A]'},
                        {id: 'pending', label: 'Pending', val: otpStats.pending, color: 'text-[#D97706]'}
                      ].map(t => (
                        <button key={t.id} onClick={() => setOtpStatusFilter(t.id)} className={cn("flex-1 py-6 flex flex-col items-center group h-[100px] transition-all relative", otpStatusFilter === t.id ? "bg-slate-50" : "hover:bg-slate-50/30")}>
                          <span className={cn("text-[32px] font-extrabold", t.color)}>{t.val}</span>
                          <span className={cn("text-[13px] font-bold uppercase", otpStatusFilter === t.id ? t.color : "text-slate-400")}>{t.label}</span>
                          {otpStatusFilter === t.id && <div className={cn("absolute bottom-0 w-full h-[3px]", t.color.replace('text-', 'bg-'))} />}
                        </button>
                      ))}
                    </div>

                    <div className="flex gap-2">
                      <button onClick={async () => {
                        const headers = ['AWB Number', 'Client Name', 'OTP Status', 'Session Status'];
                        const data = otpFilteredRows.map(r => ({'AWB Number': r.awb, 'Client Name': r.client, 'OTP Status': r.otpStatus.toUpperCase(), 'Session Status': r.sessionStatus.toUpperCase()}));
                        const success = await copyDataToClipboard(data, headers);
                        if (success) showToast(`Copied ${otpFilteredRows.length} rows`, "ok");
                      }} className="h-10 px-5 bg-blue-600 text-white rounded-lg text-[13px] font-bold">Copy Table</button>
                    </div>

                    <div className="bg-white rounded-xl border-[1.5px] border-[#F97316] shadow-2xl overflow-hidden">
                      <div className="overflow-x-auto">
                        <table className="w-full text-center border-collapse table-fixed">
                          <thead className="bg-[#0F172A] text-white h-11">
                            <tr key="otp-head">
                              <th style={{width: '32px'}} className="px-2"><input type="checkbox" /></th>
                              <th style={{width: '150px'}} className="text-[11px] font-bold">AWB Number</th>
                              <th style={{width: '180px'}} className="text-[11px] font-bold">Client Name</th>
                              <th style={{width: '160px'}} className="text-[11px] font-bold">OTP Status</th>
                              <th style={{width: '160px'}} className="text-[11px] font-bold">Session Status</th>
                              <th style={{width: '250px'}} className="text-[11px] font-bold">Return Address</th>
                            </tr>
                          </thead>
                          <tbody>
                            {Object.entries(otpFilteredRows.reduce((acc: any, row) => {
                              if (!acc[row.client]) acc[row.client] = [];
                              acc[row.client].push(row);
                              return acc;
                            }, {})).sort((a: any, b: any) => b[1].length - a[1].length).map(([client, rows]: any) => (
                              <React.Fragment key={client}>
                                <tr key={`otp-group-${client}`} className="bg-slate-800 text-white h-8">
                                  <td colSpan={6} className="text-left px-4">
                                    <div className="flex items-center justify-between">
                                      <span className="text-amber-400 font-bold uppercase text-[10px] tracking-widest">{client} — {rows.length} PKT</span>
                                      <button onClick={() => {
                                        const text = rows.map((r: any) => r.awb).join('\n');
                                        navigator.clipboard.writeText(text);
                                        showToast(`Copied ${rows.length} AWBs`, "ok");
                                      }} className="text-[9px] border border-white/20 px-2 py-0.5 rounded hover:bg-white/10 transition-all font-bold uppercase">Copy AWBs</button>
                                    </div>
                                  </td>
                                </tr>
                                {rows.map((row: any) => (
                                  <tr key={row.id} className={cn(
                                    "border-b transition-all", 
                                    row.hasMismatch ? "bg-[#FFF7ED] border-l-[3px] border-l-amber-500" : "bg-white",
                                    row.isFTPL && (row.otpStatus === 'dispatched' ? "bg-[#FFF5F5] border-l-[3px] border-l-red-600" : "bg-[#F0FDF4]")
                                  )}>
                                    <td className="px-2 py-2"><input type="checkbox" /></td>
                                    <td className="px-4 py-2 text-[13px] font-mono font-bold cursor-pointer hover:underline" onClick={() => {
                                      navigator.clipboard.writeText(row.awb);
                                      showToast("AWB Copied", "ok");
                                    }}>{row.awb}</td>
                                    <td className={cn(
                                      "px-4 py-2 text-[13px] font-semibold", 
                                      row.isFTPL && (row.otpStatus === 'dispatched' ? "text-red-600" : "text-emerald-600")
                                    )}>{row.client}</td>
                                    <td className="px-4 py-2">
                                      <div className="flex flex-col items-center gap-1">
                                        <span className={cn(
                                          "px-2 py-0.5 rounded text-[10px] font-black uppercase border", 
                                          row.otpStatus === 'dispatched' ? "bg-[#DC2626] text-white" : 
                                          row.otpStatus === 'pending' ? "bg-[#D97706] text-white" : 
                                          "bg-[#16A34A] text-white"
                                        )}>
                                          {row.otpStatus}
                                        </span>
                                        {row.logicCase > 0 && (
                                          <span className="px-1.5 py-0.5 rounded bg-amber-100 text-amber-800 text-[9px] font-black uppercase border border-amber-200">
                                            {row.logicCase === 1 ? 'RTO — Not Closed' : row.logicCase === 2 ? 'DTO — Not Closed' : 'Pending — Not Closed'}
                                          </span>
                                        )}
                                      </div>
                                    </td>
                                    <td className="px-4 py-2">
                                      <span className={cn(
                                        "px-2 py-0.5 rounded text-[10px] font-black uppercase border",
                                        row.sessionStatus === 'pending' ? "bg-amber-50 text-amber-700 border-amber-200" : 
                                        row.sessionStatus === 'rto' ? "bg-emerald-50 text-emerald-700 border-emerald-200" : 
                                        "bg-slate-50 text-slate-700 border-slate-200"
                                      )}>
                                        {row.sessionStatus}
                                      </span>
                                    </td>
                                    <td className="px-4 py-2 text-[12px] whitespace-normal break-words text-left min-w-[250px]">{row.returnAddress}</td>
                                  </tr>
                                ))}
                              </React.Fragment>
                            ))}
                            {otpFilteredRows.length === 0 && (
                              <tr key="no-otp-matches"><td colSpan={6} className="py-10 font-bold text-slate-400">No matching data in this filter</td></tr>
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        )}
      </main>
    </div>
  );
}
