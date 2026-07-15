"use client";

import React, { useState, useEffect, useMemo, useCallback, Suspense, useRef } from "react";
import { 
  Truck, 
  Copy, 
  FileSpreadsheet, 
  Settings,
  Download,
  User,
  AlertCircle,
  X,
  ArrowLeft,
  Trash2,
  RotateCcw,
  Search,
  IndianRupee,
  CheckCircle2
} from "lucide-react";
import Link from "next/link";
import * as XLSX from "xlsx";
import { cn } from "@/lib/utils";
import { useSearchParams } from "next/navigation";
import { getData, saveData } from "@/lib/storage";

/**
 * @fileOverview Delhivery POD Management Tool - EOD Rejection Page
 * Optimized for heavy Excel files with High Value (4000+) Pending shipment tracking.
 */

const REMARK_MAPPING: Record<string, string> = {
  "Incomplete address & contact details": "Return Address Not Found (Need New Contact Number)",
  "On Hold. Recipient unable to Accept Delivery": "Client Out Of Station (Receive Shipment After 3 Days)",
  "On Hold. Recipient unable to Accept Delivery for 5 days": "Client Out Of Station (Receive Shipment After 3 Days)",
  "Not Attempted": "Not Attempted To Client",
  "Seller/CWH permanently closed": "Seller/CWH Permanently Closed",
  "Recipient unavailable.Establishment closed": "Client Office Found Closed",
  "Reject but package intact": "Client Not Shared OTP",
  "Reject - RID not found": "Not Traced In Client System",
  "Barcode/QR mismatch": "Client Rejected Due To Barcode/QR Mismatch",
  "Content mismatch/missing - package tampered": "Client Rejected Due To Content Mismatch",
  "Short shipment": "Short Shipment Received By FE",
  "Recipient wants delivery at a different address": "Return Address Shifted (Need New Return Address)"
};

const STATUS_MAP: Record<string, string> = {
  "pending": "Pending",
  "dispatched": "Dispatched",
  "dispatch": "Dispatched",
  "rto": "RTO",
  "dto": "DTO",
  "delivered": "RTO"
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
  if (s.startsWith("'")) s = s.substring(1);
  if (s.toLowerCase().includes('e+') || s.includes('.')) {
    const num = Number(s);
    if (!isNaN(num)) s = String(Math.round(num));
  }
  return s.replace('.0', '');
};

const isValidAWB = (val: any): string => {
  const s = normalizeAWB(val);
  return (/^\d{8,}$/.test(s) && !isNaN(Number(s))) ? s : "";
};

interface PODRow {
  id: string;
  awb: string;
  client: string;
  orderId: string;
  status: string;
  remark: string;
  amount: number;
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
  isNotClosed: boolean;
  notClosedType: 'RTO' | 'DTO' | 'Pending' | null;
}

interface Session {
  id: string;
  feName: string;
  dspId: string;
  date: string;
  time: string;
  data: PODRow[];
  timestamp: number;
  stats?: {
    total: number;
    pending: number;
    dispatched: number;
    rto: number;
    dto: number;
    highValue: number;
  };
}

function PODToolContent() {
  const searchParams = useSearchParams();
  const [activeTab, setActiveTab] = useState<"eod" | "remark" | "otp">("eod");
  const [sessions, setSessions] = useState<Session[]>([]);
  const [selectedSessionId, setSelectedSessionId] = useState<string | null>(null);
  const [statusFilter, setStatusFilter] = useState<string>("All");
  const [searchTerm, setSearchTerm] = useState("");
  
  const [selectedRemarkChips, setSelectedRemarkChips] = useState<string[]>([]);
  const [clientFilter, setClientFilter] = useState<string>("All Clients");
  
  const [showAllPending, setShowAllPending] = useState(false);
  const [setupData, setSetupData] = useState({ feName: "", dspId: "", date: "" });
  const [isMounted, setIsMounted] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [uploadProgress, setUploadProgress] = useState<{current: number, total: number} | null>(null);
  const [selectedRowIds, setSelectedRowIds] = useState<Set<string>>(new Set());
  
  const [undoStack, setUndoStack] = useState<PODRow[][]>([]);
  const [replacerData, setReplacerData] = useState<any[]>([]);
  const [replacerMeta, setReplacerMeta] = useState<{headers: string[], remarkKey: string} | null>(null);

  const [otpData, setOtpData] = useState<OTPRow[]>([]);
  const [otpStatusFilter, setOtpStatusFilter] = useState<string>("All");
  const [otpClientFilter, setOtpClientFilter] = useState<string>("All Clients");

  const sessionsRef = useRef(sessions);
  useEffect(() => {
    sessionsRef.current = sessions;
  }, [sessions]);

  const loadSessionsFromDB = useCallback(async () => {
    try {
      const saved = await getData('pod_sessions');
      if (Array.isArray(saved)) {
        setSessions(saved);
        const sid = searchParams.get('sessionId');
        if (sid && saved.some(s => s.id === sid)) {
          setSelectedSessionId(sid);
        }
      }
    } catch (e) {}
  }, [searchParams]);

  useEffect(() => {
    setIsMounted(true);
    setSetupData(prev => ({ ...prev, date: new Date().toISOString().split('T')[0] }));
    loadSessionsFromDB();
  }, [loadSessionsFromDB]);

  useEffect(() => {
    if (isMounted) {
      saveData('pod_sessions', sessions);
    }
  }, [sessions, isMounted]);

  const currentSession = useMemo(() => sessions.find(s => s.id === selectedSessionId) || null, [sessions, selectedSessionId]);

  const stats = useMemo(() => {
    if (!currentSession) return { total: 0, pending: 0, dispatched: 0, rto: 0, dto: 0, highValue: 0 };
    return {
      total: currentSession.data.length,
      pending: currentSession.data.filter(r => r.status === 'Pending').length,
      dispatched: currentSession.data.filter(r => r.status === 'Dispatched').length,
      rto: currentSession.data.filter(r => r.status === 'RTO').length,
      dto: currentSession.data.filter(r => r.status === 'DTO').length,
      highValue: currentSession.data.filter(r => r.status === 'Pending' && (r.amount ?? 0) >= 4000).length,
    };
  }, [currentSession]);

  const otpStats = useMemo(() => {
    // Logic: Tabs count follows Session EOD CSV status exactly to match the EOD session card.
    const rto = otpData.filter(r => r.sessionStatus === 'RTO').length;
    const dto = otpData.filter(r => r.sessionStatus === 'DTO').length;
    const pending = otpData.filter(r => r.sessionStatus === 'Pending').length;
    // Dispatched in OTP but RTO/DTO in CSV (Not Closed)
    const dispatched = otpData.filter(r => r.otpStatus === 'Dispatched' && (r.sessionStatus === 'RTO' || r.sessionStatus === 'DTO')).length;

    return { total: otpData.length, dispatched, pending, rto, dto };
  }, [otpData]);

  const filteredOtpRows = useMemo(() => {
    let rows = otpData;
    if (otpStatusFilter === 'Dispatched') {
      // Show shipments that are Dispatched in OTP but RTO/DTO in CSV
      rows = rows.filter(r => r.otpStatus === 'Dispatched' && (r.sessionStatus === 'RTO' || r.sessionStatus === 'DTO'));
    } else if (otpStatusFilter === 'Pending') {
      rows = rows.filter(r => r.sessionStatus === 'Pending');
    } else if (otpStatusFilter === 'RTO') {
      rows = rows.filter(r => r.sessionStatus === 'RTO');
    } else if (otpStatusFilter === 'DTO') {
      rows = rows.filter(r => r.sessionStatus === 'DTO');
    }
    
    if (otpClientFilter !== 'All Clients') {
      rows = rows.filter(r => r.client === otpClientFilter);
    }
    return rows;
  }, [otpData, otpStatusFilter, otpClientFilter]);

  const uniqueOtpClients = useMemo(() => {
    return Array.from(new Set(filteredOtpRows.map(r => r.client))).sort();
  }, [filteredOtpRows]);

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

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!setupData.feName || !setupData.dspId) {
      showToast("Please Enter FE Name and DSP ID First", "err");
      return;
    }
    const file = e.target.files?.[0];
    if (!file) return;
    setIsProcessing(true);
    setUploadProgress(null);
    setUndoStack([]); 

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array', raw: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rawRows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        
        if (rawRows.length < 2) throw new Error("File is empty or missing data.");

        const headerRow = rawRows[0].map(h => String(h).toLowerCase().replace(/[\s_-]/g, ""));
        const findIdx = (regex: RegExp) => headerRow.findIndex(h => regex.test(h));

        const awbIdx = findIdx(/waybill|awb|awbnumber/);
        const statusIdx = findIdx(/status|currentstatus/);
        const remarkIdx = findIdx(/remark|remarks|nsl/);
        const clientIdx = findIdx(/client|clientname/);
        const orderIdx = findIdx(/order|orderid/);
        const addressIdx = findIdx(/return_address|returnaddress/);
        const amountIdx = findIdx(/amount|value|price|collectible/);

        if (awbIdx === -1) throw new Error("Could not find AWB/Waybill column.");

        const CHUNK_SIZE = 300;
        let currentIndex = 1; 
        const totalRows = rawRows.length;
        const parsedRows: PODRow[] = [];

        const processNextChunk = () => {
          try {
            const end = Math.min(currentIndex + CHUNK_SIZE, totalRows);
            for (let i = currentIndex; i < end; i++) {
              const row = rawRows[i];
              const rawAwb = row[awbIdx];
              const awb = isValidAWB(rawAwb);
              if (!awb) continue;

              const statusRaw = statusIdx !== -1 ? String(row[statusIdx]).toLowerCase().trim() : "";
              const status = STATUS_MAP[statusRaw] || "Unknown";
              if (status === "Unknown") continue;

              const remark = remarkIdx !== -1 ? String(row[remarkIdx]).trim() : "No Remark";
              const amountVal = amountIdx !== -1 ? parseFloat(String(row[amountIdx]).replace(/[^\d.]/g, '')) || 0 : 0;
              
              parsedRows.push({
                id: crypto.randomUUID(),
                awb,
                client: clientIdx !== -1 ? String(row[clientIdx]) : "Unknown",
                orderId: orderIdx !== -1 ? normalizeAWB(row[orderIdx]) : "",
                status,
                remark: remark || "No Remark",
                amount: amountVal,
                returnAddress: addressIdx !== -1 ? String(row[addressIdx]).trim() : "",
                isIntact: /reject|intact|barcode|content/i.test(remark)
              });
            }
            currentIndex = end;
            setUploadProgress({ current: end - 1, total: totalRows - 1 });

            if (currentIndex < totalRows) {
              setTimeout(processNextChunk, 0);
            } else {
              if (parsedRows.length === 0) throw new Error("No Valid Data Found In EOD Report.");
              
              const sessionStats = {
                total: parsedRows.length,
                pending: parsedRows.filter(r => r.status === 'Pending').length,
                dispatched: parsedRows.filter(r => r.status === 'Dispatched').length,
                rto: parsedRows.filter(r => r.status === 'RTO').length,
                dto: parsedRows.filter(r => r.status === 'DTO').length,
                highValue: parsedRows.filter(r => r.status === 'Pending' && (r.amount ?? 0) >= 4000).length,
              };

              const newSessionId = crypto.randomUUID();
              const creationTime = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: true });
              
              const newSession: Session = { 
                id: newSessionId, 
                feName: setupData.feName, 
                dspId: setupData.dspId, 
                date: formatDate(setupData.date), 
                time: creationTime,
                data: parsedRows, 
                timestamp: Date.now(),
                stats: sessionStats
              };

              setSessions(prev => {
                const filtered = prev.filter(s => s.dspId !== setupData.dspId || s.feName !== setupData.feName);
                return [newSession, ...filtered];
              });
              
              setSelectedSessionId(newSessionId);
              setOtpData([]); 
              showToast(`Imported ${parsedRows.length} Rows Successfully!`, "ok");
              setIsProcessing(false);
              setUploadProgress(null);
            }
          } catch (e: any) {
            showToast(e.message || "Error processing chunk", "err");
            setIsProcessing(false);
            setUploadProgress(null);
          }
        };

        processNextChunk();

      } catch (err: any) { 
        showToast(err.message || "Failed To Import File", "err"); 
        setIsProcessing(false); 
        setUploadProgress(null);
      }
    };
    reader.onerror = () => {
      showToast("Error reading file", "err");
      setIsProcessing(false);
      setUploadProgress(null);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const deleteRow = (rowId: string) => {
    if (!selectedSessionId) return;
    const currentSessions = sessionsRef.current;
    const sessionToUpdate = currentSessions.find(s => s.id === selectedSessionId);
    if (!sessionToUpdate) return;
    setUndoStack(prev => [...prev, [...sessionToUpdate.data]]);
    setSessions(prev => prev.map(s => {
      if (s.id === selectedSessionId) {
        const newData = s.data.filter(r => r.id !== rowId);
        return { 
          ...s, 
          data: newData,
          stats: {
            total: newData.length,
            pending: newData.filter(r => r.status === 'Pending').length,
            dispatched: newData.filter(r => r.status === 'Dispatched').length,
            rto: newData.filter(r => r.status === 'RTO').length,
            dto: newData.filter(r => r.status === 'DTO').length,
            highValue: newData.filter(r => r.status === 'Pending' && (r.amount ?? 0) >= 4000).length,
          }
        };
      }
      return s;
    }));
    setSelectedRowIds(prev => {
      const next = new Set(prev);
      next.delete(rowId);
      return next;
    });
  };

  const handleUndo = () => {
    if (undoStack.length === 0 || !selectedSessionId) return;
    const previousData = undoStack[undoStack.length - 1];
    setSessions(prev => prev.map(s => {
      if (s.id === selectedSessionId) {
        return {
          ...s,
          data: previousData,
          stats: {
            total: previousData.length,
            pending: previousData.filter(r => r.status === 'Pending').length,
            dispatched: previousData.filter(r => r.status === 'Dispatched').length,
            rto: previousData.filter(r => r.status === 'RTO').length,
            dto: previousData.filter(r => r.status === 'DTO').length,
            highValue: previousData.filter(r => r.status === 'Pending' && (r.amount ?? 0) >= 4000).length,
          }
        };
      }
      return s;
    }));
    setUndoStack(prev => prev.slice(0, -1));
    showToast("Reverted last deletion(s)", "ok");
  };

  const handleDeleteSelected = () => {
    if (selectedRowIds.size === 0 || !selectedSessionId) return;
    const currentSessions = sessionsRef.current;
    const sessionToUpdate = currentSessions.find(s => s.id === selectedSessionId);
    if (!sessionToUpdate) return;
    setUndoStack(prev => [...prev, [...sessionToUpdate.data]]);
    setSessions(prev => prev.map(s => {
      if (s.id === selectedSessionId) {
        const newData = s.data.filter(r => !selectedRowIds.has(r.id));
        return { 
          ...s, 
          data: newData,
          stats: {
            total: newData.length,
            pending: newData.filter(r => r.status === 'Pending').length,
            dispatched: newData.filter(r => r.status === 'Dispatched').length,
            rto: newData.filter(r => r.status === 'RTO').length,
            dto: newData.filter(r => r.status === 'DTO').length,
            highValue: newData.filter(r => r.status === 'Pending' && (r.amount ?? 0) >= 4000).length,
          }
        };
      }
      return s;
    }));
    showToast(`Deleted ${selectedRowIds.size} Selected Rows`, "ok");
    setSelectedRowIds(new Set());
  };

  const uniqueClients = useMemo(() => {
    if (!currentSession) return [];
    let rows = currentSession.data;
    if (statusFilter === 'HighValue') {
      rows = rows.filter(r => r.status === 'Pending' && (r.amount ?? 0) >= 4000);
    } else if (statusFilter !== 'All') {
      rows = rows.filter(r => r.status === statusFilter);
    }
    return Array.from(new Set(rows.map(r => r.client))).sort();
  }, [currentSession, statusFilter]);

  const filteredRows = useMemo(() => {
    if (!currentSession) return [];
    let rows = currentSession.data;
    if (statusFilter === 'HighValue') {
      rows = rows.filter(r => r.status === 'Pending' && (r.amount ?? 0) >= 4000);
    } else if (statusFilter !== 'All') {
      rows = rows.filter(r => r.status === statusFilter);
    }
    if (statusFilter === 'Pending' && !showAllPending && selectedRemarkChips.length > 0) {
      rows = rows.filter(r => selectedRemarkChips.includes(r.remark));
    }
    if (clientFilter !== 'All Clients') {
      rows = rows.filter(r => r.client === clientFilter);
    }
    if (searchTerm.trim() !== "") {
      const s = searchTerm.toLowerCase().trim();
      rows = rows.filter(r => 
        normalizeAWB(r.awb).toLowerCase().includes(s) || 
        r.orderId.toLowerCase().includes(s) || 
        r.client.toLowerCase().includes(s)
      );
    }
    return rows;
  }, [currentSession, statusFilter, selectedRemarkChips, clientFilter, showAllPending, searchTerm]);

  const isAllSelected = filteredRows.length > 0 && filteredRows.every(r => selectedRowIds.has(r.id));
  const isSomeSelected = filteredRows.some(r => selectedRowIds.has(r.id)) && !isAllSelected;

  const handleSelectAll = () => {
    if (isAllSelected) {
      const next = new Set(selectedRowIds);
      filteredRows.forEach(r => next.delete(r.id));
      setSelectedRowIds(next);
    } else {
      const next = new Set(selectedRowIds);
      filteredRows.forEach(r => next.add(r.id));
      setSelectedRowIds(next);
    }
  };

  const handleCopyTable = useCallback(async (rowsToCopy: any[]) => {
    if (!rowsToCopy.length) return;
    const headers = ['Date', 'DSP ID', 'Waybill Number', 'Client', 'Order ID', 'Remark', 'Amount', 'FE Name'];
    const exportRows = rowsToCopy.map((r, i) => ({
      'Date': formatDate(currentSession?.date),
      'DSP ID': i === 0 ? currentSession?.dspId : "",
      'Waybill Number': normalizeAWB(r.awb),
      'Client': r.client,
      'Order ID': r.orderId,
      'Remark': r.remark,
      'Amount': r.amount ?? 0,
      'FE Name': currentSession?.feName
    }));
    
    const plainText = exportRows.map(r => headers.map(h => String((r as any)[h] || "").trim()).join("\t")).join("\n");
    try {
      await navigator.clipboard.writeText(plainText);
      showToast(`Copied ${rowsToCopy.length} Rows To Clipboard`, "ok");
    } catch (e) {}
  }, [showToast, currentSession]);

  const handleCopyAWBOnly = async (rowsToCopy: any[]) => {
    if (!rowsToCopy.length) return;
    const text = rowsToCopy.map(r => normalizeAWB(r.awb)).join('\n');
    try {
      await navigator.clipboard.writeText(text);
      showToast(`Copied ${rowsToCopy.length} AWB Numbers`, "ok");
    } catch (err) { showToast("Failed To Copy AWBs", "err"); }
  };

  const downloadExcel = (rowsToDownload: any[]) => {
    if (!rowsToDownload.length) return;
    const header = ['Date', 'DSP ID', 'Waybill Number', 'Client', 'Order ID', 'Remark', 'Amount', 'FE Name'];
    const excelData = rowsToDownload.map((r, i) => [
      formatDate(currentSession?.date), 
      i === 0 ? String(currentSession?.dspId) : "", 
      String(normalizeAWB(r.awb)), 
      r.client, 
      String(r.orderId), 
      r.remark, 
      r.amount ?? 0,
      currentSession?.feName
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...excelData]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, `Report_${currentSession?.dspId || 'Export'}.xlsx`);
  };

  const deleteSession = (sessionId: string) => {
    setSessions(prev => prev.filter(s => s.id !== sessionId));
    if (selectedSessionId === sessionId) {
      setSelectedSessionId(null);
      setUndoStack([]);
      setOtpData([]);
    }
    showToast("Session Deleted", "ok");
  };

  const handleSessionClick = (s: Session) => {
    setSelectedSessionId(s.id);
    setOtpData([]); 
    setStatusFilter("All");
    setSearchTerm("");
    setSelectedRemarkChips([]);
    setClientFilter("All Clients");
    setShowAllPending(false);
    setUndoStack([]); 
    let isoDate = "";
    if (s.date && s.date.includes('-')) {
      const parts = s.date.split('-');
      if (parts.length === 3) isoDate = `${parts[2]}-${parts[1]}-${parts[0]}`;
    }
    setSetupData({ feName: s.feName, dspId: s.dspId, date: isoDate || s.date });
  };

  if (!isMounted) return null;

  return (
    <div className="min-h-screen bg-[#f8fafc]">
      <header className="bg-[#0f172a] border-b border-white/5 px-6 h-14 flex items-center justify-between sticky top-0 z-[100] shadow-lg">
        <div className="flex items-center gap-3">
          <Link href="/" className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center text-white shadow-xl hover:bg-blue-700 transition-colors">
            <ArrowLeft className="w-5 h-5" />
          </Link>
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center text-white shadow-xl">
              <Truck className="w-5 h-5" />
            </div>
            <div>
              <h1 className="text-[16px] font-extrabold text-white tracking-tight leading-none">POD Tool</h1>
              <p className="text-[10px] text-blue-400 font-bold tracking-widest leading-none mt-1">Palam Vihar RPC</p>
            </div>
          </div>
        </div>
        <div className="flex gap-8 h-full">
          {['eod', 'remark', 'otp'].map(id => (
            <button key={id} onClick={() => setActiveTab(id as any)} className={cn("h-full px-1 text-[13px] font-semibold transition-all relative border-b-2", activeTab === id ? "text-white border-blue-500" : "text-slate-400 border-transparent hover:text-white")}>
              {id === 'eod' ? 'Daily EOD Rejection' : id === 'remark' ? 'EOD Rejection Remark' : 'OTP Dispatch Check'}
            </button>
          ))}
        </div>
        <div className="text-[11px] font-black text-amber-400 tracking-widest uppercase">By Ashu</div>
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
                <input type="text" value={setupData.dspId} onChange={e => setSetupData({...setupData, dspId: e.target.value.replace(/\D/g, '')})} className="bg-[#f9fafb] border-[1.5px] border-[#d1d5db] rounded-lg px-3.5 h-[42px] text-[14px] font-bold outline-none focus:border-blue-500" placeholder="DSP ID" />
                <input type="text" value={setupData.feName} onChange={e => setSetupData({...setupData, feName: e.target.value})} className="bg-[#f9fafb] border-[1.5px] border-[#d1d5db] rounded-lg px-3.5 h-[42px] text-[14px] font-bold outline-none focus:border-blue-500" placeholder="FE Name" />
                <input type="date" value={setupData.date} onChange={e => setSetupData({...setupData, date: e.target.value})} className="bg-[#f9fafb] border-[1.5px] border-[#d1d5db] rounded-lg px-3.5 h-[42px] text-[14px] font-bold outline-none focus:border-blue-500" />
              </div>
              <div className={cn("border-2 border-dashed rounded-xl p-8 text-center transition-all relative", (!setupData.feName || !setupData.dspId || isProcessing) ? "bg-slate-100 border-slate-200 cursor-not-allowed opacity-60" : "bg-slate-50 hover:bg-white hover:border-blue-500 cursor-pointer")}>
                {setupData.feName && setupData.dspId && !isProcessing && <input type="file" onChange={handleFileUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />}
                <div className="space-y-3">
                  <FileSpreadsheet className={cn("w-6 h-6 mx-auto", isProcessing ? "text-blue-600 animate-spin" : "text-slate-400")} />
                  <p className="text-sm font-black text-slate-600 tracking-tight">
                    {isProcessing ? (uploadProgress ? `Processing... ${uploadProgress.current} of ${uploadProgress.total}` : "Reading File...") : (!setupData.feName || !setupData.dspId) ? "Enter DSP ID And FE Name To Upload" : "Import Daily EOD Report"}
                  </p>
                </div>
              </div>
            </div>

            {sessions.length > 0 && (
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-6 space-y-4">
                <div className="flex items-center justify-between mb-2">
                  <h3 className="text-sm font-black text-slate-700 tracking-tight">Recent Sessions</h3>
                  <button 
                    onClick={() => { setSessions([]); setSelectedSessionId(null); setUndoStack([]); setOtpData([]); }} 
                    className="text-[11px] font-black text-rose-600 hover:text-rose-700 transition-colors uppercase tracking-widest"
                  >
                    CLEAR ALL SESSIONS
                  </button>
                </div>
                <div className="grid grid-cols-[repeat(auto-fill,minmax(280px,1fr))] gap-4">
                  {sessions.map(s => (
                    <div key={s.id} onClick={() => handleSessionClick(s)} className={cn("relative p-4 border-[1.5px] rounded-2xl cursor-pointer transition-all shadow-sm overflow-hidden bg-white", selectedSessionId === s.id ? "border-blue-500 ring-2 ring-blue-500/10" : "hover:border-blue-300 border-slate-100")}>
                      <div className="absolute left-0 top-0 bottom-0 w-1.5 bg-blue-600" />
                      <button onClick={(e) => { e.stopPropagation(); deleteSession(s.id); }} className="absolute right-3 top-3 text-slate-300 hover:text-rose-500 transition-colors p-1">
                        <X className="w-4 h-4" />
                      </button>
                      <p className="text-[15px] font-bold text-slate-900 mb-0.5 tracking-tight">{s.feName}</p>
                      <p className="text-[11px] text-slate-400 font-bold mb-3 uppercase tracking-wider">{s.dspId} — {s.date} — {s.time || ""}</p>
                      <div className="flex flex-wrap gap-[5px]">
                        <span className="text-[10px] font-bold bg-slate-50 text-slate-600 px-2 py-0.5 rounded-[4px] border border-slate-100 uppercase">{s.stats?.total || 0} PKT</span>
                        {s.stats?.pending !== undefined && s.stats.pending > 0 && <span className="text-[10px] font-bold bg-amber-50 text-amber-600 px-2 py-0.5 rounded-[4px] border border-amber-100 uppercase">{s.stats.pending} PENDING</span>}
                        {s.stats?.highValue !== undefined && s.stats.highValue > 0 && <span className="text-[10px] font-bold bg-rose-600 text-white px-2 py-0.5 rounded-[4px] border border-rose-700 uppercase">{s.stats.highValue} HIGH VALUE</span>}
                        {s.stats?.dispatched !== undefined && s.stats.dispatched > 0 && <span className="text-[10px] font-bold bg-rose-50 text-rose-600 px-2 py-0.5 rounded-[4px] border border-rose-100 uppercase">{s.stats.dispatched} DISPATCHED</span>}
                        {s.stats?.rto !== undefined && s.stats.rto > 0 && <span className="text-[10px] font-bold bg-emerald-50 text-emerald-600 px-2 py-0.5 rounded-[4px] border border-emerald-100 uppercase">{s.stats.rto} RTO</span>}
                        {s.stats?.dto !== undefined && s.stats.dto > 0 && <span className="text-[10px] font-bold bg-emerald-50 text-emerald-600 px-2 py-0.5 rounded-[4px] border border-emerald-100 uppercase">{s.stats.dto} DTO</span>}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {currentSession && (
              <div className="space-y-6">
                <div className="bg-white rounded-xl border shadow-sm flex divide-x overflow-hidden mt-6">
                  {[
                    {id: 'All', label: 'All', color: 'text-slate-900', bgColor: 'bg-[#EFF6FF]', borderColor: 'bg-blue-500', val: stats.total},
                    {id: 'HighValue', label: 'High Value', color: 'text-rose-600', bgColor: 'bg-rose-50', borderColor: 'bg-rose-600', val: stats.highValue},
                    {id: 'Pending', label: 'Pending', color: 'text-amber-600', bgColor: 'bg-[#FFFBEB]', borderColor: 'bg-amber-500', val: stats.pending},
                    {id: 'Dispatched', label: 'Dispatched', color: 'text-rose-600', bgColor: 'bg-[#FFF5F5]', borderColor: 'bg-rose-500', val: stats.dispatched},
                    {id: 'RTO', label: 'RTO', color: 'text-emerald-600', bgColor: 'bg-[#F0FDF4]', borderColor: 'bg-emerald-500', val: stats.rto},
                    {id: 'DTO', label: 'DTO', color: 'text-emerald-600', bgColor: 'bg-[#F0FDF4]', borderColor: 'bg-emerald-500', val: stats.dto}
                  ].map(t => (
                    <button key={t.id} onClick={() => { 
                      setStatusFilter(t.id); 
                      setSelectedRemarkChips([]); 
                      setClientFilter("All Clients");
                      setShowAllPending(false); 
                    }} className={cn("flex-1 py-6 flex flex-col items-center group h-[100px] transition-all relative", statusFilter === t.id ? t.bgColor : "hover:bg-slate-50/30")}>
                      <span className={cn("text-[32px] font-extrabold leading-none mb-1", t.color)}>{t.val}</span>
                      <span className="text-[13px] font-black uppercase tracking-widest">{t.label}</span>
                      {statusFilter === t.id && <div className={cn("absolute bottom-0 w-full h-[3px]", t.borderColor)} />}
                    </button>
                  ))}
                </div>

                {statusFilter === 'Pending' && (
                  <div className="flex flex-col gap-4">
                    <div className="flex items-center justify-between">
                      <div className="flex flex-wrap gap-3 items-center">
                        {Array.from(new Set(currentSession.data.filter(r => r.status === 'Pending').map(r => r.remark))).map(remark => {
                          const count = currentSession.data.filter(r => r.status === 'Pending' && r.remark === remark).length;
                          const isSelected = selectedRemarkChips.includes(remark);
                          return (
                            <button 
                              key={`chip-${remark}`} 
                              onClick={() => {
                                setSelectedRemarkChips(prev => 
                                  prev.includes(remark) ? prev.filter(c => c !== remark) : [...prev, remark]
                                );
                                setShowAllPending(false);
                              }} 
                              className={cn(
                                "inline-flex items-center gap-3 px-4 py-2 min-h-[36px] rounded-lg text-[13px] transition-all border shadow-sm", 
                                isSelected ? "bg-blue-600 text-white border-blue-600" : "bg-white text-slate-600 border-slate-200 hover:border-blue-400"
                              )}
                            >
                              <span className="font-semibold">{remark}</span>
                              <span className={cn("px-[10px] py-[2px] rounded-full text-[12px] font-bold border", isSelected ? "bg-white/20 border-white/30" : "bg-slate-100 border-slate-200 text-slate-600")}>{count}</span>
                            </button>
                          );
                        })}
                      </div>

                      {selectedRemarkChips.length > 0 && (
                        <button 
                          onClick={() => {
                            setSelectedRemarkChips([]);
                            setShowAllPending(true);
                          }}
                          className="bg-[#1C2333] text-white px-[14px] py-[6px] rounded-[8px] text-[11.5px] font-semibold whitespace-nowrap transition-all active:scale-95 shadow-sm"
                        >
                          ← Show All Pending
                        </button>
                      )}
                    </div>
                  </div>
                )}

                <div className="flex flex-wrap items-center gap-4 mb-2">
                  <div className="flex flex-col gap-1.5">
                    <label className="text-[11px] font-black text-slate-400 uppercase tracking-widest">Client Filter</label>
                    <select
                      value={clientFilter}
                      onChange={(e) => setClientFilter(e.target.value)}
                      className="border border-slate-200 rounded-lg px-3 py-2 text-[13px] font-bold bg-white outline-none focus:border-blue-500 w-[200px] shadow-sm"
                    >
                      <option value="All Clients">All Clients</option>
                      {uniqueClients.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>
                  
                  <div className="flex flex-col gap-1.5 flex-1 min-w-[300px]">
                    <label className="text-[11px] font-black text-slate-400 uppercase tracking-widest">Search Session</label>
                    <div className="relative">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                      <input 
                        type="text" 
                        placeholder="Search AWB, Order ID or Client..." 
                        value={searchTerm}
                        onChange={(e) => setSearchTerm(e.target.value)}
                        className="w-full bg-white border border-slate-200 rounded-lg pl-10 pr-4 h-[38px] text-[13px] font-bold outline-none focus:border-blue-500 shadow-sm transition-all"
                      />
                      {searchTerm && (
                        <button onClick={() => setSearchTerm("")} className="absolute right-3 top-1/2 -translate-y-1/2 p-0.5 hover:bg-slate-100 rounded-full transition-colors">
                          <X className="w-3.5 h-3.5 text-slate-400" />
                        </button>
                      )}
                    </div>
                  </div>
                </div>

                <div className="flex flex-wrap gap-2 bg-white border rounded-xl p-1 shadow-sm items-center">
                  <button onClick={() => handleCopyAWBOnly(filteredRows)} className="h-9 px-4 bg-slate-800 text-white rounded-lg text-[11px] font-black tracking-wider flex items-center gap-2 uppercase">
                    <Copy className="w-3.5 h-3.5" /> Copy All AWB
                  </button>
                  <button onClick={() => {
                    const rows = filteredRows.filter(r => selectedRowIds.has(r.id));
                    handleCopyAWBOnly(rows);
                  }} className="h-9 px-4 bg-slate-900 text-white rounded-lg text-[11px] font-black tracking-wider flex items-center gap-2 uppercase">
                    <Copy className="w-3.5 h-3.5" /> Copy Selected AWBs
                  </button>
                  
                  {selectedRowIds.size > 0 && (
                    <button 
                      onClick={handleDeleteSelected}
                      className="h-9 px-4 bg-rose-600 text-white rounded-lg text-[11px] font-black tracking-wider flex items-center gap-2 uppercase animate-in fade-in zoom-in duration-200"
                    >
                      <Trash2 className="w-3.5 h-3.5" /> Delete Selected ({selectedRowIds.size})
                    </button>
                  )}

                  {undoStack.length > 0 && (
                    <button 
                      onClick={handleUndo}
                      className="h-9 px-4 bg-blue-600 text-white rounded-lg text-[11px] font-black tracking-wider flex items-center gap-2 uppercase animate-in slide-in-from-left-4 duration-300"
                    >
                      <RotateCcw className="w-3.5 h-3.5" /> Undo Delete {undoStack.length > 0 ? `(${undoStack.length})` : ''}
                    </button>
                  )}
                  
                  <div className="flex-1" />
                  
                  <button onClick={() => downloadExcel(filteredRows)} className="h-9 px-5 bg-emerald-600 text-white rounded-lg text-[12px] font-black uppercase">Download Excel</button>
                  <button onClick={() => handleCopyTable(filteredRows)} className="h-9 px-5 bg-blue-600 text-white rounded-lg text-[12px] font-black uppercase">Copy Table</button>
                </div>

                <div className="bg-white rounded-2xl border-[1.5px] border-slate-200 shadow-2xl overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-center border-collapse table-fixed">
                      <thead className="bg-[#0f172a] text-white h-11">
                        <tr key="main-header">
                          <th style={{width: '32px'}} className="px-2">
                            <input 
                              type="checkbox" 
                              checked={isAllSelected}
                              ref={el => { if (el) el.indeterminate = isSomeSelected; }}
                              onChange={handleSelectAll}
                            />
                          </th>
                          <th style={{width: '40px'}} className="px-2"></th>
                          <th style={{width: '80px'}} className="text-[11px] font-bold tracking-widest uppercase">DSP ID</th>
                          <th style={{width: '140px'}} className="text-[11px] font-bold tracking-widest uppercase">Waybill</th>
                          <th style={{width: '180px'}} className="text-[11px] font-bold tracking-widest uppercase">Client</th>
                          <th style={{width: '150px'}} className="text-[11px] font-bold tracking-widest uppercase">Order ID</th>
                          <th style={{width: '120px'}} className="text-[11px] font-bold tracking-widest uppercase">Amount</th>
                          <th style={{width: '160px'}} className="text-[11px] font-bold tracking-widest uppercase">Remark</th>
                          <th style={{width: '350px'}} className="text-[11px] font-bold tracking-widest text-left px-4 uppercase">Return Address</th>
                          <th style={{width: '100px'}} className="text-[11px] font-bold tracking-widest uppercase">FE Name</th>
                        </tr>
                      </thead>
                      <tbody>
                        {Object.entries(filteredRows.reduce((acc: any, row) => {
                          if (!acc[row.client]) acc[row.client] = [];
                          acc[row.client].push(row);
                          return acc;
                        }, {})).map(([client, rows]: any) => (
                          <React.Fragment key={`group-frag-${client}`}>
                            <tr className="bg-slate-800 text-white h-9">
                              <td colSpan={10} className="text-left px-4">
                                <div className="flex items-center justify-between">
                                  <div className="flex items-center gap-3">
                                    <span className="text-[10px] font-black tracking-[0.1em] text-amber-400">{client} — {rows.length} Pkt</span>
                                    {rows.some((r: any) => r.status === 'Pending' && (r.amount ?? 0) >= 4000) && (
                                      <span className="text-[9px] bg-rose-600 text-white px-2 py-0.5 rounded font-black uppercase flex items-center gap-1">
                                        <AlertCircle className="w-3 h-3" /> {rows.filter((r: any) => r.status === 'Pending' && (r.amount ?? 0) >= 4000).length} High Value Pending
                                      </span>
                                    )}
                                  </div>
                                  <button onClick={() => handleCopyAWBOnly(rows)} className="text-[9px] border border-white/20 px-2 py-0.5 rounded hover:bg-white/10 font-bold uppercase">Copy AWBs</button>
                                </div>
                              </td>
                            </tr>
                            {rows.map((row: any) => (
                              <tr key={`row-${row.id}`} className={cn("border-b hover:bg-blue-50/40", selectedRowIds.has(row.id) && "bg-blue-50/50", row.status === 'Pending' && (row.amount ?? 0) >= 4000 && "bg-rose-50/30")}>
                                <td className="px-2 py-2">
                                  <input type="checkbox" checked={selectedRowIds.has(row.id)} onChange={() => {
                                    setSelectedRowIds(prev => {
                                      const next = new Set(prev);
                                      if (next.has(row.id)) next.delete(row.id); else next.add(row.id);
                                      return next;
                                    });
                                  }} />
                                </td>
                                <td className="px-1 py-2">
                                  <button 
                                    onClick={() => deleteRow(row.id)}
                                    className="w-[26px] h-[26px] flex items-center justify-center rounded-md text-slate-300 hover:bg-[#FEF2F2] hover:text-[#DC2626] transition-all"
                                  >
                                    <X className="w-4 h-4" />
                                  </button>
                                </td>
                                <td className="px-2 py-2 text-[13px] font-bold text-slate-600">{currentSession?.dspId}</td>
                                <td className="px-2 py-2 text-[13px] font-bold font-mono text-blue-700 cursor-pointer hover:underline" onClick={() => { navigator.clipboard.writeText(normalizeAWB(row.awb)); showToast("Waybill Copied", "ok"); }}>{normalizeAWB(row.awb)}</td>
                                <td className="px-2 py-2 text-[13px] font-semibold text-slate-800">{row.client}</td>
                                <td className="px-2 py-2 text-[13px] font-medium text-slate-500 whitespace-normal break-words">{row.orderId}</td>
                                <td className={cn("px-2 py-2 text-[13px] font-black", row.status === 'Pending' && (row.amount ?? 0) >= 4000 ? "text-rose-600" : "text-slate-700")}>
                                  <span className="flex items-center justify-center gap-0.5">
                                    <IndianRupee className="w-3 h-3" /> {(row.amount ?? 0).toLocaleString()}
                                  </span>
                                </td>
                                <td className="px-2 py-2">
                                  <span className={cn("inline-block px-2 py-1 rounded text-[10px] font-black border whitespace-normal leading-normal max-w-full uppercase", row.isIntact ? "bg-rose-50 text-rose-700 border-rose-200" : "bg-amber-50 text-amber-700 border-amber-200")}>{row.remark}</span>
                                </td>
                                <td className="px-4 py-2 text-[12px] whitespace-normal break-words text-left min-w-[300px] font-medium leading-relaxed">{row.returnAddress}</td>
                                <td className="px-2 py-2 text-[13px] font-bold text-slate-700">{currentSession?.feName}</td>
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
              <div className="bg-white rounded-2xl p-8 border shadow-sm text-center border-dashed border-2 hover:border-blue-500 cursor-pointer relative">
                <input type="file" onChange={(e) => {
                  const file = e.target.files?.[0];
                  if (!file) return;
                  const reader = new FileReader();
                  reader.onload = (evt) => {
                    try {
                      const data = new Uint8Array(evt.target?.result as ArrayBuffer);
                      const wb = XLSX.read(data, { type: 'array' });
                      const ws = wb.Sheets[wb.SheetNames[0]];
                      const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });
                      const headers = Object.keys(raw[0] || {});
                      const remarkKey = headers.find(h => /remark|nsl|reason/i.test(h)) || "";
                      const replaced = raw.map(row => {
                        const original = String((row as any)[remarkKey] || "").trim();
                        const target = REMARK_MAPPING[original];
                        return { ...row, [remarkKey]: target || original, __isReplaced: !!target, __id: crypto.randomUUID() };
                      });
                      setReplacerData(replaced);
                      setReplacerMeta({ headers, remarkKey });
                      showToast(`Converted ${replaced.filter(r => r.__isReplaced).length} Remarks!`, "ok");
                    } catch (err) { showToast("Replacer Failed", "err"); }
                  };
                  reader.readAsArrayBuffer(file);
                }} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                <FileSpreadsheet className="w-10 h-10 text-slate-300 mx-auto mb-4" />
                <p className="font-black text-slate-600 tracking-tight">Drop EOD Sheet To Convert Remarks</p>
              </div>
              {replacerData.length > 0 && (
                <div className="bg-white rounded-2xl border border-emerald-500/20 shadow-xl overflow-hidden">
                  <div className="overflow-x-auto max-h-[600px]">
                    <table className="w-full text-center border-collapse">
                      <thead className="sticky top-0 bg-[#0f172a] text-white h-12">
                        <tr key="replacer-head">
                          {replacerMeta?.headers.map((h, i) => <th key={`rep-h-${i}`} className="px-4 font-bold text-[10px] tracking-widest uppercase">{h}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {replacerData.map(row => (
                          <tr key={row.__id} className={cn("h-11 border-b", row.__isReplaced ? "bg-emerald-50/40" : "bg-amber-50/40")}>
                            {replacerMeta?.headers.map((h, i) => <td key={`rep-c-${i}`} className="px-4 font-semibold text-[13px] truncate max-w-[200px]">{row[h]}</td>)}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
            <div className="col-span-4">
              <div className="bg-white rounded-2xl border shadow-sm sticky top-20 overflow-hidden">
                <div className="bg-blue-600 p-4 text-white font-black text-[12px] tracking-widest uppercase">Replacement Matrix</div>
                <div className="p-4 space-y-4 max-h-[70vh] overflow-y-auto custom-scrollbar">
                  {Object.entries(REMARK_MAPPING).map(([nsl, off]) => (
                    <div key={`matrix-${nsl}`} className="p-4 bg-slate-50 rounded-xl border space-y-2 border-slate-200">
                      <p className="text-[12px] font-bold text-amber-700">{nsl}</p>
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
            <div className="bg-white border-[1.5px] border-slate-200 rounded-2xl p-6 shadow-sm flex items-center justify-between relative overflow-hidden">
              <div className="absolute left-0 top-0 bottom-0 w-1.5 bg-blue-600" />
              {currentSession ? (
                <div className="flex items-center gap-4 w-full">
                  <div className="w-12 h-12 bg-slate-900 rounded-xl flex items-center justify-center text-white shadow-lg"><User className="w-6 h-6" /></div>
                  <div className="flex-1">
                    <p className="text-lg font-black text-slate-900 leading-tight">{currentSession.feName}</p>
                    <p className="text-[10px] font-black text-slate-400 tracking-widest uppercase mb-1">{currentSession.dspId} — {currentSession.date} — {currentSession.time || ""}</p>
                    <div className="flex flex-wrap gap-[5px]">
                      <span className="text-[10px] font-bold bg-slate-50 text-slate-600 px-2 py-0.5 rounded-[4px] border border-slate-100 uppercase">{currentSession.stats?.total || 0} PKT</span>
                      {currentSession.stats?.pending !== undefined && currentSession.stats.pending > 0 && <span className="text-[10px] font-bold bg-amber-50 text-amber-600 px-2 py-0.5 rounded-[4px] border border-amber-100 uppercase">{currentSession.stats.pending} PENDING</span>}
                      {currentSession.stats?.highValue !== undefined && currentSession.stats.highValue > 0 && <span className="text-[10px] font-bold bg-rose-600 text-white px-2 py-0.5 rounded-[4px] border border-rose-700 uppercase">{currentSession.stats.highValue} HIGH VALUE</span>}
                      {currentSession.stats?.dispatched !== undefined && currentSession.stats.dispatched > 0 && <span className="text-[10px] font-bold bg-rose-50 text-rose-600 px-2 py-0.5 rounded-[4px] border border-rose-100 uppercase">{currentSession.stats.dispatched} DISPATCHED</span>}
                      {currentSession.stats?.rto !== undefined && currentSession.stats.rto > 0 && <span className="text-[10px] font-bold bg-emerald-50 text-emerald-600 px-2 py-0.5 rounded-[4px] border border-emerald-100 uppercase">{currentSession.stats.rto} RTO</span>}
                      {currentSession.stats?.dto !== undefined && currentSession.stats.dto > 0 && <span className="text-[10px] font-bold bg-emerald-50 text-emerald-600 px-2 py-0.5 rounded-[4px] border border-emerald-100 uppercase">{currentSession.stats.dto} DTO</span>}
                    </div>
                  </div>
                </div>
              ) : (
                <div className="py-2 text-slate-400 font-bold text-sm">Select a session to start OTP check</div>
              )}
            </div>

            <div className="bg-white rounded-2xl border-[1.5px] border-dashed border-slate-300 p-8 text-center space-y-4 bg-slate-50/50">
              <div className="max-w-xl mx-auto space-y-4">
                <div className="w-14 h-14 bg-white rounded-2xl flex items-center justify-center text-blue-600 shadow-sm mx-auto border"><Download className="w-7 h-7" /></div>
                <h2 className="text-lg font-black text-slate-900">Upload Delhivery OTP Report</h2>
                <div className="relative cursor-pointer group">
                  <input type="file" onChange={(e) => {
                    const file = e.target.files?.[0];
                    if (!file || !currentSession) return;
                    setIsProcessing(true);
                    const reader = new FileReader();
                    reader.onload = (evt) => {
                      try {
                        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
                        const wb = XLSX.read(data, { type: 'array', raw: true });
                        const ws = wb.Sheets[wb.SheetNames[0]];
                        const rawRows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
                        
                        const sessionMap = new Map<string, PODRow>();
                        currentSession.data.forEach(r => sessionMap.set(normalizeAWB(r.awb), r));
                        
                        const headerRow = rawRows[0].map(h => String(h).toLowerCase().replace(/[\s_-]/g, ""));
                        const awbIdx = headerRow.findIndex(h => /waybill|awb|awbnumber/.test(h));
                        const statusIdx = headerRow.findIndex(h => /status|currentstatus/.test(h));

                        if (awbIdx === -1) throw new Error("Could not find Waybill column in OTP report.");

                        const tempOtpData: OTPRow[] = [];
                        for (let i = 1; i < rawRows.length; i++) {
                          const row = rawRows[i];
                          const awb = normalizeAWB(row[awbIdx]);
                          if (!awb || awb.length < 8) continue;
                          
                          const sessionRow = sessionMap.get(awb);
                          if (!sessionRow) continue;

                          const otpStatusRaw = String(row[statusIdx]).toLowerCase().trim();
                          let otpStatus = 'Unknown';
                          if (otpStatusRaw.includes('dispatched') || otpStatusRaw.includes('dispatch')) otpStatus = 'Dispatched';
                          else if (otpStatusRaw.includes('rto')) otpStatus = 'RTO';
                          else if (otpStatusRaw.includes('dto')) otpStatus = 'DTO';
                          else if (otpStatusRaw.includes('pending')) otpStatus = 'Pending';

                          const csvStatus = sessionRow.status;
                          // Not closed detection logic: OTP is Dispatched, but CSV is RTO/DTO/Pending
                          const isNotClosed = otpStatus === 'Dispatched' && csvStatus !== 'Dispatched';

                          tempOtpData.push({
                            id: crypto.randomUUID(),
                            awb, 
                            client: sessionRow.client, 
                            otpStatus, 
                            sessionStatus: csvStatus, 
                            returnAddress: sessionRow.returnAddress || "",
                            isNotClosed,
                            notClosedType: isNotClosed ? (csvStatus as any) : null
                          });
                        }
                        setOtpData(tempOtpData);
                        showToast(`Imported ${tempOtpData.length} Matched Records!`, "ok");
                      } catch (err: any) { showToast(err.message || "Failed To Process OTP Report", "err"); } finally { setIsProcessing(false); }
                    };
                    reader.readAsArrayBuffer(file);
                    e.target.value = "";
                  }} disabled={!selectedSessionId || isProcessing} className="absolute inset-0 opacity-0 cursor-pointer z-10 disabled:cursor-not-allowed" />
                  <div className={cn("h-12 bg-white border-2 border-slate-200 rounded-2xl flex items-center justify-center font-black transition-all uppercase tracking-widest text-[12px]", (!selectedSessionId || isProcessing) ? "opacity-50 text-slate-300" : "text-slate-600 group-hover:border-blue-500")}>
                    {isProcessing ? "Processing..." : !selectedSessionId ? "Select Session First" : "Select OTP File"}
                  </div>
                </div>
              </div>
            </div>

            {otpData.length > 0 && (
              <>
                <div className="bg-white rounded-xl border shadow-sm flex divide-x overflow-hidden mt-6">
                  {[
                    {id: 'All', label: 'All', color: 'text-slate-900', bgColor: 'bg-[#EFF6FF]', borderColor: 'bg-blue-500', val: otpStats.total},
                    {id: 'Dispatched', label: 'Dispatched', color: 'text-rose-600', bgColor: 'bg-[#FFF5F5]', borderColor: 'bg-rose-500', val: otpStats.dispatched},
                    {id: 'Pending', label: 'Pending', color: 'text-amber-600', bgColor: 'bg-[#FFFBEB]', borderColor: 'bg-amber-500', val: otpStats.pending},
                    {id: 'RTO', label: 'RTO', color: 'text-emerald-600', bgColor: 'bg-[#F0FDF4]', borderColor: 'bg-emerald-500', val: otpStats.rto},
                    {id: 'DTO', label: 'DTO', color: 'text-emerald-600', bgColor: 'bg-[#F0FDF4]', borderColor: 'bg-emerald-500', val: otpStats.dto}
                  ].map(t => (
                    <button key={t.id} onClick={() => { 
                      setOtpStatusFilter(t.id); 
                      setOtpClientFilter("All Clients");
                    }} className={cn("flex-1 py-5 flex flex-col items-center group h-[90px] transition-all relative", otpStatusFilter === t.id ? t.bgColor : "hover:bg-slate-50/30")}>
                      <span className={cn("text-[28px] font-extrabold leading-none mb-1", t.color)}>{t.val}</span>
                      <span className="text-[12px] font-black uppercase tracking-widest">{t.label}</span>
                      {otpStatusFilter === t.id && <div className={cn("absolute bottom-0 w-full h-[3px]", t.borderColor)} />}
                    </button>
                  ))}
                </div>

                <div className="flex flex-wrap items-center gap-4 mb-2">
                  <div className="flex flex-col gap-1.5">
                    <label className="text-[11px] font-black text-slate-400 uppercase tracking-widest">Client Filter</label>
                    <select
                      value={otpClientFilter}
                      onChange={(e) => setOtpClientFilter(e.target.value)}
                      className="border border-slate-200 rounded-lg px-3 py-2 text-[13px] font-bold bg-white outline-none focus:border-blue-500 w-[200px] shadow-sm"
                    >
                      <option value="All Clients">All Clients</option>
                      {uniqueOtpClients.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>
                  <div className="flex-1" />
                  <div className="flex gap-2">
                    <button onClick={() => downloadExcel(filteredOtpRows)} className="h-9 px-5 bg-emerald-600 text-white rounded-lg text-[12px] font-black uppercase">Download Excel</button>
                    <button onClick={() => handleCopyTable(filteredOtpRows)} className="h-9 px-5 bg-blue-600 text-white rounded-lg text-[12px] font-black uppercase">Copy Table</button>
                  </div>
                </div>

                <div className="bg-white rounded-2xl border-[1.5px] border-slate-200 shadow-2xl overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-center border-collapse table-fixed">
                      <thead className="bg-[#0f172a] text-white h-11">
                        <tr key="otp-main-header">
                          <th style={{width: '160px'}} className="text-[11px] font-bold tracking-widest uppercase">Waybill</th>
                          <th style={{width: '200px'}} className="text-[11px] font-bold tracking-widest uppercase">Client Name</th>
                          <th style={{width: '180px'}} className="text-[11px] font-bold tracking-widest uppercase">OTP Status</th>
                          <th style={{width: '180px'}} className="text-[11px] font-bold tracking-widest uppercase">Session Status</th>
                          <th style={{width: '350px'}} className="text-[11px] font-bold tracking-widest text-left px-4 uppercase">Return Address</th>
                        </tr>
                      </thead>
                      <tbody>
                        {Object.entries(filteredOtpRows.reduce((acc: any, row) => {
                          if (!acc[row.client]) acc[row.client] = [];
                          acc[row.client].push(row);
                          return acc;
                        }, {})).map(([client, rows]: any) => {
                          const isDispatchedTab = otpStatusFilter === 'Dispatched';

                          return (
                            <React.Fragment key={`otp-group-${client}`}>
                              {isDispatchedTab && (
                                <tr className="bg-slate-800 text-white h-9">
                                  <td colSpan={5} className="text-left px-4">
                                    <div className="flex items-center justify-between">
                                      <div className="flex items-center gap-3">
                                        <span className="text-[10px] font-black tracking-[0.1em] text-amber-400">{client} — {rows.length} Pkt</span>
                                      </div>
                                      <button onClick={() => handleCopyAWBOnly(rows)} className="text-[9px] border border-white/20 px-2 py-0.5 rounded hover:bg-white/10 font-bold uppercase">Copy AWBs</button>
                                    </div>
                                  </td>
                                </tr>
                              )}
                              {rows.map((row: any) => {
                                const isFTPL = row.client.toUpperCase().includes('FTPL');
                                const isNotClosed = row.isNotClosed;

                                return (
                                  <tr key={`otp-row-${row.id}`} className={cn(
                                    "border-b transition-colors", 
                                    isNotClosed ? "bg-[#FFF7ED] border-l-[3px] border-l-amber-500" :
                                    row.sessionStatus === 'Dispatched' && isFTPL ? "bg-[#FFF5F5] border-l-[2px] border-l-red-500" :
                                    row.sessionStatus === 'Dispatched' ? "bg-white border-l-[2px] border-l-red-500" :
                                    (row.sessionStatus === 'RTO' || row.sessionStatus === 'DTO') && isFTPL ? "bg-emerald-50/50 border-l-[2px] border-l-emerald-500" :
                                    (row.sessionStatus === 'RTO' || row.sessionStatus === 'DTO') ? "bg-white border-l-[2px] border-l-emerald-500" :
                                    "bg-white"
                                  )}>
                                    <td className="px-4 py-3 text-[13px] font-mono font-black text-blue-700 cursor-pointer hover:underline" onClick={() => { navigator.clipboard.writeText(normalizeAWB(row.awb)); showToast("Waybill Copied", "ok"); }}>{normalizeAWB(row.awb)}</td>
                                    <td className={cn("px-4 py-3 text-[13px] font-black tracking-tight", (row.sessionStatus === 'Dispatched' && isFTPL) ? "text-rose-600" : "text-slate-800")}>{row.client}</td>
                                    <td className="px-4 py-3">
                                      <div className="flex flex-col items-center gap-1">
                                        <span className={cn(
                                          "px-2.5 py-0.5 rounded text-[10px] font-black border shadow-sm uppercase", 
                                          row.otpStatus === 'Dispatched' ? "bg-rose-600 text-white border-rose-500" : 
                                          row.otpStatus === 'Pending' ? "bg-amber-500 text-white border-amber-400" : 
                                          "bg-emerald-600 text-white border-emerald-500"
                                        )}>
                                          {row.otpStatus}
                                        </span>
                                        {isNotClosed && (
                                          <span className="flex items-center gap-1 text-[9px] font-black text-amber-600 uppercase">
                                            <AlertCircle className="w-3 h-3" /> {row.sessionStatus} — Not Closed on Device
                                          </span>
                                        )}
                                      </div>
                                    </td>
                                    <td className="px-4 py-3">
                                      <span className={cn(
                                        "px-2 py-0.5 rounded text-[10px] font-black border uppercase", 
                                        row.sessionStatus === 'Pending' ? "bg-amber-50 text-amber-700 border-amber-200" : 
                                        row.sessionStatus === 'RTO' || row.sessionStatus === 'DTO' ? "bg-emerald-50 text-emerald-700 border-emerald-200" : 
                                        "bg-slate-50 text-slate-700 border-slate-200"
                                      )}>
                                        {row.sessionStatus}
                                      </span>
                                    </td>
                                    <td className="px-4 py-3 text-[12px] whitespace-normal break-words text-left min-w-[350px] font-medium leading-relaxed">{row.returnAddress}</td>
                                  </tr>
                                );
                              })}
                            </React.Fragment>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </>
            )}
          </div>
        )}
      </main>
    </div>
  );
}

export default function PODTool() {
  return (
    <Suspense fallback={
      <div className="min-h-screen bg-[#f8fafc] flex items-center justify-center">
        <div className="flex flex-col items-center gap-4">
          <Truck className="w-12 h-12 text-blue-600 animate-bounce" />
          <p className="text-sm font-bold text-slate-500">Loading Tool...</p>
        </div>
      </div>
    }>
      <PODToolContent />
    </Suspense>
  );
}