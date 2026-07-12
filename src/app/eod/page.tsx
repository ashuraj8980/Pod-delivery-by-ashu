
"use client";

import React, { useState, useEffect, useMemo, useCallback } from "react";
import { 
  Truck, 
  Copy, 
  FileSpreadsheet, 
  Settings,
  Download,
  User,
  AlertCircle,
  X,
  Trash2,
  ChevronDown,
  ChevronRight,
  Save,
  Calendar,
  ArrowLeft
} from "lucide-react";
import Link from "next/link";
import { useSearchParams } from "next/navigation";
import * as XLSX from "xlsx";
import { cn } from "@/lib/utils";

/**
 * @fileOverview Delhivery POD Management Tool - EOD Rejection Page
 * Optimized for EOD Rejection management and OTP Dispatch verification.
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

const isValidAWB = (val: any): boolean => {
  const s = normalizeAWB(val);
  return /^\d{8,}$/.test(s) && !isNaN(Number(s));
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
  isNotClosed: boolean;
  notClosedType: 'RTO' | 'DTO' | 'Pending' | null;
}

interface Session {
  id: string;
  feName: string;
  dspId: string;
  date: string;
  data: PODRow[];
  timestamp: number;
}

interface MonthlyRecord extends Session {
  stats: {
    total: number;
    pending: number;
    dispatched: number;
    rto: number;
    dto: number;
  };
}

// Structure: { "2024-03": { "2024-03-25": [Session, Session] } }
type MonthlyStorage = Record<string, Record<string, MonthlyRecord[]>>;

export default function PODTool() {
  const searchParams = useSearchParams();
  const [activeTab, setActiveTab] = useState<"eod" | "remark" | "otp">("eod");
  const [sessions, setSessions] = useState<Session[]>([]);
  const [selectedSessionId, setSelectedSessionId] = useState<string | null>(null);
  const [statusFilter, setStatusFilter] = useState<string>("All");
  
  // Monthly Records State
  const [monthlyRecords, setMonthlyRecords] = useState<MonthlyStorage>({});
  const [selectedMonth, setSelectedMonth] = useState<string>("");
  const [expandedDates, setExpandedDates] = useState<Set<string>>(new Set());

  // Filtering states
  const [selectedRemarkChips, setSelectedRemarkChips] = useState<string[]>([]);
  const [clientFilter, setClientFilter] = useState<string>("All Clients");
  const [otpClientFilter, setOtpClientFilter] = useState<string>("All Clients");
  
  const [showAllPending, setShowAllPending] = useState(false);
  const [setupData, setSetupData] = useState({ feName: "", dspId: "", date: "" });
  const [isMounted, setIsMounted] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [selectedRowIds, setSelectedRowIds] = useState<Set<string>>(new Set());
  
  const [replacerData, setReplacerData] = useState<any[]>([]);
  const [replacerMeta, setReplacerMeta] = useState<{headers: string[], remarkKey: string} | null>(null);

  const [otpData, setOtpData] = useState<OTPRow[]>([]);
  const [otpStatusFilter, setOtpStatusFilter] = useState<string>("All");

  // Load from localStorage
  useEffect(() => {
    setIsMounted(true);
    setSetupData(prev => ({ ...prev, date: new Date().toISOString().split('T')[0] }));
    
    const saved = localStorage.getItem('pod_sessions');
    if (saved) {
      try {
        setSessions(JSON.parse(saved));
      } catch (e) {}
    }

    const savedMonthly = localStorage.getItem('pod_monthly_records');
    if (savedMonthly) {
      try {
        const parsed = JSON.parse(savedMonthly);
        if (parsed && typeof parsed === 'object') {
          setMonthlyRecords(parsed);
          const months = Object.keys(parsed).sort().reverse();
          if (months.length > 0) setSelectedMonth(months[0]);
        }
      } catch (e) {
        console.error("Failed to parse monthly records", e);
        setMonthlyRecords({});
      }
    }
  }, []);

  // Handle URL Session Selection
  useEffect(() => {
    if (isMounted && sessions.length > 0) {
      const sessId = searchParams.get('sessionId');
      if (sessId && sessions.find(s => s.id === sessId)) {
        setSelectedSessionId(sessId);
      }
    }
  }, [isMounted, searchParams, sessions]);

  // Save to localStorage
  useEffect(() => {
    if (isMounted) {
      localStorage.setItem('pod_sessions', JSON.stringify(sessions));
    }
  }, [sessions, isMounted]);

  useEffect(() => {
    if (isMounted) {
      localStorage.setItem('pod_monthly_records', JSON.stringify(monthlyRecords));
    }
  }, [monthlyRecords, isMounted]);

  // Reset OTP data when session changes
  useEffect(() => {
    setOtpData([]);
    setOtpStatusFilter("All");
    setOtpClientFilter("All Clients");
  }, [selectedSessionId]);

  const currentSession = useMemo(() => sessions.find(s => s.id === selectedSessionId) || null, [sessions, selectedSessionId]);

  const stats = useMemo(() => {
    if (!currentSession) return { total: 0, pending: 0, dispatched: 0, rto: 0, dto: 0 };
    return {
      total: currentSession.data.length,
      pending: currentSession.data.filter(r => r.status === 'Pending').length,
      dispatched: currentSession.data.filter(r => r.status === 'Dispatched').length,
      rto: currentSession.data.filter(r => r.status === 'RTO').length,
      dto: currentSession.data.filter(r => r.status === 'DTO').length,
    };
  }, [currentSession]);

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

  // Shared Helper for Badge Rendering to ensure consistency
  const renderSessionBadges = useCallback((statsObj: { total: number, pending: number, rto: number, dto: number }) => {
    return (
      <div className="flex flex-wrap gap-[5px]">
        <span className="text-[10px] font-bold bg-slate-50 text-slate-600 px-2 py-0.5 rounded-[4px] border border-slate-100 uppercase">
          {statsObj.total} PKT
        </span>
        {statsObj.pending > 0 && (
          <span className="text-[10px] font-bold bg-amber-50 text-amber-600 px-2 py-0.5 rounded-[4px] border border-amber-100 uppercase">
            {statsObj.pending} PENDING
          </span>
        )}
        {statsObj.rto > 0 && (
          <span className="text-[10px] font-bold bg-rose-50 text-rose-600 px-2 py-0.5 rounded-[4px] border border-rose-100 uppercase">
            {statsObj.rto} RTO
          </span>
        )}
        {statsObj.dto > 0 && (
          <span className="text-[10px] font-bold bg-emerald-50 text-emerald-600 px-2 py-0.5 rounded-[4px] border border-emerald-100 uppercase">
            {statsObj.dto} DTO
          </span>
        )}
      </div>
    );
  }, []);

  const handleSaveAllSessions = () => {
    if (sessions.length === 0) {
      showToast("No sessions to save", "err");
      return;
    }

    try {
      setMonthlyRecords(prev => {
        const next = prev ? { ...prev } : {};
        sessions.forEach(session => {
          if (!session || !session.date) return;
          
          // Validate Date is DD-MM-YYYY using regex
          const dateRegex = /^\d{2}-\d{2}-\d{4}$/;
          if (!dateRegex.test(session.date)) return;

          const parts = session.date.split('-');
          // Safe check for parts
          if (parts.length !== 3) return;
          
          const yearMonth = `${parts[2]}-${parts[1]}`;
          const fullDateKey = `${parts[2]}-${parts[1]}-${parts[0]}`;

          if (!next[yearMonth]) next[yearMonth] = {};
          if (!next[yearMonth][fullDateKey]) next[yearMonth][fullDateKey] = [];

          const sessionStats = {
            total: session.data?.length || 0,
            pending: (session.data || []).filter(r => r.status === 'Pending').length,
            dispatched: (session.data || []).filter(r => r.status === 'Dispatched').length,
            rto: (session.data || []).filter(r => r.status === 'RTO').length,
            dto: (session.data || []).filter(r => r.status === 'DTO').length,
          };

          const newRecord: MonthlyRecord = {
            ...session,
            stats: sessionStats
          };

          // Replace if exists for the same FE/DSP on same day
          const filtered = (next[yearMonth][fullDateKey] || []).filter(s => s.dspId !== session.dspId);
          next[yearMonth][fullDateKey] = [newRecord, ...filtered];
        });
        return next;
      });

      showToast("All sessions saved to monthly records", "ok");
    } catch (error) {
      console.error("Save Daily Report failed", error);
      showToast("Application error during save", "err");
    }
  };

  const handleCopyAWBOnly = async (rowsToCopy: any[]) => {
    if (!rowsToCopy.length) return;
    const text = rowsToCopy.map(r => normalizeAWB(r.awb)).join('\n');
    try {
      await navigator.clipboard.writeText(text);
      showToast(`Copied ${rowsToCopy.length} AWB Numbers`, "ok");
    } catch (err) { showToast("Failed To Copy AWBs", "err"); }
  };

  const copyDataToClipboard = useCallback(async (rows: any[], headers: string[]) => {
    if (!rows.length) return;
    const exportHeaders = headers.filter(h => !/return address|return_address/i.test(h));
    const plainText = rows.map(r => exportHeaders.map(h => String(r[h] || "").trim()).join("\t")).join("\n");
    const rowsHtml = rows.map(r => {
      const cells = exportHeaders.map(h => {
        const val = String(r[h] || "").trim();
        const style = h.toLowerCase().includes('awb') || h.toLowerCase().includes('waybill') ? 'style=\'mso-number-format:"\\@"\'' : '';
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
      showToast("Please Enter FE Name and DSP ID First", "err");
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
          const status = STATUS_MAP[statusRaw] || "Unknown";
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
        }).filter((row): row is PODRow => row !== null && row.awb.length >= 8 && row.status !== "Unknown");
        
        if (parsedRows.length === 0) throw new Error("No Valid Data Found In EOD Report.");
        
        const newSessionId = crypto.randomUUID();
        const newSession: Session = { 
          id: newSessionId, 
          feName: setupData.feName, 
          dspId: setupData.dspId, 
          date: formatDate(setupData.date), 
          data: parsedRows, 
          timestamp: Date.now() 
        };

        setSessions(prev => {
          const filtered = prev.filter(s => s.dspId !== setupData.dspId);
          return [newSession, ...filtered];
        });
        
        setSelectedSessionId(newSessionId);
        showToast(`Imported ${parsedRows.length} Rows Locally!`, "ok");
        // TODO: Connect to Firestore later
      } catch (err: any) { showToast(err.message || "Failed To Import File", "err"); } finally { setIsProcessing(false); }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const deleteRow = (rowId: string) => {
    if (!selectedSessionId) return;
    setSessions(prev => prev.map(s => {
      if (s.id === selectedSessionId) {
        return { ...s, data: s.data.filter(r => r.id !== rowId) };
      }
      return s;
    }));
    setSelectedRowIds(prev => {
      const next = new Set(prev);
      next.delete(rowId);
      return next;
    });
  };

  const deleteSelectedRows = () => {
    if (!selectedSessionId || selectedRowIds.size === 0) return;
    setSessions(prev => prev.map(s => {
      if (s.id === selectedSessionId) {
        return { ...s, data: s.data.filter(r => !selectedRowIds.has(r.id)) };
      }
      return s;
    }));
    const count = selectedRowIds.size;
    setSelectedRowIds(new Set());
    showToast(`Deleted ${count} Rows`, "ok");
  };

  const uniqueClients = useMemo(() => {
    if (!currentSession) return [];
    let rows = currentSession.data;
    if (statusFilter !== 'All') {
      rows = rows.filter(r => r.status === statusFilter);
    }
    return Array.from(new Set(rows.map(r => r.client))).sort();
  }, [currentSession, statusFilter]);

  const filteredRows = useMemo(() => {
    if (!currentSession) return [];
    let rows = currentSession.data;
    if (statusFilter !== 'All') {
      rows = rows.filter(r => r.status === statusFilter);
    }
    if (statusFilter === 'Pending' && !showAllPending && selectedRemarkChips.length > 0) {
      rows = rows.filter(r => selectedRemarkChips.includes(r.remark));
    }
    if (clientFilter !== 'All Clients') {
      rows = rows.filter(r => r.client === clientFilter);
    }
    return rows;
  }, [currentSession, statusFilter, selectedRemarkChips, clientFilter, showAllPending]);

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
    const headers = ['Date', 'DSP ID', 'Waybill Number', 'Client', 'Order ID', 'Remark', 'FE Name'];
    const exportRows = rowsToCopy.map((r, i) => ({
      'Date': formatDate(r.date),
      'DSP ID': i === 0 ? r.dspId : "",
      'Waybill Number': normalizeAWB(r.awb),
      'Client': r.client,
      'Order ID': r.orderId,
      'Remark': r.remark,
      'FE Name': r.feName
    }));
    const success = await copyDataToClipboard(exportRows, headers);
    if (success) showToast(`Copied ${rowsToCopy.length} Rows To Clipboard`, "ok");
  }, [copyDataToClipboard, showToast]);

  const downloadExcel = (rowsToDownload: any[]) => {
    if (!rowsToDownload.length) return;
    const header = ['Date', 'DSP ID', 'Waybill Number', 'Client', 'Order ID', 'Remark', 'FE Name'];
    const excelData = rowsToDownload.map((r, i) => [
      formatDate(r.date), 
      { v: i === 0 ? String(r.dspId) : "", t: 's' }, 
      { v: String(normalizeAWB(r.awb)), t: 's' }, 
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

  const deleteSession = (sessionId: string) => {
    setSessions(prev => prev.filter(s => s.id !== sessionId));
    if (selectedSessionId === sessionId) setSelectedSessionId(null);
    showToast("Session Deleted", "ok");
  };

  const handleOTPFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (!selectedSessionId || !currentSession) return;
    setIsProcessing(true);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array', raw: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(ws, { raw: true, defval: "" });
        const sessionMap = new Map<string, PODRow>();
        currentSession.data.forEach(r => sessionMap.set(normalizeAWB(r.awb), r));
        const tempOtpData: OTPRow[] = [];
        let matchedCount = 0;
        rawData.forEach((row: any) => {
          const rawAwb = row['Waybill'] || row['AWB'] || row['waybill'] || row['awb'];
          if (!isValidAWB(rawAwb)) return;
          const awb = normalizeAWB(rawAwb);
          const sessionRow = sessionMap.get(awb);
          if (!sessionRow) return;
          matchedCount++;
          const otpStatusRaw = String(row['Status'] || row['Current Status'] || "").toLowerCase().trim();
          let otpStatus = 'Unknown';
          if (otpStatusRaw.includes('dispatched') || otpStatusRaw.includes('dispatch')) otpStatus = 'Dispatched';
          else if (otpStatusRaw.includes('rto')) otpStatus = 'RTO';
          else if (otpStatusRaw.includes('dto')) otpStatus = 'DTO';
          else if (otpStatusRaw.includes('pending')) otpStatus = 'Pending';
          const csvStatus = sessionRow.status;
          const isRTONotClosed = otpStatus === 'Dispatched' && csvStatus === 'RTO';
          const isDTONotClosed = otpStatus === 'Dispatched' && csvStatus === 'DTO';
          const isPendingNotClosed = otpStatus === 'Dispatched' && csvStatus === 'Pending';
          tempOtpData.push({
            id: crypto.randomUUID(),
            awb, 
            client: sessionRow.client, 
            otpStatus, 
            sessionStatus: csvStatus, 
            returnAddress: sessionRow.returnAddress || "",
            isNotClosed: isRTONotClosed || isDTONotClosed || isPendingNotClosed,
            notClosedType: isRTONotClosed ? 'RTO' : isDTONotClosed ? 'DTO' : isPendingNotClosed ? 'Pending' : null
          });
        });
        if (matchedCount === 0) {
          showToast("Wrong file — no AWBs match current session.", "err");
          return;
        }
        setOtpData(tempOtpData);
        showToast(`Imported ${tempOtpData.length} Matched Records Locally.`, "ok");
      } catch (err) { showToast("Failed To Process OTP Report", "err"); } finally { setIsProcessing(false); }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const getOtpBaseFilteredRows = useCallback((status: string) => {
    let rows = otpData;
    if (status !== 'All') {
      if (status === 'Dispatched') {
        rows = rows.filter(r => r.otpStatus === 'Dispatched' && (r.sessionStatus === 'Dispatched' || r.sessionStatus === 'RTO' || r.sessionStatus === 'DTO'));
      } else if (status === 'RTO') {
        rows = rows.filter(r => r.otpStatus === 'RTO' || r.notClosedType === 'RTO');
      } else if (status === 'DTO') {
        rows = rows.filter(r => r.otpStatus === 'DTO' || r.notClosedType === 'DTO');
      } else if (status === 'Pending') {
        rows = rows.filter(r => r.otpStatus === 'Pending' || r.notClosedType === 'Pending');
      }
    }
    return rows;
  }, [otpData]);

  const uniqueOtpClients = useMemo(() => {
    const rows = getOtpBaseFilteredRows(otpStatusFilter);
    return Array.from(new Set(rows.map(r => r.client))).sort();
  }, [getOtpBaseFilteredRows, otpStatusFilter]);

  const otpFilteredRows = useMemo(() => {
    let rows = getOtpBaseFilteredRows(otpStatusFilter);
    if (otpClientFilter !== 'All Clients') {
      rows = rows.filter(r => r.client === otpClientFilter);
    }
    return rows;
  }, [getOtpBaseFilteredRows, otpStatusFilter, otpClientFilter]);

  const otpStats = useMemo(() => ({
    total: otpData.length,
    dispatched: otpData.filter(r => r.otpStatus === 'Dispatched' && (r.sessionStatus === 'Dispatched' || r.sessionStatus === 'RTO' || r.sessionStatus === 'DTO')).length,
    rto: otpData.filter(r => r.otpStatus === 'RTO' || r.notClosedType === 'RTO').length,
    dto: otpData.filter(r => r.otpStatus === 'DTO' || r.notClosedType === 'DTO').length,
    pending: otpData.filter(r => r.otpStatus === 'Pending' || r.notClosedType === 'Pending').length,
  }), [otpData]);

  const downloadOTPExcel = () => {
    if (!otpFilteredRows.length) return;
    const header = ['AWB Number', 'Client Name', 'OTP Status', 'Session Status', 'Return Address'];
    const excelData = otpFilteredRows.map(r => [
      { v: String(r.awb), t: 's' },
      r.client,
      r.otpStatus,
      r.sessionStatus,
      r.returnAddress
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...excelData]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "OTP_Report");
    XLSX.writeFile(wb, `OTP_Verification_${currentSession?.dspId || 'Export'}.xlsx`);
  };

  const handleSessionClick = (s: Session) => {
    setSelectedSessionId(s.id);
    setStatusFilter("All");
    setSelectedRemarkChips([]);
    setClientFilter("All Clients");
    setShowAllPending(false);
    let isoDate = "";
    if (s.date && s.date.includes('-')) {
      const parts = s.date.split('-');
      if (parts.length === 3) isoDate = `${parts[2]}-${parts[1]}-${parts[0]}`;
    }
    setSetupData({ feName: s.feName, dspId: s.dspId, date: isoDate || s.date });
  };

  const toggleDateExpand = (date: string) => {
    setExpandedDates(prev => {
      const next = new Set(prev);
      if (next.has(date)) next.delete(date); else next.add(date);
      return next;
    });
  };

  const getMonthName = (yearMonth: string) => {
    const [y, m] = yearMonth.split('-');
    const date = new Date(parseInt(y), parseInt(m) - 1, 1);
    return date.toLocaleString('default', { month: 'long', year: 'numeric' });
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
              <div className={cn("border-2 border-dashed rounded-xl p-8 text-center transition-all relative", (!setupData.feName || !setupData.dspId) ? "bg-slate-100 border-slate-200 cursor-not-allowed opacity-60" : "bg-slate-50 hover:bg-white hover:border-blue-500 cursor-pointer")}>
                {setupData.feName && setupData.dspId && <input type="file" onChange={handleFileUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />}
                <div className="space-y-3">
                  <FileSpreadsheet className="w-6 h-6 text-slate-400 mx-auto" />
                  <p className="text-sm font-black text-slate-600 tracking-tight">{(!setupData.feName || !setupData.dspId) ? "Enter DSP ID And FE Name To Upload" : "Import Daily EOD Report"}</p>
                </div>
              </div>
            </div>

            {sessions.length > 0 && (
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-6 space-y-4">
                <div className="flex items-center justify-between mb-2">
                  <h3 className="text-sm font-black text-slate-700 tracking-tight">Recent Sessions</h3>
                  <div className="flex items-center gap-3">
                    <button 
                      onClick={handleSaveAllSessions}
                      className="bg-[#0f172a] text-white px-3 py-1.5 rounded-[8px] text-[11px] font-bold hover:bg-slate-800 transition-all shadow-sm flex items-center gap-2"
                    >
                      <Save className="w-3.5 h-3.5" />
                      SAVE DAILY REPORT
                    </button>
                    <button 
                      onClick={() => { setSessions([]); localStorage.removeItem('pod_sessions'); }} 
                      className="text-[11px] font-black text-rose-600 hover:text-rose-700 transition-colors uppercase tracking-widest"
                    >
                      CLEAR ALL HISTORY
                    </button>
                  </div>
                </div>
                <div className="grid grid-cols-[repeat(auto-fill,minmax(280px,1fr))] gap-4">
                  {sessions.map(s => {
                    const sStats = {
                      total: s.data.length,
                      pending: s.data.filter(r => r.status === 'Pending').length,
                      rto: s.data.filter(r => r.status === 'RTO').length,
                      dto: s.data.filter(r => r.status === 'DTO').length,
                    };
                    return (
                      <div key={s.id} onClick={() => handleSessionClick(s)} className={cn("relative p-4 border-[1.5px] rounded-2xl cursor-pointer transition-all shadow-sm overflow-hidden bg-white", selectedSessionId === s.id ? "border-blue-500 ring-2 ring-blue-500/10" : "hover:border-blue-300 border-slate-100")}>
                        <div className="absolute left-0 top-0 bottom-0 w-1.5 bg-blue-600" />
                        <button onClick={(e) => { e.stopPropagation(); deleteSession(s.id); }} className="absolute right-3 top-3 text-slate-300 hover:text-rose-500 transition-colors p-1">
                          <X className="w-4 h-4" />
                        </button>
                        <p className="text-[15px] font-bold text-slate-900 mb-0.5 tracking-tight">{s.feName}</p>
                        <p className="text-[11px] text-slate-400 font-bold mb-3 uppercase tracking-wider">{s.dspId} — {s.date}</p>
                        {renderSessionBadges(sStats)}
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-6 space-y-6">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <Calendar className="w-4 h-4 text-blue-600" />
                  <h3 className="text-sm font-black text-slate-700 tracking-tight uppercase tracking-widest">Monthly Records History</h3>
                </div>
                <select 
                  value={selectedMonth}
                  onChange={(e) => setSelectedMonth(e.target.value)}
                  className="border border-slate-200 rounded-lg px-3 py-1.5 text-[12px] font-bold bg-slate-50 outline-none focus:border-blue-500 shadow-sm"
                >
                  <option value="">Select Month</option>
                  {Object.keys(monthlyRecords).sort().reverse().map(month => (
                    <option key={month} value={month}>{getMonthName(month)}</option>
                  ))}
                </select>
              </div>

              {selectedMonth && monthlyRecords[selectedMonth] && (
                <div className="space-y-3">
                  {Object.keys(monthlyRecords[selectedMonth]).sort().reverse().map(dateKey => {
                    const isExpanded = expandedDates.has(dateKey);
                    const sessionsOnDate = monthlyRecords[selectedMonth][dateKey];
                    const dateDisplay = formatDate(dateKey);
                    return (
                      <div key={dateKey} className="border border-slate-100 rounded-2xl overflow-hidden bg-white shadow-sm transition-all hover:shadow-md">
                        <button 
                          onClick={() => toggleDateExpand(dateKey)}
                          className="w-full px-5 py-4 flex items-center justify-between bg-slate-50/50 hover:bg-slate-50 transition-colors"
                        >
                          <span className="text-[13px] font-bold text-slate-800">{dateDisplay} — {sessionsOnDate.length} Sessions</span>
                          {isExpanded ? <ChevronDown className="w-5 h-5 text-slate-400" /> : <ChevronRight className="w-5 h-5 text-slate-400" />}
                        </button>
                        {isExpanded && (
                          <div className="p-5 grid grid-cols-[repeat(auto-fill,minmax(300px,1fr))] gap-4 bg-white border-t border-slate-50">
                            {sessionsOnDate.map(s => {
                              const sStats = s.stats || {
                                total: s.data?.length || 0,
                                pending: (s.data || []).filter(r => r.status === 'Pending').length,
                                rto: (s.data || []).filter(r => r.status === 'RTO').length,
                                dto: (s.data || []).filter(r => r.status === 'DTO').length,
                              };
                              return (
                                <div 
                                  key={`hist-${s.id}`} 
                                  onClick={() => handleSessionClick(s as any)}
                                  className="p-4 border-[1.5px] border-slate-100 rounded-2xl cursor-pointer hover:border-blue-400 hover:shadow-md transition-all group relative overflow-hidden bg-white"
                                >
                                  <div className="absolute left-0 top-0 bottom-0 w-1.5 bg-slate-200 group-hover:bg-blue-500 transition-colors" />
                                  <p className="text-[15px] font-bold text-slate-900 leading-tight mb-0.5">{s.feName}</p>
                                  <p className="text-[11px] text-slate-400 font-bold mb-3 uppercase tracking-wider">{s.dspId} — {s.date}</p>
                                  {renderSessionBadges(sStats)}
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            {currentSession && (
              <div className="space-y-6">
                <div className="bg-white rounded-xl border shadow-sm flex divide-x overflow-hidden mt-6">
                  {[
                    {id: 'All', label: 'All', color: 'text-slate-900', bgColor: 'bg-[#EFF6FF]', borderColor: 'bg-blue-500', val: stats.total},
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

                <div className="flex items-center gap-4 mb-2">
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
                </div>

                <div className="flex gap-2 bg-white border rounded-xl p-1 shadow-sm items-center">
                  <button onClick={() => handleCopyAWBOnly(filteredRows)} className="h-9 px-4 bg-slate-800 text-white rounded-lg text-[11px] font-black tracking-wider flex items-center gap-2 uppercase">
                    <Copy className="w-3.5 h-3.5" /> Copy All AWB
                  </button>
                  <button onClick={() => {
                    const rows = filteredRows.filter(r => selectedRowIds.has(r.id));
                    handleCopyAWBOnly(rows);
                  }} className="h-9 px-4 bg-slate-900 text-white rounded-lg text-[11px] font-black tracking-wider flex items-center gap-2 uppercase">
                    <Copy className="w-3.5 h-3.5" /> Copy Selected AWBs
                  </button>
                  
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
                          <th style={{width: '200px'}} className="text-[11px] font-bold tracking-widest uppercase">Remark</th>
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
                              <td colSpan={9} className="text-left px-4">
                                <div className="flex items-center justify-between">
                                  <span className="text-[10px] font-black tracking-[0.1em] text-amber-400">{client} — {rows.length} Pkt</span>
                                  <button onClick={() => handleCopyAWBOnly(rows)} className="text-[9px] border border-white/20 px-2 py-0.5 rounded hover:bg-white/10 font-bold uppercase">Copy AWBs</button>
                                </div>
                              </td>
                            </tr>
                            {rows.map((row: any) => (
                              <tr key={`row-${row.id}`} className={cn("border-b hover:bg-blue-50/40", selectedRowIds.has(row.id) && "bg-blue-50/50")}>
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
                                <td className="px-2 py-2 text-[13px] font-bold text-slate-600">{row.dspId}</td>
                                <td className="px-2 py-2 text-[13px] font-bold font-mono text-blue-700 cursor-pointer hover:underline" onClick={() => { navigator.clipboard.writeText(normalizeAWB(row.awb)); showToast("Waybill Copied", "ok"); }}>{normalizeAWB(row.awb)}</td>
                                <td className="px-2 py-2 text-[13px] font-semibold text-slate-800">{row.client}</td>
                                <td className="px-2 py-2 text-[13px] font-medium text-slate-500 whitespace-normal break-words">{row.orderId}</td>
                                <td className="px-2 py-2">
                                  <span className={cn("inline-block px-2 py-1 rounded text-[10px] font-black border whitespace-normal leading-normal max-w-full uppercase", row.isIntact ? "bg-rose-50 text-rose-700 border-rose-200" : "bg-amber-50 text-amber-700 border-amber-200")}>{row.remark}</span>
                                </td>
                                <td className="px-4 py-2 text-[12px] whitespace-normal break-words text-left min-w-[300px] font-medium leading-relaxed">{row.returnAddress}</td>
                                <td className="px-2 py-2 text-[13px] font-bold text-slate-700">{row.feName}</td>
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
            {!selectedSessionId ? (
              <div className="bg-white rounded-2xl border p-20 text-center">
                <AlertCircle className="w-12 h-12 text-amber-500 mx-auto mb-4" />
                <h3 className="text-lg font-black text-slate-900">Session Not Selected</h3>
                <p className="text-sm font-bold text-slate-400">Please Select A Session In Daily EOD Rejection Tab First.</p>
              </div>
            ) : (
              <>
                <div className="bg-white border-[1.5px] border-slate-200 rounded-2xl p-6 shadow-sm flex items-center justify-between relative overflow-hidden">
                  <div className="absolute left-0 top-0 bottom-0 w-1.5 bg-blue-600" />
                  <div className="flex items-center gap-4">
                    <div className="w-12 h-12 bg-slate-900 rounded-xl flex items-center justify-center text-white shadow-lg"><User className="w-6 h-6" /></div>
                    <div>
                      <p className="text-lg font-black text-slate-900 leading-tight">{currentSession.feName}</p>
                      <p className="text-[10px] font-black text-slate-400 tracking-widest uppercase">{currentSession.dspId} — {currentSession.date}</p>
                    </div>
                  </div>
                  {renderSessionBadges(stats)}
                </div>

                <div className="bg-white rounded-2xl border-[1.5px] border-dashed border-slate-300 p-12 text-center space-y-4 bg-slate-50/50">
                  <div className="max-w-xl mx-auto space-y-4">
                    <div className="w-16 h-16 bg-white rounded-2xl flex items-center justify-center text-blue-600 shadow-sm mx-auto border"><Download className="w-8 h-8" /></div>
                    <div>
                      <h2 className="text-xl font-black text-slate-900">Upload Delhivery OTP Report</h2>
                      <p className="text-[12px] font-bold text-slate-400 uppercase tracking-widest">Upload The Default Delhivery Export File For This FE Session.</p>
                    </div>
                    <div className="relative cursor-pointer group">
                      <input type="file" onChange={handleOTPFileUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                      <div className="h-14 bg-white border-2 border-slate-200 rounded-2xl flex items-center justify-center font-black text-slate-600 group-hover:border-blue-500 transition-all uppercase tracking-widest">Select File</div>
                    </div>
                  </div>
                </div>

                {otpData.length > 0 && (
                  <div className="space-y-6">
                    <div className="bg-white rounded-xl border shadow-sm flex divide-x overflow-hidden">
                      {[
                        {id: 'All', label: 'All', val: otpStats.total, color: 'text-slate-900', bgColor: 'bg-[#EFF6FF]', borderColor: 'bg-blue-500'},
                        {id: 'Dispatched', label: 'Dispatched', val: otpStats.dispatched, color: 'text-rose-600', bgColor: 'bg-[#FFF5F5]', borderColor: 'bg-rose-500'},
                        {id: 'Pending', label: 'Pending', val: otpStats.pending, color: 'text-amber-600', bgColor: 'bg-[#FFFBEB]', borderColor: 'bg-amber-500'},
                        {id: 'RTO', label: 'RTO', val: otpStats.rto, color: 'text-emerald-600', bgColor: 'bg-[#F0FDF4]', borderColor: 'bg-emerald-500'},
                        {id: 'DTO', label: 'DTO', val: otpStats.dto, color: 'text-emerald-600', bgColor: 'bg-[#F0FDF4]', borderColor: 'bg-emerald-500'}
                      ].map(t => (
                        <button key={`otp-tab-${t.id}`} onClick={() => {
                          setOtpStatusFilter(t.id);
                          setOtpClientFilter("All Clients");
                        }} className={cn("flex-1 py-6 flex flex-col items-center group h-[110px] transition-all relative", otpStatusFilter === t.id ? t.bgColor : "hover:bg-slate-50/30")}>
                          <span className={cn("text-[36px] font-black leading-none mb-1", t.color)}>{t.val}</span>
                          <span className={cn("text-[13px] font-black tracking-widest uppercase", otpStatusFilter === t.id ? t.color : "text-slate-400")}>{t.label}</span>
                          {otpStatusFilter === t.id && <div className={cn("absolute bottom-0 w-full h-[4px]", t.borderColor)} />}
                        </button>
                      ))}
                    </div>

                    <div className="flex items-center gap-4 mb-2">
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
                      <button onClick={downloadOTPExcel} className="h-9 px-5 bg-emerald-600 text-white rounded-lg text-[12px] font-black flex items-center gap-2 uppercase">
                        <Download className="w-4 h-4" /> Download OTP Report
                      </button>
                    </div>

                    <div className="bg-white rounded-2xl border-[1.5px] border-slate-200 shadow-2xl overflow-hidden">
                      <div className="overflow-x-auto">
                        <table className="w-full text-center border-collapse table-fixed">
                          <thead className="bg-[#0f172a] text-white h-11">
                            <tr key="otp-main-header">
                              <th style={{width: '32px'}} className="px-2"><input type="checkbox" /></th>
                              <th style={{width: '150px'}} className="text-[11px] font-bold tracking-widest uppercase">Waybill</th>
                              <th style={{width: '200px'}} className="text-[11px] font-bold tracking-widest uppercase">Client Name</th>
                              <th style={{width: '180px'}} className="text-[11px] font-bold tracking-widest uppercase">OTP Status</th>
                              <th style={{width: '180px'}} className="text-[11px] font-bold tracking-widest uppercase">Session Status</th>
                              <th style={{width: '350px'}} className="text-[11px] font-bold tracking-widest text-left px-4 uppercase">Return Address</th>
                            </tr>
                          </thead>
                          <tbody>
                            {Object.entries(otpFilteredRows.reduce((acc: any, row) => {
                              if (!acc[row.client]) acc[row.client] = [];
                              acc[row.client].push(row);
                              return acc;
                        }, {})).sort((a: any, b: any) => b[1].length - a[1].length).map(([client, rows]: any) => (
                              <React.Fragment key={`otp-frag-${client}`}>
                                <tr className="bg-slate-800 text-white h-9">
                                  <td colSpan={6} className="text-left px-4">
                                    <div className="flex items-center justify-between">
                                      <span className="text-[10px] font-black tracking-[0.1em] text-amber-400">{client} — {rows.length} Pkt</span>
                                      <button onClick={() => handleCopyAWBOnly(rows)} className="text-[9px] border border-white/20 px-2 py-0.5 rounded hover:bg-white/10 font-bold uppercase">Copy AWBs</button>
                                    </div>
                                  </td>
                                </tr>
                                {rows.map((row: any) => (
                                  <tr key={`otp-row-${row.id}`} className={cn("border-b", row.isNotClosed ? "bg-amber-50/30 border-l-[3px] border-l-amber-500" : "bg-white")}>
                                    <td className="px-2 py-2"><input type="checkbox" /></td>
                                    <td className="px-4 py-2 text-[13px] font-mono font-black text-blue-700 cursor-pointer hover:underline" onClick={() => { navigator.clipboard.writeText(normalizeAWB(row.awb)); showToast("Waybill Copied", "ok"); }}>{normalizeAWB(row.awb)}</td>
                                    <td className="px-4 py-2 text-[13px] font-black tracking-tight">{row.client}</td>
                                    <td className="px-4 py-2">
                                      <div className="flex flex-col items-center gap-1.5 py-1">
                                        <span className={cn("px-2.5 py-0.5 rounded text-[10px] font-black border shadow-sm uppercase", row.otpStatus === 'Dispatched' ? "bg-rose-600 text-white border-rose-500" : row.otpStatus === 'Pending' ? "bg-amber-500 text-white border-amber-400" : "bg-emerald-600 text-white border-emerald-500")}>{row.otpStatus}</span>
                                        {row.isNotClosed && <span className="px-1.5 py-0.5 rounded bg-amber-100 text-amber-800 text-[9px] font-black border border-amber-300 whitespace-normal leading-tight text-center uppercase">Not closed on device</span>}
                                      </div>
                                    </td>
                                    <td className="px-4 py-2">
                                      <span className={cn("px-2 py-0.5 rounded text-[10px] font-black border uppercase", row.sessionStatus === 'Pending' ? "bg-amber-50 text-amber-700 border-amber-200" : row.sessionStatus === 'RTO' ? "bg-emerald-50 text-emerald-700 border-emerald-200" : "bg-slate-50 text-slate-700 border-slate-200")}>{row.sessionStatus}</span>
                                    </td>
                                    <td className="px-4 py-2 text-[12px] whitespace-normal break-words text-left min-w-[350px] font-medium leading-relaxed">{row.returnAddress}</td>
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
          </div>
        )}
      </main>
    </div>
  );
}
