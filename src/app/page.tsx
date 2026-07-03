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
  Check
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
 * Optimized for Excel-style Client Filtering and Remark Grouping
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
  status: string;
  returnAddress: string;
  isFTPL: boolean;
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
  const [clientSearchQuery, setClientSearchQuery] = useState("");
  const [setupData, setSetupData] = useState({ feName: "", dspId: "", date: "" });
  const [isMounted, setIsMounted] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");
  const [isPopoverOpen, setIsPopoverOpen] = useState(false);
  const [selectedRowIds, setSelectedRowIds] = useState<Set<string>>(new Set());
  
  const [replacerData, setReplacerData] = useState<any[]>([]);
  const [replacerMeta, setReplacerMeta] = useState<{headers: string[], remarkKey: string} | null>(null);

  const [otpData, setOtpData] = useState<OTPRow[]>([]);
  const [otpStatusFilter, setOtpStatusFilter] = useState<string>("all");

  useEffect(() => {
    setIsMounted(true);
    setSetupData(prev => ({ ...prev, date: new Date().toISOString().split('T')[0] }));
  }, []);

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
    const plainText = rows.map(r => 
      headers.map(h => {
        const val = String(r[h] || "").trim();
        if (h.toLowerCase().includes('awb')) return `'${val}`;
        return val;
      }).join("\t")
    ).join("\n");
    const rowsHtml = rows.map(r => {
      const cells = headers.map(h => {
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
          const awb = normalizeAWB(findVal(/waybill|awb|awbnumber/));
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
        }).filter(row => row.awb.length >= 3 && row.status !== "unknown");
        
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
        
        setSessions(prev => {
          const existingIndex = prev.findIndex(s => s.dspId === setupData.dspId);
          if (existingIndex !== -1) {
            const updatedSessions = [...prev];
            updatedSessions[existingIndex] = { ...newSession, id: prev[existingIndex].id };
            setSelectedSessionId(prev[existingIndex].id);
            return updatedSessions;
          }
          setSelectedSessionId(newSessionId);
          return [newSession, ...prev];
        });
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

  const currentSession = useMemo(() => sessions.find(s => s.id === selectedSessionId) || null, [sessions, selectedSessionId]);

  const filteredRows = useMemo(() => {
    if (!currentSession) return [];
    let rows = currentSession.data;
    if (statusFilter !== 'all') {
      rows = rows.filter(r => r.status === statusFilter);
    }
    if (activeRemarkChip) rows = rows.filter(r => r.remark === activeRemarkChip);
    
    // Client Multi-filter
    if (selectedClients.length > 0) {
      rows = rows.filter(r => selectedClients.includes(r.client));
    }

    if (searchTerm) {
      const s = searchTerm.toLowerCase();
      rows = rows.filter(r => r.awb.includes(s) || r.client.toLowerCase().includes(s) || r.orderId.toLowerCase().includes(s));
    }
    return rows;
  }, [currentSession, statusFilter, activeRemarkChip, searchTerm, selectedClients]);

  const allPossibleClients = useMemo(() => {
    if (!currentSession) return [];
    let rows = currentSession.data;
    if (statusFilter !== 'all') rows = rows.filter(r => r.status === statusFilter);
    if (activeRemarkChip) rows = rows.filter(r => r.remark === activeRemarkChip);
    return Array.from(new Set(rows.map(r => r.client))).sort();
  }, [currentSession, statusFilter, activeRemarkChip]);

  const displayedClients = useMemo(() => {
    if (!clientSearchQuery) return allPossibleClients;
    return allPossibleClients.filter(c => c.toLowerCase().includes(clientSearchQuery.toLowerCase()));
  }, [allPossibleClients, clientSearchQuery]);

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

  const handleCopyAWBOnly = async () => {
    if (!filteredRows.length) return;
    const text = filteredRows.map(r => `'${r.awb}`).join('\n');
    try {
      await navigator.clipboard.writeText(text);
      const clientMsg = selectedClients.length === 0 ? 'All Clients' : selectedClients.length === 1 ? selectedClients[0] : `${selectedClients.length} Clients`;
      showToast(`Copied ${filteredRows.length} AWB — ${clientMsg}`, "ok");
    } catch (err) {
      showToast("Failed to copy AWB", "err");
    }
  };

  const handleCopyTable = useCallback(async () => {
    if (!filteredRows.length) return;
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
    XLSX.writeFile(wb, `POD_Report_${currentSession?.dspId || 'Export'}.xlsx`);
    showToast("Report downloaded successfully", "ok");
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
            else if (h === "Awb" || h === "DSP No" || h === "Order- No") cleanRow[h] = normalizeAWB(row[h]);
            else {
              const ik = allHeaders.find(k => k.trim().toLowerCase() === h.toLowerCase());
              cleanRow[h] = ik ? row[ik] : "";
            }
          });
          return cleanRow;
        });
        setReplacerData(processed);
        setReplacerMeta({ headers: targetHeaders, remarkKey });
        showToast(`Processed ${processed.length} rows`, "ok");
      } catch (err: any) {
        showToast("Failed to process sheet", "err");
      } finally {
        setIsProcessing(false);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const handleReplacerDownload = () => {
    if (!replacerData.length || !replacerMeta) return;
    const ws = XLSX.utils.json_to_sheet(replacerData.map(r => {
      const cleanRow: any = {};
      replacerMeta.headers.forEach(h => {
        if (h === "Awb" || h === "DSP No" || h === "Order- No") {
          cleanRow[h] = { v: String(r[h]), t: 's' };
        } else {
          cleanRow[h] = r[h];
        }
      });
      return cleanRow;
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Remark Replacer");
    XLSX.writeFile(wb, "EOD_Rejection_Remarks.xlsx");
  };

  const handleOTPFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (!selectedSessionId) {
      showToast("Select a session in Daily EOD Rejection tab first!", "err");
      e.target.value = "";
      return;
    }

    setIsProcessing(true);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array', raw: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(ws, { raw: true, defval: "" });
        
        const activeAWBMap = new Map<string, string>();
        currentSession?.data.forEach(row => {
          if (row.awb) activeAWBMap.set(row.awb.trim(), row.returnAddress || "");
        });

        const filtered: OTPRow[] = [];
        rawData.forEach((row: any) => {
          const keys = Object.keys(row);
          const findVal = (regex: RegExp) => {
            const key = keys.find(k => regex.test(k.toLowerCase().replace(/[\s_-]/g, "")));
            return key ? row[key] : "";
          };
          const awb = normalizeAWB(findVal(/waybill|awb/i));
          if (activeAWBMap.has(awb)) {
            const client = String(findVal(/client/i)).trim();
            const status = String(findVal(/current status|currentstatus/i)).trim();
            filtered.push({
              id: crypto.randomUUID(),
              awb,
              client,
              status,
              returnAddress: activeAWBMap.get(awb) || "",
              isFTPL: client.toUpperCase().includes('FTPL')
            });
          }
        });

        if (filtered.length === 0) showToast("No matching AWBs found between OTP file and current session.", "err");
        else showToast(`Matched ${filtered.length} rows with current session.`, "ok");
        
        setOtpData(filtered);
      } catch (err) {
        showToast("Failed to process OTP sheet", "err");
      } finally {
        setIsProcessing(false);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const handleOtpDownload = () => {
    if (!otpFilteredRows.length) return;
    const header = ['AWB Number', 'Client Name', 'Current Status', 'Return Address'];
    const data = otpFilteredRows.map(r => [
      { v: r.awb, t: 's' },
      r.client,
      r.status,
      r.returnAddress
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...data]);
    ws['!cols'] = [{ wch: 15 }, { wch: 25 }, { wch: 15 }, { wch: 40 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "OTP Check");
    XLSX.writeFile(wb, `OTP_Check_${selectedSessionId?.slice(0, 8)}.xlsx`);
  };

  const otpFilteredRows = useMemo(() => {
    if (otpStatusFilter === 'all') return otpData;
    return otpData.filter(r => r.status.toLowerCase() === otpStatusFilter.toLowerCase());
  }, [otpData, otpStatusFilter]);

  const otpStats = useMemo(() => {
    return {
      total: otpData.length,
      dispatched: otpData.filter(r => r.status.toLowerCase() === 'dispatched').length,
      rto: otpData.filter(r => r.status.toLowerCase() === 'rto').length,
      dto: otpData.filter(r => r.status.toLowerCase() === 'dto').length,
    };
  }, [otpData]);

  const toggleClientSelection = (client: string) => {
    setSelectedClients(prev => 
      prev.includes(client) ? prev.filter(c => c !== client) : [...prev, client]
    );
  };

  const handleSelectAllClients = (checked: boolean) => {
    if (checked) {
      const visible = displayedClients;
      setSelectedClients(prev => {
        const otherSelected = prev.filter(c => !visible.includes(c));
        return [...otherSelected, ...visible];
      });
    } else {
      const visible = displayedClients;
      setSelectedClients(prev => prev.filter(c => !visible.includes(c)));
    }
  };

  const handleSearchEnter = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      const visible = displayedClients;
      setSelectedClients(prev => {
        const otherSelected = prev.filter(c => !visible.includes(c));
        return [...otherSelected, ...visible];
      });
      setIsPopoverOpen(false);
    }
  };

  // Row Selection & Bulk Actions
  const toggleRowSelection = (id: string) => {
    setSelectedRowIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const toggleGroupSelection = (rows: PODRow[], checked: boolean) => {
    setSelectedRowIds(prev => {
      const next = new Set(prev);
      rows.forEach(r => {
        if (checked) next.add(r.id);
        else next.delete(r.id);
      });
      return next;
    });
  };

  const handleDeleteSelected = () => {
    if (selectedRowIds.size === 0) return;
    setSessions(prev => prev.map(s => 
      s.id === selectedSessionId 
        ? { ...s, data: s.data.filter(r => !selectedRowIds.has(r.id)) } 
        : s
    ));
    showToast(`Deleted ${selectedRowIds.size} selected rows`, "ok");
    setSelectedRowIds(new Set());
  };

  const handleCopyGroupAWB = async (rows: PODRow[], remark: string) => {
    const text = rows.map(r => `'${r.awb}`).join('\n');
    try {
      await navigator.clipboard.writeText(text);
      showToast(`Copied ${rows.length} AWB — ${remark}`, "ok");
    } catch (err) {
      showToast("Failed to copy AWB", "err");
    }
  };

  const handleDeleteGroup = (rows: PODRow[], remark: string) => {
    const idsToRemove = new Set(rows.map(r => r.id));
    setSessions(prev => prev.map(s => 
      s.id === selectedSessionId 
        ? { ...s, data: s.data.filter(r => !idsToRemove.has(r.id)) } 
        : s
    ));
    showToast(`Deleted ${rows.length} rows — ${remark}`, "ok");
    setSelectedRowIds(prev => {
      const next = new Set(prev);
      idsToRemove.forEach(id => next.delete(id));
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
          {[
            { id: 'eod', label: 'Daily EOD Rejection' },
            { id: 'remark', label: 'EOD Rejection Remark' },
            { id: 'otp', label: 'OTP Dispatch Check' }
          ].map(tab => (
            <button 
              key={tab.id}
              onClick={() => setActiveTab(tab.id as any)}
              className={cn(
                "h-full px-1 text-[13px] font-semibold transition-all relative border-b-2",
                activeTab === tab.id ? "text-white border-blue-500" : "text-slate-400 border-transparent hover:text-white"
              )}
            >
              {tab.label}
            </button>
          ))}
        </div>
        <div>
          <span className="text-[11px] font-black text-amber-400 uppercase tracking-widest">By Ashu</span>
        </div>
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
                <div className="space-y-1.5">
                  <label className="text-[12px] font-medium text-slate-500 flex items-center gap-1.5">
                    <Hash className="w-3.5 h-3.5" /> DSP ID
                  </label>
                  <input 
                    type="text" 
                    inputMode="numeric"
                    value={setupData.dspId} 
                    onChange={(e) => setSetupData({...setupData, dspId: e.target.value.replace(/\D/g, '')})} 
                    className="w-full bg-[#F9FAFB] border-[1.5px] border-[#D1D5DB] rounded-lg px-3.5 h-[42px] text-[14px] font-bold text-[#111827] outline-none focus:border-[#1976D2] focus:ring-4 focus:ring-blue-500/5 transition-all" 
                    placeholder="Enter DSP Number" 
                  />
                </div>
                <div className="space-y-1.5">
                  <label className="text-[12px] font-medium text-slate-500 flex items-center gap-1.5">
                    <User className="w-3.5 h-3.5" /> FE Name
                  </label>
                  <input type="text" value={setupData.feName} onChange={(e) => setSetupData({...setupData, feName: e.target.value})} className="w-full bg-[#F9FAFB] border-[1.5px] border-[#D1D5DB] rounded-lg px-3.5 h-[42px] text-[14px] font-bold text-[#111827] outline-none focus:border-[#1976D2] focus:ring-4 focus:ring-blue-500/5 transition-all" placeholder="Enter Field Executive Name" />
                </div>
                <div className="space-y-1.5">
                  <label className="text-[12px] font-medium text-slate-500 flex items-center gap-1.5">
                    <Calendar className="w-3.5 h-3.5" /> Date
                  </label>
                  <input type="date" value={setupData.date} onChange={(e) => setSetupData({...setupData, date: e.target.value})} className="w-full bg-[#F9FAFB] border-[1.5px] border-[#D1D5DB] rounded-lg px-3.5 h-[42px] text-[14px] font-bold text-[#111827] outline-none focus:border-[#1976D2] focus:ring-4 focus:ring-blue-500/5 transition-all" />
                </div>
              </div>

              <div 
                onClick={() => {
                  if (!setupData.feName || !setupData.dspId) {
                    showToast("Enter DSP ID and FE Name first!", "err");
                  }
                }}
                className={cn("border-2 border-dashed rounded-xl p-8 text-center transition-all cursor-pointer relative group", isProcessing ? "bg-slate-50 opacity-80" : "bg-slate-50 hover:bg-white hover:border-blue-500")}>
                <input 
                  type="file" 
                  disabled={isProcessing || !setupData.feName || !setupData.dspId} 
                  onChange={handleFileUpload} 
                  className="absolute inset-0 opacity-0 cursor-pointer z-10" 
                />
                {isProcessing ? (
                  <div className="flex flex-col items-center gap-3"><Loader2 className="w-8 h-8 text-blue-600 animate-spin" /><p className="text-sm font-black text-blue-600 uppercase">Processing Delhivery Sheet...</p></div>
                ) : (
                  <div className="space-y-3">
                    <div className="w-12 h-12 bg-white rounded-full flex items-center justify-center mx-auto shadow-sm border border-slate-100 group-hover:scale-110 transition-transform">
                      <FileSpreadsheet className="w-6 h-6 text-slate-400" />
                    </div>
                    <div>
                      <p className="text-sm font-black text-[#111827]">Upload Delhivery Excel/CSV</p>
                      <p className="text-[11px] text-slate-400 font-bold uppercase mt-1 tracking-widest">Select file after entering DSP details</p>
                    </div>
                  </div>
                )}
              </div>
            </div>

            {sessions.length > 0 && (
              <div className="bg-white rounded-xl border border-[#E2E8F0] p-6 space-y-6 shadow-sm">
                <div className="flex items-center justify-between px-1">
                  <h3 className="text-[13px] font-bold text-[#111827] tracking-tight">Recent Sessions</h3>
                  <button onClick={() => setSessions([])} className="text-[11px] font-bold text-rose-600 hover:underline">Clear All History</button>
                </div>
                <div className="grid grid-cols-[repeat(auto-fill,minmax(260px,1fr))] gap-3">
                  {sessions.map(s => {
                    const sessionStats = {
                      total: s.data.length,
                      pending: s.data.filter(x => x.status === 'pending').length,
                      rto: s.data.filter(x => x.status === 'rto').length,
                      dto: s.data.filter(x => x.status === 'dto').length,
                    };
                    return (
                      <div 
                        key={s.id}
                        onClick={() => { 
                          setSelectedSessionId(s.id); 
                          setStatusFilter('all'); 
                          setActiveRemarkChip(null); 
                          setSelectedClients([]);
                          setClientSearchQuery("");
                          setSelectedRowIds(new Set());
                        }}
                        className={cn(
                          "bg-white border border-slate-200 rounded-xl overflow-hidden shadow-sm hover:shadow-md transition-all cursor-pointer relative pl-3 p-3 pr-4 group flex flex-col justify-between h-full min-h-[120px] max-w-[280px]",
                          selectedSessionId === s.id ? "ring-2 ring-blue-500 border-transparent" : "hover:border-blue-300"
                        )}
                      >
                        <div className="absolute left-0 top-0 bottom-0 w-1 bg-blue-600" />
                        <div className="space-y-0.5 relative">
                          <div className="flex justify-between items-start">
                             <p className="text-[14px] font-extrabold text-[#111827] truncate pr-6 leading-tight">{s.feName}</p>
                             <button 
                                onClick={(e) => { e.stopPropagation(); setSessions(prev => prev.filter(x => x.id !== s.id)); }} 
                                className="absolute top-0 right-0 text-slate-400 hover:text-rose-600 p-1 transition-colors"
                             >
                                <X className="w-3.5 h-3.5" />
                             </button>
                          </div>
                          <p className="text-[11px] text-slate-500 font-bold uppercase tracking-tight">{s.dspId} — {s.date}</p>
                        </div>
                        <div className="mt-3 flex flex-wrap gap-1">
                          <span className="px-2 py-0.5 rounded-full bg-slate-100 text-[#374151] text-[10px] font-bold uppercase">{sessionStats.total} pkt</span>
                          <span className="px-2 py-0.5 rounded-full bg-amber-50 text-[#B45309] border border-amber-100 text-[10px] font-bold uppercase">{sessionStats.pending} pending</span>
                          <span className="px-2 py-0.5 rounded-full bg-rose-50 text-[#B91C1C] border border-rose-100 text-[10px] font-bold uppercase">{sessionStats.rto} rto</span>
                          <span className="px-2 py-0.5 rounded-full bg-emerald-50 text-[#15803D] border border-emerald-100 text-[10px] font-bold uppercase">{sessionStats.dto} dto</span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {currentSession && (
              <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
                <div className="bg-white rounded-xl border border-[#E2E8F0] shadow-sm overflow-hidden flex divide-x divide-slate-100">
                  {[
                    { id: 'all', label: 'All', val: stats.total, color: 'text-blue-600' },
                    { id: 'pending', label: 'Pending', val: stats.pending, color: 'text-[#B45309]' },
                    { id: 'dispatched', label: 'Dispatch', val: stats.dispatched, color: 'text-blue-600' },
                    { id: 'rto', label: 'RTO', val: stats.rto, color: 'text-[#B91C1C]' },
                    { id: 'dto', label: 'DTO', val: stats.dto, color: 'text-[#15803D]' }
                  ].map((t) => (
                    <button 
                      key={t.id}
                      onClick={() => { 
                        setStatusFilter(t.id); 
                        setActiveRemarkChip(null); 
                        setSelectedClients([]); 
                        setClientSearchQuery("");
                        setSelectedRowIds(new Set());
                      }}
                      className={cn(
                        "flex-1 py-6 flex flex-col items-center justify-center transition-all relative group h-[100px]",
                        statusFilter === t.id ? "bg-slate-50" : "hover:bg-slate-50/30"
                      )}
                    >
                      <span className={cn("text-[32px] font-extrabold leading-none mb-1 tracking-tighter", t.color)}>{t.val}</span>
                      <span className={cn("text-[13px] font-bold", statusFilter === t.id ? t.color : "text-slate-400")}>{t.label}</span>
                      {statusFilter === t.id && <div className="absolute bottom-0 left-0 w-full h-[3px] bg-blue-600" />}
                    </button>
                  ))}
                </div>

                {statusFilter === 'pending' && (
                  <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-6">
                    <div className="flex items-center justify-between mb-5">
                      <div>
                        <h4 className="text-[13px] font-bold text-[#111827] flex items-center gap-2">
                          <BarChart3 className="w-4 h-4 text-rose-500" />
                          Remark Breakdown — Pending
                        </h4>
                        <p className="text-[11px] font-medium text-slate-400 mt-1">Click any remark chip to filter</p>
                      </div>
                      {activeRemarkChip && (
                        <button onClick={() => { 
                          setActiveRemarkChip(null); 
                          setSelectedClients([]); 
                          setClientSearchQuery("");
                          setSelectedRowIds(new Set());
                        }} className="h-8 px-4 bg-slate-900 text-white rounded-lg text-[11px] font-bold">All Pending</button>
                      )}
                    </div>
                    <div className="flex flex-wrap gap-2.5">
                      {Object.entries(
                        currentSession.data.filter(r => r.status === 'pending').reduce((acc: Record<string, number>, r) => {
                          const rem = r.remark || "No Remark";
                          acc[rem] = (acc[rem] || 0) + 1;
                          return acc;
                        }, {})
                      ).sort((a, b) => b[1] - a[1]).map(([rem, count]) => {
                        const isRed = /reject|intact|barcode|content/i.test(rem);
                        return (
                          <button 
                            key={rem}
                            onClick={() => { 
                              setActiveRemarkChip(activeRemarkChip === rem ? null : rem); 
                              setSelectedClients([]); 
                              setClientSearchQuery("");
                              setSelectedRowIds(new Set());
                            }}
                            className={cn(
                              "px-4 py-2.5 rounded-lg border flex items-center gap-3 transition-all",
                              activeRemarkChip === rem 
                                ? "bg-blue-600 border-blue-600 text-white shadow-md" 
                                : isRed 
                                  ? "bg-rose-50 border-rose-100 text-[#B91C1C] hover:border-rose-400"
                                  : "bg-slate-50 border-slate-200 text-[#374151] hover:border-blue-400"
                            )}
                          >
                            <span className="text-[13px] font-semibold">{rem}</span>
                            <span className={cn("text-[11px] font-black min-w-[22px] h-[22px] rounded flex items-center justify-center", activeRemarkChip === rem ? "bg-white/20" : "bg-slate-900 text-white")}>{count}</span>
                          </button>
                        );
                      })}
                    </div>
                  </div>
                )}

                <div className="flex items-center justify-between gap-4">
                  <div className="flex items-center gap-2 bg-white border border-slate-200 rounded-xl p-1 shadow-sm">
                    <Popover open={isPopoverOpen} onOpenChange={(open) => { setIsPopoverOpen(open); if(!open) setClientSearchQuery(""); }}>
                      <PopoverTrigger asChild>
                        <button className="h-8 min-w-[200px] px-3 flex items-center justify-between text-[12px] font-bold text-[#374151] hover:bg-slate-50 rounded-lg transition-colors border border-slate-100">
                          <span className="truncate">
                            {selectedClients.length === 0 ? 'All Clients' : 
                             selectedClients.length === 1 ? selectedClients[0] : 
                             `${selectedClients.length} Clients Selected`}
                          </span>
                          <ChevronDown className="w-3.5 h-3.5 opacity-50 ml-2" />
                        </button>
                      </PopoverTrigger>
                      <PopoverContent className="w-[280px] p-0" align="start">
                        <div className="p-2 border-b">
                          <div className="relative">
                            <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
                            <input 
                              className="w-full text-[12px] pl-8 pr-3 py-1.5 border rounded-md outline-none focus:border-blue-500" 
                              placeholder="Search client..." 
                              value={clientSearchQuery}
                              onChange={(e) => setClientSearchQuery(e.target.value)}
                              onKeyDown={handleSearchEnter}
                            />
                          </div>
                        </div>
                        <div className="p-1 border-b bg-slate-50/50">
                           <div className="flex items-center space-x-2 px-3 py-2">
                            <Checkbox 
                              id="select-all" 
                              checked={displayedClients.length > 0 && displayedClients.every(c => selectedClients.includes(c))}
                              onCheckedChange={handleSelectAllClients}
                            />
                            <label htmlFor="select-all" className="text-[11px] font-bold text-slate-900 uppercase tracking-tight cursor-pointer">
                              (Select All Search Results)
                            </label>
                           </div>
                        </div>
                        <ScrollArea className="h-[280px]">
                          <div className="p-1">
                            {displayedClients.map(c => (
                              <div 
                                key={c}
                                className="flex items-center space-x-2 px-3 py-2 hover:bg-slate-50 rounded-md transition-colors cursor-pointer"
                                onClick={() => toggleClientSelection(c)}
                              >
                                <Checkbox 
                                  checked={selectedClients.includes(c)}
                                  onCheckedChange={() => toggleClientSelection(c)}
                                  onClick={(e) => e.stopPropagation()}
                                />
                                <span className={cn(
                                  "text-[12px] font-medium truncate flex-1",
                                  selectedClients.includes(c) ? "text-blue-600 font-bold" : "text-slate-700"
                                )}>
                                  {c}
                                </span>
                              </div>
                            ))}
                            {displayedClients.length === 0 && (
                              <div className="p-4 text-center text-[11px] text-slate-400 font-bold uppercase">No matching clients</div>
                            )}
                          </div>
                        </ScrollArea>
                        <div className="p-2 border-t flex justify-end">
                           <button 
                            onClick={() => setIsPopoverOpen(false)}
                            className="text-[11px] font-black text-blue-600 uppercase tracking-widest px-4 py-1.5 hover:bg-blue-50 rounded transition-colors"
                           >
                            Done
                           </button>
                        </div>
                      </PopoverContent>
                    </Popover>
                    
                    <button onClick={handleCopyAWBOnly} className="h-8 px-4 bg-[#0F172A] hover:bg-black text-white rounded-lg text-[11px] font-black uppercase tracking-wider flex items-center gap-2 transition-all active:scale-95 whitespace-nowrap">
                      <Copy className="w-3.5 h-3.5" /> Copy Selected AWB
                    </button>

                    {selectedRowIds.size > 0 && (
                      <button 
                        onClick={handleDeleteSelected}
                        className="h-8 px-4 bg-rose-600 hover:bg-rose-700 text-white rounded-lg text-[11px] font-black uppercase tracking-wider flex items-center gap-2 transition-all active:scale-95 shadow-md whitespace-nowrap min-w-fit"
                      >
                        <Trash2 className="w-3.5 h-3.5" /> Delete Selected ({selectedRowIds.size})
                      </button>
                    )}
                  </div>

                  <div className="flex items-center gap-4">
                    <div className="flex items-center gap-2">
                      <button onClick={downloadExcel} className="h-10 px-5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-[13px] font-bold flex items-center gap-2 shadow-sm transition-all whitespace-nowrap">Download Excel</button>
                      <button onClick={handleCopyTable} className="h-10 px-5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-[13px] font-bold flex items-center gap-2 shadow-sm transition-all whitespace-nowrap">Copy Table</button>
                    </div>
                    
                    <div className="h-8 w-px bg-slate-200" />
                    
                    <div className="relative">
                      <Search className="w-4 h-4 absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" />
                      <input type="text" placeholder="Search waybill, client..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="bg-white border-[1.5px] border-slate-200 rounded-xl pl-10 pr-4 h-10 text-[13px] font-semibold text-[#111827] outline-none w-[320px] focus:border-blue-500 shadow-sm" />
                    </div>
                  </div>
                </div>

                <div className="bg-white rounded-xl border-[1.5px] border-[#F97316] shadow-2xl overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-center border-collapse table-fixed bg-white">
                      <thead className="bg-[#0F172A] text-white">
                        <tr className="h-11">
                          <th style={{ width: '32px' }} className="px-2 text-center">
                            <input 
                              type="checkbox" 
                              className="w-3.5 h-3.5 border border-slate-300 rounded" 
                              checked={filteredRows.length > 0 && filteredRows.every(r => selectedRowIds.has(r.id))}
                              onChange={(e) => toggleGroupSelection(filteredRows, e.target.checked)}
                            />
                          </th>
                          <th style={{ width: '28px' }} className="px-1 text-center"><Trash2 className="w-3.5 h-3.5 opacity-40 mx-auto" /></th>
                          <th style={{ width: '80px' }} className="px-2 text-[11px] font-bold text-center">DSP ID</th>
                          <th style={{ width: '130px' }} className="px-2 text-[11px] font-bold text-center">AWB Number</th>
                          <th style={{ width: '110px' }} className="px-2 text-[11px] font-bold text-center">Client</th>
                          <th style={{ width: '110px' }} className="px-2 text-[11px] font-bold text-center">Order ID</th>
                          <th className="px-2 text-[11px] font-bold text-center">Remark</th>
                          {statusFilter === 'pending' && <th style={{ width: '200px' }} className="px-2 text-[11px] font-bold text-center">Return Address</th>}
                          <th style={{ width: '80px' }} className="px-2 text-[11px] font-bold text-center">FE Name</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr className="bg-slate-800 text-white h-8 border-b border-white/5">
                          <td colSpan={statusFilter === 'pending' ? 9 : 8} className="px-3">
                            <div className="flex items-center gap-3 text-center justify-center">
                              <span className="text-[12px] font-bold text-amber-400">{currentSession.dspId}</span>
                              <span className="w-px h-3 bg-white/20" />
                              <span className="text-[10px] font-semibold bg-white/10 px-1.5 py-0.5 rounded uppercase tracking-tight">{filteredRows.length} pkt</span>
                              <span className="ml-4 text-[10px] text-slate-400 font-medium">{currentSession.feName} — {currentSession.date}</span>
                            </div>
                          </td>
                        </tr>
                        {filteredRows.length > 0 ? (
                          statusFilter === 'pending' ? (
                            Object.entries(
                              filteredRows.reduce((acc: Record<string, PODRow[]>, row) => {
                                const key = row.remark || "No Remark";
                                if (!acc[key]) acc[key] = [];
                                acc[key].push(row);
                                return acc;
                              }, {})
                            )
                            .sort((a, b) => b[1].length - a[1].length)
                            .map(([remark, groupRows]) => (
                              <React.Fragment key={remark}>
                                <tr className="h-10 border-b border-white/10" style={{ background: 'linear-gradient(90deg, #0D1B2E, #1A2F4A)' }}>
                                  <td className="px-2 text-center">
                                    <input 
                                      type="checkbox" 
                                      className="w-3.5 h-3.5 border border-white/50 rounded bg-transparent"
                                      checked={groupRows.every(r => selectedRowIds.has(r.id))}
                                      onChange={(e) => toggleGroupSelection(groupRows, e.target.checked)}
                                    />
                                  </td>
                                  <td colSpan={statusFilter === 'pending' ? 8 : 7} className="px-3">
                                    <div className="flex items-center justify-between">
                                      <div className="flex items-center gap-2">
                                        <span className="text-amber-400 text-[14px]">★</span>
                                        <span className="text-[11px] font-bold text-white uppercase tracking-wider">{remark}</span>
                                        <span className="w-px h-3 bg-white/20 mx-1" />
                                        <span className="text-[10px] font-black text-amber-400 border border-amber-400/30 px-1.5 py-0.5 rounded uppercase tracking-tight">{groupRows.length} pkt</span>
                                      </div>
                                      <div className="flex items-center gap-2">
                                        <button 
                                          onClick={() => handleCopyGroupAWB(groupRows, remark)}
                                          className="text-[11px] font-semibold text-white/90 border border-white/30 rounded px-2.5 py-0.5 hover:bg-white/10 transition-colors"
                                        >
                                          Copy AWB
                                        </button>
                                        <button 
                                          onClick={() => handleDeleteGroup(groupRows, remark)}
                                          className="bg-red-600/15 border border-red-600/50 text-red-300 rounded-[6px] px-[10px] py-[3px] text-[11px] font-semibold hover:bg-red-600/30 hover:text-white transition-colors"
                                        >
                                          Delete All
                                        </button>
                                      </div>
                                    </div>
                                  </td>
                                </tr>
                                {groupRows.map((row) => (
                                  <tr key={row.id} className={cn("h-11 border-b border-[#FED7AA] hover:bg-blue-50/30 transition-colors group bg-white", selectedRowIds.has(row.id) && "bg-blue-50")}>
                                    <td className="px-2 text-center">
                                      <input 
                                        type="checkbox" 
                                        className="w-3 h-3 border border-slate-300 rounded" 
                                        checked={selectedRowIds.has(row.id)}
                                        onChange={() => toggleRowSelection(row.id)}
                                      />
                                    </td>
                                    <td className="px-1 text-center"><button onClick={() => setSessions(prev => prev.map(s => s.id === selectedSessionId ? {...s, data: s.data.filter(r => r.id !== row.id)} : s))} className="text-slate-400 hover:text-red-600 hover:bg-red-50 rounded p-1 transition-colors"><Trash2 className="w-3.5 h-3.5 mx-auto" /></button></td>
                                    <td className="px-2 text-[13px] font-bold text-[#374151] truncate">{row.dspId}</td>
                                    <td 
                                      onClick={async () => {
                                        const success = await copyDataToClipboard([{ 'AWB Number': row.awb }], ['AWB Number']);
                                        if (success) showToast(`AWB ${row.awb} copied!`, "ok");
                                      }}
                                      className="px-2 text-[13px] font-bold text-[#111827] font-mono tracking-tighter truncate cursor-pointer hover:underline"
                                    >
                                      {row.awb}
                                    </td>
                                    <td className="px-2 text-[13px] font-semibold text-[#1565C0] truncate">{row.client}</td>
                                    <td className="px-2 text-[13px] font-medium text-[#111827] truncate">{row.orderId}</td>
                                    <td className="px-2">
                                      <span className={cn(
                                        "px-2 py-1 rounded text-[11px] font-semibold border shadow-sm truncate inline-block max-w-full",
                                        row.isIntact ? "bg-rose-50 text-[#B91C1C] border-rose-200" : "bg-amber-50 text-[#B45309] border-amber-200"
                                      )}>
                                        {row.remark}
                                      </span>
                                    </td>
                                    <td className="px-2 text-[12px] text-[#374151] truncate text-center">{row.returnAddress || "—"}</td>
                                    <td className="px-2 text-[13px] font-bold text-[#374151] truncate">{row.feName}</td>
                                  </tr>
                                ))}
                              </React.Fragment>
                            ))
                          ) : (
                            filteredRows.map((row) => (
                              <tr key={row.id} className={cn("h-11 border-b border-[#FED7AA] hover:bg-blue-50/30 transition-colors group bg-white", selectedRowIds.has(row.id) && "bg-blue-50")}>
                                <td className="px-2 text-center">
                                  <input 
                                    type="checkbox" 
                                    className="w-3 h-3 border border-slate-300 rounded" 
                                    checked={selectedRowIds.has(row.id)}
                                    onChange={() => toggleRowSelection(row.id)}
                                  />
                                </td>
                                <td className="px-1 text-center"><button onClick={() => setSessions(prev => prev.map(s => s.id === selectedSessionId ? {...s, data: s.data.filter(r => r.id !== row.id)} : s))} className="text-slate-400 hover:text-red-600 hover:bg-red-50 rounded p-1 transition-colors"><Trash2 className="w-3.5 h-3.5 mx-auto" /></button></td>
                                <td className="px-2 text-[13px] font-bold text-[#374151] truncate">{row.dspId}</td>
                                <td 
                                  onClick={async () => {
                                    const success = await copyDataToClipboard([{ 'AWB Number': row.awb }], ['AWB Number']);
                                    if (success) showToast(`AWB ${row.awb} copied!`, "ok");
                                  }}
                                  className="px-2 text-[13px] font-bold text-[#111827] font-mono tracking-tighter truncate cursor-pointer hover:underline"
                                >
                                  {row.awb}
                                </td>
                                <td className="px-2 text-[13px] font-semibold text-[#1565C0] truncate">{row.client}</td>
                                <td className="px-2 text-[13px] font-medium text-[#111827] truncate">{row.orderId}</td>
                                <td className="px-2">
                                  <span className={cn(
                                    "px-2 py-1 rounded text-[11px] font-semibold border shadow-sm truncate inline-block max-w-full",
                                    row.isIntact ? "bg-rose-50 text-[#B91C1C] border-rose-200" : "bg-amber-50 text-[#B45309] border-amber-200"
                                  )}>
                                    {row.remark}
                                  </span>
                                </td>
                                <td className="px-2 text-[13px] font-bold text-[#374151] truncate">{row.feName}</td>
                              </tr>
                            ))
                          )
                        ) : (
                          <tr><td colSpan={statusFilter === 'pending' ? 9 : 8} className="h-32 text-center text-[13px] font-bold text-slate-300 uppercase tracking-widest">No matching data available</td></tr>
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
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
            <div className="lg:col-span-8 space-y-6">
              <div className="bg-white rounded-xl p-8 border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between mb-8">
                  <h2 className="text-lg font-extrabold text-[#111827] flex items-center gap-3">
                    <FileSpreadsheet className="w-6 h-6 text-emerald-600" /> EOD Rejection Remark Engine
                  </h2>
                </div>
                <div className="border-2 border-dashed rounded-xl p-12 text-center transition-all cursor-pointer relative bg-slate-50 hover:bg-white hover:border-blue-500">
                  <input type="file" onChange={handleReplacerUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                  <div className="space-y-4">
                    <Download className="w-10 h-10 text-slate-300 mx-auto" />
                    <div>
                      <p className="text-base font-black text-[#111827]">Drop Master EOD Sheet Here</p>
                      <p className="text-[11px] text-slate-400 font-bold uppercase tracking-widest mt-1">Automatic Delhivery Remark Converter</p>
                    </div>
                  </div>
                </div>
              </div>

              {replacerData.length > 0 && (
                <div className="bg-white rounded-xl border border-emerald-500/20 shadow-xl overflow-hidden">
                  <div className="p-4 bg-slate-50 border-b flex items-center justify-between">
                    <div className="flex gap-2">
                      <button onClick={handleReplacerDownload} className="h-10 px-5 bg-emerald-600 text-white rounded-lg text-[13px] font-bold flex items-center gap-2 shadow-sm">
                        <Download className="w-4 h-4" /> Download Excel
                      </button>
                      <button onClick={async () => {
                        const success = await copyDataToClipboard(replacerData, replacerMeta?.headers || []);
                        if (success) showToast(`Copied ${replacerData.length} rows`, "ok");
                      }} className="h-10 px-5 bg-blue-600 text-white rounded-lg text-[13px] font-bold flex items-center gap-2 shadow-sm">
                        <Copy className="w-4 h-4" /> Copy Table
                      </button>
                    </div>
                    <button onClick={() => setReplacerData([])} className="text-[11px] font-bold text-rose-600">Discard Data</button>
                  </div>
                  <div className="overflow-x-auto max-h-[600px]">
                    <table className="w-full text-center border-collapse text-[13px] bg-white">
                      <thead className="sticky top-0 bg-[#0F172A] text-white">
                        <tr className="h-12">
                          <th style={{ width: '40px' }} className="px-4 text-center">#</th>
                          {replacerMeta?.headers.map((h, i) => <th key={i} className="px-4 font-bold text-[10px] tracking-widest">{h}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {replacerData.map((row, idx) => (
                          <tr key={row.__id} className={cn("h-11 border-b transition-colors", row.__isReplaced ? "bg-emerald-50/40" : "bg-amber-50/40")}>
                            <td className="px-4 text-center text-slate-400">{idx + 1}</td>
                            {replacerMeta?.headers.map((h, i) => (
                              <td key={i} className={cn("px-4 font-semibold truncate max-w-[200px]", h === "Remarks Of NSL" && row.__isReplaced ? "text-[#15803D] font-bold" : h === "Remarks Of NSL" ? "text-[#B45309]" : "text-[#111827]")}>
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
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden sticky top-20">
                <div className="bg-[#1976D2] p-4 text-white font-black uppercase text-[12px] tracking-widest">Replacement Matrix</div>
                <div className="p-4 space-y-4 max-h-[70vh] overflow-y-auto">
                  {Object.entries(REMARK_MAPPING).map(([nsl, off]) => (
                    <div key={nsl} className="p-4 bg-slate-50 rounded-xl border border-slate-100 space-y-2">
                      <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">Delhivery Remark</p>
                      <p className="text-[13px] font-bold text-[#B45309] leading-tight">{nsl}</p>
                      <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest pt-2">RPC Replacement</p>
                      <p className="text-[13px] font-extrabold text-[#15803D] leading-tight">{off}</p>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        )}

        {activeTab === "otp" && (
          <div className="space-y-6">
            <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-8 text-center space-y-6">
              <div className="flex flex-col items-center gap-3">
                <div className="w-12 h-12 bg-blue-50 rounded-full flex items-center justify-center text-blue-600">
                  <PackageSearch className="w-7 h-7" />
                </div>
                <div>
                  <h2 className="text-xl font-extrabold text-[#111827]">OTP Dispatch Checker</h2>
                  <p className="text-sm text-slate-500 font-medium">Upload Delhivery OTP Export to match Return Addresses from session</p>
                </div>
              </div>

              {!selectedSessionId && (
                <div className="p-4 bg-amber-50 border border-amber-200 rounded-xl max-w-2xl mx-auto flex items-center gap-3">
                  <div className="w-8 h-8 rounded-full bg-amber-100 flex items-center justify-center text-amber-600">!</div>
                  <p className="text-[13px] font-bold text-amber-800">Please select a session in Daily EOD Rejection tab first, then upload OTP file here.</p>
                </div>
              )}

              <div className={cn("border-2 border-dashed rounded-xl p-10 transition-all cursor-pointer relative bg-slate-50 hover:bg-white hover:border-blue-500 max-w-2xl mx-auto", isProcessing && "opacity-50")}>
                <input type="file" onChange={handleOTPFileUpload} disabled={isProcessing} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                <div className="space-y-3">
                  <Download className="w-8 h-8 text-slate-300 mx-auto" />
                  <p className="text-sm font-black text-[#111827]">Upload OTP Report</p>
                </div>
              </div>
            </div>

            {otpData.length > 0 && (
              <div className="space-y-6">
                <div className="bg-white rounded-xl border border-[#E2E8F0] shadow-sm overflow-hidden flex divide-x divide-slate-100">
                  {[
                    { id: 'all', label: 'All', val: otpStats.total, color: 'text-blue-600' },
                    { id: 'dispatched', label: 'Dispatched', val: otpStats.dispatched, color: 'text-[#15803D]' },
                    { id: 'rto', label: 'RTO', val: otpStats.rto, color: 'text-[#B91C1C]' },
                    { id: 'dto', label: 'DTO', val: otpStats.dto, color: 'text-blue-600' }
                  ].map((t) => (
                    <button 
                      key={t.id}
                      onClick={() => setOtpStatusFilter(t.id)}
                      className={cn(
                        "flex-1 py-6 flex flex-col items-center justify-center transition-all relative h-[100px]",
                        otpStatusFilter === t.id ? "bg-slate-50" : "hover:bg-slate-50/30"
                      )}
                    >
                      <span className={cn("text-[32px] font-extrabold leading-none mb-1", t.color)}>{t.val}</span>
                      <span className={cn("text-[13px] font-bold", otpStatusFilter === t.id ? t.color : "text-slate-400")}>{t.label}</span>
                      {otpStatusFilter === t.id && <div className="absolute bottom-0 left-0 w-full h-[3px] bg-blue-600" />}
                    </button>
                  ))}
                </div>

                <div className="flex items-center justify-between">
                  <div className="flex gap-2">
                    <button onClick={handleOtpDownload} className="h-10 px-5 bg-emerald-600 text-white rounded-lg text-[13px] font-bold flex items-center gap-2 shadow-lg">Download Excel</button>
                    <button onClick={async () => {
                      const headers = ['AWB Number', 'Client Name', 'Current Status', 'Return Address'];
                      const data = otpFilteredRows.map(r => ({
                        'AWB Number': r.awb,
                        'Client Name': r.client,
                        'Current Status': r.status,
                        'Return Address': r.returnAddress
                      }));
                      const success = await copyDataToClipboard(data, headers);
                      if (success) showToast(`Copied ${otpFilteredRows.length} rows`, "ok");
                    }} className="h-10 px-5 bg-blue-600 text-white rounded-lg text-[13px] font-bold flex items-center gap-2 shadow-lg">Copy Table</button>
                  </div>
                </div>

                <div className="bg-white rounded-xl border-[1.5px] border-[#F97316] shadow-2xl overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-center border-collapse bg-white">
                      <thead className="bg-[#0F172A] text-white">
                        <tr className="h-11">
                          <th style={{ width: '32px' }} className="px-2 text-center"><input type="checkbox" className="w-3.5 h-3.5 border border-slate-300 rounded" /></th>
                          <th className="px-4 text-[11px] font-bold text-center">AWB Number</th>
                          <th className="px-4 text-[11px] font-bold text-center">Client Name</th>
                          <th className="px-4 text-[11px] font-bold text-center">Current Status</th>
                          <th className="px-4 text-[11px] font-bold text-center">Return Address</th>
                        </tr>
                      </thead>
                      <tbody>
                        {otpFilteredRows.length > 0 ? otpFilteredRows.map((row) => (
                          <tr key={row.id} className={cn(
                            "h-11 border-b border-[#FED7AA] transition-colors bg-white",
                            row.isFTPL ? "bg-rose-50/50 border-l-[3px] border-l-rose-500" : ""
                          )}>
                            <td className="px-2 text-center"><input type="checkbox" className="w-3 h-3 border border-slate-300 rounded" /></td>
                            <td className="px-4 text-[13px] font-bold text-[#111827] font-mono tracking-tight">{row.awb}</td>
                            <td className={cn("px-4 text-[13px] font-semibold truncate", row.isFTPL ? "text-[#DC2626] font-bold" : "text-[#374151]")}>
                              {row.client}
                            </td>
                            <td className="px-4">
                              <span className={cn(
                                "px-2 py-0.5 rounded text-[10px] font-black uppercase border shadow-sm",
                                row.status.toLowerCase() === 'dispatched' ? "bg-emerald-50 text-[#15803D] border-emerald-200" :
                                row.status.toLowerCase() === 'rto' ? "bg-rose-50 text-[#B91C1C] border-rose-200" :
                                "bg-blue-50 text-blue-600 border-blue-200"
                              )}>
                                {row.status}
                              </span>
                            </td>
                            <td className="px-4 text-[13px] font-medium text-slate-500 text-center">{row.returnAddress || "—"}</td>
                          </tr>
                        )) : (
                          <tr><td colSpan={5} className="h-32 text-center text-slate-300 font-bold uppercase tracking-widest">No matching AWBs found between OTP file and current session.</td></tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}
      </main>
    </div>
  );
}
