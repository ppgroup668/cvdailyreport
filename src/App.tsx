/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback, useMemo, Fragment, ChangeEvent, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { 
  FileUp, 
  FileSpreadsheet, 
  Trash2, 
  Download, 
  LayoutDashboard, 
  AlertCircle,
  CheckCircle2,
  ChevronRight,
  TrendingUp,
  Users,
  Share2,
  Copy,
  Check,
  ExternalLink,
  Calendar
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { cn } from './lib/utils';
import { 
  HeadcountRow, 
  ProductionRow, 
  SalesRecord, 
  ManagerGroup,
  DistrictGroup
} from './types';

export default function App() {
  const [headcountFile, setHeadcountFile] = useState<File | null>(null);
  const [productionFile, setProductionFile] = useState<File | null>(null);
  const [reportData, setReportData] = useState<DistrictGroup[] | null>(null);
  const [reportDate, setReportDate] = useState<string>(new Date().toISOString().split('T')[0]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [activeView, setActiveView] = useState<'dashboard' | 'reports'>('dashboard');
  const [logoUrl, setLogoUrl] = useState<string | null>(null);
  const [availableTeams, setAvailableTeams] = useState<string[]>([]);
  const [selectedTeams, setSelectedTeams] = useState<string[]>([]);

  // Column widths state (initially: Manager, Name, FYCC, Case)
  const [colWidths, setColWidths] = useState({
    manager: 240,
    name: 290,
    fycc: 190,
    case: 140
  });

  // Self-adjust row padding (height) and font size
  const [rowPadding, setRowPadding] = useState(16); // vertical padding in pixels (matches py-4)
  const [fontSize, setFontSize] = useState(24);     // font size in pixels (default 24px)

  // --- Sharing & Viewer States ---
  const [shareId, setShareId] = useState<string | null>(() => {
    return localStorage.getItem('sales_dashboard_share_id') || null;
  });
  const [isViewerMode, setIsViewerMode] = useState(false);
  const [viewerShareId, setViewerShareId] = useState<string | null>(null);
  const [viewerAvailableDates, setViewerAvailableDates] = useState<string[]>([]);
  const [viewerSelectedDate, setViewerSelectedDate] = useState<string>('');
  const [isViewerLoading, setIsViewerLoading] = useState(false);
  const [isPublishing, setIsPublishing] = useState(false);
  const [publishSuccessMessage, setPublishSuccessMessage] = useState<string | null>(null);
  const [copied, setCopied] = useState(false);

  // Helper to load date-specific viewer data
  const loadViewerReportForDate = async (targetShareId: string, targetDate: string) => {
    setIsViewerLoading(true);
    try {
      const res = await fetch(`/api/dashboard/share/${targetShareId}/date/${targetDate}`);
      if (!res.ok) {
        throw new Error('找不到指定日期的報表數據。');
      }
      const data = await res.json();
      setReportData(data.reportData);
      setReportDate(data.date);
      if (data.selectedTeams) {
        setSelectedTeams(data.selectedTeams);
      }
      setActiveView('reports');
    } catch (err: any) {
      setError(err.message || '載入日期數據失敗。');
    } finally {
      setIsViewerLoading(false);
    }
  };

  // Helper to load whole workspace
  const loadViewerWorkspace = async (targetShareId: string, targetDate: string | null) => {
    setIsViewerLoading(true);
    setError(null);
    try {
      const res = await fetch(`/api/dashboard/share/${targetShareId}`);
      if (!res.ok) {
        throw new Error('找不到該分享的儀表板。');
      }
      const data = await res.json();
      if (data.logo) {
        setLogoUrl(data.logo);
      }
      setViewerAvailableDates(data.availableDates);
      
      let selectedD = targetDate;
      if (!selectedD || !data.availableDates.includes(selectedD)) {
        selectedD = data.availableDates[0] || '';
      }
      
      setViewerSelectedDate(selectedD);
      
      if (selectedD) {
        await loadViewerReportForDate(targetShareId, selectedD);
      } else {
        setError('該儀表板尚無發佈任何日期數據。');
      }
    } catch (err: any) {
      setError(err.message || '載入分享儀表板失敗。');
    } finally {
      setIsViewerLoading(false);
    }
  };

  // Viewer date changer
  const handleViewerDateChange = async (newDate: string) => {
    if (!viewerShareId) return;
    setViewerSelectedDate(newDate);
    const newUrl = `${window.location.pathname}?share=${viewerShareId}&date=${newDate}`;
    window.history.pushState({ path: newUrl }, '', newUrl);
    await loadViewerReportForDate(viewerShareId, newDate);
  };

  // Exit viewer mode
  const handleExitViewer = () => {
    setIsViewerMode(false);
    setViewerShareId(null);
    setReportData(null);
    setLogoUrl(null);
    setActiveView('dashboard');
    const newUrl = window.location.pathname;
    window.history.pushState({ path: newUrl }, '', newUrl);
  };

  // Helper to copy text to clipboard with iframe-safe fallback
  const copyToClipboard = (text: string) => {
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text);
      } else {
        const input = document.createElement('input');
        input.value = text;
        document.body.appendChild(input);
        input.select();
        document.execCommand('copy');
        document.body.removeChild(input);
      }
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
      return true;
    } catch (err) {
      console.error('Failed to copy: ', err);
      return false;
    }
  };

  // Publish / Share currently loaded report data
  const publishCurrentDashboard = async () => {
    if (!reportData) {
      setError('沒有可發佈的報表數據。');
      return;
    }
    setIsPublishing(true);
    setError(null);
    setPublishSuccessMessage(null);
    setCopied(false);
    
    try {
      const response = await fetch('/api/dashboard/publish', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          shareId: shareId,
          date: reportDate,
          logo: logoUrl,
          reportData: reportData,
          grandTotals: grandTotals,
          selectedTeams: selectedTeams
        })
      });
      
      if (!response.ok) {
        let errMsg = '發佈失敗，請稍後重試。';
        try {
          const errData = await response.json();
          if (errData && errData.error) errMsg = errData.error;
        } catch (_) {}
        throw new Error(errMsg);
      }
      
      const result = await response.json();
      if (result.success) {
        setShareId(result.shareId);
        localStorage.setItem('sales_dashboard_share_id', result.shareId);
        
        const shareLink = `${window.location.origin}/?share=${result.shareId}&date=${result.date}`;
        setPublishSuccessMessage(shareLink);
      }
    } catch (err: any) {
      setError(err.message || '發佈報表時出錯。');
    } finally {
      setIsPublishing(false);
    }
  };

  // URL query parameter check on mount
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const shareParam = params.get('share');
    if (shareParam) {
      setIsViewerMode(true);
      setViewerShareId(shareParam);
      loadViewerWorkspace(shareParam, params.get('date'));
    }
  }, []);

  const handleResizeStart = (col: 'manager' | 'name' | 'fycc' | 'case', e: React.MouseEvent) => {
    e.preventDefault();
    const startX = e.clientX;
    const startWidth = colWidths[col];
    
    const handleMouseMove = (moveEvent: MouseEvent) => {
      const deltaX = moveEvent.clientX - startX;
      setColWidths(prev => ({
        ...prev,
        [col]: Math.max(80, startWidth + deltaX)
      }));
    };
    
    const handleMouseUp = () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };
    
    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
  };

  const handleRowHeightResizeStart = (e: React.MouseEvent) => {
    e.preventDefault();
    const startY = e.clientY;
    const startPadding = rowPadding;
    
    const handleMouseMove = (moveEvent: MouseEvent) => {
      const deltaY = moveEvent.clientY - startY;
      // top + bottom padding change (scaled to make dragging smooth)
      setRowPadding(Math.max(4, startPadding + Math.round(deltaY / 2)));
    };
    
    const handleMouseUp = () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };
    
    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
  };

  // 當 headcountFile 改變時，異步讀取並提取所有 unique teams
  useEffect(() => {
    if (!headcountFile) {
      setAvailableTeams([]);
      setSelectedTeams([]);
      return;
    }

    const readTeams = async () => {
      try {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const data = e.target?.result;
            const workbook = XLSX.read(data, { type: 'binary' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const arrays = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
            
            const teamsSet = new Set<string>();
            arrays.forEach((row, idx) => {
              if (idx === 0) return; // skip header
              const colBRaw = String(row[1] || '').trim().toUpperCase();
              if (colBRaw) {
                teamsSet.add(colBRaw);
              }
            });

            const sortedTeams = Array.from(teamsSet).sort((a, b) => a.localeCompare(b));
            setAvailableTeams(sortedTeams);

            // 尋找包含 "CALVIN WONG" 的 Team 作為預設選取，若無則選第一個
            const calvinTeam = sortedTeams.find(t => t.includes('CALVIN WONG'));
            if (calvinTeam) {
              setSelectedTeams([calvinTeam]);
            } else if (sortedTeams.length > 0) {
              setSelectedTeams([sortedTeams[0]]);
            }
          } catch (err) {
            console.error("Error reading teams from headcount file:", err);
          }
        };
        reader.readAsBinaryString(headcountFile);
      } catch (err) {
        console.error(err);
      }
    };

    readTeams();
  }, [headcountFile]);

  const handleLogoUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        const base64String = reader.result as string;
        setLogoUrl(base64String);
      };
      reader.readAsDataURL(file);
    }
  };

  const handleFileUpload = (e: ChangeEvent<HTMLInputElement>, type: 'headcount' | 'production') => {
    const file = e.target.files?.[0];
    if (file) {
      if (type === 'headcount') setHeadcountFile(file);
      else setProductionFile(file);
      setReportData(null);
      setError(null);
    }
  };

  const clearFiles = () => {
    setHeadcountFile(null);
    setProductionFile(null);
    setReportData(null);
    setError(null);
  };

  const processFiles = useCallback(async () => {
    if (!headcountFile || !productionFile) {
      setError("請先上傳兩個 Excel 檔案。");
      return;
    }

    setIsProcessing(true);
    setError(null);

    try {
      // 讀取為 2D 陣列以便根據索引 (A, B, C...) 讀取
      const readFileAsArrays = (file: File): Promise<any[][]> => {
        return new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = (e) => {
            try {
              const data = e.target?.result;
              const workbook = XLSX.read(data, { type: 'binary' });
              const firstSheetName = workbook.SheetNames[0];
              const worksheet = workbook.Sheets[firstSheetName];
              const arrays = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
              resolve(arrays);
            } catch (err) {
              reject(err);
            }
          };
          reader.onerror = reject;
          reader.readAsBinaryString(file);
        });
      };

      const [headcountRows, productionRows] = await Promise.all([
        readFileAsArrays(headcountFile),
        readFileAsArrays(productionFile)
      ]);

      /**
       * Headcount 處理:
       * B 行 (Index 1): District
       * K 行 (Index 10): Manager (Upline Manager Name)
       * D 行 (Index 3): Name (HKID)
       */
      const normalizeName = (name: string) => name.replace(/[^a-zA-Z0-9\u4e00-\u9fa5]/g, '').toUpperCase();
      const headcountMap = new Map<string, { manager: string; originalName: string; district: string }>();
      const nameToChiMap = new Map<string, string>();
      const personTeamMap = new Map<string, string>();
      
      headcountRows.forEach((row, idx) => {
        if (idx === 0) return; // 跳過標題
        const nameRaw = String(row[3] || '').trim();
        const managerRaw = String(row[10] || '').trim();
        const chineseNameRaw = String(row[2] || '').trim();
        const colBRaw = String(row[1] || '').trim().toUpperCase();
        
        if (nameRaw && chineseNameRaw) {
          nameToChiMap.set(normalizeName(nameRaw), chineseNameRaw);
        }

        // 記錄所有人的 Team (以備經理比對)
        if (nameRaw && colBRaw) {
          personTeamMap.set(normalizeName(nameRaw), colBRaw);
        }
        if (chineseNameRaw && colBRaw) {
          personTeamMap.set(normalizeName(chineseNameRaw), colBRaw);
        }

        // 篩選 B 行是否屬於已選擇的團隊 (Team)
        const isMatchedTeam = selectedTeams && selectedTeams.length > 0
          ? selectedTeams.map(t => t.toUpperCase()).includes(colBRaw)
          : colBRaw.includes('CALVIN WONG');

        if (nameRaw && isMatchedTeam) {
          headcountMap.set(normalizeName(nameRaw), { 
            manager: managerRaw, 
            originalName: nameRaw, 
            district: colBRaw 
          });
        }
      });

      // 尋找經理資訊的輔助函式
      const findManagerInfo = (mName: string) => {
        if (!mName) return null;
        const normM = normalizeName(mName);
        
        // 1. 精確匹配
        if (personTeamMap.has(normM)) {
          return {
            team: personTeamMap.get(normM) || "",
            chineseName: nameToChiMap.get(normM) || null
          };
        }

        // 2. 模糊匹配 (包含關係，排除極短名字)
        for (const [headcountNorm, team] of personTeamMap.entries()) {
          if (headcountNorm.length >= 4 && normM.length >= 4) {
            if (headcountNorm.includes(normM) || normM.includes(headcountNorm)) {
              return {
                team,
                chineseName: nameToChiMap.get(headcountNorm) || null
              };
            }
          }
        }

        return null;
      };

      let extractedDate = new Date().toISOString().split('T')[0];
      if (productionRows.length > 1 && productionRows[1][2]) {
        const c2Val = productionRows[1][2];
        if (typeof c2Val === 'number') {
           const dateObj = new Date(Math.round((c2Val - 25569) * 86400 * 1000));
           extractedDate = dateObj.toISOString().split('T')[0];
        } else {
           extractedDate = String(c2Val);
        }
      }
      setReportDate(extractedDate);

      /**
       * Sales Production 處理:
       * E 行 (Index 4): Name (HKID)
       * O 行 (Index 14): Case
       * P 行 (Index 15): FYCC
       */
      const productionMap = new Map<string, { fyc: number; cases: number; manager: string; originalName: string }>();
      productionRows.forEach((row, idx) => {
        if (idx === 0) return;
        const nameRaw = String(row[4] || '').trim();
        const casesRaw = String(row[14] || '0').replace(/,/g, ''); // O 行
        const fycRaw = String(row[15] || '0').replace(/,/g, '');   // P 行
        const fyc = Math.round(parseFloat(fycRaw) || 0);            // P 行
        const cases = parseFloat(casesRaw) || 0;                    // O 行
        const managerFromProd = String(row[2] || '').trim(); // C 行

        // 只要大於等於 0.5 就顯示 (使用者要求: > 0.5 都顯示)
        if (nameRaw && (fyc >= 0.5 || cases >= 0.5)) {
          const key = normalizeName(nameRaw);
          // 若沒找到此人，但可能 headcount 過濾失敗，做一個 fallback 顯示所有資料
          const isMatched = headcountMap.has(key);

          if (isMatched) {
            const existing = productionMap.get(key) || { fyc: 0, cases: 0, manager: "", originalName: "" };
            const managerInfo = findManagerInfo(managerFromProd);

            // 1. MATCH 不到 MANAGER 名字，要隱藏
            // 2. 經理必須屬於已選擇的團隊之一，否則不應該出現在該團隊的報表
            let isManagerInSelectedTeams = false;
            if (managerInfo) {
              const managerTeamNorm = managerInfo.team.toUpperCase();
              if (selectedTeams && selectedTeams.length > 0) {
                isManagerInSelectedTeams = selectedTeams.map(t => t.toUpperCase()).includes(managerTeamNorm);
              } else {
                isManagerInSelectedTeams = managerTeamNorm.includes('CALVIN WONG');
              }
            }
            
            if (managerInfo && isManagerInSelectedTeams) {
              const m = managerInfo.chineseName || managerFromProd;
              const originalName = headcountMap.get(key)?.originalName || nameRaw;
              
              if (m && m.trim() !== "" && m !== "Unassigned") {
                productionMap.set(key, {
                  fyc: existing.fyc + fyc,
                  cases: existing.cases + cases,
                  manager: m,
                  originalName: originalName || existing.originalName || nameRaw
                });
              }
            }
          }
        }
      });

      const mergedRecords: SalesRecord[] = [];
      productionMap.forEach((data, key) => {
        if (data.manager && data.manager.trim() !== "" && data.manager !== "Unassigned") {
          const hc = headcountMap.get(key);
          const district = hc?.district || "";
          mergedRecords.push({
            name: data.originalName,
            manager: data.manager,
            fyc: data.fyc,
            cases: data.cases,
            district
          });
        }
      });

      // 先依 District 分類
      const districtMap = new Map<string, SalesRecord[]>();
      mergedRecords.forEach(record => {
        const dist = record.district || "OTHER";
        if (!districtMap.has(dist)) {
          districtMap.set(dist, []);
        }
        districtMap.get(dist)!.push(record);
      });

      const processedDistricts: DistrictGroup[] = [];
      districtMap.forEach((records, distName) => {
        const groups: { [key: string]: SalesRecord[] } = {};
        records.forEach(record => {
          if (!groups[record.manager]) groups[record.manager] = [];
          groups[record.manager].push(record);
        });

        const managerGroups: ManagerGroup[] = Object.entries(groups).map(([manager, recs]) => ({
          manager,
          records: recs.sort((a, b) => a.name.localeCompare(b.name)),
          totalFYC: recs.reduce((sum, r) => sum + r.fyc, 0),
          totalCases: recs.reduce((sum, r) => sum + r.cases, 0)
        })).sort((a, b) => b.totalFYC - a.totalFYC);

        const totalFYC = managerGroups.reduce((sum, mg) => sum + mg.totalFYC, 0);
        const totalCases = managerGroups.reduce((sum, mg) => sum + mg.totalCases, 0);

        processedDistricts.push({
          district: distName,
          managerGroups,
          totalFYC,
          totalCases
        });
      });

      // 按 District 總 FYC 排序
      processedDistricts.sort((a, b) => b.totalFYC - a.totalFYC);

      setReportData(processedDistricts);
      setActiveView('reports');
      // 自動滾動到頂部
      setTimeout(() => {
        window.scrollTo({ top: 0, behavior: 'smooth' });
      }, 50);
    } catch (err) {
      setError("處理檔案出錯，請確保 Excel 格式正確且包含必要欄位。");
      console.error(err);
    } finally {
      setIsProcessing(false);
    }
  }, [headcountFile, productionFile, selectedTeams]);

  const grandTotals = useMemo(() => {
    if (!reportData) return { fyc: 0, cases: 0 };
    return reportData.reduce((acc, distGroup) => ({
      fyc: acc.fyc + distGroup.totalFYC,
      cases: acc.cases + distGroup.totalCases
    }), { fyc: 0, cases: 0 });
  }, [reportData]);

  const isMultiTeam = useMemo(() => {
    return selectedTeams && selectedTeams.length > 1;
  }, [selectedTeams]);

  const theme = useMemo(() => {
    if (isMultiTeam) {
      return {
        outerBg: 'bg-[#FFB020]', // Amber/Orange-Yellow
        outerBorder: 'border-[#D97706]',
        headerBg: 'bg-[#F59E0B]',
        innerBorder: 'border-[#D97706]',
        textColor: 'text-amber-900',
        districtRowBg: 'bg-amber-50/70',
        districtRowText: 'text-amber-950',
        districtRowBorder: 'border-amber-100',
        bulletColor: 'text-[#D97706]',
        hoverBg: 'hover:bg-[#D97706]/60',
        accentText: 'text-[#D97706]',
        btnAccent: 'accent-[#D97706]',
        spinnerBorder: 'border-t-[#D97706]'
      };
    } else {
      return {
        outerBg: 'bg-[#5AC8FA]', // Classic Blue
        outerBorder: 'border-[#419CD8]',
        headerBg: 'bg-[#60A5FA]',
        innerBorder: 'border-[#419CD8]',
        textColor: 'text-blue-900',
        districtRowBg: 'bg-blue-50/70',
        districtRowText: 'text-blue-950',
        districtRowBorder: 'border-blue-100',
        bulletColor: 'text-[#419CD8]',
        hoverBg: 'hover:bg-[#419CD8]/60',
        accentText: 'text-[#419CD8]',
        btnAccent: 'accent-[#419CD8]',
        spinnerBorder: 'border-t-[#419CD8]'
      };
    }
  }, [isMultiTeam]);

  const totalTableWidth = useMemo(() => {
    return colWidths.manager + colWidths.name + colWidths.fycc + colWidths.case;
  }, [colWidths]);

  return (
    <div className="flex flex-col min-h-screen bg-[#F8FAFC] font-sans pb-20">
      {/* Viewer Mode Banner */}
      {isViewerMode && (
        <div className="bg-gradient-to-r from-blue-600 to-[#419CD8] text-white text-xs font-bold px-6 py-2.5 flex items-center justify-between shadow-md">
          <div className="flex items-center gap-2">
            <span className="flex h-2 w-2 rounded-full bg-emerald-400 animate-pulse" />
            <span>您正在瀏覽唯讀分享的「銷售業績儀表板」</span>
          </div>
          <button
            onClick={handleExitViewer}
            className="bg-white/20 hover:bg-white/30 text-white px-3 py-1 rounded-md text-[10px] font-bold transition-all border border-white/20 uppercase cursor-pointer"
          >
            建立我的報表 ➜
          </button>
        </div>
      )}

      {/* Header */}
      <header className="h-16 bg-white border-b border-slate-200 flex items-center justify-between px-6 flex-shrink-0 z-10 shadow-sm">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 bg-[#419CD8] rounded-lg flex items-center justify-center shadow-lg shadow-blue-100">
            <LayoutDashboard className="w-5 h-5 text-white" />
          </div>
          <h1 className="text-xl font-bold text-slate-800 tracking-tight">
            {isViewerMode ? "銷售分享儀表板" : "每日銷售報告系統"}
          </h1>
        </div>
        {reportData && !isViewerMode && (
          <div className="flex bg-slate-100 p-1 rounded-xl">
            <button
              onClick={() => setActiveView('dashboard')}
              className={cn(
                "px-4 py-1.5 rounded-lg text-xs font-bold transition-all",
                activeView === 'dashboard'
                  ? "bg-white text-slate-800 shadow-sm"
                  : "text-slate-500 hover:text-slate-800"
              )}
            >
              數據上傳
            </button>
            <button
              onClick={() => setActiveView('reports')}
              className={cn(
                "px-4 py-1.5 rounded-lg text-xs font-bold transition-all",
                activeView === 'reports'
                  ? "bg-white text-slate-800 shadow-sm"
                  : "text-slate-500 hover:text-slate-800"
              )}
            >
              報表結果
            </button>
          </div>
        )}
      </header>

      <main className="flex-1 p-4 md:p-8">
        <div className="max-w-5xl mx-auto space-y-8">
          
          {isViewerLoading ? (
            <div className="flex flex-col items-center justify-center py-24 space-y-4 bg-white border border-slate-200 rounded-3xl shadow-sm">
              <div className={cn("w-12 h-12 border-4 border-slate-200 rounded-full animate-spin", theme.spinnerBorder)}></div>
              <p className="text-sm font-bold text-slate-600 animate-pulse">正在載入專屬分享數據，請稍候...</p>
            </div>
          ) : (
            <>
              {/* Upload Section */}
              {activeView === 'dashboard' && !isViewerMode && (
            <section className="bg-white rounded-3xl border border-slate-200 shadow-sm p-8">
            <div className="mb-8">
              <h2 className="text-2xl font-extrabold text-slate-800">上傳原始數據</h2>
              <p className="text-sm text-slate-500 mt-1">請上傳 Headcount 及 Sales Production Excel 檔案進行分析。</p>
            </div>

            <div className="grid md:grid-cols-2 gap-6">
              {/* File Slot 1 */}
              <div className={cn(
                "relative border-2 border-dashed rounded-3xl flex flex-col items-center justify-center p-8 transition-all group",
                headcountFile ? "bg-emerald-50 border-emerald-200" : "bg-slate-50 border-slate-200 hover:border-blue-400 cursor-pointer"
              )}>
                {!headcountFile && (
                  <input
                    type="file"
                    accept=".xlsx, .xls"
                    onChange={(e) => handleFileUpload(e, 'headcount')}
                    className="absolute inset-0 opacity-0 cursor-pointer z-10"
                  />
                )}
                
                {headcountFile && (
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      setHeadcountFile(null);
                      setReportData(null);
                    }}
                    className="absolute top-4 right-4 p-2 bg-white rounded-full shadow-sm border border-slate-200 text-slate-400 hover:text-red-500 hover:border-red-200 transition-colors z-20"
                    title="移除檔案"
                  >
                    <Trash2 className="w-5 h-5" />
                  </button>
                )}

                <div className={cn(
                  "w-14 h-14 bg-white rounded-2xl flex items-center justify-center shadow-md mb-4 border transition-transform group-hover:scale-110 relative z-0",
                  headcountFile ? "border-emerald-100" : "border-slate-100"
                )}>
                  {headcountFile ? <CheckCircle2 className="w-7 h-7 text-emerald-500" /> : <Users className="w-7 h-7 text-blue-500" />}
                </div>
                <p className="font-bold text-slate-700 text-lg relative z-0 text-center">{headcountFile ? headcountFile.name : "Headcount Excel (D/J行)"}</p>
                <div className="mt-4 flex items-center gap-2 text-[10px] font-black text-slate-400 uppercase tracking-widest relative z-0">
                  {headcountFile ? "已選擇檔案" : "點擊或拖放檔案"}
                </div>
              </div>

              {/* File Slot 2 */}
              <div className={cn(
                "relative border-2 border-dashed rounded-3xl flex flex-col items-center justify-center p-8 transition-all group",
                productionFile ? "bg-emerald-50 border-emerald-200" : "bg-slate-50 border-slate-200 hover:border-blue-400 cursor-pointer"
              )}>
                {!productionFile && (
                  <input
                    type="file"
                    accept=".xlsx, .xls"
                    onChange={(e) => handleFileUpload(e, 'production')}
                    className="absolute inset-0 opacity-0 cursor-pointer z-10"
                  />
                )}
                
                {productionFile && (
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      setProductionFile(null);
                      setReportData(null);
                    }}
                    className="absolute top-4 right-4 p-2 bg-white rounded-full shadow-sm border border-slate-200 text-slate-400 hover:text-red-500 hover:border-red-200 transition-colors z-20"
                    title="移除檔案"
                  >
                    <Trash2 className="w-5 h-5" />
                  </button>
                )}

                <div className={cn(
                  "w-14 h-14 bg-white rounded-2xl flex items-center justify-center shadow-md mb-4 border transition-transform group-hover:scale-110 relative z-0",
                  productionFile ? "border-emerald-100" : "border-slate-100"
                )}>
                  {productionFile ? <CheckCircle2 className="w-7 h-7 text-emerald-500" /> : <TrendingUp className="w-7 h-7 text-blue-500" />}
                </div>
                <p className="font-bold text-slate-700 text-lg relative z-0 text-center">{productionFile ? productionFile.name : "Production Excel (B/E行)"}</p>
                <div className="mt-4 flex items-center gap-2 text-[10px] font-black text-slate-400 uppercase tracking-widest relative z-0">
                  {productionFile ? "已選擇檔案" : "點擊或拖放檔案"}
                </div>
              </div>
            </div>

            {headcountFile && availableTeams.length > 0 && (
              <motion.div 
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                className="mt-8 p-6 bg-blue-50/40 rounded-2xl border border-blue-100 max-w-xl mx-auto shadow-sm animate-fade-in"
              >
                <div className="flex items-center justify-between mb-4">
                  <label className="text-sm font-extrabold text-slate-700 flex items-center gap-2">
                    <Users className="w-4 h-4 text-[#419CD8]" />
                    選擇要生成報表的團隊 (District / Team)
                  </label>
                  <span className="text-xs bg-blue-100/80 text-[#419CD8] font-black px-2.5 py-1 rounded-full shadow-sm">
                    已選擇 {selectedTeams.length} / {availableTeams.length} 個
                  </span>
                </div>

                {/* Quick actions */}
                <div className="flex flex-wrap gap-2 mb-4">
                  <button
                    onClick={() => {
                      setSelectedTeams([...availableTeams]);
                      setReportData(null);
                    }}
                    type="button"
                    className="text-xs font-bold text-[#419CD8] bg-white hover:bg-blue-50 border border-blue-200 px-3 py-1.5 rounded-lg shadow-sm transition-all cursor-pointer"
                  >
                    選擇所有團隊 (All Teams)
                  </button>
                  <button
                    onClick={() => {
                      const calvinTeam = availableTeams.find(t => t.includes('CALVIN WONG'));
                      if (calvinTeam) {
                        setSelectedTeams([calvinTeam]);
                      } else {
                        setSelectedTeams([]);
                      }
                      setReportData(null);
                    }}
                    type="button"
                    className="text-xs font-bold text-[#419CD8] bg-white hover:bg-blue-50 border border-blue-200 px-3 py-1.5 rounded-lg shadow-sm transition-all cursor-pointer"
                  >
                    只選 CALVIN WONG
                  </button>
                  <button
                    onClick={() => {
                      setSelectedTeams([]);
                      setReportData(null);
                    }}
                    type="button"
                    className="text-xs font-bold text-slate-500 bg-white hover:bg-slate-50 border border-slate-200 px-3 py-1.5 rounded-lg shadow-sm transition-all cursor-pointer"
                  >
                    清除選擇
                  </button>
                </div>

                {/* Scrollable list of checkboxes */}
                <div className="max-h-48 overflow-y-auto border border-slate-200 rounded-xl bg-white p-2.5 space-y-1.5 shadow-inner">
                  {availableTeams.map((team) => {
                    const isChecked = selectedTeams.includes(team);
                    return (
                      <label 
                        key={team} 
                        className={cn(
                          "flex items-center gap-3 px-3 py-1.5 rounded-lg cursor-pointer transition-colors text-xs font-bold select-none",
                          isChecked 
                            ? "bg-blue-50/70 text-slate-800 border border-blue-100" 
                            : "hover:bg-slate-50 text-slate-600 border border-transparent"
                        )}
                      >
                        <input
                          type="checkbox"
                          checked={isChecked}
                          onChange={() => {
                            if (isChecked) {
                              setSelectedTeams(selectedTeams.filter(t => t !== team));
                            } else {
                              setSelectedTeams([...selectedTeams, team]);
                            }
                            setReportData(null); // 當更換團隊時，清除舊的報告
                          }}
                          className="w-4 h-4 rounded text-[#419CD8] border-slate-300 focus:ring-[#419CD8] cursor-pointer"
                        />
                        <span className="flex-1 truncate">{team}</span>
                      </label>
                    );
                  })}
                </div>

                <p className="text-[11px] text-slate-500 mt-3 leading-relaxed">
                  系統已自動讀取 Headcount 檔案 B 行的 <b>District</b>。您可以剔選多個團隊以合併生成報表，或點擊「選擇所有團隊」一次生成。
                </p>
              </motion.div>
            )}

            <div className="mt-8 flex justify-center gap-4">
              <button
                onClick={processFiles}
                disabled={!headcountFile || !productionFile || isProcessing}
                className={cn(
                  "px-10 py-3 rounded-xl font-bold text-lg transition-all shadow-xl flex items-center gap-3",
                  headcountFile && productionFile 
                    ? "bg-[#419CD8] text-white hover:bg-[#3587bd] shadow-blue-100 scale-105" 
                    : "bg-slate-100 text-slate-400 cursor-not-allowed"
                )}
              >
                {isProcessing ? "正在處理..." : "生成報表結果"}
              </button>
              <button
                onClick={() => window.location.reload()}
                className="px-8 py-3 rounded-xl font-bold text-lg transition-all shadow-xl bg-slate-800 text-white hover:bg-slate-700 hover:scale-105 flex items-center gap-3"
              >
                重新上傳
              </button>
            </div>

            {error && (
              <motion.div 
                initial={{ opacity: 0 }} 
                animate={{ opacity: 1 }} 
                className="mt-6 p-4 bg-red-50 text-red-600 rounded-xl flex items-center gap-3 border border-red-100 text-sm font-medium"
              >
                <AlertCircle className="w-5 h-5 flex-shrink-0" />
                {error}
              </motion.div>
            )}
          </section>
          )}

          {/* Report Result Section */}
          {activeView === 'reports' && reportData && (
            <motion.div
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="space-y-6 mx-auto transition-all duration-300"
              style={{ maxWidth: `${totalTableWidth + 24}px`, width: '100%' }}
            >
              {isViewerMode ? (
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-white border border-slate-200 px-6 py-4 rounded-3xl shadow-sm">
                  <div className="flex items-center gap-2">
                    <Calendar className="w-5 h-5 text-[#419CD8]" />
                    <span className="text-sm font-extrabold text-slate-700">選擇報告日期 (Date)：</span>
                    <select
                      value={viewerSelectedDate}
                      onChange={(e) => handleViewerDateChange(e.target.value)}
                      className="bg-slate-50 hover:bg-slate-100 border border-slate-300 rounded-xl px-4 py-2 text-sm font-black text-slate-800 focus:outline-none focus:ring-2 focus:ring-[#419CD8] cursor-pointer"
                    >
                      {viewerAvailableDates.map(date => (
                        <option key={date} value={date}>{date}</option>
                      ))}
                    </select>
                  </div>
                  <div className="text-xs text-slate-400 font-medium italic">
                    唯讀分享版 • 共 {viewerAvailableDates.length} 個歷史日期
                  </div>
                </div>
              ) : (
                <div className="flex justify-between items-center bg-white border border-slate-200 px-6 py-4 rounded-3xl shadow-sm">
                  <button
                    onClick={() => setActiveView('dashboard')}
                    className="px-5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-sm rounded-xl transition-all flex items-center gap-2 border border-slate-200 hover:scale-[1.02] cursor-pointer"
                  >
                    ← 返回上傳數據頁面
                  </button>
                  <div className="text-xs text-slate-500 font-medium">
                    可隨時點擊返回修改數據或重新上傳
                  </div>
                </div>
              )}

              {/* Publish & Share Card (Only for admin) */}
              {!isViewerMode && (
                <div className="bg-white border border-slate-200 rounded-3xl p-6 shadow-sm space-y-4">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                    <div>
                      <h4 className="text-base font-extrabold text-slate-800 flex items-center gap-2">
                        <Share2 className="w-5 h-5 text-[#419CD8]" />
                        發佈並產生分享連結 (Publish & Share)
                      </h4>
                      <p className="text-xs text-slate-500 mt-1">
                        發佈後會建立一個專屬的唯讀儀表板連結，可以分享給其他人，且支援「選擇日期看」功能！
                      </p>
                    </div>
                    <button
                      onClick={publishCurrentDashboard}
                      disabled={isPublishing}
                      className={cn(
                        "px-5 py-2.5 rounded-xl font-bold text-xs text-white shadow-md transition-all active:scale-95 cursor-pointer flex items-center gap-2",
                        isPublishing ? "bg-slate-400 cursor-not-allowed" : "bg-gradient-to-r from-[#419CD8] to-blue-500 hover:from-[#3587bd] hover:to-blue-600"
                      )}
                    >
                      {isPublishing ? "正在發佈..." : shareId ? "🔄 更新發佈報表" : "🚀 首次發佈並產生連結"}
                    </button>
                  </div>

                  {publishSuccessMessage && (
                    <motion.div 
                      initial={{ opacity: 0, y: -10 }}
                      animate={{ opacity: 1, y: 0 }}
                      className="bg-emerald-50 border border-emerald-200 rounded-2xl p-4 flex flex-col sm:flex-row sm:items-center gap-3 justify-between"
                    >
                      <div className="space-y-1">
                        <span className="text-xs font-black text-emerald-800 flex items-center gap-1.5">
                          <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                          報表發佈成功！此連結已整合您的 Logo 數據，且支援歷史日期選擇。
                        </span>
                        <div className="text-[11px] text-slate-500 font-mono select-all truncate max-w-lg">
                          {publishSuccessMessage}
                        </div>
                      </div>
                      <button
                        onClick={() => copyToClipboard(publishSuccessMessage)}
                        className={cn(
                          "px-4 py-2 rounded-xl text-xs font-black transition-all flex items-center gap-1.5 cursor-pointer",
                          copied ? "bg-emerald-600 text-white" : "bg-white border border-emerald-300 text-emerald-700 hover:bg-emerald-50"
                        )}
                      >
                        {copied ? (
                          <>
                            <Check className="w-3.5 h-3.5" />
                            已複製
                          </>
                        ) : (
                          <>
                            <Copy className="w-3.5 h-3.5" />
                            複製連結
                          </>
                        )}
                      </button>
                    </motion.div>
                  )}
                </div>
              )}

              <div className={cn(theme.outerBg, "p-1 shadow-2xl rounded-sm overflow-hidden")}>
                <div className={cn(theme.outerBg, "relative border-[6px]", theme.outerBorder, "rounded-xl overflow-hidden")}>
                  {/* Header Section */}
                  <div className={cn("flex items-stretch h-32 border-b-2 border-black/10", theme.outerBg)}>
                    <div className={cn("flex-1 bg-white m-3 rounded-2xl border-4 shadow-inner flex flex-col items-center justify-center relative overflow-hidden", theme.outerBorder)}>
                      <div className="absolute top-0 left-0 w-8 h-8 rounded-full border-4 border-yellow-400 -m-3"></div>
                      <h3 className="text-3xl font-black text-black tracking-tighter italic">DAILY REPORT</h3>
                      <div className="bg-[#F27D26] text-white px-6 py-0.5 rounded-full text-[10px] font-bold mt-2 shadow-md border-white/30 border">
                        每日及時更新準時送達
                      </div>
                    </div>
                    
                    {/* Logo Section */}
                    <div className={cn("w-32 bg-white m-3 rounded-2xl border-4 shadow-inner flex items-center justify-center flex-shrink-0 relative group overflow-hidden", theme.outerBorder)}>
                      {logoUrl ? (
                        <div className="relative w-full h-full flex items-center justify-center p-2 bg-white">
                          <img 
                            src={logoUrl} 
                            alt="Uploaded Logo" 
                            className="max-w-full max-h-full object-contain" 
                            referrerPolicy="no-referrer" 
                          />
                          {!isViewerMode && (
                            <div className="absolute inset-0 bg-black/60 opacity-0 group-hover:opacity-100 transition-opacity flex flex-col items-center justify-center gap-1.5 z-10">
                              <label className={cn("cursor-pointer text-white text-[10px] font-bold px-2 py-1 rounded shadow-sm transition-colors text-center", isMultiTeam ? "bg-[#D97706] hover:bg-[#b45309]" : "bg-[#419CD8] hover:bg-[#3587bd]")}>
                                更換 LOGO
                                <input 
                                  type="file" 
                                  accept="image/*" 
                                  className="hidden" 
                                  onChange={handleLogoUpload} 
                                />
                              </label>
                              <button 
                                type="button"
                                onClick={(e) => { e.stopPropagation(); setLogoUrl(null); }}
                                className="bg-red-500 hover:bg-red-600 text-white text-[10px] font-bold px-2 py-1 rounded shadow-sm transition-colors text-center"
                              >
                                移除 LOGO
                              </button>
                            </div>
                          )}
                        </div>
                      ) : isViewerMode ? (
                        <div className="flex flex-col items-center justify-center text-center p-2">
                          <TrendingUp className={cn("w-6 h-6 mb-1", theme.bulletColor)} />
                          <span className={cn("text-[10px] font-black tracking-widest uppercase", theme.bulletColor)}>SALES REPORT</span>
                        </div>
                      ) : (
                        <label className="flex flex-col items-center justify-center w-full h-full cursor-pointer hover:bg-slate-50 transition-colors p-2 text-center select-none">
                          <FileUp className={cn("w-5 h-5 mb-1 group-hover:scale-110 transition-transform", theme.bulletColor)} />
                          <span className={cn("text-[10px] font-black uppercase tracking-wider", theme.bulletColor)}>上傳 LOGO</span>
                          <input 
                            type="file" 
                            accept="image/*" 
                            className="hidden" 
                            onChange={handleLogoUpload} 
                          />
                        </label>
                      )}
                    </div>
                  </div>

                  {/* Info Row */}
                  <div className="bg-white px-4 py-1.5 flex justify-between items-center text-[11px] border-b border-gray-200">
                    <span className="font-semibold text-slate-700">Source by Daily Submission Report</span>
                    <div className="flex gap-12">
                       <span className="font-bold text-slate-800 text-lg">as of</span>
                       <span className="font-black text-slate-900 text-lg">{reportDate}</span>
                    </div>
                  </div>

                  {/* Help Tip & Control Panel */}
                  <div className="bg-slate-50 p-5 border-b border-slate-200">
                    <div className="flex flex-col xl:flex-row xl:items-center justify-between gap-4">
                      <div className="flex items-center gap-2 text-sm text-slate-700 font-medium">
                        <span className="text-base">💡</span>
                        <span>
                          <b>自助排版面板</b>：您可以<b>按住表頭邊緣拖曳</b>（左右調寬、上下調高），或使用滑桿設定：
                        </span>
                      </div>
                      
                      {/* Sliders Container */}
                      <div className="flex flex-wrap items-center gap-5 text-xs font-semibold text-slate-600 bg-white px-4 py-2.5 rounded-xl shadow-sm border border-slate-100">
                        {/* Font Size Slider */}
                        <div className="flex items-center gap-1.5">
                          <span>🔤 字體大小:</span>
                          <input 
                            type="range" 
                            min="14" 
                            max="40" 
                            value={fontSize} 
                            onChange={(e) => setFontSize(parseInt(e.target.value))}
                            className={cn("w-20 cursor-pointer", theme.btnAccent)}
                          />
                          <span className="text-slate-400 font-mono w-7">{fontSize}px</span>
                        </div>

                        {/* Row Height Slider */}
                        <div className="flex items-center gap-1.5">
                          <span>↕️ 行高(間距):</span>
                          <input 
                            type="range" 
                            min="4" 
                            max="40" 
                            value={rowPadding} 
                            onChange={(e) => setRowPadding(parseInt(e.target.value))}
                            className={cn("w-20 cursor-pointer", theme.btnAccent)}
                          />
                          <span className="text-slate-400 font-mono w-7">{rowPadding}px</span>
                        </div>

                        {/* Reset Button */}
                        <button 
                          onClick={() => {
                            setColWidths({ manager: 240, name: 290, fycc: 190, case: 140 });
                            setRowPadding(16);
                            setFontSize(24);
                          }}
                          className="bg-slate-100 hover:bg-slate-200 text-slate-700 px-2.5 py-1 rounded text-[10px] font-bold transition-all active:scale-95"
                        >
                          重設樣式
                        </button>
                      </div>
                    </div>
                  </div>

                  {/* Table area */}
                  <div className="bg-white overflow-x-auto w-full no-scrollbar">
                    <table 
                      className="text-left border-collapse table-fixed select-none"
                      style={{ width: colWidths.manager + colWidths.name + colWidths.fycc + colWidths.case }}
                    >
                      <colgroup>
                        <col style={{ width: colWidths.manager }} />
                        <col style={{ width: colWidths.name }} />
                        <col style={{ width: colWidths.fycc }} />
                        <col style={{ width: colWidths.case }} />
                      </colgroup>
                      <thead>
                        <tr className={cn("text-white", theme.headerBg)}>
                          <th 
                            className={cn("relative px-3 border-r font-black select-none group", theme.innerBorder)}
                            style={{ paddingTop: `${rowPadding}px`, paddingBottom: `${rowPadding}px`, fontSize: `${fontSize}px` }}
                          >
                            <div className="truncate pr-3">Manager</div>
                            <div 
                              onMouseDown={(e) => handleResizeStart('manager', e)}
                              className={cn("absolute right-0 top-0 bottom-0 w-2.5 cursor-col-resize z-10 transition-colors", theme.hoverBg)}
                              title="左右拖動調整寬度"
                            />
                            <div 
                              onMouseDown={handleRowHeightResizeStart}
                              className={cn("absolute left-0 right-0 bottom-0 h-2 cursor-row-resize z-20 transition-colors", theme.hoverBg)}
                              title="上下拖動調整行高"
                            />
                          </th>
                          <th 
                            className={cn("relative px-3 border-r font-black select-none group", theme.innerBorder)}
                            style={{ paddingTop: `${rowPadding}px`, paddingBottom: `${rowPadding}px`, fontSize: `${fontSize}px` }}
                          >
                            <div className="truncate pr-3">Name (HKID)</div>
                            <div 
                              onMouseDown={(e) => handleResizeStart('name', e)}
                              className={cn("absolute right-0 top-0 bottom-0 w-2.5 cursor-col-resize z-10 transition-colors", theme.hoverBg)}
                              title="左右拖動調整寬度"
                            />
                            <div 
                              onMouseDown={handleRowHeightResizeStart}
                              className={cn("absolute left-0 right-0 bottom-0 h-2 cursor-row-resize z-20 transition-colors", theme.hoverBg)}
                              title="上下拖動調整行高"
                            />
                          </th>
                          <th 
                            className={cn("relative px-3 border-r text-center font-black select-none group", theme.innerBorder)}
                            style={{ paddingTop: `${rowPadding}px`, paddingBottom: `${rowPadding}px`, fontSize: `${fontSize}px` }}
                          >
                            <div className="truncate px-1">FYCC</div>
                            <div 
                              onMouseDown={(e) => handleResizeStart('fycc', e)}
                              className={cn("absolute right-0 top-0 bottom-0 w-2.5 cursor-col-resize z-10 transition-colors", theme.hoverBg)}
                              title="左右拖動調整寬度"
                            />
                            <div 
                              onMouseDown={handleRowHeightResizeStart}
                              className={cn("absolute left-0 right-0 bottom-0 h-2 cursor-row-resize z-20 transition-colors", theme.hoverBg)}
                              title="上下拖動調整行高"
                            />
                          </th>
                          <th 
                            className="relative px-3 font-black select-none group text-center"
                            style={{ paddingTop: `${rowPadding}px`, paddingBottom: `${rowPadding}px`, fontSize: `${fontSize}px` }}
                          >
                            <div className="truncate px-1">Case</div>
                            <div 
                              onMouseDown={(e) => handleResizeStart('case', e)}
                              className={cn("absolute right-0 top-0 bottom-0 w-2.5 cursor-col-resize z-10 transition-colors", theme.hoverBg)}
                              title="左右拖動調整寬度"
                            />
                            <div 
                              onMouseDown={handleRowHeightResizeStart}
                              className={cn("absolute left-0 right-0 bottom-0 h-2 cursor-row-resize z-20 transition-colors", theme.hoverBg)}
                              title="上下拖動調整行高"
                            />
                          </th>
                        </tr>
                      </thead>
                      <tbody className="font-medium" style={{ fontSize: `${fontSize}px` }}>
                        {reportData.map((distGroup) => (
                           <Fragment key={distGroup.district}>
                             {/* District Classification Row */}
                             <tr className={cn(theme.districtRowBg, theme.districtRowText, "border-b uppercase font-black", theme.districtRowBorder)}>
                               <td 
                                 colSpan={2} 
                                 className={cn("px-3 border-r truncate", theme.districtRowBorder)}
                                 style={{ paddingTop: `${rowPadding}px`, paddingBottom: `${rowPadding}px` }}
                               >
                                 <div className="flex items-center gap-1.5">
                                   <span className={theme.bulletColor}>■</span>
                                   <span className="truncate">{distGroup.district}</span>
                                 </div>
                               </td>
                               <td 
                                 className={cn("px-3 border-r text-right font-black truncate", theme.districtRowBorder, theme.textColor)}
                                 style={{ paddingTop: `${rowPadding}px`, paddingBottom: `${rowPadding}px` }}
                               >
                                 {distGroup.totalFYC.toLocaleString()}
                               </td>
                               <td 
                                 className={cn("px-3 text-right font-black truncate", theme.textColor)}
                                 style={{ paddingTop: `${rowPadding}px`, paddingBottom: `${rowPadding}px` }}
                               >
                                 {distGroup.totalCases.toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 })}
                               </td>
                             </tr>

                             {/* Manager Groups */}
                             {distGroup.managerGroups.map((group) => (
                               <Fragment key={group.manager}>
                                 {group.records.map((record, idx) => (
                                   <tr key={`${record.name}-${idx}`} className="border-b border-gray-100 hover:bg-sky-50 transition-colors uppercase whitespace-nowrap">
                                     <td 
                                       className="px-3 border-r border-gray-100 font-bold truncate" 
                                       title={group.manager}
                                       style={{ paddingTop: `${rowPadding * 0.8}px`, paddingBottom: `${rowPadding * 0.8}px` }}
                                     >
                                       {idx === 0 ? `- ${group.manager}` : ""}
                                     </td>
                                     <td 
                                       className="px-3 border-r border-gray-100 truncate" 
                                       title={record.name}
                                       style={{ paddingTop: `${rowPadding * 0.8}px`, paddingBottom: `${rowPadding * 0.8}px` }}
                                     >
                                       {record.name}
                                     </td>
                                     <td 
                                       className="px-3 border-r border-gray-100 text-right font-semibold truncate"
                                       style={{ paddingTop: `${rowPadding * 0.8}px`, paddingBottom: `${rowPadding * 0.8}px` }}
                                     >
                                       {record.fyc.toLocaleString()}
                                     </td>
                                     <td 
                                       className="px-3 text-right truncate"
                                       style={{ paddingTop: `${rowPadding * 0.8}px`, paddingBottom: `${rowPadding * 0.8}px` }}
                                     >
                                       {record.cases.toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 })}
                                     </td>
                                   </tr>
                                 ))}
                               </Fragment>
                             ))}
                           </Fragment>
                         ))}
                       </tbody>
                       <tfoot className="font-black bg-white border-t-2 border-gray-300" style={{ fontSize: `${fontSize}px` }}>
                          <tr className="bg-slate-50">
                            <td 
                              colSpan={2} 
                              className="px-3 border-r border-gray-100 text-right font-semibold text-slate-500"
                              style={{ paddingTop: `${rowPadding}px`, paddingBottom: `${rowPadding}px` }}
                            >
                              TOTAL
                            </td>
                            <td 
                              className={cn("px-3 border-r border-gray-100 text-right underline decoration-double underline-offset-4 truncate", theme.textColor)}
                              style={{ paddingTop: `${rowPadding}px`, paddingBottom: `${rowPadding}px` }}
                            >
                              {grandTotals.fyc.toLocaleString()}
                            </td>
                            <td 
                              className={cn("px-3 text-right underline decoration-double underline-offset-4 truncate", theme.textColor)}
                              style={{ paddingTop: `${rowPadding}px`, paddingBottom: `${rowPadding}px` }}
                            >
                              {grandTotals.cases.toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 })}
                            </td>
                          </tr>
                       </tfoot>
                    </table>
                  </div>
                </div>
              </div>
            </motion.div>
          )}
            </>
          )}
        </div>
      </main>
    </div>
  );
}
