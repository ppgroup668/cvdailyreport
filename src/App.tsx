/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useCallback, useMemo, Fragment, ChangeEvent, useEffect } from 'react';
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
  Users
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { cn } from './lib/utils';
import { 
  HeadcountRow, 
  ProductionRow, 
  SalesRecord, 
  ManagerGroup 
} from './types';

export default function App() {
  const [headcountFile, setHeadcountFile] = useState<File | null>(null);
  const [productionFile, setProductionFile] = useState<File | null>(null);
  const [reportData, setReportData] = useState<ManagerGroup[] | null>(null);
  const [reportDate, setReportDate] = useState<string>(new Date().toISOString().split('T')[0]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [activeView, setActiveView] = useState<'dashboard' | 'reports'>('dashboard');
  const [logoUrl, setLogoUrl] = useState<string | null>(null);
  const [availableTeams, setAvailableTeams] = useState<string[]>([]);
  const [selectedTeams, setSelectedTeams] = useState<string[]>([]);

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
      const url = URL.createObjectURL(file);
      setLogoUrl(url);
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
      const headcountMap = new Map<string, { manager: string; originalName: string }>();
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
          headcountMap.set(normalizeName(nameRaw), { manager: managerRaw, originalName: nameRaw });
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
       * O 行 (Index 14): FYC
       * P 行 (Index 15): Case
       */
      const productionMap = new Map<string, { fyc: number; cases: number; manager: string; originalName: string }>();
      productionRows.forEach((row, idx) => {
        if (idx === 0) return;
        const nameRaw = String(row[4] || '').trim();
        const fycRaw = String(row[14] || '0').replace(/,/g, '');
        const casesRaw = String(row[15] || '0').replace(/,/g, '');
        const fyc = Math.round(parseFloat(fycRaw) || 0); // O 行
        const cases = parseFloat(casesRaw) || 0; // P 行
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
      productionMap.forEach((data) => {
        if (data.manager && data.manager.trim() !== "" && data.manager !== "Unassigned") {
          mergedRecords.push({
            name: data.originalName,
            manager: data.manager,
            fyc: data.fyc,
            cases: data.cases
          });
        }
      });

      const groups: { [key: string]: SalesRecord[] } = {};
      mergedRecords.forEach(record => {
        if (!groups[record.manager]) groups[record.manager] = [];
        groups[record.manager].push(record);
      });

      const processedGroups: ManagerGroup[] = Object.entries(groups).map(([manager, records]) => ({
        manager,
        records: records.sort((a, b) => a.name.localeCompare(b.name)),
        totalFYC: records.reduce((sum, r) => sum + r.fyc, 0),
        totalCases: records.reduce((sum, r) => sum + r.cases, 0)
      })).sort((a, b) => b.totalFYC - a.totalFYC);

      setReportData(processedGroups);
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
    return reportData.reduce((acc, group) => ({
      fyc: acc.fyc + group.totalFYC,
      cases: acc.cases + group.totalCases
    }), { fyc: 0, cases: 0 });
  }, [reportData]);

  return (
    <div className="flex flex-col min-h-screen bg-[#F8FAFC] font-sans pb-20">
      {/* Header */}
      <header className="h-16 bg-white border-b border-slate-200 flex items-center justify-between px-6 flex-shrink-0 z-10 shadow-sm">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 bg-[#419CD8] rounded-lg flex items-center justify-center shadow-lg shadow-blue-100">
            <LayoutDashboard className="w-5 h-5 text-white" />
          </div>
          <h1 className="text-xl font-bold text-slate-800 tracking-tight">每日銷售報告系統</h1>
        </div>
        {reportData && (
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
          
          {/* Upload Section */}
          {activeView === 'dashboard' && (
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
              className="space-y-6 max-w-xl mx-auto"
            >
              <div className="flex justify-between items-center bg-white border border-slate-200 px-6 py-4 rounded-3xl shadow-sm">
                <button
                  onClick={() => setActiveView('dashboard')}
                  className="px-5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-sm rounded-xl transition-all flex items-center gap-2 border border-slate-200 hover:scale-[1.02]"
                >
                  ← 返回上傳數據頁面
                </button>
                <div className="text-xs text-slate-500 font-medium">
                  可隨時點擊返回修改數據或重新上傳
                </div>
              </div>

              <div className="bg-[#5AC8FA] p-1 shadow-2xl rounded-sm overflow-hidden">
                <div className="bg-[#5AC8FA] relative border-[6px] border-[#419CD8] rounded-xl overflow-hidden">
                  {/* Header Section */}
                  <div className="flex bg-[#5AC8FA] items-stretch h-32 border-b-2 border-black/10">
                    <div className="flex-1 bg-white m-3 rounded-2xl border-4 border-[#419CD8] shadow-inner flex flex-col items-center justify-center relative overflow-hidden">
                      <div className="absolute top-0 left-0 w-8 h-8 rounded-full border-4 border-yellow-400 -m-3"></div>
                      <h3 className="text-3xl font-black text-black tracking-tighter italic">DAILY REPORT</h3>
                      <div className="bg-[#F27D26] text-white px-6 py-0.5 rounded-full text-[10px] font-bold mt-2 shadow-md border-white/30 border">
                        每日及時更新準時送達
                      </div>
                    </div>
                    
                    {/* Logo Section */}
                    <div className="w-32 bg-white m-3 rounded-2xl border-4 border-[#419CD8] shadow-inner flex items-center justify-center flex-shrink-0 relative group overflow-hidden">
                      {logoUrl ? (
                        <div className="relative w-full h-full flex items-center justify-center p-2 bg-white">
                          <img 
                            src={logoUrl} 
                            alt="Uploaded Logo" 
                            className="max-w-full max-h-full object-contain" 
                            referrerPolicy="no-referrer" 
                          />
                          <div className="absolute inset-0 bg-black/60 opacity-0 group-hover:opacity-100 transition-opacity flex flex-col items-center justify-center gap-1.5 z-10">
                            <label className="cursor-pointer bg-[#419CD8] hover:bg-[#3587bd] text-white text-[10px] font-bold px-2 py-1 rounded shadow-sm transition-colors text-center">
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
                        </div>
                      ) : (
                        <label className="flex flex-col items-center justify-center w-full h-full cursor-pointer hover:bg-slate-50 transition-colors p-2 text-center select-none">
                          <FileUp className="w-5 h-5 text-[#419CD8] mb-1 group-hover:scale-110 transition-transform" />
                          <span className="text-[10px] font-black text-[#419CD8] uppercase tracking-wider">上傳 LOGO</span>
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

                  {/* Table area */}
                  <div className="bg-white overflow-x-auto">
                    <table className="w-full text-left border-collapse table-fixed">
                      <thead>
                        <tr className="bg-[#60A5FA] text-white text-base">
                          <th className="px-2.5 py-1 border-r border-[#419CD8] w-[140px] text-sm">Manager</th>
                          <th className="px-2.5 py-1 border-r border-[#419CD8] w-[180px] text-sm">Name (HKID)</th>
                          <th className="px-2.5 py-1 border-r border-[#419CD8] text-center w-[90px] text-sm">FYC</th>
                          <th className="px-2.5 py-1 text-center w-[60px] text-sm">Case</th>
                        </tr>
                      </thead>
                      <tbody className="text-xs font-medium">
                        {reportData.map((group) => (
                          <Fragment key={group.manager}>
                            {group.records.map((record, idx) => (
                              <tr key={`${record.name}-${idx}`} className="border-b border-gray-100 hover:bg-sky-50 transition-colors uppercase">
                                <td className="px-2.5 py-1 border-r border-gray-100 font-bold truncate w-[140px]" title={group.manager}>
                                  {idx === 0 ? `- ${group.manager}` : ""}
                                </td>
                                <td className="px-2.5 py-1 border-r border-gray-100 truncate w-[180px]" title={record.name}>
                                  {record.name}
                                </td>
                                <td className="px-2.5 py-1 border-r border-gray-100 text-right font-semibold w-[90px]">
                                  {record.fyc.toLocaleString()}
                                </td>
                                <td className="px-2.5 py-1 text-right w-[60px]">
                                  {record.cases}
                                </td>
                              </tr>
                            ))}
                          </Fragment>
                        ))}
                      </tbody>
                      <tfoot className="font-black text-base bg-white border-t border-gray-300">
                         <tr>
                           <td colSpan={2} className="px-2.5 py-1 border-r border-gray-100"></td>
                           <td className="px-2.5 py-1 border-r border-gray-100 text-right underline decoration-double underline-offset-4">
                             {grandTotals.fyc.toLocaleString()}
                           </td>
                           <td className="px-2.5 py-1 text-right underline decoration-double underline-offset-4">
                             {grandTotals.cases.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 1 })}
                           </td>
                         </tr>
                      </tfoot>
                    </table>
                  </div>
                </div>
              </div>
            </motion.div>
          )}
        </div>
      </main>
    </div>
  );
}
