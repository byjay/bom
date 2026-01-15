import React, { useState, useEffect, useMemo } from 'react';
import { createRoot } from 'react-dom/client';
import { Download, Upload, Filter, Calculator, Layers, FileSpreadsheet, Trash2, GripHorizontal, AlignLeft, PieChart, BarChart3, X, Loader2 } from 'lucide-react';
import * as XLSX from 'xlsx';

// --- Constants & Types ---

const COL_DEFINITIONS = [
  { key: 'FILENAME', label: 'FILENAME', type: 'text' },
  { key: 'BLOCK', label: 'BLOCK', type: 'text' },
  { key: 'MOD', label: 'MOD. NO', type: 'text' },
  { key: 'OFF', label: 'OFF', type: 'text' },
  { key: 'WLEG', label: 'WLEG', type: 'text' },
  { key: 'WNO', label: 'WNO', type: 'text' },
  { key: 'P. NO', label: 'P. NO', type: 'text' },
  { key: 'WELD. LENG.', label: 'WELD. LENG.', type: 'number' },
  { key: 'DETAIL VIEW', label: 'DETAIL VIEW', type: 'text' },
  { key: 'SIDE', label: 'SIDE', type: 'text' },
  { key: 'WELD UNIQUE ID', label: 'WELD UNIQUE ID', type: 'id' },
  { key: 'MATCH_MATNO', label: 'MATNO', type: 'text' }, // The matched MatNo
  { key: 'STEEL NO', label: 'STEEL NO', type: 'text' },
  { key: 'NESTING DWG', label: 'NESTING DWG', type: 'text' },
  { key: 'ea', label: 'ea', type: 'number' },
  { key: 'total', label: 'total', type: 'number' },
  { key: 'MIX', label: 'MIX', type: 'text' },
  { key: 'Grade', label: 'Grade', type: 'text' },
  { key: 'T', label: 'T', type: 'number' },
  { key: 'B', label: 'B', type: 'number' },
  { key: 'L(OD)', label: 'L(OD)', type: 'number' },
  { key: 'WEIGHT', label: 'WEIGHT', type: 'number' },
  { key: 'TPYE', label: 'TYPE', type: 'text' }
];

const GROUP_OPTIONS = [
  { id: 'FILENAME', label: 'FILENAME' },
  { id: 'WELD UNIQUE ID', label: 'WELD UNIQUE ID' },
  { id: 'MATCH_MATNO', label: 'MAT NO' },
  { id: 'STEEL NO', label: 'STEEL NO' },
  { id: 'NESTING DWG', label: 'NESTING DWG' },
  { id: 'Grade', label: 'Grade' },
];

// --- Default Data from Real Excel Files (bom (1).xlsx + bom (2).xlsx) ---
const DEFAULT_WELD_DATA = [
  { "FILENAME": "bom (2).xlsx", "행 레이블": "A1", "WELD UNIQUE ID": "A1_SA01_001_01", "MATNO1": "A1-SA01-001", "": "A1-SA01-001" },
  { "FILENAME": "bom (2).xlsx", "행 레이블": "", "WELD UNIQUE ID": "A1_SA01_001_01", "MATNO1": "A1-SA01-010", "": "A1-SA01-010" },
  { "FILENAME": "bom (2).xlsx", "행 레이블": "", "WELD UNIQUE ID": "A1_SA01_001_01", "MATNO1": "A1-SA01-011", "": "A1-SA01-011" },
  { "FILENAME": "bom (2).xlsx", "행 레이블": "", "WELD UNIQUE ID": "A1_SA01_001_01", "MATNO1": "A1-SA01-012", "": "A1-SA01-012" },
  { "FILENAME": "bom (2).xlsx", "행 레이블": "", "WELD UNIQUE ID": "A1_SA01_001_01", "MATNO1": "A1-SA01-013", "": "A1-SA01-013" },
  { "FILENAME": "bom (2).xlsx", "행 레이블": "", "WELD UNIQUE ID": "A1_SA01_002_02", "MATNO1": "A1-SA01-002", "": "A1-SA01-002" },
  { "FILENAME": "bom (2).xlsx", "행 레이블": "", "WELD UNIQUE ID": "A1_SA01_002_02", "MATNO1": "A1-SA01-010", "": "A1-SA01-010" },
  { "FILENAME": "bom (2).xlsx", "행 레이블": "", "WELD UNIQUE ID": "A1_SA01_002_02", "MATNO1": "A1-SA01-011", "": "A1-SA01-011" },
  { "FILENAME": "bom (2).xlsx", "행 레이블": "", "WELD UNIQUE ID": "A1_SA01_002_02", "MATNO1": "A1-SA01-012", "": "A1-SA01-012" },
  { "FILENAME": "bom (2).xlsx", "행 레이블": "", "WELD UNIQUE ID": "A1_SA01_002_02", "MATNO1": "A1-SA01-013", "": "A1-SA01-013" }
];
const DEFAULT_MAT_DATA = [
  { "FILENAME": "bom (1).xlsx", "BLOCK": "A1", "MOD": "SA01", "MATNO": "A1-SA01-001", "STEEL NO": "FC89439801", "NESTING DWG": "BA21-A1A2CNX72", "ea": "1", "total": "1", "MIX": "단독", "no": "", "Grade": "S420MLO", "T": "80", "B": "1120", "L(OD)": "1920", "WEIGHT": "1111.1", "TPYE": "PLATE" },
  { "FILENAME": "bom (1).xlsx", "BLOCK": "A1", "MOD": "SA01", "MATNO": "A1-SA01-002", "STEEL NO": "FC89026801", "NESTING DWG": "BA21-A1A2CNX23", "ea": "2", "total": "2", "MIX": "단독", "no": "a b", "Grade": "S420M+OPT30", "T": "50", "B": "1120", "L(OD)": "1920", "WEIGHT": "1389", "TPYE": "PLATE" },
  { "FILENAME": "bom (1).xlsx", "BLOCK": "A1", "MOD": "SA01", "MATNO": "A1-SA01-003", "STEEL NO": "FC89037801", "NESTING DWG": "BA21-A1A2CNX25", "ea": "2", "total": "2", "MIX": "단독", "no": "a b", "Grade": "S420M+OPT30", "T": "50", "B": "1080", "L(OD)": "1880", "WEIGHT": "1294.8", "TPYE": "PLATE" },
  { "FILENAME": "bom (1).xlsx", "BLOCK": "A1", "MOD": "SA01", "MATNO": "A1-SA01-004", "STEEL NO": "FD35919801", "NESTING DWG": "BA21-A1A2CNX02", "ea": "3", "total": "3", "MIX": "단독", "no": "a b c", "Grade": "S420M", "T": "20", "B": "375", "L(OD)": "1917", "WEIGHT": "338.4", "TPYE": "PLATE" },
  { "FILENAME": "bom (1).xlsx", "BLOCK": "A1", "MOD": "SA01", "MATNO": "A1-SA01-005", "STEEL NO": "FD35919901", "NESTING DWG": "BA21-A1A2CNX03", "ea": "3", "total": "3", "MIX": "단독", "no": "a b c", "Grade": "S420M", "T": "20", "B": "1120", "L(OD)": "1920", "WEIGHT": "833.4", "TPYE": "PLATE" },
  { "FILENAME": "bom (1).xlsx", "BLOCK": "A1", "MOD": "SA01", "MATNO": "A1-SA01-006", "STEEL NO": "FD35919801", "NESTING DWG": "BA21-A1A2CNX02", "ea": "3", "total": "3", "MIX": "단독", "no": "a b c", "Grade": "S420M", "T": "20", "B": "375", "L(OD)": "1917", "WEIGHT": "338.4", "TPYE": "PLATE" },
  { "FILENAME": "bom (1).xlsx", "BLOCK": "A1", "MOD": "SA01", "MATNO": "A1-SA01-007", "STEEL NO": "FD35919801", "NESTING DWG": "BA21-A1A2CNX02", "ea": "3", "total": "3", "MIX": "단독", "no": "a b c", "Grade": "S420M", "T": "20", "B": "375", "L(OD)": "1877", "WEIGHT": "330.9", "TPYE": "PLATE" },
  { "FILENAME": "bom (1).xlsx", "BLOCK": "A1", "MOD": "SA01", "MATNO": "A1-SA01-008", "STEEL NO": "FD35919901", "NESTING DWG": "BA21-A1A2CNX03", "ea": "3", "total": "3", "MIX": "단독", "no": "a b c", "Grade": "S420M", "T": "20", "B": "1080", "L(OD)": "1880", "WEIGHT": "776.7", "TPYE": "PLATE" },
  { "FILENAME": "bom (1).xlsx", "BLOCK": "A1", "MOD": "SA01", "MATNO": "A1-SA01-009", "STEEL NO": "FD35919801", "NESTING DWG": "BA21-A1A2CNX02", "ea": "3", "total": "3", "MIX": "단독", "no": "a b c", "Grade": "S420M", "T": "20", "B": "375", "L(OD)": "1877", "WEIGHT": "330.9", "TPYE": "PLATE" },
  { "FILENAME": "bom (1).xlsx", "BLOCK": "A1", "MOD": "SA01", "MATNO": "A1-SA01-010", "STEEL NO": "FC89119501", "NESTING DWG": "BA21-A1A2CNX15", "ea": "1", "total": "1", "MIX": "단독", "no": "", "Grade": "S420M+Z35+OPT30", "T": "40", "B": "2000", "L(OD)": "14999.5", "WEIGHT": "9419.7", "TPYE": "PLATE" }
];

const App = () => {
  // --- State ---
  const [weldRaw, setWeldRaw] = useState<any[]>(DEFAULT_WELD_DATA);
  const [matRaw, setMatRaw] = useState<any[]>(DEFAULT_MAT_DATA);
  const [matchedData, setMatchedData] = useState<any[]>([]);
  const [showAnalysis, setShowAnalysis] = useState(true);
  const [isLoading, setIsLoading] = useState(false);
  
  // View Controls
  const [groupBy, setGroupBy] = useState<string>('WELD UNIQUE ID');
  const [filters, setFilters] = useState({ block: '', id: '', mat: '', grade: '' });

  // --- 1. Robust Excel Parsing (Auto-Detect, Multiple Files) ---
  const processExcelFiles = async (files: FileList | File[]) => {
    setIsLoading(true);
    // Use setTimeout to allow UI to update with Loading state
    setTimeout(async () => {
      const newWeldData: any[] = [];
      const newMatData: any[] = [];
      let processedCount = 0;

      for (let fIndex = 0; fIndex < files.length; fIndex++) {
          const file = files[fIndex];
          try {
              const buffer = await file.arrayBuffer();
              // ENHANCED: 테이블 형식 지원
              const wb = XLSX.read(buffer, {
                  cellStyles: false, // 스타일 무시로 속도 향상
                  cellDates: true,   // 날짜 자동 변환
                  sheetStubs: true   // 빈 셀도 인식
              });
              const sheet = wb.Sheets[wb.SheetNames[0]];

              if (!sheet) {
                  console.warn(`Sheet empty in ${file.name}`);
                  continue;
              }

              // A. Handle Merges - ENHANCED (빠른 병합셀 처리)
              if (sheet['!merges']) {
                  sheet['!merges'].forEach(range => {
                      const startCell = sheet[XLSX.utils.encode_cell(range.s)];
                      const value = startCell?.v ?? startCell?.w ?? '';
                      if (!value) return; // 빈 병합셀 건너뛰기

                      for (let R = range.s.r; R <= range.e.r; ++R) {
                          for (let C = range.s.c; C <= range.e.c; ++C) {
                              const addr = XLSX.utils.encode_cell({r: R, c: C});
                              if (!sheet[addr]) {
                                  sheet[addr] = { t: 's', v: value, w: String(value) };
                              } else if (!sheet[addr].v) {
                                  sheet[addr].v = value;
                              }
                          }
                      }
                  });
              }

              // B. Find Header Row - ENHANCED (빈 행 건너뛰기, 점수 시스템 강화)
              const rawJson = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false }) as any[][];
              if (rawJson.length === 0) continue;

              let headerRowIdx = 0;
              let maxScore = -1;
              const WELD_KEYS = ['WELD', 'UNIQUE', 'BLOCK', 'MATNO', 'WLEG', 'LENGTH', 'ID', 'WNO'];
              const MAT_KEYS = ['STEEL', 'NESTING', 'WEIGHT', 'GRADE', 'EA', 'QTY', 'MAT', 'TYPE', 'TPYE'];

              for(let i=0; i < Math.min(rawJson.length, 30); i++) {
                  const row = rawJson[i];

                  // 빈 행 건너뛰기
                  const nonEmpty = row.filter(c => c && String(c).trim().length > 0);
                  if (nonEmpty.length === 0) continue;

                  const rowStr = row.map(c => String(c || '').toUpperCase()).join(' ');

                  // 키워드 점수 (가중치 적용)
                  let score = 0;
                  WELD_KEYS.forEach(k => { if (rowStr.includes(k)) score += 3; });
                  MAT_KEYS.forEach(k => { if (rowStr.includes(k)) score += 3; });

                  // 비어있지 않은 셀 보너스
                  score += nonEmpty.length * 0.3;

                  // 문자열 비율 (숫자보다 문자열이 많으면 헤더)
                  const stringCount = nonEmpty.filter(c => typeof c === 'string' && isNaN(Number(c))).length;
                  if (stringCount / nonEmpty.length > 0.6) score += 5;

                  // 너무 긴 값이 있으면 감점
                  if (row.some(c => String(c || '').length > 100)) score -= 20;

                  if (score > maxScore) {
                      maxScore = score;
                      headerRowIdx = i;
                  }
              }

              // C. Extract Data & Normalize Headers
              const headers = rawJson[headerRowIdx].map(h => String(h).trim());
              // Helper to fuzzy match column name
              const findKey = (candidates: string[]) => headers.find(h => candidates.some(c => h.toUpperCase().includes(c)));

              const dataRows = [];
              for(let i = headerRowIdx + 1; i < rawJson.length; i++) {
                  const row = rawJson[i];
                  // ENHANCED: 빈 행 건너뛰기 (더 엄격하게)
                  const hasData = row.some(c => c !== null && c !== undefined && String(c).trim() !== '');
                  if (!hasData) continue;

                  const obj: any = { FILENAME: file.name };
                  headers.forEach((h, idx) => {
                      const val = row[idx];
                      // 빈 값 정규화
                      obj[h] = (val === null || val === undefined || val === '') ? '' : val;
                  });
                  dataRows.push(obj);
              }

              if (dataRows.length === 0) continue;

              // D. Fill Down Logic
              // Identify critical ID column for this specific file
              const idColName = findKey(['WELD UNIQUE ID', 'WELD ID', 'UNIQUE ID']) || 'WELD UNIQUE ID';
              const matColName = findKey(['MATNO', 'MAT NO']) || 'MATNO';
              
              const FILL_KEYS = ['BLOCK', 'MOD', 'DWG. Title', 'WELD. LENG.', 'SIDE', idColName];
              
              let lastVals: any = {};
              const filledData = dataRows.map(row => {
                  const currentId = row[idColName];
                  // Is this a material line? (check for mat column existence and value)
                  const hasMat = row[matColName]; 

                  if (currentId) {
                      FILL_KEYS.forEach(k => { if(row[k]) lastVals[k] = row[k]; });
                      // Ensure normalized ID key exists
                      row['WELD UNIQUE ID'] = currentId;
                      return row;
                  } else if (hasMat && lastVals[idColName]) {
                      const newRow = { ...row };
                      FILL_KEYS.forEach(k => {
                          if(!newRow[k] && lastVals[k]) newRow[k] = lastVals[k];
                      });
                      // Ensure normalized ID key exists
                      newRow['WELD UNIQUE ID'] = lastVals[idColName];
                      return newRow;
                  }
                  // Fallback: if we haven't seen an ID yet, just return row
                  return row;
              });

              // E. Auto-Detect Type
              const headerStr = headers.join(' ').toUpperCase();
              const isMaterial = (headerStr.includes('STEEL') && headerStr.includes('NESTING')) || headerStr.includes('WEIGHT') || headerStr.includes('GRADE');
              const isWeld = headerStr.includes('WELD') || headerStr.includes('BLOCK') || headerStr.includes('MATNO');

              if (isMaterial) {
                  newMatData.push(...filledData);
              } else if (isWeld) {
                  newWeldData.push(...filledData);
              } else {
                  // Fallback: If ambiguous, put in Weld if it looks like a BOM
                  if (headerStr.includes('NO') || headerStr.includes('ID')) newWeldData.push(...filledData);
              }
              processedCount++;

          } catch (err) {
              console.error(`Error parsing ${file.name}:`, err);
          }
      }

      // Update State (Accumulate)
      if (newWeldData.length > 0) {
          setWeldRaw(prev => [...prev, ...newWeldData]);
      }
      if (newMatData.length > 0) {
          setMatRaw(prev => [...prev, ...newMatData]);
      }
      
      setIsLoading(false);

      if (processedCount === 0) {
        alert("Failed to load files. Please check if the Excel files are valid and contain headers like 'WELD UNIQUE ID', 'STEEL NO', or 'BLOCK'.");
      } else if (newWeldData.length + newMatData.length === 0) {
        alert("Files loaded but no data rows were found. Please check header detection.");
      } else {
        // Success
      }
    }, 100);
  };

  // --- 2. Matching Logic (Wide-to-Tall + Join) ---
  useEffect(() => {
    if (weldRaw.length === 0) return;

    // Index Material
    const matMap = new Map();
    // Try to find the best MATNO key in material data
    const sampleMat = matRaw[0] || {};
    const matKey = Object.keys(sampleMat).find(k => k.toUpperCase().includes('MATNO') || k.toUpperCase().includes('MAT NO')) || 'MATNO';
    
    matRaw.forEach(row => {
      const rawKey = row[matKey];
      if (rawKey) {
          const key = String(rawKey).replace(/\s+/g,'').toUpperCase();
          matMap.set(key, row);
      }
    });

    // Process Welds
    const results: any[] = [];
    weldRaw.forEach((wRow, idx) => {
       // Find all MATNO columns (MATNO1, MATNO2...)
       const keys = Object.keys(wRow).filter(k => /^MAT.*NO/i.test(k));
       // If none, check plain 'MATNO'
       if(keys.length === 0) {
         const plain = Object.keys(wRow).find(k => k.toUpperCase() === 'MATNO');
         if (plain) keys.push(plain);
       }

       let foundAny = false;
       keys.forEach(k => {
          const val = wRow[k];
          if(val) {
             foundAny = true;
             const normalized = String(val).replace(/\s+/g,'').toUpperCase();
             const matInfo = matMap.get(normalized) || {};
             
             // Merge row data. Prefer Weld data, then Material data.
             // Ensure STEEL NO and NESTING DWG come from MatInfo if available.
             // Normalize keys for the UI
             results.push({
               ...wRow,
               ...matInfo,
               // Explicitly ensure we use the matched MATNO for display logic
               'MATCH_MATNO': val, 
               'STEEL NO': matInfo['STEEL NO'] || matInfo['STEEL_NO'] || '',
               'NESTING DWG': matInfo['NESTING DWG'] || matInfo['NESTING'] || '',
               'Grade': matInfo['Grade'] || matInfo['GRADE'] || '',
               'WEIGHT': matInfo['WEIGHT'] || matInfo['Weight'] || 0,
               _originIdx: idx
             });
          }
       });

       // Keep orphan welds (rows with no material match or no material number)
       if(!foundAny) {
         results.push({ ...wRow, 'MATCH_MATNO': '', _originIdx: idx });
       }
    });

    setMatchedData(results);
  }, [weldRaw, matRaw]);

  // --- 3. View Logic (Filtering, Sorting, Aggregation) ---
  const { viewData, stats, groupStats } = useMemo(() => {
    let data = [...matchedData];

    // Filter
    if (filters.block) data = data.filter(r => String(r.BLOCK||'').toLowerCase().includes(filters.block.toLowerCase()));
    if (filters.id) data = data.filter(r => String(r['WELD UNIQUE ID']||'').toLowerCase().includes(filters.id.toLowerCase()));
    if (filters.mat) data = data.filter(r => String(r['MATCH_MATNO']||'').toLowerCase().includes(filters.mat.toLowerCase()));
    if (filters.grade) data = data.filter(r => String(r.Grade||'').toLowerCase().includes(filters.grade.toLowerCase()));

    // Global Stats
    const stats = data.reduce((acc, r) => ({
      count: acc.count + 1,
      weight: acc.weight + (parseFloat(r.WEIGHT) || 0),
      length: acc.length + (parseFloat(r['WELD. LENG.']) || 0),
      ea: acc.ea + (parseFloat(r.ea) || 0)
    }), { count: 0, weight: 0, length: 0, ea: 0 });

    // Group Distribution Analysis
    const groupDist = data.reduce((acc, row) => {
        const key = row[groupBy] || '(Empty)';
        if (!acc[key]) {
            acc[key] = { name: key, count: 0, weight: 0, length: 0 };
        }
        acc[key].count += 1;
        acc[key].weight += (parseFloat(row.WEIGHT) || 0);
        acc[key].length += (parseFloat(row['WELD. LENG.']) || 0);
        return acc;
    }, {} as Record<string, any>);

    // Sort groups by Weight Descending
    const groupStatsArray = Object.values(groupDist).sort((a: any, b: any) => b.weight - a.weight);

    // Sort Data by Group Key for table
    data.sort((a, b) => {
       const valA = String(a[groupBy] || '');
       const valB = String(b[groupBy] || '');
       return valA.localeCompare(valB, undefined, { numeric: true });
    });

    return { viewData: data, stats, groupStats: groupStatsArray };
  }, [matchedData, filters, groupBy]);

  // --- 4. Render Helpers ---
  const exportFile = () => {
    const ws = XLSX.utils.json_to_sheet(viewData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Matched_Data");
    XLSX.writeFile(wb, "Weld_Mat_Report.xlsx");
  };

  const clearAll = () => { setWeldRaw([]); setMatRaw([]); setMatchedData([]); };

  // Organize Columns: Group Key First
  const orderedCols = useMemo(() => {
    const others = COL_DEFINITIONS.filter(c => c.key !== groupBy);
    const primary = COL_DEFINITIONS.find(c => c.key === groupBy) || { key: groupBy, label: groupBy, type: 'text' };
    return [primary, ...others];
  }, [groupBy]);

  return (
    <div className="h-screen flex flex-col bg-slate-100 font-sans text-slate-800 overflow-hidden relative">
      
      {/* Loading Overlay */}
      {isLoading && (
        <div className="absolute inset-0 bg-white/80 z-50 flex flex-col items-center justify-center backdrop-blur-sm">
           <Loader2 size={48} className="text-emerald-600 animate-spin mb-4" />
           <p className="text-lg font-semibold text-slate-700">Processing Excel Files...</p>
           <p className="text-sm text-slate-500">Parsing headers and matching data</p>
        </div>
      )}

      {/* 1. Top Bar: Controls */}
      <div className="bg-white border-b border-slate-200 shadow-sm z-30 flex-none">
        <div className="px-4 py-3 flex items-center justify-between">
           {/* Left: Branding & Upload */}
           <div className="flex items-center gap-6">
              <div className="flex items-center gap-2 text-emerald-700">
                 <Layers size={24} />
                 <div>
                    <h1 className="font-bold text-lg leading-tight">Weld Matcher Pro</h1>
                    <div className="text-[10px] text-slate-400 font-medium">MULTI-FILE PARSER</div>
                 </div>
              </div>

              <div className="h-8 w-px bg-slate-200"></div>

              <label className="flex items-center gap-2 cursor-pointer bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg shadow-sm transition-all active:scale-95">
                 <Upload size={16} />
                 <span className="text-sm font-semibold">Load Excel Files</span>
                 <input 
                    type="file" 
                    multiple 
                    accept=".xlsx,.xls" 
                    className="hidden" 
                    onChange={(e) => {
                        if(e.target.files && e.target.files.length > 0) {
                            processExcelFiles(e.target.files);
                        }
                        e.target.value = '';
                    }} 
                 />
              </label>

              {(weldRaw.length > 0 || matRaw.length > 0) && (
                 <button onClick={clearAll} className="text-slate-400 hover:text-red-500 transition-colors p-2">
                    <Trash2 size={18} />
                 </button>
              )}
           </div>

           {/* Right: Export & Layout Toggle */}
           <div className="flex items-center gap-2">
               <button 
                  onClick={() => setShowAnalysis(!showAnalysis)} 
                  className={`flex items-center gap-2 px-3 py-2 rounded-lg text-sm font-semibold transition-colors border
                     ${showAnalysis ? 'bg-indigo-50 text-indigo-700 border-indigo-200' : 'text-slate-500 border-transparent hover:bg-slate-50'}`}
               >
                  <BarChart3 size={18} /> Stats Panel
               </button>
               <button onClick={exportFile} disabled={viewData.length===0} className="flex items-center gap-2 text-emerald-600 hover:bg-emerald-50 px-3 py-2 rounded-lg text-sm font-semibold disabled:opacity-50 transition-colors border border-transparent hover:border-emerald-200">
                  <FileSpreadsheet size={18} /> Export Report
               </button>
           </div>
        </div>

        {/* Grouping Toolbar */}
        <div className="px-4 py-2 bg-slate-50 border-t border-slate-200 flex items-center gap-4 overflow-x-auto">
           <div className="flex items-center gap-2 text-slate-500 text-xs font-bold uppercase tracking-wider shrink-0">
              <GripHorizontal size={14} /> Group By:
           </div>
           <div className="flex gap-2">
              {GROUP_OPTIONS.map(opt => (
                <button
                  key={opt.id}
                  onClick={() => setGroupBy(opt.id)}
                  className={`px-3 py-1.5 rounded-full text-xs font-semibold border transition-all whitespace-nowrap
                    ${groupBy === opt.id 
                       ? 'bg-emerald-600 text-white border-emerald-600 shadow-md transform scale-105' 
                       : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-100'}`}
                >
                  {opt.label}
                </button>
              ))}
           </div>
        </div>
      </div>

      {/* 2. Stats Dashboard (Sticky Top) */}
      <div className="bg-emerald-900 text-white px-6 py-3 grid grid-cols-4 gap-4 shadow-inner flex-none z-20">
          <StatItem label="Total Rows" value={stats.count} icon={<AlignLeft size={16}/>} />
          <StatItem label="Total Weight (kg)" value={stats.weight.toFixed(1)} icon={<Calculator size={16}/>} />
          <StatItem label="Total Weld Length" value={stats.length.toLocaleString()} unit="mm" icon={<Calculator size={16}/>} />
          <StatItem label="Total Quantity" value={stats.ea} unit="EA" icon={<Calculator size={16}/>} />
      </div>

      {/* 3. Filter Bar */}
      <div className="bg-white px-4 py-2 flex items-center gap-4 border-b border-slate-200 flex-none z-20">
         <Filter size={16} className="text-slate-400" />
         <FilterInput placeholder="Block..." value={filters.block} onChange={v => setFilters({...filters, block: v})} />
         <FilterInput placeholder="Weld ID..." value={filters.id} onChange={v => setFilters({...filters, id: v})} />
         <FilterInput placeholder="Mat No..." value={filters.mat} onChange={v => setFilters({...filters, mat: v})} />
         <FilterInput placeholder="Grade..." value={filters.grade} onChange={v => setFilters({...filters, grade: v})} />
      </div>

      {/* Main Content Area (Split View) */}
      <div className="flex-1 flex overflow-hidden">
        
        {/* LEFT: Main Table */}
        <div className="flex-1 overflow-auto bg-slate-100 p-4">
            <div className="bg-white rounded-lg shadow-sm border border-slate-200 overflow-hidden min-h-[300px]">
            {viewData.length > 0 ? (
                <div className="overflow-x-auto">
                    <table className="w-full text-xs text-left border-collapse">
                    <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm">
                        <tr>
                            <th className="px-3 py-2 border-b border-r border-slate-200 w-10 text-center font-bold text-slate-400">#</th>
                            {orderedCols.map((col, idx) => (
                            <th key={col.key} className={`px-3 py-2 border-b border-r border-slate-200 font-bold whitespace-nowrap
                                ${idx === 0 ? 'bg-emerald-50 text-emerald-800 border-r-emerald-200' : 'text-slate-600'}
                                ${['STEEL NO','NESTING DWG','Grade','WEIGHT'].includes(col.key) ? 'bg-blue-50/50' : ''}
                            `}>
                                {col.label}
                            </th>
                            ))}
                        </tr>
                    </thead>
                    <tbody>
                        {renderTableBody(viewData, orderedCols, groupBy)}
                    </tbody>
                    </table>
                </div>
            ) : (
                <div className="flex flex-col items-center justify-center h-64 text-slate-400">
                    <FileSpreadsheet size={48} className="mb-2 opacity-20" />
                    <p>Load Excel files to see data</p>
                </div>
            )}
            </div>
        </div>

        {/* RIGHT: Analysis Sidebar */}
        {showAnalysis && (
            <div className="w-96 bg-white border-l border-slate-200 flex flex-col shadow-xl z-20 transition-all">
                <div className="p-4 border-b border-slate-100 bg-slate-50 flex items-center justify-between">
                    <div>
                        <h3 className="text-sm font-bold text-slate-700 flex items-center gap-2">
                             <PieChart size={16} className="text-indigo-500" />
                             Analysis by {groupBy}
                        </h3>
                        <p className="text-[10px] text-slate-400 mt-0.5">Distribution Analysis</p>
                    </div>
                    <button onClick={() => setShowAnalysis(false)} className="text-slate-400 hover:text-red-500"><X size={16}/></button>
                </div>
                
                <div className="flex-1 overflow-y-auto p-4 space-y-6">
                    {groupStats.length === 0 ? (
                        <div className="text-center text-slate-400 py-10 text-xs">No data to analyze</div>
                    ) : (
                        groupStats.map((grp: any, idx: number) => {
                            const weightPercent = (grp.weight / (stats.weight || 1)) * 100;
                            const countPercent = (grp.count / (stats.count || 1)) * 100;
                            
                            return (
                                <div key={idx} className="group">
                                    <div className="flex justify-between items-baseline mb-1">
                                        <span className="text-xs font-bold text-slate-700 truncate w-40" title={grp.name}>{grp.name}</span>
                                        <span className="text-[10px] font-mono text-slate-500">{weightPercent.toFixed(1)}% Weight</span>
                                    </div>
                                    
                                    {/* Bar Graph Container */}
                                    <div className="w-full h-1.5 bg-slate-100 rounded-full mb-2 overflow-hidden">
                                        <div 
                                            className="h-full bg-gradient-to-r from-indigo-500 to-purple-500 rounded-full" 
                                            style={{ width: `${Math.max(weightPercent, 1)}%` }}
                                        ></div>
                                    </div>

                                    {/* Metrics Grid */}
                                    <div className="grid grid-cols-3 gap-2 mt-2">
                                        <div className="bg-slate-50 p-1.5 rounded border border-slate-100 text-center">
                                            <div className="text-[9px] text-slate-400 uppercase">Count</div>
                                            <div className="text-xs font-semibold text-slate-700">{grp.count}</div>
                                        </div>
                                        <div className="bg-slate-50 p-1.5 rounded border border-slate-100 text-center">
                                            <div className="text-[9px] text-slate-400 uppercase">Weight</div>
                                            <div className="text-xs font-semibold text-slate-700">{grp.weight.toFixed(0)}</div>
                                        </div>
                                        <div className="bg-slate-50 p-1.5 rounded border border-slate-100 text-center">
                                            <div className="text-[9px] text-slate-400 uppercase">Length</div>
                                            <div className="text-xs font-semibold text-slate-700">{grp.length.toFixed(0)}</div>
                                        </div>
                                    </div>
                                    
                                    <div className="h-px bg-slate-100 mt-4 group-last:hidden"></div>
                                </div>
                            );
                        })
                    )}
                </div>
                
                <div className="p-3 bg-slate-50 border-t border-slate-200 text-[10px] text-center text-slate-400">
                    Showing top {groupStats.length} groups
                </div>
            </div>
        )}

      </div>
    </div>
  );
};

// --- Helper Components ---

const StatItem = ({ label, value, unit, icon }: any) => (
  <div className="flex items-center gap-3 bg-white/10 rounded-lg px-4 py-2 backdrop-blur-sm border border-white/10 min-w-[140px]">
     <div className="p-1.5 bg-emerald-500/20 rounded text-emerald-200">{icon}</div>
     <div>
        <div className="text-[10px] text-emerald-200 font-bold uppercase tracking-wider">{label}</div>
        <div className="text-lg font-mono font-semibold truncate">{value} <span className="text-xs font-sans text-emerald-300">{unit}</span></div>
     </div>
  </div>
);

const FilterInput = ({ placeholder, value, onChange }: any) => (
  <input 
    type="text" 
    placeholder={placeholder} 
    value={value} 
    onChange={(e) => onChange(e.target.value)}
    className="bg-slate-50 border border-slate-200 rounded px-2 py-1 text-xs w-32 focus:outline-none focus:border-emerald-500 transition-all"
  />
);

// --- Table Body Renderer with Group Merging ---
function renderTableBody(data: any[], columns: any[], groupKey: string) {
  const rows: React.ReactNode[] = [];
  let rowSpanCount = 0;
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const isStartOfGroup = i === 0 || String(data[i][groupKey]) !== String(data[i-1][groupKey]);

    // Calculate RowSpan if it's the start of a group
    if (isStartOfGroup) {
       rowSpanCount = 1;
       for (let j = i + 1; j < data.length; j++) {
         if (String(data[j][groupKey]) === String(row[groupKey])) {
           rowSpanCount++;
         } else {
           break;
         }
       }
    }

    // Check if Steel No is missing (indicates no match found)
    const isNoMatch = !row['STEEL NO'] || row['STEEL NO'] === 'NO. 없음';

    rows.push(
      <tr key={i} className={`hover:bg-slate-50 border-b border-slate-100 ${isNoMatch ? 'bg-amber-50' : ''}`}>
         {/* Index Column - Merged for group */}
         {isStartOfGroup && (
           <td rowSpan={rowSpanCount} className="px-3 py-2 text-center text-slate-400 font-mono border-r border-slate-200 align-top pt-3 bg-white">
             {i + 1}
           </td>
         )}

         {/* Data Columns */}
         {columns.map((col, cIdx) => {
            // If this is the grouping column, merge it
            if (col.key === groupKey) {
               return isStartOfGroup ? (
                 <td key={col.key} rowSpan={rowSpanCount} className="px-3 py-2 border-r border-slate-200 align-top pt-3 font-bold text-emerald-800 bg-emerald-50/30 whitespace-nowrap">
                    {row[col.key]}
                 </td>
               ) : null;
            }

            // Other columns - Regular cells
            // Special highlighting for Mat columns
            const isMatCol = ['STEEL NO','NESTING DWG','Grade','WEIGHT'].includes(col.key);
            return (
               <td key={col.key} className={`px-3 py-2 border-r border-slate-200 whitespace-nowrap text-slate-700
                  ${isMatCol ? 'bg-blue-50/10' : ''}
                  ${col.key === 'MATCH_MATNO' ? 'font-semibold text-blue-700' : ''}
               `}>
                  {row[col.key]}
               </td>
            );
         })}
      </tr>
    );
  }
  return rows;
}

const container = document.getElementById('root');
const root = createRoot(container!);
root.render(<App />);