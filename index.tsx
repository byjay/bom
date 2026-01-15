import React, { useState, useEffect, useMemo, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import { Download, Upload, Filter, Calculator, Layers, FileSpreadsheet, Trash2, GripHorizontal, AlignLeft, PieChart, BarChart3, X, Loader2, Settings, ChevronDown, ChevronUp, CheckSquare, Square, AlertTriangle } from 'lucide-react';
import * as XLSX from 'xlsx';
import * as echarts from 'echarts';
import 'echarts-gl';

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
  { key: 'MATCH_MATNO', label: 'MATNO', type: 'text' },
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
  { id: 'BLOCK', label: 'BLOCK' }
];

type SortConfig = { key: string; direction: 'asc' | 'desc' } | null;

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

// --- Data Validation System ---
interface ValidationIssue {
  row: number;
  field: string;
  message: string;
  severity: 'error' | 'warning';
}

const validateData = (data: any[]): ValidationIssue[] => {
  const issues: ValidationIssue[] = [];

  data.forEach((row, idx) => {
    // 1. 필수 필드 검증
    if (!row['WELD UNIQUE ID']) {
      issues.push({ row: idx + 1, field: 'WELD UNIQUE ID', message: 'Missing WELD UNIQUE ID', severity: 'error' });
    }

    // 2. MATNO 매칭 검증
    if (!row['MATCH_MATNO']) {
      issues.push({ row: idx + 1, field: 'MATCH_MATNO', message: 'No material number matched', severity: 'warning' });
    }

    // 3. STEEL NO 검증
    if (!row['STEEL NO'] || row['STEEL NO'] === '') {
      issues.push({ row: idx + 1, field: 'STEEL NO', message: 'Missing STEEL NO - Material not found', severity: 'warning' });
    }

    // 4. 수치 검증
    const weight = parseFloat(row['WEIGHT']);
    if (row['WEIGHT'] && (isNaN(weight) || weight < 0)) {
      issues.push({ row: idx + 1, field: 'WEIGHT', message: 'Invalid weight value', severity: 'error' });
    }

    const weldLength = parseFloat(row['WELD. LENG.']);
    if (row['WELD. LENG.'] && (isNaN(weldLength) || weldLength < 0)) {
      issues.push({ row: idx + 1, field: 'WELD. LENG.', message: 'Invalid weld length', severity: 'error' });
    }

    // 5. Grade 검증
    if (row['Grade'] && !/^[A-Z0-9\+\-]+$/i.test(row['Grade'])) {
      issues.push({ row: idx + 1, field: 'Grade', message: 'Invalid Grade format', severity: 'warning' });
    }
  });

  return issues;
};

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

  // NEW: Settings Panel
  const [showSettings, setShowSettings] = useState(false);
  const [visibleColumns, setVisibleColumns] = useState<Record<string, boolean>>(
    COL_DEFINITIONS.reduce((acc, col) => ({ ...acc, [col.key]: true }), {})
  );

  // NEW: Sorting
  const [sortConfig, setSortConfig] = useState<SortConfig>(null);

  // NEW: Selection
  const [selectedRows, setSelectedRows] = useState<Set<number>>(new Set());
  const [selectAll, setSelectAll] = useState(false);

  // NEW: 3D Chart
  const [show3DChart, setShow3DChart] = useState(false);
  const [chartType, setChartType] = useState<'bar3d' | 'pie3d' | 'scatter3d' | 'mixed'>('bar3d');
  const chartRef = useRef<HTMLDivElement>(null);

  // NEW: Validation
  const [validationIssues, setValidationIssues] = useState<ValidationIssue[]>([]);
  const [showValidation, setShowValidation] = useState(false);

  // --- 1. Robust Excel Parsing (Auto-Detect, Multiple Files) ---
  const processExcelFiles = async (files: FileList | File[]) => {
    setIsLoading(true);
    setTimeout(async () => {
      const newWeldData: any[] = [];
      const newMatData: any[] = [];
      let processedCount = 0;

      for (let fIndex = 0; fIndex < files.length; fIndex++) {
          const file = files[fIndex];
          try {
              const buffer = await file.arrayBuffer();
              const wb = XLSX.read(buffer, {
                  cellStyles: false,
                  cellDates: true,
                  sheetStubs: true
              });
              const sheet = wb.Sheets[wb.SheetNames[0]];

              if (!sheet) {
                  console.warn(`Sheet empty in ${file.name}`);
                  continue;
              }

              // A. Handle Merges
              if (sheet['!merges']) {
                  sheet['!merges'].forEach(range => {
                      const startCell = sheet[XLSX.utils.encode_cell(range.s)];
                      const value = startCell?.v ?? startCell?.w ?? '';
                      if (!value) return;

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

              // B. Find Header Row
              const rawJson = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false }) as any[][];
              if (rawJson.length === 0) continue;

              let headerRowIdx = 0;
              let maxScore = -1;
              const WELD_KEYS = ['WELD', 'UNIQUE', 'BLOCK', 'MATNO', 'WLEG', 'LENGTH', 'ID', 'WNO'];
              const MAT_KEYS = ['STEEL', 'NESTING', 'WEIGHT', 'GRADE', 'EA', 'QTY', 'MAT', 'TYPE', 'TPYE'];

              for(let i=0; i < Math.min(rawJson.length, 30); i++) {
                  const row = rawJson[i];
                  const nonEmpty = row.filter(c => c && String(c).trim().length > 0);
                  if (nonEmpty.length === 0) continue;

                  const rowStr = row.map(c => String(c || '').toUpperCase()).join(' ');

                  let score = 0;
                  WELD_KEYS.forEach(k => { if (rowStr.includes(k)) score += 3; });
                  MAT_KEYS.forEach(k => { if (rowStr.includes(k)) score += 3; });
                  score += nonEmpty.length * 0.3;

                  const stringCount = nonEmpty.filter(c => typeof c === 'string' && isNaN(Number(c))).length;
                  if (stringCount / nonEmpty.length > 0.6) score += 5;
                  if (row.some(c => String(c || '').length > 100)) score -= 20;

                  if (score > maxScore) {
                      maxScore = score;
                      headerRowIdx = i;
                  }
              }

              // C. Extract Data & Normalize Headers
              const headers = rawJson[headerRowIdx].map(h => String(h).trim());
              const findKey = (candidates: string[]) => headers.find(h => candidates.some(c => h.toUpperCase().includes(c)));

              const dataRows = [];
              for(let i = headerRowIdx + 1; i < rawJson.length; i++) {
                  const row = rawJson[i];
                  const hasData = row.some(c => c !== null && c !== undefined && String(c).trim() !== '');
                  if (!hasData) continue;

                  const obj: any = { FILENAME: file.name };
                  headers.forEach((h, idx) => {
                      const val = row[idx];
                      obj[h] = (val === null || val === undefined || val === '') ? '' : val;
                  });
                  dataRows.push(obj);
              }

              if (dataRows.length === 0) continue;

              // D. Fill Down Logic
              const idColName = findKey(['WELD UNIQUE ID', 'WELD ID', 'UNIQUE ID']) || 'WELD UNIQUE ID';
              const matColName = findKey(['MATNO', 'MAT NO']) || 'MATNO';

              const FILL_KEYS = ['BLOCK', 'MOD', 'DWG. Title', 'WELD. LENG.', 'SIDE', idColName];

              let lastVals: any = {};
              const filledData = dataRows.map(row => {
                  const currentId = row[idColName];
                  const hasMat = row[matColName];

                  if (currentId) {
                      FILL_KEYS.forEach(k => { if(row[k]) lastVals[k] = row[k]; });
                      row['WELD UNIQUE ID'] = currentId;
                      return row;
                  } else if (hasMat && lastVals[idColName]) {
                      const newRow = { ...row };
                      FILL_KEYS.forEach(k => {
                          if(!newRow[k] && lastVals[k]) newRow[k] = lastVals[k];
                      });
                      newRow['WELD UNIQUE ID'] = lastVals[idColName];
                      return newRow;
                  }
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
                  if (headerStr.includes('NO') || headerStr.includes('ID')) newWeldData.push(...filledData);
              }
              processedCount++;

          } catch (err) {
              console.error(`Error parsing ${file.name}:`, err);
          }
      }

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
      }
    }, 100);
  };

  // --- 2. Matching Logic (Wide-to-Tall + Join) ---
  useEffect(() => {
    if (weldRaw.length === 0) return;

    const matMap = new Map();
    const sampleMat = matRaw[0] || {};
    const matKey = Object.keys(sampleMat).find(k => k.toUpperCase().includes('MATNO') || k.toUpperCase().includes('MAT NO')) || 'MATNO';

    matRaw.forEach(row => {
      const rawKey = row[matKey];
      if (rawKey) {
          const key = String(rawKey).replace(/\s+/g,'').toUpperCase();
          matMap.set(key, row);
      }
    });

    const results: any[] = [];
    weldRaw.forEach((wRow, idx) => {
       const keys = Object.keys(wRow).filter(k => /^MAT.*NO/i.test(k));
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

             results.push({
               ...wRow,
               ...matInfo,
               'MATCH_MATNO': val,
               'STEEL NO': matInfo['STEEL NO'] || matInfo['STEEL_NO'] || '',
               'NESTING DWG': matInfo['NESTING DWG'] || matInfo['NESTING'] || '',
               'Grade': matInfo['Grade'] || matInfo['GRADE'] || '',
               'WEIGHT': matInfo['WEIGHT'] || matInfo['Weight'] || 0,
               _originIdx: idx,
               _rowId: results.length
             });
          }
       });

       if(!foundAny) {
         results.push({ ...wRow, 'MATCH_MATNO': '', _originIdx: idx, _rowId: results.length });
       }
    });

    setMatchedData(results);

    // Run validation
    const issues = validateData(results);
    setValidationIssues(issues);
  }, [weldRaw, matRaw]);

  // --- 3. View Logic (Filtering, Sorting, Aggregation) ---
  const { viewData, stats, groupStats } = useMemo(() => {
    let data = [...matchedData];

    // Filter
    if (filters.block) data = data.filter(r => String(r.BLOCK||'').toLowerCase().includes(filters.block.toLowerCase()));
    if (filters.id) data = data.filter(r => String(r['WELD UNIQUE ID']||'').toLowerCase().includes(filters.id.toLowerCase()));
    if (filters.mat) data = data.filter(r => String(r['MATCH_MATNO']||'').toLowerCase().includes(filters.mat.toLowerCase()));
    if (filters.grade) data = data.filter(r => String(r.Grade||'').toLowerCase().includes(filters.grade.toLowerCase()));

    // Sort
    if (sortConfig) {
      data.sort((a, b) => {
        const aVal = a[sortConfig.key];
        const bVal = b[sortConfig.key];

        // Handle numeric sorting
        const aNum = parseFloat(aVal);
        const bNum = parseFloat(bVal);
        if (!isNaN(aNum) && !isNaN(bNum)) {
          return sortConfig.direction === 'asc' ? aNum - bNum : bNum - aNum;
        }

        // String sorting
        const aStr = String(aVal || '');
        const bStr = String(bVal || '');
        const result = aStr.localeCompare(bStr, undefined, { numeric: true });
        return sortConfig.direction === 'asc' ? result : -result;
      });
    } else {
      // Default: sort by group key
      data.sort((a, b) => {
         const valA = String(a[groupBy] || '');
         const valB = String(b[groupBy] || '');
         return valA.localeCompare(valB, undefined, { numeric: true });
      });
    }

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

    const groupStatsArray = Object.values(groupDist).sort((a: any, b: any) => b.weight - a.weight);

    return { viewData: data, stats, groupStats: groupStatsArray };
  }, [matchedData, filters, groupBy, sortConfig]);

  // --- 4. Export Functions ---
  const exportMultiSheet = () => {
    const wb = XLSX.utils.book_new();

    // Sheet 1: All Data
    const ws1 = XLSX.utils.json_to_sheet(viewData);
    XLSX.utils.book_append_sheet(wb, ws1, "All Data");

    // Sheet 2: By Block
    const blockGroups = viewData.reduce((acc, row) => {
      const block = row.BLOCK || 'Unknown';
      if (!acc[block]) acc[block] = [];
      acc[block].push(row);
      return acc;
    }, {} as Record<string, any[]>);

    Object.entries(blockGroups).slice(0, 10).forEach(([block, rows]) => {
      const ws = XLSX.utils.json_to_sheet(rows);
      const safeSheetName = block.substring(0, 31).replace(/[:\\/?*\[\]]/g, '_');
      XLSX.utils.book_append_sheet(wb, ws, `Block_${safeSheetName}`);
    });

    // Sheet 3: By Grade
    const gradeGroups = viewData.reduce((acc, row) => {
      const grade = row.Grade || 'Unknown';
      if (!acc[grade]) acc[grade] = [];
      acc[grade].push(row);
      return acc;
    }, {} as Record<string, any[]>);

    Object.entries(gradeGroups).slice(0, 10).forEach(([grade, rows]) => {
      const ws = XLSX.utils.json_to_sheet(rows);
      const safeSheetName = grade.substring(0, 31).replace(/[:\\/?*\[\]]/g, '_');
      XLSX.utils.book_append_sheet(wb, ws, `Grade_${safeSheetName}`);
    });

    // Sheet 4: Statistics
    const statsData = [
      { Metric: 'Total Rows', Value: stats.count },
      { Metric: 'Total Weight (kg)', Value: stats.weight.toFixed(2) },
      { Metric: 'Total Weld Length (mm)', Value: stats.length.toFixed(2) },
      { Metric: 'Total Quantity', Value: stats.ea },
      { Metric: '', Value: '' },
      { Metric: 'Group', Value: 'Count' },
      ...groupStats.map((g: any) => ({ Metric: g.name, Value: g.count }))
    ];
    const ws4 = XLSX.utils.json_to_sheet(statsData);
    XLSX.utils.book_append_sheet(wb, ws4, "Statistics");

    // Sheet 5: Selected Items
    if (selectedRows.size > 0) {
      const selectedData = viewData.filter(row => selectedRows.has(row._rowId));
      const ws5 = XLSX.utils.json_to_sheet(selectedData);
      XLSX.utils.book_append_sheet(wb, ws5, "Selected Items");
    }

    XLSX.writeFile(wb, `BOM_Report_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  // --- 5. Selection Handlers ---
  const handleSelectAll = () => {
    if (selectAll) {
      setSelectedRows(new Set());
    } else {
      setSelectedRows(new Set(viewData.map(row => row._rowId)));
    }
    setSelectAll(!selectAll);
  };

  const handleRowSelect = (rowId: number) => {
    const newSelected = new Set(selectedRows);
    if (newSelected.has(rowId)) {
      newSelected.delete(rowId);
    } else {
      newSelected.add(rowId);
    }
    setSelectedRows(newSelected);
    setSelectAll(newSelected.size === viewData.length);
  };

  // --- 6. Sort Handler ---
  const handleSort = (key: string) => {
    let direction: 'asc' | 'desc' = 'asc';
    if (sortConfig && sortConfig.key === key && sortConfig.direction === 'asc') {
      direction = 'desc';
    }
    setSortConfig({ key, direction });
  };

  // --- 7. 3D Chart Rendering ---
  useEffect(() => {
    if (!show3DChart || !chartRef.current || groupStats.length === 0) return;

    const chart = echarts.init(chartRef.current);

    const topGroups = groupStats.slice(0, 10);

    let option: any = {};

    switch (chartType) {
      case 'bar3d':
        option = {
          tooltip: {},
          visualMap: {
            max: Math.max(...topGroups.map((g: any) => g.weight)),
            inRange: { color: ['#50a3ba', '#eac736', '#d94e5d'] }
          },
          xAxis3D: { type: 'category', data: topGroups.map((g: any) => g.name) },
          yAxis3D: { type: 'value', name: 'Weight' },
          zAxis3D: { type: 'value', name: 'Count' },
          grid3D: {
            viewControl: { autoRotate: true, autoRotateSpeed: 5 }
          },
          series: [{
            type: 'bar3D',
            data: topGroups.map((g: any, idx: number) => [idx, g.weight, g.count]),
            shading: 'lambert',
            label: { show: false },
            emphasis: { label: { show: true } }
          }]
        };
        break;

      case 'pie3d':
        option = {
          tooltip: {},
          series: [{
            type: 'pie3D',
            data: topGroups.map((g: any) => ({ name: g.name, value: g.weight })),
            itemStyle: {
              opacity: 0.8
            },
            label: {
              show: true,
              formatter: '{b}: {c}'
            }
          }]
        };
        break;

      case 'scatter3d':
        option = {
          tooltip: {},
          visualMap: {
            max: Math.max(...topGroups.map((g: any) => g.weight)),
            inRange: { color: ['#50a3ba', '#eac736', '#d94e5d'] }
          },
          xAxis3D: { type: 'value', name: 'Count' },
          yAxis3D: { type: 'value', name: 'Weight' },
          zAxis3D: { type: 'value', name: 'Length' },
          grid3D: {
            viewControl: { autoRotate: true, autoRotateSpeed: 3 }
          },
          series: [{
            type: 'scatter3D',
            data: topGroups.map((g: any) => [g.count, g.weight, g.length]),
            symbolSize: 12,
            itemStyle: { opacity: 0.8 }
          }]
        };
        break;

      case 'mixed':
        option = {
          tooltip: {},
          legend: { data: ['Weight', 'Count'] },
          xAxis3D: { type: 'category', data: topGroups.map((g: any) => g.name) },
          yAxis3D: { type: 'value' },
          zAxis3D: { type: 'value' },
          grid3D: {
            viewControl: { autoRotate: true }
          },
          series: [
            {
              name: 'Weight',
              type: 'bar3D',
              data: topGroups.map((g: any, idx: number) => [idx, g.weight, 0]),
              shading: 'lambert'
            },
            {
              name: 'Count',
              type: 'bar3D',
              data: topGroups.map((g: any, idx: number) => [idx, 0, g.count]),
              shading: 'lambert'
            }
          ]
        };
        break;
    }

    chart.setOption(option);

    return () => {
      chart.dispose();
    };
  }, [show3DChart, chartType, groupStats]);

  const clearAll = () => {
    setWeldRaw([]);
    setMatRaw([]);
    setMatchedData([]);
    setSelectedRows(new Set());
    setSelectAll(false);
    setValidationIssues([]);
  };

  // Organize Columns
  const orderedCols = useMemo(() => {
    const visibleColDefs = COL_DEFINITIONS.filter(c => visibleColumns[c.key]);
    const others = visibleColDefs.filter(c => c.key !== groupBy);
    const primary = visibleColDefs.find(c => c.key === groupBy) || { key: groupBy, label: groupBy, type: 'text' };
    return [primary, ...others];
  }, [groupBy, visibleColumns]);

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

      {/* 3D Chart Modal */}
      {show3DChart && (
        <div className="absolute inset-0 bg-black/50 z-50 flex items-center justify-center p-4" onClick={() => setShow3DChart(false)}>
          <div className="bg-white rounded-lg shadow-2xl w-full max-w-5xl h-[80vh] flex flex-col" onClick={e => e.stopPropagation()}>
            <div className="p-4 border-b flex items-center justify-between">
              <h3 className="text-lg font-bold">3D Visualization - {chartType.toUpperCase()}</h3>
              <div className="flex gap-2">
                {(['bar3d', 'pie3d', 'scatter3d', 'mixed'] as const).map(type => (
                  <button
                    key={type}
                    onClick={() => setChartType(type)}
                    className={`px-3 py-1 text-xs rounded ${chartType === type ? 'bg-emerald-600 text-white' : 'bg-slate-200'}`}
                  >
                    {type.toUpperCase()}
                  </button>
                ))}
                <button onClick={() => setShow3DChart(false)} className="ml-4 text-slate-400 hover:text-red-500">
                  <X size={20} />
                </button>
              </div>
            </div>
            <div ref={chartRef} className="flex-1"></div>
          </div>
        </div>
      )}

      {/* Settings Panel */}
      {showSettings && (
        <div className="absolute inset-0 bg-black/50 z-50 flex items-center justify-center p-4" onClick={() => setShowSettings(false)}>
          <div className="bg-white rounded-lg shadow-2xl w-full max-w-2xl max-h-[80vh] flex flex-col" onClick={e => e.stopPropagation()}>
            <div className="p-4 border-b flex items-center justify-between">
              <h3 className="text-lg font-bold">Column Settings</h3>
              <button onClick={() => setShowSettings(false)} className="text-slate-400 hover:text-red-500">
                <X size={20} />
              </button>
            </div>
            <div className="flex-1 overflow-y-auto p-4 grid grid-cols-2 gap-2">
              {COL_DEFINITIONS.map(col => (
                <label key={col.key} className="flex items-center gap-2 p-2 hover:bg-slate-50 rounded cursor-pointer">
                  <input
                    type="checkbox"
                    checked={visibleColumns[col.key]}
                    onChange={(e) => setVisibleColumns({ ...visibleColumns, [col.key]: e.target.checked })}
                    className="w-4 h-4"
                  />
                  <span className="text-sm">{col.label}</span>
                </label>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* Validation Panel */}
      {showValidation && (
        <div className="absolute inset-0 bg-black/50 z-50 flex items-center justify-center p-4" onClick={() => setShowValidation(false)}>
          <div className="bg-white rounded-lg shadow-2xl w-full max-w-4xl max-h-[80vh] flex flex-col" onClick={e => e.stopPropagation()}>
            <div className="p-4 border-b flex items-center justify-between">
              <h3 className="text-lg font-bold flex items-center gap-2">
                <AlertTriangle className="text-amber-500" size={20} />
                Data Validation Issues ({validationIssues.length})
              </h3>
              <button onClick={() => setShowValidation(false)} className="text-slate-400 hover:text-red-500">
                <X size={20} />
              </button>
            </div>
            <div className="flex-1 overflow-y-auto p-4">
              {validationIssues.length === 0 ? (
                <div className="text-center text-slate-400 py-10">No validation issues found</div>
              ) : (
                <table className="w-full text-xs">
                  <thead className="bg-slate-50 sticky top-0">
                    <tr>
                      <th className="px-3 py-2 text-left border-b">Row</th>
                      <th className="px-3 py-2 text-left border-b">Field</th>
                      <th className="px-3 py-2 text-left border-b">Message</th>
                      <th className="px-3 py-2 text-left border-b">Severity</th>
                    </tr>
                  </thead>
                  <tbody>
                    {validationIssues.map((issue, idx) => (
                      <tr key={idx} className="hover:bg-slate-50 border-b">
                        <td className="px-3 py-2">{issue.row}</td>
                        <td className="px-3 py-2 font-semibold">{issue.field}</td>
                        <td className="px-3 py-2">{issue.message}</td>
                        <td className="px-3 py-2">
                          <span className={`px-2 py-0.5 rounded text-xs ${issue.severity === 'error' ? 'bg-red-100 text-red-700' : 'bg-amber-100 text-amber-700'}`}>
                            {issue.severity}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Top Bar */}
      <div className="bg-white border-b border-slate-200 shadow-sm z-30 flex-none">
        <div className="px-4 py-3 flex items-center justify-between">
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

           <div className="flex items-center gap-2">
               {validationIssues.length > 0 && (
                 <button
                    onClick={() => setShowValidation(true)}
                    className="flex items-center gap-2 px-3 py-2 rounded-lg text-sm font-semibold bg-amber-50 text-amber-700 border border-amber-200"
                 >
                    <AlertTriangle size={18} /> {validationIssues.length} Issues
                 </button>
               )}
               <button
                  onClick={() => setShow3DChart(true)}
                  className="flex items-center gap-2 px-3 py-2 rounded-lg text-sm font-semibold text-purple-600 hover:bg-purple-50 border border-transparent hover:border-purple-200"
               >
                  <PieChart size={18} /> 3D Chart
               </button>
               <button
                  onClick={() => setShowSettings(true)}
                  className="flex items-center gap-2 px-3 py-2 rounded-lg text-sm font-semibold text-slate-600 hover:bg-slate-50"
               >
                  <Settings size={18} /> Settings
               </button>
               <button
                  onClick={() => setShowAnalysis(!showAnalysis)}
                  className={`flex items-center gap-2 px-3 py-2 rounded-lg text-sm font-semibold transition-colors border
                     ${showAnalysis ? 'bg-indigo-50 text-indigo-700 border-indigo-200' : 'text-slate-500 border-transparent hover:bg-slate-50'}`}
               >
                  <BarChart3 size={18} /> Stats
               </button>
               <button
                  onClick={exportMultiSheet}
                  disabled={viewData.length===0}
                  className="flex items-center gap-2 text-emerald-600 hover:bg-emerald-50 px-3 py-2 rounded-lg text-sm font-semibold disabled:opacity-50 transition-colors border border-transparent hover:border-emerald-200"
               >
                  <FileSpreadsheet size={18} /> Export
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

      {/* Stats Dashboard */}
      <div className="bg-emerald-900 text-white px-6 py-3 grid grid-cols-4 gap-4 shadow-inner flex-none z-20">
          <StatItem label="Total Rows" value={stats.count} icon={<AlignLeft size={16}/>} />
          <StatItem label="Total Weight (kg)" value={stats.weight.toFixed(1)} icon={<Calculator size={16}/>} />
          <StatItem label="Total Weld Length" value={stats.length.toLocaleString()} unit="mm" icon={<Calculator size={16}/>} />
          <StatItem label="Selected" value={selectedRows.size} unit="rows" icon={<CheckSquare size={16}/>} />
      </div>

      {/* Filter Bar */}
      <div className="bg-white px-4 py-2 flex items-center gap-4 border-b border-slate-200 flex-none z-20">
         <Filter size={16} className="text-slate-400" />
         <FilterInput placeholder="Block..." value={filters.block} onChange={(v: string) => setFilters({...filters, block: v})} />
         <FilterInput placeholder="Weld ID..." value={filters.id} onChange={(v: string) => setFilters({...filters, id: v})} />
         <FilterInput placeholder="Mat No..." value={filters.mat} onChange={(v: string) => setFilters({...filters, mat: v})} />
         <FilterInput placeholder="Grade..." value={filters.grade} onChange={(v: string) => setFilters({...filters, grade: v})} />
      </div>

      {/* Main Content */}
      <div className="flex-1 flex overflow-hidden">

        {/* Table */}
        <div className="flex-1 overflow-auto bg-slate-100 p-4">
            <div className="bg-white rounded-lg shadow-sm border border-slate-200 overflow-hidden min-h-[300px]">
            {viewData.length > 0 ? (
                <div className="overflow-x-auto">
                    <table className="w-full text-xs text-left border-collapse">
                    <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm">
                        <tr>
                            <th className="px-3 py-2 border-b border-r border-slate-200 w-10 text-center">
                              <button onClick={handleSelectAll} className="hover:text-emerald-600">
                                {selectAll ? <CheckSquare size={16} /> : <Square size={16} />}
                              </button>
                            </th>
                            <th className="px-3 py-2 border-b border-r border-slate-200 w-10 text-center font-bold text-slate-400">#</th>
                            {orderedCols.map((col, idx) => (
                            <th
                              key={col.key}
                              className={`px-3 py-2 border-b border-r border-slate-200 font-bold whitespace-nowrap cursor-pointer hover:bg-slate-100
                                ${idx === 0 ? 'bg-emerald-50 text-emerald-800 border-r-emerald-200' : 'text-slate-600'}
                                ${['STEEL NO','NESTING DWG','Grade','WEIGHT'].includes(col.key) ? 'bg-blue-50/50' : ''}
                            `}
                              onClick={() => handleSort(col.key)}
                            >
                              <div className="flex items-center gap-1">
                                {col.label}
                                {sortConfig?.key === col.key && (
                                  sortConfig.direction === 'asc' ? <ChevronUp size={12} /> : <ChevronDown size={12} />
                                )}
                              </div>
                            </th>
                            ))}
                        </tr>
                    </thead>
                    <tbody>
                        {renderTableBody(viewData, orderedCols, groupBy, selectedRows, handleRowSelect)}
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

        {/* Analysis Sidebar */}
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

                            return (
                                <div key={idx} className="group">
                                    <div className="flex justify-between items-baseline mb-1">
                                        <span className="text-xs font-bold text-slate-700 truncate w-40" title={grp.name}>{grp.name}</span>
                                        <span className="text-[10px] font-mono text-slate-500">{weightPercent.toFixed(1)}%</span>
                                    </div>

                                    <div className="w-full h-1.5 bg-slate-100 rounded-full mb-2 overflow-hidden">
                                        <div
                                            className="h-full bg-gradient-to-r from-indigo-500 to-purple-500 rounded-full"
                                            style={{ width: `${Math.max(weightPercent, 1)}%` }}
                                        ></div>
                                    </div>

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

// --- Table Body Renderer ---
function renderTableBody(
  data: any[],
  columns: any[],
  groupKey: string,
  selectedRows: Set<number>,
  handleRowSelect: (id: number) => void
) {
  const rows: React.ReactNode[] = [];
  let rowSpanCount = 0;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const isStartOfGroup = i === 0 || String(data[i][groupKey]) !== String(data[i-1][groupKey]);

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

    const isNoMatch = !row['STEEL NO'] || row['STEEL NO'] === 'NO. 없음';
    const isSelected = selectedRows.has(row._rowId);

    rows.push(
      <tr key={i} className={`hover:bg-slate-50 border-b border-slate-100 ${isNoMatch ? 'bg-amber-50' : ''} ${isSelected ? 'bg-blue-50' : ''}`}>
         <td className="px-3 py-2 text-center border-r border-slate-200">
           <button onClick={() => handleRowSelect(row._rowId)} className="hover:text-emerald-600">
             {isSelected ? <CheckSquare size={16} /> : <Square size={16} />}
           </button>
         </td>

         {isStartOfGroup && (
           <td rowSpan={rowSpanCount} className="px-3 py-2 text-center text-slate-400 font-mono border-r border-slate-200 align-top pt-3 bg-white">
             {i + 1}
           </td>
         )}

         {columns.map((col, cIdx) => {
            if (col.key === groupKey) {
               return isStartOfGroup ? (
                 <td key={col.key} rowSpan={rowSpanCount} className="px-3 py-2 border-r border-slate-200 align-top pt-3 font-bold text-emerald-800 bg-emerald-50/30 whitespace-nowrap">
                    {row[col.key]}
                 </td>
               ) : null;
            }

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
