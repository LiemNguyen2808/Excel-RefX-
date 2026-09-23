import * as XLSX from 'xlsx';
import {
  ColumnMeta,
  SheetData,
  UploadedFileInfo,
  ExtractionConfig,
  ExtractionResult,
  MatchResultRow,
  FileMatchStat,
} from '../types/excel';

/**
 * Convert 0-based column index to Excel column letter (0 -> "A", 25 -> "Z", 26 -> "AA")
 */
export function indexToColumnLetter(index: number): string {
  let letter = '';
  let temp = index;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

/**
 * Convert Excel column letter to 0-based column index ("A" -> 0, "C" -> 2, "AA" -> 26)
 */
export function columnLetterToIndex(letter: string): number {
  const upper = letter.toUpperCase().trim();
  let index = 0;
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index - 1;
}

/**
 * High-performance Excel parsing with Dense mode & minimal memory overhead for massive files
 */
export function parseExcelWorkbook(
  data: ArrayBuffer,
  defaultHeaderRow: number = 0,
  defaultDataRow?: number
): { sheetNames: string[]; getSheetData: (name: string, headerRow?: number, dataRow?: number) => SheetData } {
  // Dense mode is 5x faster and uses up to 80% less memory on huge files
  const workbook = XLSX.read(data, {
    type: 'array',
    dense: true,
    cellDates: false,
    cellNF: false,
    cellHTML: false,
    cellText: false,
    cellFormula: false,
  });

  const sheetNames = workbook.SheetNames;

  const getSheetData = (name: string, headerRow = defaultHeaderRow, dataRow?: number): SheetData => {
    const worksheet = workbook.Sheets[name];
    if (!worksheet) {
      throw new Error(`Sheet ${name} not found`);
    }

    let rawRows: (string | number | boolean | null | undefined)[][];

    // If dense mode populated !data, read directly without allocating millions of cell keys
    const denseData = (worksheet as unknown as { '!data'?: Array<Array<{ v?: string | number | boolean } | null>> })['!data'];
    if (denseData && Array.isArray(denseData)) {
      rawRows = denseData.map((row) =>
        row ? row.map((cell) => (cell && cell.v !== undefined ? cell.v : null)) : []
      );
    } else {
      rawRows = XLSX.utils.sheet_to_json<(string | number | boolean | null | undefined)[]>(worksheet, {
        header: 1,
        raw: true,
        defval: null,
        blankrows: true,
      });
    }

    const totalRows = rawRows.length;
    const finalHeaderRow = Math.max(0, Math.min(headerRow, Math.max(0, totalRows - 1)));
    const finalDataRow = dataRow !== undefined ? dataRow : finalHeaderRow + 1;

    // Detect maximum column count by checking top rows
    let maxCols = 0;
    const inspectionLimit = Math.min(100, totalRows);
    for (let i = 0; i < inspectionLimit; i++) {
      if (rawRows[i] && rawRows[i].length > maxCols) {
        maxCols = rawRows[i].length;
      }
    }
    if (rawRows[finalHeaderRow] && rawRows[finalHeaderRow].length > maxCols) {
      maxCols = rawRows[finalHeaderRow].length;
    }

    const headerLine = rawRows[finalHeaderRow] || [];
    const sampleLine = rawRows[finalDataRow] || [];

    const columns: ColumnMeta[] = [];
    for (let c = 0; c < maxCols; c++) {
      const letter = indexToColumnLetter(c);
      const rawHeader = headerLine[c];
      let headerName = '';
      if (rawHeader !== undefined && rawHeader !== null) {
        headerName = String(rawHeader).trim();
      }
      if (!headerName) {
        headerName = `[Cột ${letter}]`;
      }

      const sampleVal = sampleLine[c] !== undefined && sampleLine[c] !== null ? sampleLine[c] : null;

      columns.push({
        index: c,
        letter,
        name: headerName,
        sampleValue: sampleVal,
        selected: true,
        outputName: headerName,
      });
    }

    return {
      sheetName: name,
      rawRows,
      headerRowIndex: finalHeaderRow,
      dataStartRowIndex: finalDataRow,
      columns,
      totalRows,
    };
  };

  return { sheetNames, getSheetData };
}

/**
 * Standardize code for matching
 */
function normalizeCode(val: unknown, mode: ExtractionConfig['matchMode']): string {
  if (val === null || val === undefined) return '';
  const str = String(val);
  if (mode === 'exact') {
    return str;
  }
  return str.trim().toLowerCase();
}

/**
 * Perform reference extraction on a single sheet
 */
export function extractDataByCodes(
  sheetData: SheetData,
  searchCodes: string[],
  config: ExtractionConfig,
  sourceFileName: string = 'Tệp hiện tại',
  sourceSheetName: string = sheetData.sheetName
): ExtractionResult {
  return extractDataFromMultipleSheets(
    [
      {
        fileName: sourceFileName,
        sheetName: sourceSheetName,
        sheetData,
      },
    ],
    searchCodes,
    config
  );
}

/**
 * Perform reference extraction across multiple files and sheets simultaneously
 */
export function extractDataFromMultipleFiles(
  files: UploadedFileInfo[],
  searchCodes: string[],
  config: ExtractionConfig,
  onProgress?: (progressText: string, percent: number) => void
): ExtractionResult {
  const targets = files
    .filter((f) => f.enabled && f.currentSheetData)
    .map((f) => ({
      fileId: f.id,
      fileName: f.name,
      sheetName: f.activeSheetName,
      sheetData: f.currentSheetData!,
    }));

  return extractDataFromMultipleSheets(targets, searchCodes, config, onProgress);
}

interface SheetTarget {
  fileId?: string;
  fileName: string;
  sheetName: string;
  sheetData: SheetData;
}

function extractDataFromMultipleSheets(
  targets: SheetTarget[],
  searchCodes: string[],
  config: ExtractionConfig,
  onProgress?: (progressText: string, percent: number) => void
): ExtractionResult {
  const startTime = performance.now();

  // Filter and normalize lookup codes
  const cleanCodes = searchCodes
    .map((c) => String(c).trim())
    .filter((c) => c.length > 0);

  const normalizedTargets = cleanCodes.map((code) => ({
    original: code,
    normalized: normalizeCode(code, config.matchMode),
  }));

  const targetMap = new Map<string, string[]>();
  for (const t of normalizedTargets) {
    if (!targetMap.has(t.normalized)) {
      targetMap.set(t.normalized, []);
    }
    targetMap.get(t.normalized)!.push(t.original);
  }

  const matchedRows: MatchResultRow[] = [];
  const matchedCodesSet = new Set<string>();
  const duplicateCodesCount: Record<string, number> = {};
  const fileStatsMap = new Map<string, { fileName: string; matchedCount: number; uniqueCodes: Set<string> }>();

  let totalRowsScanned = 0;

  for (let tIndex = 0; tIndex < targets.length; tIndex++) {
    const target = targets[tIndex];
    const sheetData = target.sheetData;
    const fileKey = target.fileId || target.fileName;

    if (!fileStatsMap.has(fileKey)) {
      fileStatsMap.set(fileKey, {
        fileName: target.fileName,
        matchedCount: 0,
        uniqueCodes: new Set<string>(),
      });
    }
    const currentFileStat = fileStatsMap.get(fileKey)!;

    if (onProgress) {
      const pct = Math.round((tIndex / targets.length) * 100);
      onProgress(`Đang quét tệp ${tIndex + 1}/${targets.length}: ${target.fileName}...`, pct);
    }

    // Determine reference column index in this sheet
    let refColIndex = columnLetterToIndex(config.referenceColumnLetter);

    // If matching by header name across files, find matching column name
    if (config.matchBy === 'header_name' && config.referenceColumnName) {
      const targetName = config.referenceColumnName.toLowerCase().trim();
      const colByName = sheetData.columns.find(
        (c) => c.name.toLowerCase().trim() === targetName
      );
      if (colByName) {
        refColIndex = colByName.index;
      }
    }

    // Map output columns for this sheet
    const colMap = config.selectedColumns.map((col) => {
      let resolvedIndex = columnLetterToIndex(col.letter);

      // If matching by header name, try finding column by original name in this sheet
      if (config.matchBy === 'header_name' && col.originalName) {
        const found = sheetData.columns.find(
          (c) => c.name.toLowerCase().trim() === col.originalName.toLowerCase().trim()
        );
        if (found) {
          resolvedIndex = found.index;
        }
      }

      return {
        colIndex: resolvedIndex,
        outputKey: col.outputName || col.originalName || `Cột ${col.letter}`,
        letter: col.letter,
      };
    });

    const dataRows = sheetData.rawRows.slice(sheetData.dataStartRowIndex);

    for (let rIndex = 0; rIndex < dataRows.length; rIndex++) {
      totalRowsScanned++;
      const row = dataRows[rIndex];
      if (!row) continue;

      const cellVal = row[refColIndex];
      if (cellVal === null || cellVal === undefined) continue;

      const normalizedCell = normalizeCode(cellVal, config.matchMode);
      if (!normalizedCell) continue;

      let matchedOriginalCodes: string[] = [];

      if (config.matchMode === 'contains') {
        for (const t of normalizedTargets) {
          if (normalizedCell.includes(t.normalized)) {
            matchedOriginalCodes.push(t.original);
          }
        }
      } else {
        if (targetMap.has(normalizedCell)) {
          matchedOriginalCodes = targetMap.get(normalizedCell)!;
        }
      }

      if (matchedOriginalCodes.length > 0) {
        for (const code of matchedOriginalCodes) {
          matchedCodesSet.add(code);
          duplicateCodesCount[code] = (duplicateCodesCount[code] || 0) + 1;

          if (config.matchStrategy === 'first_row_only' && duplicateCodesCount[code] > 1) {
            continue;
          }

          const rowData: Record<string, string | number | boolean | null | undefined> = {};
          for (const c of colMap) {
            const val = row[c.colIndex];
            rowData[c.outputKey] = val !== undefined && val !== null ? val : '';
          }

          matchedRows.push({
            originalRowIndex: sheetData.dataStartRowIndex + rIndex + 1,
            matchedCode: code,
            sourceFileName: target.fileName,
            sourceSheetName: target.sheetName,
            data: rowData,
          });

          currentFileStat.matchedCount++;
          currentFileStat.uniqueCodes.add(code);
        }
      }
    }
  }

  // Find unmatched codes across all files
  const unmatchedCodes = cleanCodes.filter((c) => !matchedCodesSet.has(c));
  const endTime = performance.now();

  const fileStats: FileMatchStat[] = Array.from(fileStatsMap.entries()).map(([fId, stat]) => ({
    fileId: fId,
    fileName: stat.fileName,
    matchedRowCount: stat.matchedCount,
    uniqueCodesFound: stat.uniqueCodes.size,
  }));

  if (onProgress) {
    onProgress('Hoàn thành trích xuất dữ liệu!', 100);
  }

  return {
    matchedRows,
    matchedCodes: Array.from(matchedCodesSet),
    unmatchedCodes,
    duplicateCodesCount,
    totalRowsScanned,
    totalFilesProcessed: targets.length,
    fileStats,
    durationMs: Math.round(endTime - startTime),
  };
}

/**
 * Export results to XLSX file with multiple sheets:
 * Sheet 1: Consolidated Extracted Data
 * Sheet 2: File Statistics (if multi-file)
 * Sheet 3: Unmatched Codes (if requested)
 */
export function exportResultsToExcel(
  results: ExtractionResult,
  config: ExtractionConfig,
  fileName: string = 'Ket_Qua_Trich_Xuat_Tong_Hop.xlsx'
): void {
  const wb = XLSX.utils.book_new();

  // Prepare sheet 1: Consolidated Matched Data
  const sheet1Data: Record<string, unknown>[] = [];
  for (const row of results.matchedRows) {
    const item: Record<string, unknown> = {};

    if (config.includeSourceFileName) {
      item['Tệp nguồn'] = row.sourceFileName;
    }
    if (config.includeSheetName) {
      item['Sheet nguồn'] = row.sourceSheetName;
    }

    item['Dòng nguồn (Row)'] = row.originalRowIndex;
    item['Mã tham chiếu khớp'] = row.matchedCode;

    Object.assign(item, row.data);
    sheet1Data.push(item);
  }

  const wsData = XLSX.utils.json_to_sheet(sheet1Data);

  // Set column widths dynamically
  const colKeys = sheet1Data.length > 0 ? Object.keys(sheet1Data[0]) : [];
  const colWidths = colKeys.map((key) => {
    let maxLen = key.length;
    for (let i = 0; i < Math.min(100, sheet1Data.length); i++) {
      const val = sheet1Data[i][key];
      if (val !== undefined && val !== null) {
        const str = String(val);
        if (str.length > maxLen) maxLen = str.length;
      }
    }
    return { wch: Math.min(65, Math.max(12, maxLen + 3)) };
  });
  wsData['!cols'] = colWidths;

  XLSX.utils.book_append_sheet(wb, wsData, 'Dữ liệu trích xuất');

  // Prepare sheet 2: File Statistics Summary (if more than 1 file)
  if (results.fileStats && results.fileStats.length > 1) {
    const statsData = results.fileStats.map((f, idx) => ({
      STT: idx + 1,
      'Tên tệp Excel': f.fileName,
      'Số dòng khớp': f.matchedRowCount,
      'Số mã tìm thấy': f.uniqueCodesFound,
    }));
    const wsStats = XLSX.utils.json_to_sheet(statsData);
    wsStats['!cols'] = [{ wch: 8 }, { wch: 45 }, { wch: 18 }, { wch: 18 }];
    XLSX.utils.book_append_sheet(wb, wsStats, 'Thống kê theo tệp');
  }

  // Prepare sheet 3: Unmatched Codes
  if (config.includeUnmatchedSummary && results.unmatchedCodes.length > 0) {
    const sheet3Data = results.unmatchedCodes.map((code, idx) => ({
      STT: idx + 1,
      'Mã không tìm thấy': code,
      'Trạng thái': 'Không có dữ liệu trong bất kỳ tệp nguồn nào',
    }));
    const wsUnmatched = XLSX.utils.json_to_sheet(sheet3Data);
    wsUnmatched['!cols'] = [{ wch: 8 }, { wch: 30 }, { wch: 45 }];
    XLSX.utils.book_append_sheet(wb, wsUnmatched, 'Mã không tìm thấy');
  }

  // Trigger download
  XLSX.writeFile(wb, fileName);
}

/**
 * Export results to CSV (UTF-8 with BOM)
 */
export function exportResultsToCsv(
  results: ExtractionResult,
  config: ExtractionConfig,
  fileName: string = 'Ket_Qua_Trich_Xuat.csv'
): void {
  if (results.matchedRows.length === 0) {
    alert('Không có dữ liệu để xuất CSV');
    return;
  }

  const headers: string[] = [];
  if (config.includeSourceFileName) headers.push('Tệp nguồn');
  if (config.includeSheetName) headers.push('Sheet nguồn');
  headers.push('Dòng nguồn', 'Mã tham chiếu', ...Object.keys(results.matchedRows[0].data));

  const escapeCsv = (str: unknown) => {
    if (str === null || str === undefined) return '""';
    const s = String(str).replace(/"/g, '""');
    return `"${s}"`;
  };

  const csvRows: string[] = [];
  csvRows.push(headers.map(escapeCsv).join(','));

  for (const r of results.matchedRows) {
    const rowValues: unknown[] = [];
    if (config.includeSourceFileName) rowValues.push(r.sourceFileName);
    if (config.includeSheetName) rowValues.push(r.sourceSheetName);
    rowValues.push(r.originalRowIndex, r.matchedCode, ...Object.values(r.data));
    csvRows.push(rowValues.map(escapeCsv).join(','));
  }

  const csvContent = '\uFEFF' + csvRows.join('\r\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * Export TSV format for clipboard copying (paste directly into Excel)
 */
export function getTsvDataForClipboard(results: ExtractionResult, config: ExtractionConfig): string {
  if (results.matchedRows.length === 0) return '';

  const headers: string[] = [];
  if (config.includeSourceFileName) headers.push('Tệp nguồn');
  if (config.includeSheetName) headers.push('Sheet nguồn');
  headers.push('Dòng nguồn', 'Mã tham chiếu', ...Object.keys(results.matchedRows[0].data));

  const rows: string[] = [headers.join('\t')];

  for (const r of results.matchedRows) {
    const rowValues: unknown[] = [];
    if (config.includeSourceFileName) rowValues.push(r.sourceFileName);
    if (config.includeSheetName) rowValues.push(r.sourceSheetName);
    rowValues.push(
      r.originalRowIndex,
      r.matchedCode,
      ...Object.values(r.data).map((v) =>
        v === null || v === undefined ? '' : String(v).replace(/\t|\r|\n/g, ' ')
      )
    );
    rows.push(rowValues.join('\t'));
  }

  return rows.join('\n');
}
