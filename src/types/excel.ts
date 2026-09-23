export interface ColumnMeta {
  index: number; // 0-based column index
  letter: string; // A, B, C...
  name: string; // Header text
  sampleValue?: string | number | boolean | null;
  selected: boolean;
  outputName?: string;
}

export interface SheetData {
  sheetName: string;
  rawRows: (string | number | boolean | null | undefined)[][];
  headerRowIndex: number; // 0-based
  dataStartRowIndex: number; // 0-based
  columns: ColumnMeta[];
  totalRows: number;
}

export interface UploadedFileInfo {
  id: string;
  name: string;
  size: number;
  status: 'pending' | 'parsing' | 'ready' | 'error';
  errorMsg?: string;
  sheetNames: string[];
  activeSheetName: string;
  headerRowIndex: number;
  dataStartRowIndex: number;
  currentSheetData?: SheetData;
  totalRows: number;
  enabled: boolean; // whether to include this file in multi-file extraction
  progressPercent?: number;
  buffer?: ArrayBuffer;
}

export interface MatchResultRow {
  originalRowIndex: number;
  matchedCode: string;
  sourceFileName: string;
  sourceSheetName: string;
  data: Record<string, string | number | boolean | null | undefined>; // keyed by column letter or outputName
}

export interface FileMatchStat {
  fileId: string;
  fileName: string;
  matchedRowCount: number;
  uniqueCodesFound: number;
}

export interface ExtractionResult {
  matchedRows: MatchResultRow[];
  matchedCodes: string[];
  unmatchedCodes: string[];
  duplicateCodesCount: Record<string, number>;
  totalRowsScanned: number;
  totalFilesProcessed: number;
  fileStats: FileMatchStat[];
  durationMs: number;
}

export interface ExtractionConfig {
  headerRowIndex: number; // 0-based (default/primary)
  dataStartRowIndex: number; // 0-based (default/primary)
  referenceColumnLetter: string; // e.g. "C"
  referenceColumnName?: string; // e.g. "PR/ INDENT"
  matchBy: 'letter' | 'header_name'; // match by column letter across files, or search for column name in each file
  selectedColumns: {
    letter: string;
    originalName: string;
    outputName: string;
  }[];
  matchMode: 'exact' | 'case_insensitive_trim' | 'contains';
  matchStrategy: 'all_rows' | 'first_row_only';
  includeUnmatchedSummary: boolean;
  includeSourceFileName: boolean; // include "Tên tệp nguồn" column in export
  includeSheetName: boolean; // include "Tên Sheet" column in export
}

export interface ExtractionPreset {
  id: string;
  name: string;
  description?: string;
  createdAt: string;
  referenceColumnLetter: string;
  referenceColumnName?: string;
  matchBy?: 'letter' | 'header_name';
  headerRowIndex: number;
  dataStartRowIndex: number;
  selectedColumnLetters: string[];
  columnAliases: Record<string, string>;
  matchMode: 'exact' | 'case_insensitive_trim' | 'contains';
  matchStrategy: 'all_rows' | 'first_row_only';
}
