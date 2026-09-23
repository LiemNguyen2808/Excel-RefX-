import React, { useState, useEffect } from 'react';
import { TopBar } from './components/TopBar';
import { FileUploader } from './components/FileUploader';
import { ConfigStep } from './components/ConfigStep';
import { CodeInputStep } from './components/CodeInputStep';
import { ResultsTable } from './components/ResultsTable';
import { GuideModal } from './components/GuideModal';
import {
  SheetData,
  UploadedFileInfo,
  ExtractionConfig,
  ExtractionResult,
  ExtractionPreset,
} from './types/excel';
import {
  parseExcelWorkbook,
  extractDataFromMultipleFiles,
} from './utils/excelParser';
import {
  createMultiSampleWorkbooks,
  SAMPLE_QUERY_CODES,
} from './utils/sampleData';

const PRESETS_STORAGE_KEY = 'excel_refx_presets_v2';

const DEFAULT_PRESETS: ExtractionPreset[] = [
  {
    id: 'preset-yarn-plan',
    name: 'Mẫu Kế hoạch Sợi Đa Tệp (Theo ảnh minh họa)',
    description: 'Tham chiếu theo Cột C: PR/INDENT, lấy Tên sợi, Nhu cầu, Tồn kho, Khách hàng, PO, Lot',
    createdAt: new Date().toISOString(),
    referenceColumnLetter: 'C',
    referenceColumnName: 'PR/ INDENT',
    matchBy: 'letter',
    headerRowIndex: 0,
    dataStartRowIndex: 2,
    selectedColumnLetters: ['C', 'D', 'F', 'G', 'I', 'K', 'M', 'P'],
    columnAliases: {
      C: 'Mã PR/ INDENT',
      D: 'Loại sợi (Yarn Kind)',
      F: 'Nhu cầu (KG)',
      G: 'Tồn kho W/H (KG)',
      I: 'Đã sử dụng (KG)',
      K: 'Khách hàng (Buyer)',
      M: 'Đơn hàng (P/O)',
      P: 'Số Lot',
    },
    matchMode: 'case_insensitive_trim',
    matchStrategy: 'all_rows',
  },
];

export default function App() {
  const [files, setFiles] = useState<UploadedFileInfo[]>([]);
  const [inspectedFileId, setInspectedFileId] = useState<string | null>(null);
  const [isBatchReading, setIsBatchReading] = useState(false);
  const [batchReadProgress, setBatchReadProgress] = useState('');

  // Primary configuration state (applied to active/all files)
  const [referenceColumnLetter, setReferenceColumnLetter] = useState<string>('C');
  const [referenceColumnName, setReferenceColumnName] = useState<string>('PR/ INDENT');
  const [matchBy, setMatchBy] = useState<'letter' | 'header_name'>('letter');
  const [selectedColumnLetters, setSelectedColumnLetters] = useState<string[]>([]);
  const [columnAliases, setColumnAliases] = useState<Record<string, string>>({});
  const [includeSourceFileName, setIncludeSourceFileName] = useState<boolean>(true);
  const [includeSheetName, setIncludeSheetName] = useState<boolean>(false);
  const [rawCodeInput, setRawCodeInput] = useState<string>('');

  const [config, setConfig] = useState<ExtractionConfig>({
    headerRowIndex: 0,
    dataStartRowIndex: 2,
    referenceColumnLetter: 'C',
    referenceColumnName: 'PR/ INDENT',
    matchBy: 'letter',
    selectedColumns: [],
    matchMode: 'case_insensitive_trim',
    matchStrategy: 'all_rows',
    includeUnmatchedSummary: true,
    includeSourceFileName: true,
    includeSheetName: false,
  });

  // Results & Execution state
  const [results, setResults] = useState<ExtractionResult | null>(null);
  const [isExecuting, setIsExecuting] = useState(false);
  const [isGuideOpen, setIsGuideOpen] = useState(false);

  // Presets
  const [presets, setPresets] = useState<ExtractionPreset[]>(() => {
    try {
      const saved = localStorage.getItem(PRESETS_STORAGE_KEY);
      if (saved) return JSON.parse(saved);
    } catch {
      // Fallback
    }
    return DEFAULT_PRESETS;
  });

  // Active file for column preview and inspector
  const primaryFile = files.find((f) => f.id === inspectedFileId) || files.find((f) => f.enabled) || files[0];
  const primarySheetData = primaryFile?.currentSheetData;

  // Save presets
  useEffect(() => {
    try {
      localStorage.setItem(PRESETS_STORAGE_KEY, JSON.stringify(presets));
    } catch {
      // Ignore
    }
  }, [presets]);

  // Synchronize columns when primary file changes
  useEffect(() => {
    if (primarySheetData && selectedColumnLetters.length === 0) {
      const defaultCols = primarySheetData.columns.slice(0, 10).map((c) => c.letter);
      setSelectedColumnLetters(defaultCols);
    }
  }, [primarySheetData]);

  // Read a single file asynchronously
  const parseSingleFile = async (
    file: File | { name: string; buffer: ArrayBuffer; size: number },
    defaultHeaderRow: number = 0,
    defaultDataRow: number = 2
  ): Promise<UploadedFileInfo> => {
    let buffer: ArrayBuffer;
    let name: string;
    let size: number;

    if (file instanceof File) {
      buffer = await file.arrayBuffer();
      name = file.name;
      size = file.size;
    } else {
      buffer = file.buffer;
      name = file.name;
      size = file.size;
    }

    const parsed = parseExcelWorkbook(buffer, defaultHeaderRow, defaultDataRow);
    const firstSheetName = parsed.sheetNames[0] || 'Sheet1';
    const sheetData = parsed.getSheetData(firstSheetName, defaultHeaderRow, defaultDataRow);

    return {
      id: 'file-' + Math.random().toString(36).substring(2, 9),
      name,
      size,
      status: 'ready',
      sheetNames: parsed.sheetNames,
      activeSheetName: firstSheetName,
      headerRowIndex: defaultHeaderRow,
      dataStartRowIndex: defaultDataRow,
      currentSheetData: sheetData,
      totalRows: sheetData.totalRows,
      enabled: true,
      buffer,
    };
  };

  // Upload multiple files concurrently / chunked
  const handleUploadMultipleFiles = async (newFiles: FileList | File[]) => {
    const fileArray = Array.from(newFiles);
    if (fileArray.length === 0) return;

    setIsBatchReading(true);
    setBatchReadProgress(`Đang tải và tối ưu bộ nhớ cho ${fileArray.length} tệp...`);

    const parsedList: UploadedFileInfo[] = [];

    for (let i = 0; i < fileArray.length; i++) {
      const f = fileArray[i];
      setBatchReadProgress(`Đang đọc tệp ${i + 1}/${fileArray.length}: ${f.name}...`);
      try {
        // Yield to browser UI
        await new Promise((resolve) => setTimeout(resolve, 10));
        const fileInfo = await parseSingleFile(f, config.headerRowIndex, config.dataStartRowIndex);
        parsedList.push(fileInfo);
      } catch (err) {
        parsedList.push({
          id: 'file-' + Math.random().toString(36).substring(2, 9),
          name: f.name,
          size: f.size,
          status: 'error',
          errorMsg: (err as Error).message,
          sheetNames: [],
          activeSheetName: '',
          headerRowIndex: 0,
          dataStartRowIndex: 1,
          totalRows: 0,
          enabled: false,
        });
      }
    }

    setFiles((prev) => {
      const updated = [...prev, ...parsedList];
      if (!inspectedFileId && parsedList.length > 0) {
        setInspectedFileId(parsedList[0].id);
      }
      return updated;
    });

    setIsBatchReading(false);
    setBatchReadProgress('');
    setResults(null);
  };

  // Remove a single file
  const handleRemoveFile = (fileId: string) => {
    setFiles((prev) => {
      const updated = prev.filter((f) => f.id !== fileId);
      if (inspectedFileId === fileId) {
        setInspectedFileId(updated[0]?.id || null);
      }
      return updated;
    });
    setResults(null);
  };

  // Clear all files
  const handleClearAllFiles = () => {
    setFiles([]);
    setInspectedFileId(null);
    setResults(null);
  };

  // Toggle single file enable/disable
  const handleToggleFileEnabled = (fileId: string) => {
    setFiles((prev) =>
      prev.map((f) => (f.id === fileId ? { ...f, enabled: !f.enabled } : f))
    );
    setResults(null);
  };

  // Select/Deselect all files
  const handleSelectAllFiles = (selectAll: boolean) => {
    setFiles((prev) => prev.map((f) => ({ ...f, enabled: selectAll })));
    setResults(null);
  };

  // Change active sheet for a file
  const handleSelectFileSheet = (fileId: string, sheetName: string) => {
    setFiles((prev) =>
      prev.map((f) => {
        if (f.id !== fileId || !f.buffer) return f;
        const parsed = parseExcelWorkbook(f.buffer, f.headerRowIndex, f.dataStartRowIndex);
        const nextSheetData = parsed.getSheetData(sheetName, f.headerRowIndex, f.dataStartRowIndex);
        return {
          ...f,
          activeSheetName: sheetName,
          currentSheetData: nextSheetData,
          totalRows: nextSheetData.totalRows,
        };
      })
    );
    setResults(null);
  };

  // Update header row for a specific file
  const handleUpdateFileHeaderRow = (fileId: string, rowIndex: number) => {
    setFiles((prev) =>
      prev.map((f) => {
        if (f.id !== fileId || !f.buffer) return f;
        const nextDataRow = Math.max(rowIndex + 1, f.dataStartRowIndex);
        const parsed = parseExcelWorkbook(f.buffer, rowIndex, nextDataRow);
        const nextSheetData = parsed.getSheetData(f.activeSheetName, rowIndex, nextDataRow);
        return {
          ...f,
          headerRowIndex: rowIndex,
          dataStartRowIndex: nextDataRow,
          currentSheetData: nextSheetData,
        };
      })
    );
    setResults(null);
  };

  // Update data start row for a specific file
  const handleUpdateFileDataRow = (fileId: string, rowIndex: number) => {
    setFiles((prev) =>
      prev.map((f) => {
        if (f.id !== fileId || !f.buffer) return f;
        const parsed = parseExcelWorkbook(f.buffer, f.headerRowIndex, rowIndex);
        const nextSheetData = parsed.getSheetData(f.activeSheetName, f.headerRowIndex, rowIndex);
        return {
          ...f,
          dataStartRowIndex: rowIndex,
          currentSheetData: nextSheetData,
        };
      })
    );
    setResults(null);
  };

  // Apply row configuration from inspected file to ALL uploaded files
  const handleApplyRowsToAllFiles = (headerRow: number, dataRow: number) => {
    setIsBatchReading(true);
    setBatchReadProgress('Đang đồng bộ dòng tiêu đề cho tất cả các tệp...');

    setTimeout(() => {
      setFiles((prev) =>
        prev.map((f) => {
          if (!f.buffer) return f;
          try {
            const parsed = parseExcelWorkbook(f.buffer, headerRow, dataRow);
            const nextSheetData = parsed.getSheetData(f.activeSheetName, headerRow, dataRow);
            return {
              ...f,
              headerRowIndex: headerRow,
              dataStartRowIndex: dataRow,
              currentSheetData: nextSheetData,
              totalRows: nextSheetData.totalRows,
            };
          } catch {
            return f;
          }
        })
      );
      setIsBatchReading(false);
      setBatchReadProgress('');
      setResults(null);
    }, 50);
  };

  // Change reference column
  const handleChangeReferenceColumn = (letter: string) => {
    setReferenceColumnLetter(letter);
    const colName = primarySheetData?.columns.find((c) => c.letter === letter)?.name || '';
    setReferenceColumnName(colName);
    if (!selectedColumnLetters.includes(letter)) {
      setSelectedColumnLetters([letter, ...selectedColumnLetters]);
    }
    setResults(null);
  };

  // Execute extraction across multiple files
  const handleExecuteExtraction = () => {
    const enabledFiles = files.filter((f) => f.enabled && f.currentSheetData);
    if (enabledFiles.length === 0) {
      alert('Vui lòng bật ít nhất 1 tệp Excel để trích xuất.');
      return;
    }

    const parsedCodes = rawCodeInput
      .split(/[\r\n,;\t]+/)
      .map((c) => c.trim())
      .filter((c) => c.length > 0);

    if (parsedCodes.length === 0) {
      alert('Vui lòng nhập ít nhất 1 mã để tra cứu.');
      return;
    }

    setIsExecuting(true);

    const activeConfig: ExtractionConfig = {
      ...config,
      headerRowIndex: primaryFile?.headerRowIndex ?? 0,
      dataStartRowIndex: primaryFile?.dataStartRowIndex ?? 2,
      referenceColumnLetter,
      referenceColumnName: referenceColumnName || primarySheetData?.columns.find((c) => c.letter === referenceColumnLetter)?.name,
      matchBy,
      selectedColumns: selectedColumnLetters.map((letter) => {
        const colMeta = primarySheetData?.columns.find((c) => c.letter === letter);
        const originalName = colMeta?.name || `Cột ${letter}`;
        const outputName = columnAliases[letter]?.trim() || originalName;
        return { letter, originalName, outputName };
      }),
      includeSourceFileName,
      includeSheetName,
    };
    setConfig(activeConfig);

    setTimeout(() => {
      try {
        const res = extractDataFromMultipleFiles(enabledFiles, parsedCodes, activeConfig);
        setResults(res);
      } catch (err) {
        alert('Lỗi trong quá trình trích xuất: ' + (err as Error).message);
      } finally {
        setIsExecuting(false);
      }
    }, 60);
  };

  // Load 3 sample workbooks matching user's textile yarn plan
  const handleLoadMultiSample = async () => {
    setIsBatchReading(true);
    setBatchReadProgress('Đang nạp 3 tệp mẫu kế hoạch sợi (Tháng 8, Tháng 9, Tháng 10)...');

    const sampleWorkbooks = createMultiSampleWorkbooks();
    const parsedList: UploadedFileInfo[] = [];

    for (const item of sampleWorkbooks) {
      const parsed = await parseSingleFile(item, 0, 2);
      parsedList.push(parsed);
    }

    setFiles(parsedList);
    setInspectedFileId(parsedList[1].id); // Inspect Month 9

    setReferenceColumnLetter('C');
    setReferenceColumnName('PR/ INDENT');
    setMatchBy('letter');
    setSelectedColumnLetters(['C', 'D', 'F', 'G', 'I', 'K', 'M', 'P']);
    setColumnAliases({
      C: 'PR/ INDENT',
      D: 'YARN KIND (Loại sợi)',
      F: 'YARN DEMAND (KG)',
      G: 'STATUS IN W/H (KG)',
      I: 'USED (KG)',
      K: 'BUYER (Khách hàng)',
      M: 'P/O (Đơn mua hàng)',
      P: 'LOT (Số Lot sợi)',
    });
    setIncludeSourceFileName(true);
    setIncludeSheetName(false);
    setRawCodeInput(SAMPLE_QUERY_CODES.join('\n'));

    const activeConfig: ExtractionConfig = {
      headerRowIndex: 0,
      dataStartRowIndex: 2,
      referenceColumnLetter: 'C',
      referenceColumnName: 'PR/ INDENT',
      matchBy: 'letter',
      selectedColumns: ['C', 'D', 'F', 'G', 'I', 'K', 'M', 'P'].map((letter) => {
        const colMeta = parsedList[1].currentSheetData?.columns.find((c) => c.letter === letter);
        const originalName = colMeta?.name || `Cột ${letter}`;
        return {
          letter,
          originalName,
          outputName: originalName,
        };
      }),
      matchMode: 'case_insensitive_trim',
      matchStrategy: 'all_rows',
      includeUnmatchedSummary: true,
      includeSourceFileName: true,
      includeSheetName: false,
    };
    setConfig(activeConfig);

    // Auto-extract
    const res = extractDataFromMultipleFiles(parsedList, SAMPLE_QUERY_CODES, activeConfig);
    setResults(res);

    setIsBatchReading(false);
    setBatchReadProgress('');
  };

  // Presets handling
  const handleSavePreset = (name: string) => {
    const newPreset: ExtractionPreset = {
      id: 'preset-' + Date.now(),
      name,
      createdAt: new Date().toISOString(),
      referenceColumnLetter,
      referenceColumnName,
      matchBy,
      headerRowIndex: primaryFile?.headerRowIndex ?? 0,
      dataStartRowIndex: primaryFile?.dataStartRowIndex ?? 2,
      selectedColumnLetters,
      columnAliases,
      matchMode: config.matchMode,
      matchStrategy: config.matchStrategy,
    };
    setPresets((prev) => [newPreset, ...prev]);
  };

  const handleApplyPreset = (preset: ExtractionPreset) => {
    setReferenceColumnLetter(preset.referenceColumnLetter);
    if (preset.referenceColumnName) setReferenceColumnName(preset.referenceColumnName);
    if (preset.matchBy) setMatchBy(preset.matchBy);
    setSelectedColumnLetters(preset.selectedColumnLetters);
    setColumnAliases(preset.columnAliases || {});
    setConfig((prev) => ({
      ...prev,
      referenceColumnLetter: preset.referenceColumnLetter,
      referenceColumnName: preset.referenceColumnName,
      matchBy: preset.matchBy || 'letter',
      matchMode: preset.matchMode,
      matchStrategy: preset.matchStrategy,
    }));
    if (preset.headerRowIndex !== undefined && preset.dataStartRowIndex !== undefined) {
      handleApplyRowsToAllFiles(preset.headerRowIndex, preset.dataStartRowIndex);
    }
  };

  const handleDeletePreset = (id: string) => {
    setPresets((prev) => prev.filter((p) => p.id !== id));
  };

  const handleReset = () => {
    setFiles([]);
    setInspectedFileId(null);
    setResults(null);
    setRawCodeInput('');
  };

  // Load sample on first mount
  useEffect(() => {
    handleLoadMultiSample();
  }, []);

  return (
    <div className="min-h-screen bg-neutral-100/70 text-neutral-900 flex flex-col font-sans">
      <TopBar
        onLoadSample={handleLoadMultiSample}
        onOpenGuide={() => setIsGuideOpen(true)}
        onReset={handleReset}
        hasData={files.length > 0}
      />

      <main className="flex-1 mx-auto w-full max-w-7xl px-4 sm:px-6 lg:px-8 py-6 space-y-6">
        {/* Step 1: Upload / Manage Multiple Files */}
        <FileUploader
          files={files}
          onUploadMultipleFiles={handleUploadMultipleFiles}
          onRemoveFile={handleRemoveFile}
          onClearAllFiles={handleClearAllFiles}
          onToggleFileEnabled={handleToggleFileEnabled}
          onSelectAllFiles={handleSelectAllFiles}
          onSelectFileSheet={handleSelectFileSheet}
          onUpdateFileHeaderRow={handleUpdateFileHeaderRow}
          onUpdateFileDataRow={handleUpdateFileDataRow}
          onApplyRowsToAllFiles={handleApplyRowsToAllFiles}
          onLoadMultiSample={handleLoadMultiSample}
          inspectedFileId={inspectedFileId}
          onSetInspectedFileId={setInspectedFileId}
          isBatchReading={isBatchReading}
          batchReadProgress={batchReadProgress}
        />

        {/* Step 2 & 3: Only when at least 1 file with sheetData is ready */}
        {primarySheetData && (
          <>
            {/* Step 2: Configure Reference & Output Columns */}
            <ConfigStep
              columns={primarySheetData.columns}
              referenceColumnLetter={referenceColumnLetter}
              onChangeReferenceColumn={handleChangeReferenceColumn}
              referenceColumnName={referenceColumnName}
              selectedColumnLetters={selectedColumnLetters}
              onChangeSelectedColumns={setSelectedColumnLetters}
              columnAliases={columnAliases}
              onChangeColumnAlias={(letter, alias) =>
                setColumnAliases((prev) => ({ ...prev, [letter]: alias }))
              }
              onReorderColumns={setSelectedColumnLetters}
              matchBy={matchBy}
              onChangeMatchBy={setMatchBy}
              includeSourceFileName={includeSourceFileName}
              onChangeIncludeSourceFileName={setIncludeSourceFileName}
              includeSheetName={includeSheetName}
              onChangeIncludeSheetName={setIncludeSheetName}
              presets={presets}
              onSavePreset={handleSavePreset}
              onApplyPreset={handleApplyPreset}
              onDeletePreset={handleDeletePreset}
              filesCount={files.length}
              enabledFilesCount={files.filter((f) => f.enabled).length}
            />

            {/* Step 3: Enter Search Codes & Matching Rules */}
            <CodeInputStep
              sheetData={primarySheetData}
              referenceColumnLetter={referenceColumnLetter}
              rawCodeInput={rawCodeInput}
              onChangeRawCodeInput={setRawCodeInput}
              config={config}
              onChangeConfig={(changes) => setConfig((prev) => ({ ...prev, ...changes }))}
              onExecuteExtraction={handleExecuteExtraction}
              isExecuting={isExecuting}
            />
          </>
        )}

        {/* Step 4: Results Table & Export */}
        {results && primarySheetData && (
          <ResultsTable
            results={results}
            config={config}
            sheetData={primarySheetData}
            sourceFileName={primaryFile?.name || null}
            multiFilesCount={files.filter((f) => f.enabled).length}
          />
        )}
      </main>

      <footer className="border-t border-neutral-200 bg-white py-4 mt-12 text-center text-xs text-neutral-500">
        <div className="mx-auto max-w-7xl px-4 flex flex-col sm:flex-row items-center justify-between gap-2">
          <span>Excel RefX · Xử lý đa tệp Excel dung lượng lớn &amp; Trích xuất dữ liệu tự động</span>
          <span>Công nghệ Dense Mode tiết kiệm 80% RAM · 100% Cục bộ &amp; Bảo mật trên trình duyệt</span>
        </div>
      </footer>

      <GuideModal isOpen={isGuideOpen} onClose={() => setIsGuideOpen(false)} />
    </div>
  );
}
