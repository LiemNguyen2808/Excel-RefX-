import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { SheetData, ExtractionConfig } from '../types/excel';
import { FileCode, Upload, ListFilter, Play, Sparkles, AlertCircle, Copy, Check } from 'lucide-react';
import { SAMPLE_QUERY_CODES } from '../utils/sampleData';

interface CodeInputStepProps {
  sheetData: SheetData | null;
  referenceColumnLetter: string;
  rawCodeInput: string;
  onChangeRawCodeInput: (val: string) => void;
  config: ExtractionConfig;
  onChangeConfig: (newConfig: Partial<ExtractionConfig>) => void;
  onExecuteExtraction: () => void;
  isExecuting: boolean;
}

export const CodeInputStep: React.FC<CodeInputStepProps> = ({
  sheetData,
  referenceColumnLetter,
  rawCodeInput,
  onChangeRawCodeInput,
  config,
  onChangeConfig,
  onExecuteExtraction,
  isExecuting,
}) => {
  const [activeTab, setActiveTab] = useState<'paste' | 'upload' | 'select'>('paste');
  const [copiedSample, setCopiedSample] = useState(false);
  const [uniqueValuesSearch, setUniqueValuesSearch] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Parse input codes into clean array
  const parsedCodes = rawCodeInput
    .split(/[\r\n,;\t]+/)
    .map((c) => c.trim())
    .filter((c) => c.length > 0);

  const uniqueCodesCount = new Set(parsedCodes).size;

  // Extract unique values from the source sheet's reference column for the "select" tab
  const getUniqueSourceValues = (): string[] => {
    if (!sheetData) return [];
    const refCol = sheetData.columns.find((c) => c.letter === referenceColumnLetter);
    if (!refCol) return [];
    const colIdx = refCol.index;
    const valuesSet = new Set<string>();

    const rows = sheetData.rawRows.slice(sheetData.dataStartRowIndex);
    for (const r of rows) {
      if (r && r[colIdx] !== null && r[colIdx] !== undefined) {
        const str = String(r[colIdx]).trim();
        if (str) valuesSet.add(str);
      }
    }
    return Array.from(valuesSet);
  };

  const uniqueSourceValues = getUniqueSourceValues();
  const filteredUniqueValues = uniqueSourceValues.filter((v) =>
    v.toLowerCase().includes(uniqueValuesSearch.toLowerCase())
  );

  // Handle uploading a secondary file containing lookup codes
  const handleUploadCodeFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = evt.target?.result as ArrayBuffer;
        const wb = XLSX.read(data, { type: 'array' });
        const firstSheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json<(string | number)[]>(firstSheet, { header: 1 });

        // Collect all non-empty strings/numbers from the first column or all columns
        const extracted: string[] = [];
        for (const row of json) {
          if (row && row.length > 0) {
            // Check first column
            const val = row[0];
            if (val !== undefined && val !== null && String(val).trim()) {
              extracted.push(String(val).trim());
            }
          }
        }

        if (extracted.length > 0) {
          onChangeRawCodeInput(extracted.join('\n'));
          setActiveTab('paste');
        } else {
          alert('Không tìm thấy mã nào trong tệp vừa tải lên.');
        }
      } catch (err) {
        alert('Lỗi đọc tệp mã: ' + (err as Error).message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleInsertSampleCodes = () => {
    onChangeRawCodeInput(SAMPLE_QUERY_CODES.join('\n'));
    setCopiedSample(true);
    setTimeout(() => setCopiedSample(false), 2000);
  };

  const handleToggleSelectUniqueValue = (val: string) => {
    const existing = new Set(parsedCodes);
    if (existing.has(val)) {
      existing.delete(val);
    } else {
      existing.add(val);
    }
    onChangeRawCodeInput(Array.from(existing).join('\n'));
  };

  return (
    <div className="bg-white border border-neutral-200 rounded-lg p-5 shadow-xs space-y-5">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-4 border-b border-neutral-100">
        <div>
          <h2 className="text-base font-semibold text-neutral-900">
            3. Nhập Danh Sách Mã Cần Tra Cứu &amp; Thiết Lập Khớp
          </h2>
          <p className="text-xs text-neutral-500 mt-0.5">
            Cung cấp các mã (ví dụ mã PR/INDENT, mã đơn, mã hàng...) để đối chiếu với Cột {referenceColumnLetter}.
          </p>
        </div>

        {/* Input Method Tabs */}
        <div className="flex items-center gap-1 p-1 bg-neutral-100 rounded-lg">
          <button
            type="button"
            onClick={() => setActiveTab('paste')}
            className={`px-3 py-1 text-xs font-medium rounded-md transition-colors ${
              activeTab === 'paste'
                ? 'bg-white text-neutral-900 shadow-xs'
                : 'text-neutral-600 hover:text-neutral-900'
            }`}
          >
            Dán danh sách mã
          </button>
          <button
            type="button"
            onClick={() => setActiveTab('upload')}
            className={`px-3 py-1 text-xs font-medium rounded-md transition-colors ${
              activeTab === 'upload'
                ? 'bg-white text-neutral-900 shadow-xs'
                : 'text-neutral-600 hover:text-neutral-900'
            }`}
          >
            Nạp từ tệp mã khác
          </button>
          <button
            type="button"
            onClick={() => setActiveTab('select')}
            className={`px-3 py-1 text-xs font-medium rounded-md transition-colors ${
              activeTab === 'select'
                ? 'bg-white text-neutral-900 shadow-xs'
                : 'text-neutral-600 hover:text-neutral-900'
            }`}
          >
            Chọn từ tệp nguồn ({uniqueSourceValues.length})
          </button>
        </div>
      </div>

      {/* Tab 1: Textarea Paste */}
      {activeTab === 'paste' && (
        <div className="space-y-2">
          <div className="flex items-center justify-between">
            <label className="text-xs font-medium text-neutral-700">
              Nhập hoặc dán danh sách mã (Mỗi mã 1 dòng hoặc cách nhau bởi dấu phẩy / Tab):
            </label>
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={handleInsertSampleCodes}
                className="text-xs text-emerald-700 hover:text-emerald-900 flex items-center gap-1 underline font-medium"
              >
                {copiedSample ? (
                  <>
                    <Check className="h-3 w-3" />
                    <span>Đã điền mã mẫu!</span>
                  </>
                ) : (
                  <>
                    <Sparkles className="h-3 w-3" />
                    <span>Điền mẫu: TCG-120895-001, ...</span>
                  </>
                )}
              </button>
              {parsedCodes.length > 0 && (
                <button
                  type="button"
                  onClick={() => onChangeRawCodeInput('')}
                  className="text-xs text-neutral-400 hover:text-neutral-600 underline"
                >
                  Xóa trắng
                </button>
              )}
            </div>
          </div>

          <div className="relative">
            <textarea
              rows={6}
              value={rawCodeInput}
              onChange={(e) => onChangeRawCodeInput(e.target.value)}
              placeholder={`Ví dụ:\nTCG-120895-001\nTCG-120899-001\nTCG-120904-001\nTCG-120906-001\n...`}
              className="w-full text-xs font-mono p-3 bg-neutral-50/70 border border-neutral-300 rounded-md focus:bg-white focus:outline-none focus:ring-2 focus:ring-emerald-500 text-neutral-900 placeholder:text-neutral-400"
            />
          </div>

          <div className="flex items-center justify-between text-xs text-neutral-500">
            <span>
              Đã nhận diện: <strong className="font-mono tabular-nums text-neutral-900 font-bold">{parsedCodes.length}</strong> mã 
              {parsedCodes.length !== uniqueCodesCount && (
                <span className="text-amber-600 ml-1">
                  ({uniqueCodesCount} mã không trùng lặp)
                </span>
              )}
            </span>
            <span className="text-[11px] text-neutral-400">
              Có thể copy nguyên 1 cột từ Excel và dán trực tiếp vào đây
            </span>
          </div>
        </div>
      )}

      {/* Tab 2: Upload File containing codes */}
      {activeTab === 'upload' && (
        <div className="p-6 border-2 border-dashed border-neutral-300 rounded-lg text-center bg-neutral-50/50">
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx, .xls, .csv"
            onChange={handleUploadCodeFile}
            className="hidden"
          />
          <Upload className="mx-auto h-8 w-8 text-neutral-400 mb-2" />
          <div className="text-xs font-medium text-neutral-800">
            Tải lên tệp Excel/CSV chứa danh sách mã cần tìm kiếm
          </div>
          <p className="text-[11px] text-neutral-500 mt-1 max-w-sm mx-auto">
            Hệ thống sẽ tự động đọc các mã ở cột đầu tiên của tệp và nạp vào danh sách tra cứu.
          </p>
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            className="mt-3 px-3 py-1.5 text-xs font-medium text-neutral-800 bg-white border border-neutral-300 hover:bg-neutral-100 rounded shadow-xs"
          >
            Chọn tệp chứa mã
          </button>
        </div>
      )}

      {/* Tab 3: Select from source values */}
      {activeTab === 'select' && (
        <div className="space-y-2">
          <div className="flex items-center justify-between">
            <span className="text-xs text-neutral-700 font-medium">
              Tìm và chọn trực tiếp từ các mã duy nhất đang có tại Cột {referenceColumnLetter}:
            </span>
            <input
              type="text"
              value={uniqueValuesSearch}
              onChange={(e) => setUniqueValuesSearch(e.target.value)}
              placeholder="Tìm kiếm mã trong bảng..."
              className="text-xs px-2.5 py-1 border border-neutral-300 rounded w-52"
            />
          </div>

          <div className="border border-neutral-200 rounded-md p-2 max-h-48 overflow-y-auto grid grid-cols-2 sm:grid-cols-3 gap-1.5 bg-neutral-50/50">
            {filteredUniqueValues.length === 0 ? (
              <div className="col-span-full text-center py-4 text-xs text-neutral-400">
                Không tìm thấy mã nào
              </div>
            ) : (
              filteredUniqueValues.map((val) => {
                const isSelected = parsedCodes.includes(val);
                return (
                  <button
                    key={val}
                    type="button"
                    onClick={() => handleToggleSelectUniqueValue(val)}
                    className={`flex items-center justify-between px-2.5 py-1 rounded text-xs text-left truncate transition-colors border ${
                      isSelected
                        ? 'bg-emerald-100/70 border-emerald-300 text-emerald-950 font-medium'
                        : 'bg-white border-neutral-200 text-neutral-700 hover:bg-neutral-100'
                    }`}
                  >
                    <span className="truncate font-mono">{val}</span>
                    {isSelected && <Check className="h-3 w-3 text-emerald-700 shrink-0 ml-1" />}
                  </button>
                );
              })
            )}
          </div>
        </div>
      )}

      {/* Matching Rules & Duplicate Strategy Options */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-3 p-3 bg-neutral-50 border border-neutral-200 rounded-md text-xs">
        <div>
          <label className="block font-medium text-neutral-800 mb-1">
            Quy tắc so khớp (Match Mode):
          </label>
          <select
            value={config.matchMode}
            onChange={(e) =>
              onChangeConfig({ matchMode: e.target.value as ExtractionConfig['matchMode'] })
            }
            className="w-full bg-white border border-neutral-300 rounded px-2.5 py-1.5 font-medium text-neutral-800"
          >
            <option value="case_insensitive_trim">
              Không phân biệt hoa/thường &amp; xóa khoảng trắng (Khuyên dùng)
            </option>
            <option value="exact">Chính xác tuyệt đối 100%</option>
            <option value="contains">Chứa chuỗi mã (Contains)</option>
          </select>
          <p className="text-[10px] text-neutral-500 mt-1">
            Tự động lọc các khoảng trắng thừa thường gặp khi copy từ Excel.
          </p>
        </div>

        <div>
          <label className="block font-medium text-neutral-800 mb-1">
            Xử lý khi 1 mã có nhiều dòng:
          </label>
          <select
            value={config.matchStrategy}
            onChange={(e) =>
              onChangeConfig({ matchStrategy: e.target.value as ExtractionConfig['matchStrategy'] })
            }
            className="w-full bg-white border border-neutral-300 rounded px-2.5 py-1.5 font-medium text-neutral-800"
          >
            <option value="all_rows">Lấy tất cả các dòng khớp (Đầy đủ)</option>
            <option value="first_row_only">Chỉ lấy dòng đầu tiên</option>
          </select>
          <p className="text-[10px] text-neutral-500 mt-1">
            Ví dụ 1 mã PR có nhiều dòng PO, LOT, ngày giao khác nhau.
          </p>
        </div>

        <div>
          <label className="block font-medium text-neutral-800 mb-1">
            Tùy chọn file xuất mới:
          </label>
          <div className="mt-2 space-y-1.5">
            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={config.includeUnmatchedSummary}
                onChange={(e) => onChangeConfig({ includeUnmatchedSummary: e.target.checked })}
                className="rounded border-neutral-300 text-emerald-600 focus:ring-emerald-500"
              />
              <span className="text-neutral-700">
                Thêm sheet &quot;Mã không tìm thấy&quot;
              </span>
            </label>
          </div>
          <p className="text-[10px] text-neutral-500 mt-1">
            Tiện đối soát các mã bị sai sót hoặc chưa có trong hệ thống.
          </p>
        </div>
      </div>

      {/* Primary Execution Action */}
      <div className="pt-2 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div className="text-xs text-neutral-500">
          {parsedCodes.length === 0 ? (
            <span className="text-neutral-500 flex items-center gap-1">
              <AlertCircle className="h-3.5 w-3.5 text-neutral-400" />
              Vui lòng nhập ít nhất 1 mã để bắt đầu tra cứu.
            </span>
          ) : (
            <span className="text-neutral-700">
              Sẵn sàng tra cứu <strong className="font-mono text-neutral-900">{parsedCodes.length}</strong> mã trên cột <strong className="text-emerald-700 font-bold">{referenceColumnLetter}</strong>
            </span>
          )}
        </div>

        <button
          type="button"
          onClick={onExecuteExtraction}
          disabled={parsedCodes.length === 0 || isExecuting}
          className="flex items-center justify-center gap-2 px-6 py-2.5 text-sm font-semibold text-white bg-emerald-600 hover:bg-emerald-700 disabled:opacity-40 disabled:hover:bg-emerald-600 rounded-lg shadow-sm transition-all"
        >
          <Play className="h-4 w-4 fill-white" />
          <span>{isExecuting ? 'Đang trích xuất...' : 'Trích Xuất & Lọc Dữ Liệu Tự Động'}</span>
        </button>
      </div>
    </div>
  );
};
