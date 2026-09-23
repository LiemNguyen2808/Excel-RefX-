import React, { useState, useMemo } from 'react';
import { ExtractionResult, ExtractionConfig, SheetData } from '../types/excel';
import {
  Download,
  FileSpreadsheet,
  FileText,
  Copy,
  Check,
  Search,
  ChevronLeft,
  ChevronRight,
  AlertTriangle,
  ArrowUpDown,
  Layers,
  Filter
} from 'lucide-react';
import {
  exportResultsToExcel,
  exportResultsToCsv,
  getTsvDataForClipboard,
} from '../utils/excelParser';

interface ResultsTableProps {
  results: ExtractionResult;
  config: ExtractionConfig;
  sheetData: SheetData;
  sourceFileName: string | null;
  multiFilesCount?: number;
}

export const ResultsTable: React.FC<ResultsTableProps> = ({
  results,
  config,
  sheetData,
  sourceFileName,
  multiFilesCount = 1,
}) => {
  const [activeTab, setActiveTab] = useState<'matched' | 'unmatched' | 'stats' | 'source'>('matched');
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedFileFilter, setSelectedFileFilter] = useState<string>('all');
  const [copiedTsv, setCopiedTsv] = useState(false);
  const [copiedMissing, setCopiedMissing] = useState(false);
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(25);
  const [sortCol, setSortCol] = useState<string | null>(null);
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('asc');

  // Filtered rows for matched data
  const filteredRows = useMemo(() => {
    let rows = results.matchedRows;

    // Filter by specific file
    if (selectedFileFilter !== 'all') {
      rows = rows.filter((r) => r.sourceFileName === selectedFileFilter);
    }

    if (searchTerm.trim()) {
      const term = searchTerm.toLowerCase();
      rows = rows.filter((r) => {
        if (r.matchedCode.toLowerCase().includes(term)) return true;
        if (r.sourceFileName && r.sourceFileName.toLowerCase().includes(term)) return true;
        if (r.sourceSheetName && r.sourceSheetName.toLowerCase().includes(term)) return true;
        if (String(r.originalRowIndex).includes(term)) return true;
        for (const val of Object.values(r.data)) {
          if (String(val || '').toLowerCase().includes(term)) return true;
        }
        return false;
      });
    }

    if (sortCol) {
      rows = [...rows].sort((a, b) => {
        let valA: unknown =
          sortCol === '__row__'
            ? a.originalRowIndex
            : sortCol === '__code__'
            ? a.matchedCode
            : sortCol === '__file__'
            ? a.sourceFileName
            : a.data[sortCol];
        let valB: unknown =
          sortCol === '__row__'
            ? b.originalRowIndex
            : sortCol === '__code__'
            ? b.matchedCode
            : sortCol === '__file__'
            ? b.sourceFileName
            : b.data[sortCol];

        if (typeof valA === 'number' && typeof valB === 'number') {
          return sortDirection === 'asc' ? valA - valB : valB - valA;
        }
        const strA = String(valA || '');
        const strB = String(valB || '');
        return sortDirection === 'asc' ? strA.localeCompare(strB) : strB.localeCompare(strA);
      });
    }

    return rows;
  }, [results.matchedRows, selectedFileFilter, searchTerm, sortCol, sortDirection]);

  // Pagination calculation
  const totalPages = Math.max(1, Math.ceil(filteredRows.length / pageSize));
  const paginatedRows = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredRows.slice(start, start + pageSize);
  }, [filteredRows, currentPage, pageSize]);

  // Headers list for matched table
  const tableHeaders = useMemo(() => {
    if (results.matchedRows.length === 0) {
      return config.selectedColumns.map((c) => c.outputName || c.originalName);
    }
    return Object.keys(results.matchedRows[0].data);
  }, [results.matchedRows, config.selectedColumns]);

  const handleSort = (colName: string) => {
    if (sortCol === colName) {
      setSortDirection(sortDirection === 'asc' ? 'desc' : 'asc');
    } else {
      setSortCol(colName);
      setSortDirection('asc');
    }
  };

  const handleCopyTsv = async () => {
    const tsv = getTsvDataForClipboard(results, config);
    if (!tsv) return;
    try {
      await navigator.clipboard.writeText(tsv);
      setCopiedTsv(true);
      setTimeout(() => setCopiedTsv(false), 2000);
    } catch {
      // Fallback
    }
  };

  const handleCopyMissingCodes = async () => {
    if (results.unmatchedCodes.length === 0) return;
    try {
      await navigator.clipboard.writeText(results.unmatchedCodes.join('\n'));
      setCopiedMissing(true);
      setTimeout(() => setCopiedMissing(false), 2000);
    } catch {
      // Fallback
    }
  };

  const outputFileName = multiFilesCount > 1
    ? `Trich_Xuat_Tong_Hop_${multiFilesCount}_Tep.xlsx`
    : sourceFileName
    ? `Trich_Xuat_${sourceFileName.replace(/\.[^/.]+$/, '')}.xlsx`
    : 'Ket_Qua_Trich_Xuat.xlsx';

  return (
    <div className="bg-white border border-neutral-200 rounded-lg p-5 shadow-xs space-y-5">
      {/* Header and Stats */}
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 pb-4 border-b border-neutral-100">
        <div>
          <h2 className="text-base font-semibold text-neutral-900 flex items-center gap-2">
            <span>4. Kết Quả Trích Xuất &amp; Xuất File Mới</span>
            {multiFilesCount > 1 && (
              <span className="text-xs font-normal text-emerald-800 bg-emerald-100 px-2 py-0.5 rounded font-mono">
                Đã gộp từ {multiFilesCount} tệp
              </span>
            )}
          </h2>
          <div className="flex flex-wrap items-center gap-2 text-xs text-neutral-500 mt-1">
            <span>Tra cứu theo: <strong className="text-neutral-800">Cột {config.referenceColumnLetter} {config.referenceColumnName ? `(${config.referenceColumnName})` : ''}</strong></span>
            <span aria-hidden="true">·</span>
            <span>Tổng số cột xuất: <strong className="text-neutral-800">{config.selectedColumns.length}</strong></span>
            <span aria-hidden="true">·</span>
            <span>
              Quét <strong className="font-mono tabular-nums text-neutral-800">{results.totalRowsScanned.toLocaleString()}</strong> dòng qua <strong className="font-mono text-neutral-800">{results.totalFilesProcessed || 1}</strong> tệp trong <strong className="font-mono tabular-nums text-neutral-800">{results.durationMs}ms</strong>
            </span>
          </div>
        </div>

        {/* Action Export Buttons */}
        <div className="flex flex-wrap items-center gap-2">
          <button
            type="button"
            onClick={handleCopyTsv}
            disabled={results.matchedRows.length === 0}
            className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-medium text-neutral-700 bg-neutral-100 hover:bg-neutral-200 rounded-md transition-colors whitespace-nowrap disabled:opacity-40"
            title="Sao chép dữ liệu dạng bảng để dán trực tiếp vào Excel đang mở"
          >
            {copiedTsv ? <Check className="h-4 w-4 text-emerald-600" /> : <Copy className="h-4 w-4 text-neutral-500" />}
            <span>{copiedTsv ? 'Đã sao chép!' : 'Sao chép để dán Excel'}</span>
          </button>

          <button
            type="button"
            onClick={() => exportResultsToCsv(results, config, outputFileName.replace('.xlsx', '.csv'))}
            disabled={results.matchedRows.length === 0}
            className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-medium text-neutral-700 hover:text-neutral-900 border border-neutral-300 hover:border-neutral-400 rounded-md transition-colors whitespace-nowrap disabled:opacity-40"
          >
            <FileText className="h-4 w-4 text-neutral-500" />
            <span>Xuất CSV</span>
          </button>

          <button
            type="button"
            onClick={() => exportResultsToExcel(results, config, outputFileName)}
            disabled={results.matchedRows.length === 0}
            className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-emerald-600 hover:bg-emerald-700 rounded-md shadow-xs transition-colors whitespace-nowrap disabled:opacity-40"
          >
            <Download className="h-4 w-4" />
            <span>Lưu file Excel mới ({multiFilesCount > 1 ? 'Tổng hợp đa tệp' : '.xlsx'})</span>
          </button>
        </div>
      </div>

      {/* KPI Cards / Statistical Summary */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <div className="p-3 bg-neutral-50 border border-neutral-200 rounded-md">
          <span className="text-[11px] text-neutral-500 block uppercase font-medium">Số mã cần tìm</span>
          <span className="text-xl font-bold font-mono tabular-nums text-neutral-900 mt-0.5 block">
            {results.matchedCodes.length + results.unmatchedCodes.length}
          </span>
          <span className="text-[11px] text-neutral-500">Mã yêu cầu</span>
        </div>

        <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-md">
          <span className="text-[11px] text-emerald-700 block uppercase font-medium">Mã tìm thấy</span>
          <span className="text-xl font-bold font-mono tabular-nums text-emerald-900 mt-0.5 block">
            {results.matchedCodes.length}
          </span>
          <span className="text-[11px] text-emerald-700">Khớp thành công</span>
        </div>

        <div className={`p-3 rounded-md border ${results.unmatchedCodes.length > 0 ? 'bg-amber-50 border-amber-200' : 'bg-neutral-50 border-neutral-200'}`}>
          <span className={`text-[11px] block uppercase font-medium ${results.unmatchedCodes.length > 0 ? 'text-amber-800' : 'text-neutral-500'}`}>
            Mã không tìm thấy
          </span>
          <span className={`text-xl font-bold font-mono tabular-nums mt-0.5 block ${results.unmatchedCodes.length > 0 ? 'text-amber-900' : 'text-neutral-900'}`}>
            {results.unmatchedCodes.length}
          </span>
          <span className={`text-[11px] ${results.unmatchedCodes.length > 0 ? 'text-amber-700' : 'text-neutral-500'}`}>
            Chưa có trong các file
          </span>
        </div>

        <div className="p-3 bg-blue-50 border border-blue-200 rounded-md">
          <span className="text-[11px] text-blue-700 block uppercase font-medium">Dòng dữ liệu lấy ra</span>
          <span className="text-xl font-bold font-mono tabular-nums text-blue-900 mt-0.5 block">
            {results.matchedRows.length}
          </span>
          <span className="text-[11px] text-blue-700">
            {multiFilesCount > 1 ? `Gộp từ ${multiFilesCount} tệp` : 'Dòng trong file mới'}
          </span>
        </div>
      </div>

      {/* Multi-file Distribution breakdown (if multi-file) */}
      {results.fileStats && results.fileStats.length > 1 && (
        <div className="p-3 bg-neutral-50 border border-neutral-200 rounded-md space-y-2">
          <div className="flex items-center justify-between text-xs">
            <span className="font-semibold text-neutral-800 flex items-center gap-1.5">
              <Layers className="h-3.5 w-3.5 text-emerald-700" />
              Phân bố dữ liệu trích xuất theo từng tệp Excel nguồn:
            </span>
            <span className="text-neutral-500 text-[11px]">
              Click vào tệp bên dưới để lọc nhanh bảng
            </span>
          </div>

          <div className="flex flex-wrap gap-2">
            <button
              type="button"
              onClick={() => {
                setSelectedFileFilter('all');
                setCurrentPage(1);
              }}
              className={`px-2.5 py-1 rounded text-xs transition-colors border ${
                selectedFileFilter === 'all'
                  ? 'bg-neutral-900 text-white border-neutral-900 font-medium'
                  : 'bg-white text-neutral-700 border-neutral-300 hover:bg-neutral-100'
              }`}
            >
              Tất cả tệp ({results.matchedRows.length} dòng)
            </button>

            {results.fileStats.map((f) => (
              <button
                key={f.fileId}
                type="button"
                onClick={() => {
                  setSelectedFileFilter(f.fileName);
                  setCurrentPage(1);
                }}
                className={`px-2.5 py-1 rounded text-xs transition-colors border flex items-center gap-1.5 ${
                  selectedFileFilter === f.fileName
                    ? 'bg-emerald-600 text-white border-emerald-600 font-medium'
                    : 'bg-white text-neutral-700 border-neutral-300 hover:bg-neutral-100'
                }`}
              >
                <span className="truncate max-w-44">{f.fileName}</span>
                <span className={`px-1 py-0.2 rounded text-[10px] font-mono ${selectedFileFilter === f.fileName ? 'bg-emerald-700 text-white' : 'bg-neutral-100 text-neutral-800'}`}>
                  {f.matchedRowCount} dòng
                </span>
              </button>
            ))}
          </div>
        </div>
      )}

      {/* Tabs and Controls */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pt-2">
        <div className="flex items-center gap-1 border-b border-neutral-200 sm:border-b-0 pb-1 sm:pb-0">
          <button
            type="button"
            onClick={() => setActiveTab('matched')}
            className={`px-3 py-1.5 text-xs font-medium border-b-2 transition-colors ${
              activeTab === 'matched'
                ? 'border-emerald-600 text-emerald-900 font-semibold'
                : 'border-transparent text-neutral-500 hover:text-neutral-800'
            }`}
          >
            Dữ liệu trích xuất ({results.matchedRows.length} dòng)
          </button>

          {results.unmatchedCodes.length > 0 && (
            <button
              type="button"
              onClick={() => setActiveTab('unmatched')}
              className={`px-3 py-1.5 text-xs font-medium border-b-2 transition-colors flex items-center gap-1 ${
                activeTab === 'unmatched'
                  ? 'border-amber-600 text-amber-900 font-semibold'
                  : 'border-transparent text-amber-700 hover:text-amber-900'
              }`}
            >
              <AlertTriangle className="h-3 w-3 text-amber-600" />
              Mã không tìm thấy ({results.unmatchedCodes.length})
            </button>
          )}

          <button
            type="button"
            onClick={() => setActiveTab('source')}
            className={`px-3 py-1.5 text-xs font-medium border-b-2 transition-colors ${
              activeTab === 'source'
                ? 'border-neutral-900 text-neutral-900 font-semibold'
                : 'border-transparent text-neutral-500 hover:text-neutral-800'
            }`}
          >
            Xem trước tệp nguồn ({sheetData.totalRows} dòng)
          </button>
        </div>

        {activeTab === 'matched' && (
          <div className="flex items-center gap-2">
            <div className="relative">
              <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 h-3.5 w-3.5 text-neutral-400" />
              <input
                type="text"
                value={searchTerm}
                onChange={(e) => {
                  setSearchTerm(e.target.value);
                  setCurrentPage(1);
                }}
                placeholder="Tìm nhanh trong kết quả..."
                className="pl-8 pr-2.5 py-1 text-xs bg-neutral-50 border border-neutral-300 rounded focus:bg-white focus:outline-none focus:ring-1 focus:ring-neutral-400 w-52"
              />
            </div>

            <select
              value={pageSize}
              onChange={(e) => {
                setPageSize(Number(e.target.value));
                setCurrentPage(1);
              }}
              className="text-xs bg-white border border-neutral-300 rounded px-2 py-1 text-neutral-700"
            >
              <option value={25}>25 dòng/trang</option>
              <option value={50}>50 dòng/trang</option>
              <option value={100}>100 dòng/trang</option>
              <option value={1000}>Tất cả</option>
            </select>
          </div>
        )}
      </div>

      {/* Tab 1: Matched Data Table */}
      {activeTab === 'matched' && (
        <div>
          {results.matchedRows.length === 0 ? (
            <div className="text-center py-12 border border-dashed border-neutral-300 rounded-lg">
              <FileSpreadsheet className="h-10 w-10 text-neutral-300 mx-auto mb-2" />
              <div className="text-sm font-medium text-neutral-800">Không tìm thấy dòng dữ liệu nào</div>
              <p className="text-xs text-neutral-500 mt-1 max-w-md mx-auto">
                Không có dòng nào trong các file nguồn khớp với danh sách mã đã nhập. Bạn hãy kiểm tra lại cột tham chiếu hoặc thử chọn chế độ &quot;Không phân biệt hoa/thường &amp; xóa khoảng trắng&quot;.
              </p>
            </div>
          ) : (
            <div className="border border-neutral-200 rounded-md overflow-hidden">
              <div className="overflow-x-auto max-h-[520px]">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-neutral-100 text-neutral-700 sticky top-0 border-b border-neutral-200 z-10 font-semibold select-none">
                    <tr>
                      {/* Source file name column if enabled or multi-file */}
                      {(config.includeSourceFileName || multiFilesCount > 1) && (
                        <th
                          onClick={() => handleSort('__file__')}
                          className="p-2.5 border-r border-neutral-200 min-w-36 cursor-pointer hover:bg-neutral-200/70"
                        >
                          <div className="flex items-center justify-between gap-1">
                            <span>Tệp nguồn</span>
                            <ArrowUpDown className="h-3 w-3 text-neutral-400" />
                          </div>
                        </th>
                      )}

                      <th
                        onClick={() => handleSort('__row__')}
                        className="p-2.5 border-r border-neutral-200 w-24 text-center cursor-pointer hover:bg-neutral-200/70"
                      >
                        <div className="flex items-center justify-center gap-1">
                          <span>Dòng Excel</span>
                          <ArrowUpDown className="h-3 w-3 text-neutral-400" />
                        </div>
                      </th>

                      <th
                        onClick={() => handleSort('__code__')}
                        className="p-2.5 border-r border-neutral-200 min-w-36 cursor-pointer hover:bg-neutral-200/70 text-emerald-900 bg-emerald-100/60"
                      >
                        <div className="flex items-center gap-1">
                          <span>Mã tham chiếu</span>
                          <ArrowUpDown className="h-3 w-3 text-emerald-700" />
                        </div>
                      </th>

                      {tableHeaders.map((header) => (
                        <th
                          key={header}
                          onClick={() => handleSort(header)}
                          className="p-2.5 border-r border-neutral-200 min-w-32 whitespace-nowrap cursor-pointer hover:bg-neutral-200/70"
                        >
                          <div className="flex items-center justify-between gap-1">
                            <span className="truncate">{header}</span>
                            <ArrowUpDown className="h-3 w-3 text-neutral-400 shrink-0" />
                          </div>
                        </th>
                      ))}
                    </tr>
                  </thead>

                  <tbody className="divide-y divide-neutral-100 bg-white">
                    {paginatedRows.map((row, idx) => (
                      <tr key={idx} className="hover:bg-neutral-50/80 transition-colors">
                        {(config.includeSourceFileName || multiFilesCount > 1) && (
                          <td className="p-2 border-r border-neutral-200 font-medium text-neutral-700 whitespace-nowrap">
                            <div className="flex items-center gap-1 max-w-44 truncate" title={row.sourceFileName}>
                              <FileSpreadsheet className="h-3.5 w-3.5 text-neutral-400 shrink-0" />
                              <span className="truncate">{row.sourceFileName}</span>
                            </div>
                          </td>
                        )}

                        <td className="p-2 border-r border-neutral-200 text-center font-mono text-[11px] text-neutral-400">
                          #{row.originalRowIndex}
                        </td>

                        <td className="p-2 border-r border-neutral-200 font-mono font-semibold text-emerald-950 bg-emerald-50/30">
                          {row.matchedCode}
                        </td>

                        {tableHeaders.map((h) => {
                          const val = row.data[h];
                          const isNumber = typeof val === 'number';
                          return (
                            <td
                              key={h}
                              className={`p-2 border-r border-neutral-200 whitespace-nowrap text-neutral-800 ${
                                isNumber ? 'font-mono text-right tabular-nums' : ''
                              }`}
                            >
                              {val !== undefined && val !== null ? String(val) : ''}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {/* Pagination bar */}
              <div className="flex items-center justify-between p-3 bg-neutral-50 border-t border-neutral-200 text-xs text-neutral-600">
                <div>
                  Hiển thị {(currentPage - 1) * pageSize + 1} -{' '}
                  {Math.min(currentPage * pageSize, filteredRows.length)} trên tổng số{' '}
                  <strong className="font-mono tabular-nums text-neutral-900">{filteredRows.length}</strong> dòng
                  {selectedFileFilter !== 'all' && (
                    <span className="text-emerald-700 font-medium ml-2">
                      (Đang lọc riêng tệp: {selectedFileFilter})
                    </span>
                  )}
                </div>

                <div className="flex items-center gap-1.5">
                  <button
                    type="button"
                    onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                    disabled={currentPage === 1}
                    className="p-1 rounded border border-neutral-300 bg-white hover:bg-neutral-100 disabled:opacity-30"
                  >
                    <ChevronLeft className="h-4 w-4" />
                  </button>
                  <span className="px-2 font-mono">
                    Trang {currentPage} / {totalPages}
                  </span>
                  <button
                    type="button"
                    onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                    disabled={currentPage === totalPages}
                    className="p-1 rounded border border-neutral-300 bg-white hover:bg-neutral-100 disabled:opacity-30"
                  >
                    <ChevronRight className="h-4 w-4" />
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Tab 2: Unmatched Codes */}
      {activeTab === 'unmatched' && (
        <div className="p-4 bg-amber-50/50 border border-amber-200 rounded-lg space-y-3">
          <div className="flex items-center justify-between">
            <div>
              <h3 className="text-xs font-bold text-amber-950 uppercase tracking-wide">
                Danh sách {results.unmatchedCodes.length} mã không tìm thấy trong bất kỳ tệp nguồn nào:
              </h3>
              <p className="text-xs text-amber-800 mt-0.5">
                Các mã này được nhập vào tra cứu nhưng không có dòng dữ liệu nào tương ứng trong tất cả các tệp Excel đã nạp.
              </p>
            </div>

            <button
              type="button"
              onClick={handleCopyMissingCodes}
              className="inline-flex items-center gap-1 px-3 py-1.5 text-xs font-medium text-amber-900 bg-amber-100 hover:bg-amber-200 rounded transition-colors"
            >
              {copiedMissing ? <Check className="h-3.5 w-3.5" /> : <Copy className="h-3.5 w-3.5" />}
              <span>{copiedMissing ? 'Đã sao chép!' : 'Sao chép danh sách mã thiếu'}</span>
            </button>
          </div>

          <div className="border border-amber-200 rounded-md bg-white p-3 max-h-60 overflow-y-auto font-mono text-xs text-neutral-800 space-y-1">
            {results.unmatchedCodes.map((code, idx) => (
              <div key={idx} className="flex items-center justify-between py-1 border-b border-neutral-100 last:border-0">
                <span className="text-neutral-500 w-8">{idx + 1}.</span>
                <span className="font-semibold text-rose-700 flex-1">{code}</span>
                <span className="text-[11px] text-neutral-400">Không tìm thấy</span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Tab 3: Full Source Sheet Preview */}
      {activeTab === 'source' && (
        <div className="border border-neutral-200 rounded-md overflow-hidden">
          <div className="p-2.5 bg-neutral-100 border-b border-neutral-200 text-xs text-neutral-600 flex items-center justify-between">
            <span>
              Xem tệp nguồn: <strong>{sourceFileName || sheetData.sheetName}</strong> ({sheetData.totalRows} dòng x {sheetData.columns.length} cột)
            </span>
            <span className="text-neutral-400 text-[11px]">
              Hiển thị 50 dòng đầu tiên
            </span>
          </div>

          <div className="overflow-x-auto max-h-[450px]">
            <table className="w-full text-left text-xs border-collapse">
              <thead className="bg-neutral-50 text-neutral-600 sticky top-0 border-b border-neutral-200">
                <tr>
                  <th className="p-2 border-r border-neutral-200 text-center w-16">#</th>
                  {sheetData.columns.map((c) => (
                    <th key={c.letter} className="p-2 border-r border-neutral-200 whitespace-nowrap min-w-28 font-medium">
                      <span className="font-mono text-neutral-400 font-bold mr-1">[{c.letter}]</span>
                      {c.name}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-neutral-100 bg-white">
                {sheetData.rawRows.slice(0, 50).map((row, rIdx) => {
                  const isHeader = rIdx === sheetData.headerRowIndex;
                  return (
                    <tr key={rIdx} className={isHeader ? 'bg-emerald-50 font-semibold' : 'hover:bg-neutral-50'}>
                      <td className="p-2 border-r border-neutral-200 text-center font-mono text-neutral-400 text-[11px]">
                        {rIdx + 1}
                      </td>
                      {sheetData.columns.map((c) => (
                        <td key={c.letter} className="p-2 border-r border-neutral-200 whitespace-nowrap text-neutral-700 truncate max-w-xs">
                          {row && row[c.index] !== undefined && row[c.index] !== null ? String(row[c.index]) : ''}
                        </td>
                      ))}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
};
