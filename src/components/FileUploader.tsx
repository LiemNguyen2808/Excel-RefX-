import React, { useRef, useState } from 'react';
import {
  UploadCloud,
  FileSpreadsheet,
  Trash2,
  Plus,
  Sparkles,
  Layers,
  ChevronDown,
  CheckCircle2,
  AlertCircle,
  FileText,
  CopyCheck
} from 'lucide-react';
import { SheetData, UploadedFileInfo } from '../types/excel';

interface FileUploaderProps {
  files: UploadedFileInfo[];
  onUploadMultipleFiles: (files: FileList | File[]) => void;
  onRemoveFile: (fileId: string) => void;
  onClearAllFiles: () => void;
  onToggleFileEnabled: (fileId: string) => void;
  onSelectAllFiles: (selectAll: boolean) => void;
  onSelectFileSheet: (fileId: string, sheetName: string) => void;
  onUpdateFileHeaderRow: (fileId: string, rowIndex: number) => void;
  onUpdateFileDataRow: (fileId: string, rowIndex: number) => void;
  onApplyRowsToAllFiles: (headerRow: number, dataRow: number) => void;
  onLoadMultiSample: () => void;
  inspectedFileId: string | null;
  onSetInspectedFileId: (fileId: string | null) => void;
  isBatchReading: boolean;
  batchReadProgress: string;
}

export const FileUploader: React.FC<FileUploaderProps> = ({
  files,
  onUploadMultipleFiles,
  onRemoveFile,
  onClearAllFiles,
  onToggleFileEnabled,
  onSelectAllFiles,
  onSelectFileSheet,
  onUpdateFileHeaderRow,
  onUpdateFileDataRow,
  onApplyRowsToAllFiles,
  onLoadMultiSample,
  inspectedFileId,
  onSetInspectedFileId,
  isBatchReading,
  batchReadProgress,
}) => {
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const addMoreInputRef = useRef<HTMLInputElement>(null);

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      onUploadMultipleFiles(e.dataTransfer.files);
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const formatFileSize = (bytes: number) => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
  };

  const totalBytes = files.reduce((acc, f) => acc + f.size, 0);
  const totalRows = files.reduce((acc, f) => acc + (f.enabled ? f.totalRows : 0), 0);
  const enabledFilesCount = files.filter((f) => f.enabled).length;

  const inspectedFile = files.find((f) => f.id === inspectedFileId) || files[0];

  return (
    <div className="bg-white border border-neutral-200 rounded-lg p-5 shadow-xs space-y-4">
      {/* Top Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-4 border-b border-neutral-100">
        <div>
          <div className="flex items-center gap-2">
            <h2 className="text-base font-semibold text-neutral-900">
              1. Tệp Excel nguồn (Nạp đồng thời nhiều tệp dung lượng cực lớn)
            </h2>
            {files.length > 0 && (
              <span className="text-xs font-mono font-semibold bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded">
                {files.length} tệp ({enabledFilesCount} đang chọn)
              </span>
            )}
          </div>
          <p className="text-xs text-neutral-500 mt-0.5">
            Tải lên cùng lúc 1 hoặc hàng chục tệp Excel (.xlsx, .xls) hoặc CSV. Công nghệ Dense Mode tối ưu bộ nhớ, đọc mượt cả các tệp phức tạp hàng trăm ngàn dòng.
          </p>
        </div>

        <div className="flex items-center gap-2">
          {files.length === 0 && (
            <button
              onClick={onLoadMultiSample}
              type="button"
              className="inline-flex items-center gap-2 px-3 py-2 text-xs font-medium text-emerald-800 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 rounded-md transition-colors whitespace-nowrap"
            >
              <Sparkles className="h-4 w-4 text-emerald-600" />
              Nạp thử 3 tệp mẫu (T8, T9, T10)
            </button>
          )}

          {files.length > 0 && (
            <button
              type="button"
              onClick={() => addMoreInputRef.current?.click()}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-white bg-neutral-900 hover:bg-neutral-800 rounded-md transition-colors whitespace-nowrap shadow-xs"
            >
              <Plus className="h-3.5 w-3.5" />
              <span>Thêm tệp khác</span>
            </button>
          )}
          <input
            ref={addMoreInputRef}
            type="file"
            multiple
            accept=".xlsx, .xls, .csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
            className="hidden"
            onChange={(e) => {
              if (e.target.files && e.target.files.length > 0) {
                onUploadMultipleFiles(e.target.files);
              }
            }}
          />
        </div>
      </div>

      {/* Progress banner when parsing large files */}
      {isBatchReading && (
        <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-md text-xs text-emerald-900 flex items-center justify-between animate-pulse">
          <div className="flex items-center gap-2">
            <div className="h-3.5 w-3.5 border-2 border-emerald-600 border-t-transparent rounded-full animate-spin"></div>
            <span className="font-medium">{batchReadProgress || 'Đang xử lý tệp lớn...'}</span>
          </div>
          <span className="text-[11px] text-emerald-700">Tối ưu bộ nhớ RAM</span>
        </div>
      )}

      {/* Empty State: Dropzone */}
      {files.length === 0 ? (
        <div
          onDrop={handleDrop}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onClick={() => fileInputRef.current?.click()}
          className={`border-2 border-dashed rounded-lg p-9 text-center cursor-pointer transition-all ${
            isDragging
              ? 'border-emerald-500 bg-emerald-50/50'
              : 'border-neutral-300 hover:border-neutral-400 bg-neutral-50/50'
          }`}
        >
          <input
            ref={fileInputRef}
            type="file"
            multiple
            accept=".xlsx, .xls, .csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
            className="hidden"
            onChange={(e) => {
              if (e.target.files && e.target.files.length > 0) {
                onUploadMultipleFiles(e.target.files);
              }
            }}
          />
          <div className="mx-auto flex h-14 w-14 items-center justify-center rounded-full bg-neutral-100 text-neutral-600 mb-3">
            <UploadCloud className="h-7 w-7" />
          </div>
          <div className="text-sm font-semibold text-neutral-900">
            Kéo thả 1 hoặc nhiều tệp Excel / CSV vào đây, hoặc click để chọn hàng loạt tệp
          </div>
          <p className="mt-1 text-xs text-neutral-500 max-w-xl mx-auto">
            Hỗ trợ nạp cùng lúc nhiều tệp .xlsx, .xls, .csv có dung lượng lớn. Tự động liên kết và trích xuất dữ liệu đa tệp theo danh sách mã bạn chỉ định.
          </p>
        </div>
      ) : (
        /* Populated State: File Management Queue */
        <div className="space-y-4">
          {/* Batch Metrics & Global Actions */}
          <div className="flex flex-wrap items-center justify-between gap-3 p-3 bg-neutral-50 border border-neutral-200 rounded-md text-xs">
            <div className="flex flex-wrap items-center gap-4 text-neutral-700">
              <span>
                Tổng số tệp: <strong className="font-mono text-neutral-900">{files.length}</strong>
              </span>
              <span className="text-neutral-300">|</span>
              <span>
                Đang kích hoạt: <strong className="font-mono text-emerald-800">{enabledFilesCount} tệp</strong>
              </span>
              <span className="text-neutral-300">|</span>
              <span>
                Tổng dung lượng: <strong className="font-mono text-neutral-900">{formatFileSize(totalBytes)}</strong>
              </span>
              <span className="text-neutral-300">|</span>
              <span>
                Tổng dòng quét dự kiến: <strong className="font-mono tabular-nums text-neutral-900">{totalRows.toLocaleString()}</strong> dòng
              </span>
            </div>

            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => onSelectAllFiles(enabledFilesCount !== files.length)}
                className="px-2.5 py-1 text-xs font-medium text-neutral-700 bg-white border border-neutral-300 hover:bg-neutral-100 rounded transition-colors"
              >
                {enabledFilesCount === files.length ? 'Bỏ chọn tất cả' : 'Chọn tất cả tệp'}
              </button>
              <button
                type="button"
                onClick={onClearAllFiles}
                className="px-2.5 py-1 text-xs font-medium text-rose-700 hover:text-rose-800 underline transition-colors"
              >
                Xóa toàn bộ
              </button>
            </div>
          </div>

          {/* Files List Table */}
          <div className="border border-neutral-200 rounded-md overflow-hidden">
            <div className="overflow-x-auto max-h-72">
              <table className="w-full text-left text-xs border-collapse">
                <thead className="bg-neutral-100 text-neutral-700 sticky top-0 border-b border-neutral-200 z-10 font-semibold">
                  <tr>
                    <th className="p-2.5 w-12 text-center">Bật</th>
                    <th className="p-2.5 min-w-56">Tên tệp Excel</th>
                    <th className="p-2.5 w-24">Dung lượng</th>
                    <th className="p-2.5 min-w-36">Sheet đang chọn</th>
                    <th className="p-2.5 w-28 text-right">Tổng số dòng</th>
                    <th className="p-2.5 w-24 text-center">Dòng Tiêu đề</th>
                    <th className="p-2.5 w-24 text-center">Dòng Bắt đầu</th>
                    <th className="p-2.5 w-32 text-center">Thao tác</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-neutral-100 bg-white">
                  {files.map((file) => {
                    const isInspected = file.id === inspectedFile?.id;
                    const isReady = file.status === 'ready';

                    return (
                      <tr
                        key={file.id}
                        className={`hover:bg-neutral-50/70 transition-colors ${
                          !file.enabled ? 'opacity-50 bg-neutral-50/40' : isInspected ? 'bg-emerald-50/30' : ''
                        }`}
                      >
                        {/* Checkbox toggle enabled */}
                        <td className="p-2.5 text-center">
                          <input
                            type="checkbox"
                            checked={file.enabled}
                            onChange={() => onToggleFileEnabled(file.id)}
                            className="rounded border-neutral-300 text-emerald-600 focus:ring-emerald-500 cursor-pointer"
                          />
                        </td>

                        {/* File Name & Status */}
                        <td className="p-2.5 font-medium text-neutral-900">
                          <div className="flex items-center gap-2">
                            <FileSpreadsheet className="h-4 w-4 text-emerald-600 shrink-0" />
                            <span className="truncate max-w-xs" title={file.name}>
                              {file.name}
                            </span>
                            {file.status === 'parsing' && (
                              <span className="text-[10px] text-amber-600 animate-pulse">
                                (Đang nạp...)
                              </span>
                            )}
                            {file.status === 'error' && (
                              <span className="text-[10px] text-rose-600 flex items-center gap-0.5" title={file.errorMsg}>
                                <AlertCircle className="h-3 w-3" /> Lỗi
                              </span>
                            )}
                          </div>
                        </td>

                        {/* File Size */}
                        <td className="p-2.5 text-neutral-500 font-mono">
                          {formatFileSize(file.size)}
                        </td>

                        {/* Sheet Selection */}
                        <td className="p-2.5">
                          {file.sheetNames && file.sheetNames.length > 1 ? (
                            <select
                              value={file.activeSheetName}
                              onChange={(e) => onSelectFileSheet(file.id, e.target.value)}
                              className="text-xs bg-white border border-neutral-300 rounded px-2 py-1 text-neutral-800 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                            >
                              {file.sheetNames.map((sn) => (
                                <option key={sn} value={sn}>
                                  {sn}
                                </option>
                              ))}
                            </select>
                          ) : (
                            <span className="text-neutral-700 font-mono text-[11px]">
                              {file.activeSheetName || 'Sheet1'}
                            </span>
                          )}
                        </td>

                        {/* Total Rows */}
                        <td className="p-2.5 text-right font-mono font-semibold tabular-nums text-neutral-800">
                          {file.totalRows ? file.totalRows.toLocaleString() : '-'}
                        </td>

                        {/* Header Row Index */}
                        <td className="p-2.5 text-center">
                          <select
                            value={file.headerRowIndex}
                            onChange={(e) => onUpdateFileHeaderRow(file.id, Number(e.target.value))}
                            className="text-xs bg-white border border-neutral-300 rounded px-1.5 py-0.5 text-neutral-800 font-mono"
                          >
                            {Array.from({ length: 10 }, (_, i) => (
                              <option key={i} value={i}>
                                Dòng {i + 1}
                              </option>
                            ))}
                          </select>
                        </td>

                        {/* Data Start Row Index */}
                        <td className="p-2.5 text-center">
                          <select
                            value={file.dataStartRowIndex}
                            onChange={(e) => onUpdateFileDataRow(file.id, Number(e.target.value))}
                            className="text-xs bg-white border border-neutral-300 rounded px-1.5 py-0.5 text-neutral-800 font-mono"
                          >
                            {Array.from({ length: 12 }, (_, i) => {
                              const r = i + 1;
                              if (r <= file.headerRowIndex) return null;
                              return (
                                <option key={r} value={r}>
                                  Dòng {r + 1}
                                </option>
                              );
                            })}
                          </select>
                        </td>

                        {/* Actions */}
                        <td className="p-2.5 text-center">
                          <div className="flex items-center justify-center gap-1.5">
                            <button
                              type="button"
                              onClick={() =>
                                onSetInspectedFileId(isInspected ? null : file.id)
                              }
                              className={`px-2 py-0.5 text-[11px] font-medium rounded border transition-colors ${
                                isInspected
                                  ? 'bg-emerald-600 text-white border-emerald-600'
                                  : 'bg-white text-neutral-700 border-neutral-300 hover:bg-neutral-100'
                              }`}
                            >
                              {isInspected ? 'Đang soi' : 'Soi dòng'}
                            </button>
                            <button
                              type="button"
                              onClick={() => onRemoveFile(file.id)}
                              className="p-1 text-neutral-400 hover:text-rose-600 rounded transition-colors"
                              title="Xóa tệp này"
                            >
                              <Trash2 className="h-3.5 w-3.5" />
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>

          {/* Quick Row Sync Button (Apply header/data rows from inspected file to all files) */}
          {files.length > 1 && inspectedFile && (
            <div className="flex items-center justify-between p-2.5 bg-neutral-50 border border-neutral-200 rounded-md text-xs">
              <span className="text-neutral-600">
                Đang cấu hình: <strong>{inspectedFile.name}</strong> (Tiêu đề: <em>Dòng {inspectedFile.headerRowIndex + 1}</em>, Bắt đầu DL: <em>Dòng {inspectedFile.dataStartRowIndex + 1}</em>)
              </span>
              <button
                type="button"
                onClick={() =>
                  onApplyRowsToAllFiles(inspectedFile.headerRowIndex, inspectedFile.dataStartRowIndex)
                }
                className="inline-flex items-center gap-1.5 px-3 py-1 font-medium text-emerald-800 bg-white border border-emerald-300 hover:bg-emerald-50 rounded shadow-2xs transition-colors"
              >
                <CopyCheck className="h-3.5 w-3.5 text-emerald-600" />
                <span>Áp dụng dòng này cho TẤT CẢ {files.length} tệp</span>
              </button>
            </div>
          )}

          {/* Row Inspector for Inspected File */}
          {inspectedFile && inspectedFile.currentSheetData && (
            <div className="border border-neutral-200 rounded-md overflow-hidden bg-neutral-50/50 p-3 space-y-2">
              <div className="flex items-center justify-between pb-2 border-b border-neutral-200">
                <span className="text-xs font-semibold text-neutral-800">
                  Bảng soi 10 dòng đầu của tệp: <strong className="text-emerald-800">{inspectedFile.name}</strong> (Sheet: {inspectedFile.activeSheetName})
                </span>
                <span className="text-[11px] text-neutral-500">
                  Click nút bên trái để gán nhanh dòng Tiêu đề hoặc dòng Dữ liệu
                </span>
              </div>

              <div className="overflow-x-auto max-h-56 border border-neutral-200 rounded bg-white">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-neutral-100 text-neutral-600 sticky top-0 border-b border-neutral-200">
                    <tr>
                      <th className="p-2 border-r border-neutral-200 w-24 text-center">
                        Số dòng
                      </th>
                      <th className="p-2 border-r border-neutral-200 w-28 text-center">
                        Thiết lập
                      </th>
                      {inspectedFile.currentSheetData.rawRows[0]?.slice(0, 10).map((_, colIdx) => (
                        <th key={colIdx} className="p-2 border-r border-neutral-200 whitespace-nowrap font-mono text-[11px]">
                          Cột {String.fromCharCode(65 + colIdx)}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {inspectedFile.currentSheetData.rawRows.slice(0, 10).map((row, rIdx) => {
                      const isHeader = rIdx === inspectedFile.headerRowIndex;
                      const isDataStart = rIdx === inspectedFile.dataStartRowIndex;

                      return (
                        <tr
                          key={rIdx}
                          className={`border-b border-neutral-100 transition-colors ${
                            isHeader
                              ? 'bg-emerald-50 text-emerald-950 font-medium'
                              : isDataStart
                              ? 'bg-amber-50 text-amber-950'
                              : 'hover:bg-neutral-50'
                          }`}
                        >
                          <td className="p-2 border-r border-neutral-200 font-mono text-center text-neutral-500">
                            Dòng {rIdx + 1}
                          </td>
                          <td className="p-2 border-r border-neutral-200 text-center whitespace-nowrap">
                            {isHeader ? (
                              <span className="text-[11px] font-semibold text-emerald-700 bg-emerald-100 px-1.5 py-0.5 rounded">
                                Tiêu đề
                              </span>
                            ) : isDataStart ? (
                              <span className="text-[11px] font-semibold text-amber-700 bg-amber-100 px-1.5 py-0.5 rounded">
                                Bắt đầu DL
                              </span>
                            ) : (
                              <div className="flex items-center justify-center gap-1">
                                <button
                                  type="button"
                                  onClick={() => onUpdateFileHeaderRow(inspectedFile.id, rIdx)}
                                  className="text-[10px] text-neutral-600 hover:text-emerald-700 underline"
                                  title="Đặt làm dòng tiêu đề cho tệp này"
                                >
                                  Header
                                </button>
                                <span className="text-neutral-300">·</span>
                                <button
                                  type="button"
                                  onClick={() => onUpdateFileDataRow(inspectedFile.id, rIdx)}
                                  className="text-[10px] text-neutral-600 hover:text-amber-700 underline"
                                  title="Đặt làm dòng bắt đầu dữ liệu cho tệp này"
                                >
                                  Dữ liệu
                                </button>
                              </div>
                            )}
                          </td>
                          {Array.from({ length: 10 }, (_, cIdx) => (
                            <td
                              key={cIdx}
                              className="p-2 border-r border-neutral-200 whitespace-nowrap text-neutral-700 max-w-44 truncate"
                            >
                              {row && row[cIdx] !== undefined && row[cIdx] !== null
                                ? String(row[cIdx])
                                : ''}
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
      )}
    </div>
  );
};
